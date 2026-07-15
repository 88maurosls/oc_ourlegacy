import io
import re
import zipfile
from pathlib import Path

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


APP_VERSION = "2026-07-15_V6"

OUTPUT_COLUMNS = [
    "SKU",
    "Product Name",
    "Variant",
    "Size",
    "Quantity",
    "Unit Price",
    "RRP",
]

CATEGORY_RULES = [
    ("M", "MEN"),
    ("W", "WOMEN"),
    ("A", "UNISEX"),
]


def clean_text(value):
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()

    return str(value).strip()


def key(value):
    return re.sub(r"[^a-z0-9]", "", clean_text(value).lower())


def normalize_size(value):
    text = clean_text(value)

    if key(text) == "os":
        return "UNICA"

    return text


def to_number(value):
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, (int, float)):
        if float(value).is_integer():
            return int(value)
        return float(value)

    text = clean_text(value)

    if not text or text.startswith("="):
        return None

    text = re.sub(r"[^0-9,.-]", "", text)

    if text in ("", "-", ".", ","):
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        number = float(text)
    except ValueError:
        return None

    if number.is_integer():
        return int(number)

    return number


ALIASES = {
    "sku": {
        "sku",
        "stylesku",
        "stylecode",
        "itemcode",
        "codice",
    },
    "variant_sku": {
        "variantsku",
        "variantskucode",
        "itemsku",
        "skuvariant",
    },
    "product_name": {
        "productname",
        "product",
        "model",
        "categoria",
        "category",
    },
    "variant": {
        "variant",
        "description",
        "color",
        "colour",
        "variantname",
        "descrizione",
    },
    "sc": {
        "sc",
        "sizescale",
        "scala",
        "scalataglie",
        "sizetable",
    },
    "unit_price": {
        "unitprice",
        "wholesaleprice",
        "wholesale",
        "cost",
        "buyingprice",
        "price",
    },
    "rrp": {
        "rrp",
        "retailprice",
        "retail",
        "msrp",
        "sellingprice",
    },
    "q": {
        "q",
        "qty",
        "quantity",
        "totalquantity",
        "totqty",
        "total",
    },
    "size": {
        "size",
        "taglia",
    },
}


def map_headers(ws, row):
    found = {}

    for col in range(1, ws.max_column + 1):
        header = key(ws.cell(row, col).value)

        if not header:
            continue

        for field, names in ALIASES.items():
            if header in names and field not in found:
                found[field] = col

    return found


def find_linear_sheet(wb):
    required = {
        "product_name",
        "variant",
        "size",
        "q",
        "unit_price",
        "rrp",
    }

    for ws in wb.worksheets:
        max_check_row = min(ws.max_row, 20)

        for row in range(1, max_check_row + 1):
            found = map_headers(ws, row)

            if required.issubset(found):
                return ws, row, found

    return None, None, None


def read_linear_sheet(ws, header_row, col):
    records = []

    for row in range(header_row + 1, ws.max_row + 1):
        product_name = clean_text(ws.cell(row, col["product_name"]).value)
        variant = clean_text(ws.cell(row, col["variant"]).value)
        size = normalize_size(ws.cell(row, col["size"]).value)
        quantity = to_number(ws.cell(row, col["q"]).value)

        if not product_name and not variant and not size and quantity is None:
            continue

        if quantity is None or quantity == 0:
            continue

        sku = ""

        if "sku" in col:
            sku = clean_text(ws.cell(row, col["sku"]).value)

        if not sku and "variant_sku" in col:
            sku = clean_text(ws.cell(row, col["variant_sku"]).value)

        records.append(
            {
                "SKU": sku,
                "Product Name": product_name,
                "Variant": variant,
                "Size": size,
                "Quantity": int(quantity) if float(quantity).is_integer() else quantity,
                "Unit Price": to_number(ws.cell(row, col["unit_price"]).value),
                "RRP": to_number(ws.cell(row, col["rrp"]).value),
            }
        )

    if not records:
        raise ValueError("Il file sembra già lineare, ma non contiene righe utili.")

    notes = [
        f"File già lineare letto dal foglio '{ws.title}'.",
        f"Riga intestazioni: {header_row}.",
    ]

    return records, notes


def find_matrix_header(wb):
    required = {
        "product_name",
        "variant",
        "sc",
        "unit_price",
        "rrp",
    }

    best = None

    for ws in wb.worksheets:
        for row in range(1, ws.max_row + 1):
            found = map_headers(ws, row)
            score = len(required.intersection(found))

            if best is None or score > best[0]:
                sample_values = []

                for col in range(1, min(ws.max_column, 12) + 1):
                    sample_values.append(clean_text(ws.cell(row, col).value))

                sample = " | ".join(sample_values)
                best = (score, ws.title, row, sample, found)

            if required.issubset(found):
                return ws, row, found

    if best:
        score, sheet_name, row, sample, found = best

        raise ValueError(
            "Non trovo la riga intestazioni della commessa originale. "
            f"Miglior candidato: foglio '{sheet_name}', riga {row}. "
            f"Campi riconosciuti: {', '.join(sorted(found.keys())) or 'nessuno'}. "
            f"Contenuto: {sample}"
        )

    raise ValueError("Non trovo nessuna riga utilizzabile nel file.")


def read_size_scales(ws, header_row, sc_col, qty_cols):
    scales = {}

    for row in range(1, header_row):
        scale_code = key(ws.cell(row, sc_col).value)

        if not scale_code:
            continue

        sizes = {}

        for col in qty_cols:
            size = normalize_size(ws.cell(row, col).value)

            if size:
                sizes[col] = size

        if sizes:
            scales[scale_code] = sizes

    if not scales:
        raise ValueError("Non trovo la scala taglie sopra la tabella nella colonna SC.")

    return scales


def convert_excel(file_bytes):
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)

    linear_ws, linear_header_row, linear_cols = find_linear_sheet(workbook)

    if linear_ws is not None:
        return read_linear_sheet(linear_ws, linear_header_row, linear_cols)

    ws, header_row, col = find_matrix_header(workbook)

    qty_cols = list(range(col["sc"] + 1, col["unit_price"]))

    if not qty_cols:
        raise ValueError("Non trovo colonne quantità tra SC e Unit Price.")

    scales = read_size_scales(ws, header_row, col["sc"], qty_cols)

    records = []
    notes = [
        f"Foglio letto: '{ws.title}'.",
        f"Riga intestazioni: {header_row}.",
        f"Scale taglie trovate: {', '.join(sorted(scale.upper() for scale in scales))}.",
    ]

    for row in range(header_row + 1, ws.max_row + 1):
        product_name = clean_text(ws.cell(row, col["product_name"]).value)
        variant = clean_text(ws.cell(row, col["variant"]).value)

        if not product_name and not variant:
            continue

        scale_raw = ws.cell(row, col["sc"]).value
        scale_code = key(scale_raw)

        if not scale_code:
            notes.append(f"Riga {row}: SC mancante, saltata.")
            continue

        if scale_code not in scales:
            notes.append(f"Riga {row}: scala '{clean_text(scale_raw)}' non trovata, saltata.")
            continue

        sku = ""

        if "sku" in col:
            sku = clean_text(ws.cell(row, col["sku"]).value)

        if not sku and "variant_sku" in col:
            sku = clean_text(ws.cell(row, col["variant_sku"]).value)

        unit_price = to_number(ws.cell(row, col["unit_price"]).value)
        rrp = to_number(ws.cell(row, col["rrp"]).value)

        row_total = 0

        for qty_col in qty_cols:
            quantity = to_number(ws.cell(row, qty_col).value)

            if quantity is None or quantity == 0:
                continue

            if quantity < 0:
                notes.append(f"Riga {row}: quantità negativa ignorata.")
                continue

            size = scales[scale_code].get(qty_col, "")

            if not size:
                notes.append(
                    f"Riga {row}: quantità presente ma taglia assente nella scala {clean_text(scale_raw)}."
                )
                continue

            quantity_output = int(quantity) if float(quantity).is_integer() else quantity

            records.append(
                {
                    "SKU": sku,
                    "Product Name": product_name,
                    "Variant": variant,
                    "Size": size,
                    "Quantity": quantity_output,
                    "Unit Price": unit_price,
                    "RRP": rrp,
                }
            )

            row_total += float(quantity)

        if "q" in col:
            expected_total = to_number(ws.cell(row, col["q"]).value)

            if expected_total is not None and abs(row_total - float(expected_total)) > 0.00001:
                notes.append(
                    f"Riga {row}: totale taglie {row_total:g}, colonna Q {float(expected_total):g}."
                )

    if not records:
        raise ValueError("Non ho trovato quantità maggiori di zero da esportare.")

    return records, notes


def category_from_sku(sku):
    code = clean_text(sku).upper()

    for prefix, label in CATEGORY_RULES:
        if code.startswith(prefix):
            return label

    return None


def split_by_category(records):
    split = {
        "MEN": [],
        "WOMEN": [],
        "UNISEX": [],
    }

    skipped = []

    for record in records:
        label = category_from_sku(record.get("SKU"))

        if label:
            split[label].append(record)
        else:
            skipped.append(record)

    return split, skipped


def make_output_excel(records, sheet_title):
    workbook = Workbook()
    ws = workbook.active
    ws.title = sheet_title[:31]

    ws.append(OUTPUT_COLUMNS)

    for record in records:
        ws.append([record.get(column, "") for column in OUTPUT_COLUMNS])

    header_fill = PatternFill("solid", fgColor="1F4E78")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{ws.max_row}"

    widths = {
        "A": 18,
        "B": 28,
        "C": 42,
        "D": 12,
        "E": 10,
        "F": 12,
        "G": 12,
    }

    for column_letter, width in widths.items():
        ws.column_dimensions[column_letter].width = width

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 5).number_format = "0"
        ws.cell(row, 6).number_format = "0.00"
        ws.cell(row, 7).number_format = "0.00"

    output = io.BytesIO()
    workbook.save(output)

    data = output.getvalue()

    with zipfile.ZipFile(io.BytesIO(data), "r") as zipped:
        bad_file = zipped.testzip()

        if bad_file:
            raise ValueError(f"Excel generato non valido: {bad_file}")

        if any(name.startswith("xl/tables/") for name in zipped.namelist()):
            raise ValueError("Excel generato non valido: contiene tabelle strutturate.")

    return data


def prepare_files(records, base_name):
    split, skipped = split_by_category(records)

    files = {}

    for label, rows in split.items():
        if not rows:
            continue

        files[label] = {
            "rows": rows,
            "file_name": f"{base_name}_{label}.xlsx",
            "data": make_output_excel(rows, label),
        }

    return files, skipped


def main():
    st.set_page_config(
        page_title="Ordine Excel in formato lineare",
        page_icon="📄",
        layout="centered",
    )

    st.title("Conversione ordine Our Legacy in formato lineare")
    st.caption(f"Versione app: {APP_VERSION}")

    st.write(
        "L'app legge la scala taglie sopra la tabella, converte O/S in UNICA "
        "e prepara tre file separati: MEN, WOMEN e UNISEX."
        "File UNISEX esportato da verificare alla fine."
    )

    uploaded_file = st.file_uploader(
        "Carica un file .xlsx o .xlsm",
        type=["xlsx", "xlsm"],
    )

    if uploaded_file is None:
        return

    try:
        records, notes = convert_excel(uploaded_file.getvalue())

        base_name = Path(uploaded_file.name).stem
        files, skipped = prepare_files(records, base_name)

        total_quantity = sum(float(row.get("Quantity") or 0) for row in records)
        exported_quantity = sum(
            float(row.get("Quantity") or 0)
            for item in files.values()
            for row in item["rows"]
        )

        metric_col_1, metric_col_2, metric_col_3 = st.columns(3)

        metric_col_1.metric("Righe lette", len(records))
        metric_col_2.metric("Quantità letta", f"{total_quantity:g}")
        metric_col_3.metric("Quantità esportata", f"{exported_quantity:g}")

        st.subheader("Download separati")

        download_col_1, download_col_2, download_col_3 = st.columns(3)

        download_columns = {
            "MEN": download_col_1,
            "WOMEN": download_col_2,
            "UNISEX": download_col_3,
        }

        for label, column in download_columns.items():
            item = files.get(label)

            with column:
                st.markdown(f"### {label}")

                if item:
                    quantity = sum(float(row.get("Quantity") or 0) for row in item["rows"])

                    st.write(f"{len(item['rows'])} righe")
                    st.write(f"Quantità {quantity:g}")

                    st.download_button(
                        label=f"Scarica {label}",
                        data=item["data"],
                        file_name=item["file_name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{label}",
                        use_container_width=True,
                    )
                else:
                    st.write("Nessuna riga")
                    st.download_button(
                        label=f"Scarica {label}",
                        data=b"",
                        file_name=f"{base_name}_{label}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{label}_disabled",
                        disabled=True,
                        use_container_width=True,
                    )

        if skipped:
            st.warning(
                f"{len(skipped)} righe non esportate perché lo SKU non inizia con M, W o A."
            )

        with st.expander("Anteprima dati"):
            st.dataframe(
                records,
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Dettagli conversione"):
            for note in notes:
                st.write("•", note)

    except Exception as exc:
        st.error(f"Impossibile convertire il file: {exc}")
        st.info(f"Versione app in uso: {APP_VERSION}")


if __name__ == "__main__":
    main()
