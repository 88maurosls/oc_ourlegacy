import io
import re
import zipfile
from pathlib import Path

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

APP_VERSION = "2026-07-15_SIMPLE_V4"
OUTPUT_COLUMNS = ["SKU", "Product Name", "Variant", "Size", "Quantity", "Unit Price", "RRP"]


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def key(value):
    return re.sub(r"[^a-z0-9]", "", clean_text(value).lower())


def to_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else float(value)

    s = clean_text(value)
    if not s or s.startswith("="):
        return None

    s = re.sub(r"[^0-9,.-]", "", s)
    if s in ("", "-", ".", ","):
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        n = float(s)
    except ValueError:
        return None

    return int(n) if n.is_integer() else n


ALIASES = {
    "sku": {"sku", "stylesku", "stylecode", "itemcode", "codice"},
    "variant_sku": {"variantsku", "variantskucode", "itemsku", "skuvariant"},
    "product_name": {"productname", "product", "model", "categoria", "category"},
    "variant": {"variant", "description", "color", "colour", "variantname", "descrizione"},
    "sc": {"sc", "sizescale", "scala", "scalataglie", "sizetable"},
    "unit_price": {"unitprice", "wholesaleprice", "wholesale", "cost", "buyingprice", "price"},
    "rrp": {"rrp", "retailprice", "retail", "msrp", "sellingprice"},
    "q": {"q", "qty", "quantity", "totalquantity", "totqty", "total"},
    "size": {"size", "taglia"},
}


def map_headers(ws, row):
    found = {}
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        h = key(ws.cell(row, col).value)
        if not h:
            continue
        for field, names in ALIASES.items():
            if h in names and field not in found:
                found[field] = col
    return found


def find_linear_sheet(wb):
    # Serve per non andare in errore se viene ricaricato un file già linearizzato.
    required = {"product_name", "variant", "size", "q", "unit_price", "rrp"}
    for ws in wb.worksheets:
        for row in range(1, min(ws.max_row, 20) + 1):
            found = map_headers(ws, row)
            if required.issubset(found):
                return ws, row, found
    return None, None, None


def read_linear_sheet(ws, header_row, col):
    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        product = clean_text(ws.cell(r, col["product_name"]).value)
        variant = clean_text(ws.cell(r, col["variant"]).value)
        size = clean_text(ws.cell(r, col["size"]).value)
        qty = to_number(ws.cell(r, col["q"]).value)

        if not product and not variant and not size and qty is None:
            continue
        if qty is None or qty == 0:
            continue

        sku_value = ""
        if "sku" in col:
            sku_value = clean_text(ws.cell(r, col["sku"]).value)
        if not sku_value and "variant_sku" in col:
            sku_value = clean_text(ws.cell(r, col["variant_sku"]).value)

        records.append({
            "SKU": sku_value,
            "Product Name": product,
            "Variant": variant,
            "Size": size,
            "Quantity": int(qty) if float(qty).is_integer() else qty,
            "Unit Price": to_number(ws.cell(r, col["unit_price"]).value),
            "RRP": to_number(ws.cell(r, col["rrp"]).value),
        })

    if not records:
        raise ValueError("Il file sembra già lineare, ma non contiene righe utili.")
    return records, [f"File già lineare letto dal foglio '{ws.title}', riga intestazioni {header_row}."]


def find_matrix_header(wb):
    # La commessa originale deve avere Product Name, Variant, SC, Unit Price, RRP.
    # SKU è comodo ma non obbligatorio: se manca uso Variant SKU.
    required = {"product_name", "variant", "sc", "unit_price", "rrp"}
    best = None

    for ws in wb.worksheets:
        for row in range(1, ws.max_row + 1):
            found = map_headers(ws, row)
            score = len(required.intersection(found))
            if best is None or score > best[0]:
                sample = " | ".join(clean_text(ws.cell(row, c).value) for c in range(1, min(ws.max_column, 12) + 1))
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
    for r in range(1, header_row):
        scale_code = key(ws.cell(r, sc_col).value)
        if not scale_code:
            continue

        sizes = {}
        for c in qty_cols:
            size = clean_text(ws.cell(r, c).value)
            if size:
                sizes[c] = size

        if sizes:
            scales[scale_code] = sizes

    if not scales:
        raise ValueError("Non trovo la scala taglie sopra la tabella nella colonna SC.")
    return scales


def convert_excel(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)

    linear_ws, linear_header_row, linear_cols = find_linear_sheet(wb)
    if linear_ws is not None:
        return read_linear_sheet(linear_ws, linear_header_row, linear_cols)

    ws, header_row, col = find_matrix_header(wb)

    qty_cols = list(range(col["sc"] + 1, col["unit_price"]))
    if not qty_cols:
        raise ValueError("Non trovo colonne quantità tra SC e Unit Price.")

    scales = read_size_scales(ws, header_row, col["sc"], qty_cols)
    records = []
    notes = [
        f"Foglio letto: '{ws.title}'.",
        f"Riga intestazioni: {header_row}.",
        f"Scale taglie trovate: {', '.join(sorted(s.upper() for s in scales))}.",
    ]

    for r in range(header_row + 1, ws.max_row + 1):
        product = clean_text(ws.cell(r, col["product_name"]).value)
        variant = clean_text(ws.cell(r, col["variant"]).value)
        if not product and not variant:
            continue

        scale_raw = ws.cell(r, col["sc"]).value
        scale_code = key(scale_raw)
        if not scale_code:
            notes.append(f"Riga {r}: SC mancante, saltata.")
            continue
        if scale_code not in scales:
            notes.append(f"Riga {r}: scala '{clean_text(scale_raw)}' non trovata, saltata.")
            continue

        sku_value = ""
        if "sku" in col:
            sku_value = clean_text(ws.cell(r, col["sku"]).value)
        if not sku_value and "variant_sku" in col:
            sku_value = clean_text(ws.cell(r, col["variant_sku"]).value)

        unit_price = to_number(ws.cell(r, col["unit_price"]).value)
        rrp = to_number(ws.cell(r, col["rrp"]).value)
        row_total = 0

        for c in qty_cols:
            qty = to_number(ws.cell(r, c).value)
            if qty is None or qty == 0:
                continue
            if qty < 0:
                notes.append(f"Riga {r}: quantità negativa ignorata.")
                continue

            size = scales[scale_code].get(c, "")
            if not size:
                notes.append(f"Riga {r}: quantità presente ma taglia assente nella scala {clean_text(scale_raw)}.")
                continue

            qty_out = int(qty) if float(qty).is_integer() else qty
            records.append({
                "SKU": sku_value,
                "Product Name": product,
                "Variant": variant,
                "Size": size,
                "Quantity": qty_out,
                "Unit Price": unit_price,
                "RRP": rrp,
            })
            row_total += float(qty)

        if "q" in col:
            expected = to_number(ws.cell(r, col["q"]).value)
            if expected is not None and abs(row_total - float(expected)) > 0.00001:
                notes.append(f"Riga {r}: totale taglie {row_total:g}, colonna Q {float(expected):g}.")

    if not records:
        raise ValueError("Non ho trovato quantità maggiori di zero da esportare.")

    return records, notes


def make_output_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Linear"

    ws.append(OUTPUT_COLUMNS)
    for rec in records:
        ws.append([rec.get(col, "") for col in OUTPUT_COLUMNS])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{ws.max_row}"

    widths = [18, 26, 42, 12, 10, 12, 12]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + idx)].width = width

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 5).number_format = "0"
        ws.cell(row, 6).number_format = "0.00"
        ws.cell(row, 7).number_format = "0.00"

    output = io.BytesIO()
    wb.save(output)
    data = output.getvalue()

    # Controllo rapido: zip valido e niente tabelle strutturate.
    with zipfile.ZipFile(io.BytesIO(data), "r") as z:
        bad = z.testzip()
        if bad:
            raise ValueError(f"Excel generato non valido: {bad}")
        if any(name.startswith("xl/tables/") for name in z.namelist()):
            raise ValueError("Excel generato non valido: contiene tabelle strutturate.")

    return data


st.set_page_config(page_title="Ordine Excel in formato lineare", page_icon="📄")
st.title("Conversione ordine Excel in formato lineare")
st.caption(f"Versione app: {APP_VERSION}")
st.write("Carica l'Excel originale della commessa. L'app legge la scala taglie sopra la tabella e genera un file lineare.")

uploaded = st.file_uploader("Carica un file .xlsx o .xlsm", type=["xlsx", "xlsm"])

if uploaded is not None:
    try:
        records, notes = convert_excel(uploaded.getvalue())
        output = make_output_excel(records)
        total_qty = sum(float(row.get("Quantity") or 0) for row in records)

        c1, c2 = st.columns(2)
        c1.metric("Righe generate", len(records))
        c2.metric("Quantità totale", f"{total_qty:g}")

        st.dataframe(records, use_container_width=True, hide_index=True)

        output_name = f"{Path(uploaded.name).stem}_linear.xlsx"
        st.download_button(
            "Scarica Excel lineare",
            data=output,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander("Dettagli conversione"):
            for note in notes:
                st.write("•", note)

    except Exception as exc:
        st.error(f"Impossibile convertire il file: {exc}")
        st.info(f"Versione app in uso: {APP_VERSION}")
