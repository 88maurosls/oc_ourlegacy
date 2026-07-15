import io
import re
from pathlib import Path
from typing import Any

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_COLUMNS = [
    "SKU",
    "Product Name",
    "Variant",
    "Size",
    "Quantity",
    "Unit Price",
    "RRP",
]


def normalize(value: Any) -> str:
    """Normalizza testi e intestazioni per confronti robusti."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def excel_text(value: Any) -> str:
    """Converte un valore Excel in testo senza aggiungere .0 alle taglie intere."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_number(value: Any) -> float | int | None:
    """Legge numeri Excel o stringhe come '1.234,56 EUR'."""
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, (int, float)):
        number = float(value)
        return int(number) if number.is_integer() else number

    text = str(value).strip()
    if not text:
        return None

    text = re.sub(r"[^0-9,\.\-]", "", text)
    if not text or text in {"-", ".", ","}:
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

    return int(number) if number.is_integer() else number


def find_header_row_and_sheet(workbook):
    """Cerca automaticamente il foglio e la riga con le intestazioni prodotto."""
    required_headers = {
        "product name",
        "variant",
        "sc",
        "unit price",
        "rrp",
    }

    for worksheet in workbook.worksheets:
        max_scan_rows = min(worksheet.max_row, 150)
        for row_number in range(1, max_scan_rows + 1):
            row_headers = {
                normalize(worksheet.cell(row_number, col).value)
                for col in range(1, worksheet.max_column + 1)
            }
            if required_headers.issubset(row_headers):
                return worksheet, row_number

    raise ValueError(
        "Non trovo la riga con le intestazioni Product Name, Variant, SC, "
        "Unit Price e RRP."
    )


def column_positions(worksheet, header_row: int) -> dict[str, int]:
    positions: dict[str, int] = {}
    for col in range(1, worksheet.max_column + 1):
        key = normalize(worksheet.cell(header_row, col).value)
        if key and key not in positions:
            positions[key] = col
    return positions


def build_size_scales(
    worksheet,
    header_row: int,
    scale_column: int,
    quantity_columns: list[int],
) -> dict[str, dict[int, str]]:
    """
    Legge le scale taglie sopra la tabella.

    Esempio:
    SC A -> O/S
    SC C -> 44, 46, 48, 50, 52, 54
    """
    scales: dict[str, dict[int, str]] = {}

    for row_number in range(1, header_row):
        raw_code = worksheet.cell(row_number, scale_column).value
        code = normalize(raw_code)
        if not code:
            continue

        sizes: dict[int, str] = {}
        for col in quantity_columns:
            size = excel_text(worksheet.cell(row_number, col).value)
            if size:
                sizes[col] = size

        if sizes:
            scales[code] = sizes

    if not scales:
        raise ValueError("Non è stata rilevata alcuna scala taglie sopra la tabella.")

    return scales


def convert_excel(file_bytes: bytes):
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)
    worksheet, header_row = find_header_row_and_sheet(workbook)
    positions = column_positions(worksheet, header_row)

    def get_column(*names: str, required: bool = True) -> int | None:
        for name in names:
            col = positions.get(normalize(name))
            if col is not None:
                return col
        if required:
            raise ValueError(f"Colonna mancante: {' / '.join(names)}")
        return None

    # Usa prima la colonna SKU. Se vuota, usa Variant SKU come fallback.
    sku_column = get_column("SKU", required=False)
    variant_sku_column = get_column("Variant SKU", required=False)
    if sku_column is None and variant_sku_column is None:
        raise ValueError("Non trovo né la colonna SKU né la colonna Variant SKU.")

    product_name_column = get_column("Product Name")
    variant_column = get_column("Variant")
    scale_column = get_column("SC")
    unit_price_column = get_column("Unit Price")
    rrp_column = get_column("RRP")
    total_quantity_column = get_column("Q", "Quantity", required=False)

    if unit_price_column <= scale_column + 1:
        raise ValueError("Non trovo le colonne quantità comprese tra SC e Unit Price.")

    quantity_columns = list(range(scale_column + 1, unit_price_column))
    size_scales = build_size_scales(
        worksheet,
        header_row,
        scale_column,
        quantity_columns,
    )

    records: list[dict[str, Any]] = []
    warnings: list[str] = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        sku = worksheet.cell(row_number, sku_column).value if sku_column else None
        if (sku is None or str(sku).strip() == "") and variant_sku_column:
            sku = worksheet.cell(row_number, variant_sku_column).value

        product_name = worksheet.cell(row_number, product_name_column).value
        variant = worksheet.cell(row_number, variant_column).value
        scale_raw = worksheet.cell(row_number, scale_column).value

        # Ignora righe vuote, subtotali, note e condizioni di vendita.
        if all(
            value is None or str(value).strip() == ""
            for value in (sku, product_name, variant, scale_raw)
        ):
            continue

        if product_name is None or str(product_name).strip() == "":
            continue

        scale_code = normalize(scale_raw)
        if not scale_code:
            warnings.append(f"Riga {row_number}: SC mancante, riga ignorata.")
            continue

        scale = size_scales.get(scale_code)
        if scale is None:
            warnings.append(
                f"Riga {row_number}: scala taglie SC '{excel_text(scale_raw)}' non trovata."
            )
            continue

        unit_price_raw = worksheet.cell(row_number, unit_price_column).value
        rrp_raw = worksheet.cell(row_number, rrp_column).value
        unit_price = parse_number(unit_price_raw)
        rrp = parse_number(rrp_raw)
        if unit_price is None:
            unit_price = unit_price_raw
        if rrp is None:
            rrp = rrp_raw

        row_quantity = 0.0

        for col in quantity_columns:
            quantity = parse_number(worksheet.cell(row_number, col).value)
            if quantity is None or quantity == 0:
                continue
            if quantity < 0:
                warnings.append(
                    f"Riga {row_number}: quantità negativa nella colonna {col}, ignorata."
                )
                continue

            size = scale.get(col)
            if not size:
                warnings.append(
                    f"Riga {row_number}: quantità presente ma taglia assente "
                    f"nella scala SC '{excel_text(scale_raw)}'."
                )
                continue

            records.append(
                {
                    "SKU": excel_text(sku),
                    "Product Name": excel_text(product_name),
                    "Variant": excel_text(variant),
                    "Size": size,
                    "Quantity": quantity,
                    "Unit Price": unit_price,
                    "RRP": rrp,
                }
            )
            row_quantity += float(quantity)

        if total_quantity_column:
            expected_quantity = parse_number(
                worksheet.cell(row_number, total_quantity_column).value
            )
            if expected_quantity is not None and abs(
                row_quantity - float(expected_quantity)
            ) > 0.000001:
                warnings.append(
                    f"Riga {row_number}: totale taglie {row_quantity:g}, "
                    f"ma la colonna Q indica {float(expected_quantity):g}."
                )

    if not records:
        raise ValueError("Non è stata trovata alcuna quantità da esportare.")

    return records, warnings, worksheet.title, sorted(
        excel_text(code).upper() for code in size_scales.keys()
    )


def create_output_excel(records: list[dict[str, Any]]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Linear"

    worksheet.append(OUTPUT_COLUMNS)
    for record in records:
        worksheet.append([record[column] for column in OUTPUT_COLUMNS])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    # Formati numerici.
    for row in range(2, worksheet.max_row + 1):
        worksheet.cell(row, 5).number_format = "0.##"
        worksheet.cell(row, 6).number_format = "#,##0.00"
        worksheet.cell(row, 7).number_format = "#,##0.00"

    # Tabella Excel con filtri e righe alternate.
    table = Table(displayName="LinearTable", ref=worksheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)

    widths = {
        "A": 18,
        "B": 24,
        "C": 42,
        "D": 12,
        "E": 12,
        "F": 14,
        "G": 14,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


st.set_page_config(page_title="Excel Size Matrix to Linear", page_icon="📊")
st.title("Conversione ordine Excel in formato lineare")
st.write(
    "Carica il file ordine. L'app legge automaticamente la scala taglie indicata "
    "nella colonna **SC** e crea una riga per ogni taglia con quantità maggiore di zero."
)

uploaded_file = st.file_uploader(
    "Carica un file .xlsx o .xlsm",
    type=["xlsx", "xlsm"],
)

if uploaded_file is not None:
    try:
        records, warnings, sheet_name, detected_scales = convert_excel(
            uploaded_file.getvalue()
        )
        output_bytes = create_output_excel(records)

        total_quantity = sum(float(record["Quantity"]) for record in records)
        col1, col2, col3 = st.columns(3)
        col1.metric("Righe esportate", len(records))
        col2.metric("Quantità totale", f"{total_quantity:g}")
        col3.metric("Foglio letto", sheet_name)

        st.caption("Scale taglie rilevate: " + ", ".join(detected_scales))
        st.dataframe(records[:200], use_container_width=True, hide_index=True)

        source_name = Path(uploaded_file.name).stem
        st.download_button(
            label="Scarica Excel lineare",
            data=output_bytes,
            file_name=f"{source_name}_linear.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        if warnings:
            with st.expander(f"Avvisi di controllo ({len(warnings)})"):
                for warning in warnings[:100]:
                    st.write("•", warning)
                if len(warnings) > 100:
                    st.write(f"Altri {len(warnings) - 100} avvisi non mostrati.")

    except Exception as exc:
        st.error(f"Impossibile convertire il file: {exc}")
