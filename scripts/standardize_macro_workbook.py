from __future__ import annotations

from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT_DIR = Path(__file__).resolve().parents[1]
SOURCE = Path(r"C:\Users\julio\Downloads\Arriendos base de datos Macro Final.xlsm")
OUTPUT_DIR = ROOT_DIR / "data" / "standardized_workbooks"
OUTPUT = OUTPUT_DIR / "Arriendos_base_de_datos_estandarizada.xlsm"


BLUE = "1F4E78"
LIGHT_BLUE = "D9EAF7"
GREEN = "70AD47"
LIGHT_GREEN = "E2F0D9"
YELLOW = "FFF2CC"
ORANGE = "F4B183"
GRAY = "D9E1F2"
WHITE = "FFFFFF"
DARK = "1F2937"


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    workbook = load_workbook(SOURCE, keep_vba=True)
    todos_values = pd.read_excel(SOURCE, sheet_name="TODOS", header=None, engine="openpyxl")

    for sheet_name in ["INICIO", "BASE_ESTANDAR", "LUCY_EXPORT", "DICCIONARIO"]:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    build_inicio(workbook)
    base_rows = extract_standard_base(todos_values)
    build_base_estandar(workbook, base_rows)
    build_lucy_export(workbook, base_rows)
    build_diccionario(workbook)
    reorder_sheets(workbook)
    apply_common_settings(workbook)

    workbook.save(OUTPUT)
    print(OUTPUT.resolve())


def build_inicio(workbook) -> None:
    ws = workbook.create_sheet("INICIO", 0)
    ws.sheet_properties.tabColor = BLUE
    ws["A1"] = "ARRIENDOS - BASE DE DATOS ESTANDARIZADA"
    ws["A1"].font = Font(bold=True, color=WHITE, size=16)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE)
    ws.merge_cells("A1:H1")

    sections = [
        ("Objetivo", "Conservar el archivo macro actual, pero agregar tablas limpias para usuario y automatizacion."),
        ("Flujo recomendado", "1) Usar TODOS para buscar/filtrar como antes. 2) Revisar BASE_ESTANDAR. 3) Usar LUCY_EXPORT para copiar/cargar a Lucy o CSV."),
        ("Hojas originales", "Se conservan lucy, MACHOTE, TODOS, contratos, PRORRATEO, CUENTAS, ADMON, arriendo1 y compenzaciones."),
        ("Hojas nuevas", "BASE_ESTANDAR organiza las lineas contables. LUCY_EXPORT replica el layout operativo de Lucy. DICCIONARIO explica columnas."),
        ("Macro existente", "Se conserva el VBA original para no romper el archivo. La macro actual sigue exportando CSV desde lucy/MACHOTE."),
        ("Nota tecnica", "Se evita usar CECO textuales como VARIABLE en la base limpia cuando existe un CECO real en la cabecera del grupo."),
    ]

    row = 3
    for title, text in sections:
        ws.cell(row=row, column=1).value = title
        ws.cell(row=row, column=1).font = Font(bold=True, color=WHITE)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=GREEN)
        ws.cell(row=row, column=2).value = text
        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        row += 2

    ws["A17"] = "Orden sugerido"
    ws["A17"].font = Font(bold=True, color=WHITE)
    ws["A17"].fill = PatternFill("solid", fgColor=BLUE)
    ws["B17"] = "INICIO -> BASE_ESTANDAR -> LUCY_EXPORT -> TODOS -> MACHOTE -> lucy -> maestros"
    ws.merge_cells("B17:H17")
    set_widths(ws, {"A": 24, "B": 90})


def extract_standard_base(todos_values: pd.DataFrame) -> list[dict]:
    rows: list[dict] = []
    current: dict = {}

    for idx, series in todos_values.iterrows():
        row_idx = idx + 1
        if row_idx < 9:
            continue
        values = series.reindex(range(25)).tolist()
        category = clean(values[0])
        vendor = values[1]
        account_or_name = values[2]
        amount = values[3]
        tax_code = values[4]
        posting_key = values[5]
        text = values[6]
        store = values[6]
        group_ceco = values[7]
        line_ceco = values[8]

        if category in {"ADMINISTRACION", "ARRIENDOS"}:
            current = {
                "categoria": category,
                "vendor": vendor,
                "razon_social": account_or_name,
                "tienda": store,
                "ceco_grupo": group_ceco,
                "end_of_term": values[9],
                "status_en_rem": values[10],
                "rent_min": values[11],
                "valor_actual_canon": values[14],
                "iva_canon": values[15],
                "iva_deducible_vw": values[16],
                "iva_prorrateo_vq": values[17],
                "porcentaje_iva_vw": values[18],
                "porcentaje_iva_vq": values[19],
            }
            continue

        if not current:
            continue

        if clean(account_or_name) == "TOTAL":
            continue

        if not is_line_row(vendor, amount, posting_key, text):
            continue

        ceco = line_ceco
        if pd.isna(ceco) or ceco in (None, ""):
            ceco = current.get("ceco_grupo")
        profit_center = ceco
        rows.append(
            {
                "categoria": current.get("categoria"),
                "vendor": current.get("vendor"),
                "razon_social": current.get("razon_social"),
                "tienda": current.get("tienda"),
                "ceco": ceco,
                "profit_center": profit_center,
                "account": vendor,
                "account_name": account_or_name,
                "amount": amount,
                "tax_code": tax_code,
                "posting_key": posting_key,
                "text": text,
                "end_of_term": current.get("end_of_term"),
                "status_en_rem": current.get("status_en_rem"),
                "rent_min": current.get("rent_min"),
                "valor_actual_canon": current.get("valor_actual_canon"),
                "iva_canon": current.get("iva_canon"),
                "iva_deducible_vw": current.get("iva_deducible_vw"),
                "iva_prorrateo_vq": current.get("iva_prorrateo_vq"),
                "porcentaje_iva_vw": current.get("porcentaje_iva_vw"),
                "porcentaje_iva_vq": current.get("porcentaje_iva_vq"),
                "source_sheet": "TODOS",
                "source_row": row_idx,
            }
        )

    return rows


def build_base_estandar(workbook, rows: list[dict]) -> None:
    ws = workbook.create_sheet("BASE_ESTANDAR", 1)
    ws.sheet_properties.tabColor = GREEN
    headers = [
        "CATEGORIA",
        "VENDOR",
        "RAZON SOCIAL",
        "TIENDA",
        "CECO",
        "PROFIT CENTER",
        "CUENTA",
        "NOMBRE CUENTA",
        "IMPORTE",
        "CODIGO IMPUESTO",
        "D/H",
        "TEXTO",
        "END OF TERM",
        "STATUS EN REM",
        "RENT MIN",
        "VALOR ACTUAL CANON",
        "IVA CANON",
        "IVA DEDUCIBLE VW",
        "IVA PRORRATEO VQ",
        "PORCENTAJE IVA VW",
        "PORCENTAJE IVA VQ",
        "SOURCE SHEET",
        "SOURCE ROW",
    ]
    write_table(ws, headers, [[row.get(key) for key in row_keys()] for row in rows], "BaseEstandarTable")
    ws["A1"].comment = None
    set_widths(
        ws,
        {
            "A": 18,
            "B": 14,
            "C": 34,
            "D": 14,
            "E": 12,
            "F": 14,
            "G": 16,
            "H": 28,
            "I": 16,
            "J": 18,
            "K": 10,
            "L": 32,
        },
    )


def build_lucy_export(workbook, rows: list[dict]) -> None:
    ws = workbook.create_sheet("LUCY_EXPORT", 2)
    ws.sheet_properties.tabColor = ORANGE
    headers = [
        "D/A$S/H",
        "Imp$Amount",
        "Cod_Iva$Vat Code",
        "Conto$GL Account",
        "CdC$Cost Center",
        "ProfitCenter$Profit Center",
        "Text$Text",
        "Ordine$Internal Order",
        "Wbs$WBS",
        "Attribuzione$Assignment",
        "Network$Network",
        "OpNetwork$OpNetwork",
        "MaterialNumber$Material Number",
        "Quantity$Quantity",
        "TransactionType$Transaction Type",
        "POR$POR#",
        "CompanyGL$CompanyGL",
        "Brand$Brand",
        "TradingPartner$Trading Partner",
        "PosOrd$Order Position",
        "OdaOdv$PO/SR",
        "SOURCE ROW",
    ]
    export_rows = []
    for row in rows:
        export_rows.append(
            [
                row.get("posting_key"),
                row.get("amount"),
                row.get("tax_code"),
                row.get("account"),
                None,
                row.get("profit_center"),
                row.get("text"),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                row.get("source_row"),
            ]
        )
    write_table(ws, headers, export_rows, "LucyExportTable")
    set_widths(ws, {"A": 12, "B": 16, "C": 18, "D": 18, "E": 18, "F": 24, "G": 32})


def build_diccionario(workbook) -> None:
    ws = workbook.create_sheet("DICCIONARIO", 3)
    ws.sheet_properties.tabColor = YELLOW
    headers = ["HOJA", "COLUMNA", "USO", "FUENTE / COMENTARIO"]
    rows = [
        ["BASE_ESTANDAR", "CATEGORIA", "Identifica si la linea viene de ADMINISTRACION o ARRIENDOS.", "Tomado de la cabecera de grupo en TODOS."],
        ["BASE_ESTANDAR", "VENDOR", "Vendor SAP del grupo.", "Tomado de TODOS."],
        ["BASE_ESTANDAR", "CECO", "Centro de costo limpio para cada linea.", "Usa la linea si existe, si no toma el CECO de cabecera."],
        ["BASE_ESTANDAR", "CUENTA", "Cuenta contable de la linea.", "Tomada de las lineas filtrables de TODOS."],
        ["BASE_ESTANDAR", "IMPORTE", "Valor que se copia a Lucy/SAP.", "Tomado de la linea contable."],
        ["BASE_ESTANDAR", "CODIGO IMPUESTO", "Codigo fiscal VQ, VW, V0, etc.", "Tomado de TODOS/MACHOTE."],
        ["BASE_ESTANDAR", "D/H", "40 debe/cargo, 50 haber/abono, 31 proveedor.", "Segun notas del archivo original."],
        ["BASE_ESTANDAR", "TEXTO", "Texto operativo para Lucy.", "Tomado de la formula/texto de TODOS."],
        ["LUCY_EXPORT", "Todas", "Layout final para copiar/exportar a Lucy.", "Replica headers actuales de la hoja lucy."],
        ["TODOS", "Original", "Hoja operativa original para filtro manual.", "Se conserva para no romper el uso actual."],
        ["MACHOTE/lucy", "Original", "Salida actual del macro.", "Se conserva junto con el VBA original."],
    ]
    write_table(ws, headers, rows, "DiccionarioTable")
    set_widths(ws, {"A": 18, "B": 24, "C": 62, "D": 62})


def write_table(ws, headers: list[str], rows: list[list], table_name: str) -> None:
    ws.append(headers)
    for row in rows:
        ws.append(row)

    style_header(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    if rows:
        table = Table(displayName=table_name, ref=ws.dimensions)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)


def row_keys() -> list[str]:
    return [
        "categoria",
        "vendor",
        "razon_social",
        "tienda",
        "ceco",
        "profit_center",
        "account",
        "account_name",
        "amount",
        "tax_code",
        "posting_key",
        "text",
        "end_of_term",
        "status_en_rem",
        "rent_min",
        "valor_actual_canon",
        "iva_canon",
        "iva_deducible_vw",
        "iva_prorrateo_vq",
        "porcentaje_iva_vw",
        "porcentaje_iva_vq",
        "source_sheet",
        "source_row",
    ]


def reorder_sheets(workbook) -> None:
    preferred = [
        "INICIO",
        "BASE_ESTANDAR",
        "LUCY_EXPORT",
        "DICCIONARIO",
        "TODOS",
        "MACHOTE",
        "lucy",
        "PRORRATEO",
        "contratos",
        "CUENTAS",
        "compenzaciones",
        "ADMON",
        "arriendo1",
    ]
    workbook._sheets.sort(key=lambda ws: preferred.index(ws.title) if ws.title in preferred else len(preferred))


def apply_common_settings(workbook) -> None:
    for ws in workbook.worksheets:
        if ws.title in {"BASE_ESTANDAR", "LUCY_EXPORT", "DICCIONARIO"}:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(vertical="center", wrap_text=False)
        if ws.title == "TODOS":
            ws.sheet_properties.tabColor = BLUE
        if ws.title == "MACHOTE":
            ws.sheet_properties.tabColor = ORANGE
        if ws.title == "lucy":
            ws.sheet_properties.tabColor = ORANGE


def style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_widths(ws, widths: dict[str, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for idx in range(1, ws.max_column + 1):
        letter = get_column_letter(idx)
        if ws.column_dimensions[letter].width is None:
            ws.column_dimensions[letter].width = 14


def clean(value) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text.upper() if text else None


def is_line_row(account, amount, posting_key, text) -> bool:
    if pd.isna(account) or pd.isna(amount) or pd.isna(posting_key):
        return False
    if account in (None, "") or amount in (None, "") or posting_key in (None, ""):
        return False
    if clean(account) == "TOTAL" or clean(text) == "TOTAL":
        return False
    try:
        float(str(amount).replace(",", "").replace("-", ""))
    except ValueError:
        return False
    return True


if __name__ == "__main__":
    main()
