import csv
import logging
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import holidays
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from .pdf_reader import InvoiceSupportReader
from .money import allocate_cop, round_cop


LOGGER = logging.getLogger(__name__)


LEARNING_MONTH_PATTERN = (
    r"ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|"
    r"SEPTIEMBRE|SETIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE|"
    r"ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|SEPT|OCT|NOV|DIC"
)


def normalize_learning_phrase(value) -> str:
    """Remove invoice-period tokens and per-invoice numeric readings so learned
    phrases remain reusable across months."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return ""

    text = re.sub(r"\b20\d{2}\s*[-/.]\s*(?:0?[1-9]|1[0-2])\s*[-/.]\s*(?:0?[1-9]|[12]\d|3[01])\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:0?[1-9]|[12]\d|3[01])\s*[-/.]\s*(?:0?[1-9]|1[0-2])\s*[-/.]\s*20\d{2}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b20\d{2}\s*[-/.]\s*(?:0?[1-9]|1[0-2])\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:0?[1-9]|1[0-2])\s*[-/.]\s*(?:20)?\d{2}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(rf"\b(?:{LEARNING_MONTH_PATTERN})\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b20\d{2}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:MES|PER[IÍ]ODO|A[NÑ]O)\b", " ", text, flags=re.IGNORECASE)

    text = re.sub(r"\b\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?\b", " ", text)
    text = re.sub(r"\b\d+[.,]\d+\b", " ", text)
    text = re.sub(r"\b\d{3,}\b", " ", text)
    text = re.sub(r"\b(FACTOR|CONSUMO|CONTADOR)\s+\d+\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(LECTURA\s+(?:ANT|ACT)[\.\:]?)\s*\d*\b", r"\1", text, flags=re.IGNORECASE)

    text = re.sub(r"(?:\s+\b(?:DE|DEL)\b)+\s*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^[\s\-_/.,;:|\u2013\u2014]+", "", text)
    text = re.sub(r"[\s\-_/.,;:|\u2013\u2014]+$", "", text)
    return re.sub(r"\s+", " ", text).strip()


class PipelineError(Exception):
    """Controlled error for pipeline and upload issues."""


@dataclass
class PipelineResult:
    output_csv: str
    invoice_csv_bundle: str
    manual_style_csv_bundle: str
    output_rows: int
    header_rows: int
    validation_rows: int
    processing_issues: list[dict]
    logs: list[str]
    warnings: list[str]

    def to_dict(self) -> dict:
        return {
            "output_csv": self.output_csv,
            "invoice_csv_bundle": self.invoice_csv_bundle,
            "manual_style_csv_bundle": self.manual_style_csv_bundle,
            "output_rows": self.output_rows,
            "header_rows": self.header_rows,
            "validation_rows": self.validation_rows,
            "processing_issues": self.processing_issues,
            "logs": self.logs,
            "warnings": self.warnings,
        }


class LeaseAccountingPipeline:
    TEXT_MAX_LENGTH = 50
    PAYMENT_TERMS_DEFAULT = "0001"
    PAYMENT_MODE_DEFAULT = "B - AR-Customer ACH(PRL Franchise)"
    PARTNER_BANK_DEFAULT = "_ - _"
    CONCEPT_RULES = []
    ACCOUNT_FALLBACK_CONCEPTS = {}
    ITEM_CONCEPT_RULES = [
        ("VIGILANCIA", [r"\bVIGILANCIA\b"]),
        ("ENERGÍA", [r"\bENERGIA\b", r"\bELECTRICA(?:S)?\b", r"\bKWH\b", r"\bLUZ\b"]),
        ("ACUED", [r"\bAGUA\b", r"\bACUEDUCTO\b", r"\bALCANTARILLADO\b", r"\bWATER\b", r"\bSEWER\b", r"\bDRENAJE\b", r"\bAGUACONTADOR\b"]),
        ("MTTO", [r"\bMANTENIMIENTO\b", r"\bMANTO\b", r"\bMTTO\b", r"\bREPARACION(?:ES)?\b"]),
        ("ASEO /AIRE ACON/FUMIGACION", [r"\bASEO\b", r"\bAIRE\b", r"\bACONDICIONADO\b", r"\bFUMIGACION\b"]),
        ("INTERES", [r"\bINTERES(?:ES)?\b", r"\bMORA\b"]),
        ("FP", [r"\bMERCADEO\b", r"\bFONDO\b", r"\bIMPREVISTO(?:S)?\b", r"\bPROMOCION\b", r"\bRESERVA\b"]),
        ("GV", [r"\bPUBLICIDAD\b", r"\bPUBLICIDADO\b", r"\bPARQUEADERO\b"]),
        ("GC", [r"\bEXPENSAS?\b", r"\bEXPENSAS?\s+COMUNES?\b", r"\bGASTO\s+COMUN\b", r"\bADMIN", r"\bADMON\b", r"\bCUOTADEADMINISTRACION\b", r"\bCOPROPIEDAD\b", r"\bPROPIEDAD\s+HORIZONTAL\b", r"\bMODULO\s+GENERAL\b", r"\bCOMUNES?\s+GENERALES?\b"]),
    ]
    RENT_ITEM_PATTERNS = [
        r"\bESPACIO\s+FIJO\b",
        r"\bCONCESION\s+DE\s+ESPACIO\s+FIJO\b",
        r"\bFIJ[OA]\b",
        r"\bVALOR\s+MINIMO\b",
        r"\bMINIM[OA]\b",
        r"\bMINIMO\s+MENSUAL\b",
        r"\bSUMA\s+MINIMA\b",
        r"\bSUMA\s+MINIMA\s+GARANTIZADA\b",
        r"\bGARANTIZADA\b",
        r"\bCONCESION\s+MINIMO\b",
        r"\bVARIABLE\b",
        r"\bVBLE\b",
        r"\bESPACIO\s+VARIABLE\b",
        r"\bCONCESION\s+DE\s+ESPACIO\s+VARIABLE\b",
        r"\bCANON\s+VARIABLE\b",
        r"\bARRENDAMIENTO\s+VARIABLE\b",
        r"\bARRIENDO\s+VARIABLE\b",
        r"\bVALOR\s+PORCENTUAL\b",
        r"\bPORCENTUAL\b",
        r"\bAJUSTE\s+ARRENDAMIENTO\s+VARIABLE\b",
        r"\bAJUSTE\s+CANON\b",
        r"\bAJUSTE\s+CANON\s+ARRENDAMIENTO\b",
        r"\bVENTAS\s+AL\s+DETALLE\s+VBL(?:E)?\b",
        r"\bARRENDAMIENTO\b",
        r"\bARRIENDO\b",
        r"\bCANON\b",
        r"\bCONCESION\b",
        r"\bINMUEBLE(?:S)?\b",
    ]
    MACRO_LUCY_COLUMNS = [
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
    ]
    HEADER_EXPORT_COLUMNS = [
        "Supplier Email",
        "Worfklow Closing Causal",
        "Total Vat Amount",
        "checkSkipSubjectChannel",
        "Withholding Tax Amount",
        "Legal entity",
        "Payment Block",
        "Corporate Approval",
        "Backup Archive Counter",
        "Vat Exempt 3",
        "Extra costs sum",
        "Withholding Tax Code",
        "Sap Protocol",
        "File Type",
        "Transaction",
        "Indirect Costs",
        "Withholding Base Amount",
        "RetryWfDocumentsReprocessingCounter",
        "Ind. Ret.",
        "Original Invoice Number",
        "Extract Flag",
        "Autoref3",
        "Assignment 1",
        "Mail Subject",
        "Post date",
        "System",
        "Period",
        "Order Proposal",
        "Bank",
        "Po Item",
        "Vat Rate 3",
        "Vat Rate 2",
        "Vat Rate 1",
        "Invoice Date",
        "Original Invoice Ref Number",
        "Invoice Number",
        "Document Type Sap",
        "PO",
        "original invoice SeeNeo POs",
        "Delivery Note",
        "Currency",
        "Vat Code",
        "Invoice Total Amount",
        "Total Impuestos Retenidos",
        "MIRO Type",
        "Total Taxable Amount",
        "Total VAT Amount",
        "Vat Exempt 1",
        "Vat Exempt 2",
        "Payment Terms",
        "Payment Mode",
        "Allocated Costs",
        "System.1",
        "Withholding Tax",
        "Partner Bank",
        "Custom Workflow Type",
        "Action",
        "POR Creator",
        "Taxable 4",
        "Iban",
        "Vat Amount 4",
        "Taxable 10",
        "Vat Amount 10",
        "Balance",
        "Vat Reporting Date",
        "Accounting Counterpart",
        "UUID",
        "Exchange",
        "APS",
        "Exchange Date",
        "Sdidate",
        "Parking Date",
        "Line Item",
        "Doc in Latency",
        "Business Place",
        "Section Code",
        "Ksef ID",
        "Ksef Date",
        "LETTINT",
        "CONAI",
        "GUID",
        "FISCAL_CODE",
        "Block Simulate",
        "Litigation Reason",
        "Total From SAP",
        "POR PROCESSED",
        "Wf Banca",
        "Wf Payment Methods",
        "Freight",
        "EC Autorization",
        "EC Tipo de Pago",
        "EC Pais Recep.Pago",
        "EC Pais Exento",
        "CH QR IBAN",
        "CH QR REF",
        "EC Aplica Convinio",
        "Kid Number",
        "EC Forma de Pago",
        "EC CODIGO",
        "EC ZEICE",
        "SCB Indicator",
        "Natura",
        "Detraction",
        "OCR Vendor Mail",
        "EC Pais Recep.Pago2",
        "Tax Stamp",
        "Transformed invoice reference number ",
        "MarkID",
        "EC Aplica Convinio 2",
        "Predictive AI Response",
        "Result",
        "Calculate Tax",
        "Receipt Date",
        "Last Sap Execution",
        "SAP_KOSTL",
        "SAP_PROJK",
        "Due Date",
        "Sap Posting Date",
        "Sap Payment Block",
        "SdiId",
        "Sap message",
        "Archive Link Done",
        "Litigation Causal",
        "Group Company",
        "ZCTT",
        "SAP_ARCHIVE",
        "Parking Protocol",
        "Archived Reg. Protocol",
        "Archived Reg. Date",
    ]
    GROUP_COLORS = [
        ("6F42C1", ["E9D8FD", "F3E8FF", "FAF5FF"]),
        ("0F766E", ["CCFBF1", "DCFCE7", "F0FDFA"]),
        ("C2410C", ["FFEDD5", "FEF3C7", "FFF7ED"]),
        ("1D4ED8", ["DBEAFE", "E0F2FE", "EFF6FF"]),
        ("BE123C", ["FFE4E6", "FCE7F3", "FFF1F2"]),
        ("3F6212", ["ECFCCB", "FEF9C3", "F7FEE7"]),
    ]
    HEADER_FILL = "1F2937"
    HEADER_FONT = "FFFFFF"

    DEFAULT_LEARNING_DICTIONARY = Path(__file__).resolve().parents[3] / "Diccionario_Conceptos_Simple.xlsx"

    def __init__(self, output_dir: Path, learning_dictionary_path: Path | None = None, account_reference_path: Path | None = None):
        self.base_output_dir = Path(output_dir)
        self.base_output_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir = self.base_output_dir
        self.logs: list[str] = []
        self.warnings: list[str] = []
        self.contract_review_items: list[dict] = []
        self.discount_review_items: list[dict] = []
        self.beneficiary_review_items: list[dict] = []
        self.preflight_processing_issues: list[dict] = []
        self.distribution_rules = pd.DataFrame()
        self.learning_dictionary_path = Path(learning_dictionary_path) if learning_dictionary_path else self.DEFAULT_LEARNING_DICTIONARY
        self.learning_dictionary = pd.DataFrame()
        self.account_reference_path = Path(account_reference_path) if account_reference_path else self._default_account_reference_path()
        self.FIXED_ACCOUNT = None
        self.VARIABLE_ACCOUNT = None
        self.VAT_ACCOUNT = None
        self.account_reference = self._load_account_reference(self.account_reference_path)
        self._configure_account_reference()

    def _configure_account_reference(self) -> None:
        self.FIXED_ACCOUNT = self._account_code_for_concept("RF")
        self.VARIABLE_ACCOUNT = self._account_code_for_concept("RV")
        self.ADMIN_SELL_CAM_ACCOUNT = self._account_code_for_description("ADMINISTRACION SELL CAM")
        self.ADMIN_VARIABLE_ACCOUNT = self._account_code_for_description("ADMINISTRACION VARIABLE")
        self.VAT_ACCOUNT = (
            self._account_code_for_concept("IVA")
            or self._account_code_for_concept("VAT")
            or self._account_code_for_concept("IMPUESTO AL VALOR AGREGADO")
        )

    def _default_account_reference_path(self) -> Path:
        candidates = [
            Path(__file__).resolve().parents[3] / "data" / "input" / "FACTURAS_CONTABILIZADAS.xlsx",
            Path(__file__).resolve().parents[3] / "FACTURAS_CONTABILIZADAS.xlsx",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]

    def _load_account_reference(self, path: Path | None) -> pd.DataFrame:
        target_path = Path(path) if path else self._default_account_reference_path()
        if not target_path.exists():
            self.CONCEPT_RULES = []
            self.ACCOUNT_FALLBACK_CONCEPTS = {}
            return pd.DataFrame()

        try:
            sheets = pd.read_excel(target_path, sheet_name=None, dtype=str)
        except Exception:
            self.CONCEPT_RULES = []
            self.ACCOUNT_FALLBACK_CONCEPTS = {}
            return pd.DataFrame()

        account_aliases = {"cuenta", "cuenta_contable", "gl_account", "account", "codigo", "codigo_cuenta", "n_cuenta"}
        description_aliases = {"descripcion", "description", "nombre_cuenta", "nombre_de_cuenta", "texto", "detalle", "nombre"}
        concept_aliases = {"concepto", "concept", "tipo", "rent_type", "serie", "codigo_concepto", "tipo_cuenta", "texto_corto_lucy"}
        reference_tables = []
        for sheet_name, sheet in sheets.items():
            if sheet.empty:
                continue
            columns = {self._normalize_column_name(column): column for column in sheet.columns}
            account_column = next((columns[key] for key in account_aliases if key in columns), None)
            concept_column = next((columns[key] for key in concept_aliases if key in columns), None)
            description_column = next((columns[key] for key in description_aliases if key in columns), None)
            if not account_column:
                header_row = self._find_reference_header_row(sheet, account_aliases | description_aliases | concept_aliases)
                if header_row is not None:
                    sheet = sheet.iloc[header_row + 1:].copy()
                    sheet.columns = [self._normalize_column_name(value) for value in sheets[sheet_name].iloc[header_row]]
                    columns = {str(column): column for column in sheet.columns}
                    account_column = next((columns[key] for key in account_aliases if key in columns), None)
                    concept_column = next((columns[key] for key in concept_aliases if key in columns), None)
                    description_column = next((columns[key] for key in description_aliases if key in columns), None)
            if not account_column or not concept_column:
                if not account_column or not description_column:
                    continue

            selected = sheet[[account_column]].copy()
            selected.columns = ["account"]
            selected["description"] = sheet[description_column]
            selected["concept"] = (
                sheet[concept_column]
                if concept_column
                else selected["description"].map(self._infer_reference_concept)
            )
            reference_tables.append(selected)

        if not reference_tables:
            self.CONCEPT_RULES = []
            self.ACCOUNT_FALLBACK_CONCEPTS = {}
            return pd.DataFrame()

        normalized = pd.concat(reference_tables, ignore_index=True).fillna("")

        normalized["account"] = normalized["account"].map(self._clean_numeric_code)
        normalized["concept"] = normalized["concept"].map(self._clean_text)
        normalized["description"] = normalized["description"].map(self._clean_text)
        normalized = normalized[normalized["account"].notna() & normalized["concept"].astype(str).str.strip().ne("")]
        normalized = normalized.drop_duplicates(subset=["account", "concept", "description"])
        self.CONCEPT_RULES = [
            (row["account"], row["description"], row["concept"])
            for _, row in normalized.iterrows()
            if row["description"]
        ]
        self.ACCOUNT_FALLBACK_CONCEPTS = {
            row["account"]: row["concept"]
            for _, row in normalized.drop_duplicates(subset=["account"], keep="first").iterrows()
            if row["account"] and row["concept"]
        }
        return normalized

    def _find_reference_header_row(self, sheet: pd.DataFrame, aliases: set[str]) -> int | None:
        for row_index, row in sheet.iterrows():
            values = {self._normalize_column_name(value) for value in row.tolist() if self._clean_text(value)}
            if values & aliases:
                return int(row_index)
        return None

    def _infer_reference_concept(self, value) -> str:
        text = self._concept_key(value)
        if not text:
            return ""
        if "ARRIENDO FIJO" in text or "ARRIENDOS FIJOS" in text:
            return "RF"
        if "ARRIENDO VARIABLE" in text:
            return "RV"
        if "CUENTA DEL IVA" in text or text == "IVA" or " IMPUESTO AL VALOR " in f" {text} ":
            return "IVA"
        if "INTERES" in text or "MORA" in text:
            return "INTERES"
        if "ADMINISTRACION" in text or "ADMON" in text:
            return "GC"
        if "FONDO" in text or "PROMOCION" in text:
            return "FP"
        if "ENERGIA" in text or "ELECTR" in text:
            return "ENERGÍA"
        if any(token in text for token in ("AGUA", "ACUEDUCTO", "ALCANTARILLADO", "WATER", "SEWER", "DRENAJE")):
            return "ACUED"
        if "MANTENIMIENTO" in text:
            return "MTTO"
        if any(token in text for token in ("ASEO", "AIRE", "FUMIGACION")):
            return "ASEO /AIRE ACON/FUMIGACION"
        if "VIG" in text:
            return "VIGILANCIA"
        if any(token in text for token in ("GTO VTA", "GASTOS", "RECUPERACION", "PARQUEADERO")):
            return "GV"
        return ""

    def _account_code_for_concept(self, concept: str | None) -> str | None:
        target = self._concept_key(concept)
        if not target:
            return None
        for account_code, mapped_concept in self.ACCOUNT_FALLBACK_CONCEPTS.items():
            if self._concept_key(mapped_concept) == target:
                return account_code
        return None

    def _account_code_for_description(self, description: str | None) -> str | None:
        target = self._concept_key(description)
        if not target or self.account_reference.empty:
            return None
        matches = self.account_reference[
            self.account_reference["description"].map(self._concept_key).eq(target)
        ]
        accounts = matches["account"].dropna().astype(str).unique()
        if len(accounts) > 1:
            raise PipelineError(
                f"La descripción de cuenta '{description}' está asociada a más de una cuenta contable."
            )
        return accounts[0] if len(accounts) == 1 else None

    def run(
        self,
        invoices_path: Path,
        control_path: Path,
        contracts_path: Path,
        prorateo_path: Path,
        distribution_path: Path,
        history_path: Path | None = None,
        support_paths: list[Path] | None = None,
        support_split_factors: dict[str, int] | None = None,
        support_split_weights: dict[str, list[str | float]] | None = None,
        macro_template_path: Path | None = None,
        period: str | None = None,
        contract_selections: dict[str, str] | None = None,
        discount_selections: dict[str, str] | None = None,
        beneficiary_selections: dict[str, str] | None = None,
    ) -> dict:
        self.logs = []
        self.warnings = []
        self.contract_review_items = []
        self.discount_review_items = []
        self.beneficiary_review_items = []
        self.preflight_processing_issues = []
        run_folder_name = pd.Timestamp.now().strftime("%Y.%m.%d_%H.%M")
        self.output_dir = self.base_output_dir / run_folder_name
        self.output_dir.mkdir(parents=True, exist_ok=True)
        data, output_df, validation_df = self.prepare(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths,
            support_split_factors=support_split_factors,
            support_split_weights=support_split_weights,
            macro_template_path=macro_template_path,
            period=period,
            contract_selections=contract_selections,
            discount_selections=discount_selections,
            beneficiary_selections=beneficiary_selections,
        )
        return self.generate_output(
            output_df,
            validation_df,
            macro_database_df=data.get("macro_database", pd.DataFrame()),
            macro_template_path=macro_template_path,
        ).to_dict()

    def prepare(
        self,
        invoices_path: Path,
        control_path: Path,
        contracts_path: Path,
        prorateo_path: Path,
        distribution_path: Path,
        history_path: Path | None = None,
        support_paths: list[Path] | None = None,
        support_split_factors: dict[str, int] | None = None,
        support_split_weights: dict[str, list[str | float]] | None = None,
        macro_template_path: Path | None = None,
        period: str | None = None,
        contract_selections: dict[str, str] | None = None,
        discount_selections: dict[str, str] | None = None,
        beneficiary_selections: dict[str, str] | None = None,
    ) -> tuple[dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame]:
        self.warnings = []
        self.contract_review_items = []
        self.discount_review_items = []
        self.beneficiary_review_items = []
        self.preflight_processing_issues = []
        if history_path is not None:
            self.account_reference_path = Path(history_path)
            self.account_reference = self._load_account_reference(self.account_reference_path)
            self._configure_account_reference()
        if not self.FIXED_ACCOUNT or not self.VARIABLE_ACCOUNT:
            raise PipelineError(
                "No se encontraron las cuentas RF y RV en FACTURAS_CONTABILIZADAS.xlsx. "
                "Revisa que el consolidado incluya las columnas de cuenta y concepto/serie."
            )
        self.learning_dictionary = self._load_learning_dictionary(self.learning_dictionary_path)
        data = self.load_data(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths,
            support_split_factors=support_split_factors,
            support_split_weights=support_split_weights,
            macro_template_path=macro_template_path,
        )
        data = self.clean_data(data, beneficiary_selections=beneficiary_selections or {})
        data = self._exclude_manual_discount_invoices(data, discount_selections or {})
        merged = self.merge_data(data, contract_selections=contract_selections)
        output_df, validation_df = self.apply_rules(
            merged,
            period=period,
            discount_selections=discount_selections or {},
        )
        return data, output_df, validation_df

    def load_data(
        self,
        invoices_path: Path,
        control_path: Path,
        contracts_path: Path,
        prorateo_path: Path,
        distribution_path: Path,
        history_path: Path | None = None,
        support_paths: list[Path] | None = None,
        support_split_factors: dict[str, int] | None = None,
        support_split_weights: dict[str, list[str | float]] | None = None,
        macro_template_path: Path | None = None,
    ) -> dict[str, pd.DataFrame]:
        self._log("Loading input files.")
        return {
            "invoices": self._load_invoices(invoices_path),
            "control": self._load_control(control_path),
            "contracts": self._load_contracts(contracts_path),
            "prorateo": self._load_prorateo(prorateo_path),
            "distribution": self._load_distribution(distribution_path),
            "history": self._safe_read_table(history_path) if history_path else pd.DataFrame(),
            "invoice_supports": self._load_invoice_supports(
                support_paths or [],
                support_split_factors=support_split_factors,
                support_split_weights=support_split_weights,
            ),
            "macro_database": self._load_macro_database(macro_template_path) if macro_template_path else pd.DataFrame(),
        }

    def clean_data(
        self,
        data: dict[str, pd.DataFrame],
        beneficiary_selections: dict[str, str] | None = None,
    ) -> dict[str, pd.DataFrame]:
        self._log("Cleaning and standardizing data.")
        cleaned: dict[str, pd.DataFrame] = {}
        for name, frame in data.items():
            if frame.empty:
                cleaned[name] = frame.copy()
                continue
            df = frame.copy()
            df.columns = [self._normalize_column_name(col) for col in df.columns]
            for col in df.columns:
                if col in {
                    "invoice_line_items",
                    "line_items",
                    "support_line_items",
                    "invoice_discounts",
                    "discounts",
                    "support_discounts",
                    "split_weights",
                    "support_split_weights",
                }:
                    continue
                if df[col].dtype == object:
                    df[col] = df[col].map(self._clean_text)
            cleaned[name] = df

        invoices = cleaned["invoices"]
        for col in [
            "invoice_id",
            "vendor",
            "supplier_nit",
            "store",
            "reference",
            "total",
            "vat",
            "invoice_date",
            "due_date",
            "ceco",
            "withholding_tax",
            "payment_terms_text",
            "item_count_hint",
            "ubl_document_type",
            "invoice_line_items",
            "invoice_discounts",
            "payable_rounding",
        ]:
            if col not in invoices.columns:
                invoices[col] = (
                    [[] for _ in range(len(invoices))]
                    if col in {"invoice_line_items", "invoice_discounts"}
                    else None
                )
        invoices["invoice_date"] = pd.to_datetime(invoices["invoice_date"], errors="coerce")
        invoices["due_date"] = pd.to_datetime(invoices["due_date"], errors="coerce")
        invoices["total"] = pd.to_numeric(invoices["total"], errors="coerce").map(round_cop)
        invoices["vat"] = pd.to_numeric(invoices["vat"], errors="coerce").map(lambda value: round_cop(value, 0))
        invoices["withholding_tax"] = pd.to_numeric(invoices["withholding_tax"], errors="coerce").map(round_cop)
        invoices["item_count_hint"] = pd.to_numeric(invoices["item_count_hint"], errors="coerce")
        invoices["payable_rounding"] = pd.to_numeric(invoices["payable_rounding"], errors="coerce").map(round_cop)
        invoices["invoice_id"] = invoices["invoice_id"].map(self._normalize_invoice_id)
        invoices["vendor"] = invoices["vendor"].map(self._clean_numeric_code)
        invoices["amount"] = (invoices["total"].fillna(0) - invoices["vat"].fillna(0)).map(
            lambda value: round_cop(value, 0)
        )
        invoices["vendor_key"] = invoices["vendor"].map(self._text_key)
        invoices["store_key"] = invoices["store"].map(self._text_key)
        invoices["ceco_key"] = invoices["ceco"].map(self._ceco_key)
        invoices["invoice_vendor_key"] = invoices["invoice_id"].map(self._clean_text).fillna("") + "|" + invoices["vendor_key"].fillna("")
        invoices["invoice_pdf_key"] = invoices["invoice_id"].map(self._invoice_key)
        if "invoice_input_type" not in invoices.columns:
            invoices["invoice_input_type"] = "table"
        else:
            invoices["invoice_input_type"] = invoices["invoice_input_type"].fillna("table")
        table_input = invoices["invoice_input_type"].eq("table").any()

        blank_invoice_rows = int(invoices["invoice_pdf_key"].isna().sum())
        if blank_invoice_rows:
            invoices = invoices[invoices["invoice_pdf_key"].notna()].copy()
            self._log(f"Ignored {blank_invoice_rows} blank rows from the invoices file.")

        invoice_supports = cleaned["invoice_supports"]
        if not invoice_supports.empty:
            invoice_supports["invoice_key"] = invoice_supports["invoice_id"].map(self._invoice_key)
            invoice_supports["supplier_nit"] = invoice_supports["supplier_id"]
            invoice_supports["split_factor"] = pd.to_numeric(invoice_supports.get("split_factor", 1), errors="coerce").fillna(1)
            invoice_keys = set(invoices["invoice_pdf_key"].dropna())
            unmatched_numbers = invoice_supports[
                ~invoice_supports["invoice_key"].isin(invoice_keys)
            ]
            if not unmatched_numbers.empty:
                for _, support in unmatched_numbers.iterrows():
                    invoice_id = self._clean_text(support.get("invoice_id")) or "SIN_NUMERO"
                    source_file = self._clean_text(support.get("source_file")) or "archivo adjunto"
                    self.preflight_processing_issues.append(
                        {
                            "invoice_id": invoice_id,
                            "source_file": source_file,
                            "store": None,
                            "location": source_file,
                            "problem": "El numero de factura no existe en el invoices file; no se genero su CSV",
                            "possible_solution": (
                                "Agrega el numero exacto al invoices file si la factura es valida, o retira el XML "
                                "del lote si fue enviado por error."
                            ),
                            "issue_type": "invoice_number_mismatch",
                        }
                    )
                self._record_warning(
                    f"Se omitieron {len(unmatched_numbers)} XML/PDF porque su numero de factura no existe "
                    "en el invoices file. El resto del lote continuo."
                )
                invoice_supports = invoice_supports.drop(index=unmatched_numbers.index).copy()

            if table_input:
                xml_supports = invoice_supports[
                    invoice_supports.get("source_type", pd.Series(index=invoice_supports.index, dtype=object))
                    .fillna("")
                    .str.lower()
                    .eq("xml")
                ].copy()
                if xml_supports.empty:
                    raise PipelineError(
                        "El invoices file no tiene facturas XML coincidentes. "
                        "Adjunta al menos un XML para generar archivos CSV."
                    )

                matched_xml_keys = set(xml_supports["invoice_key"].dropna())
                omitted = invoices[~invoices["invoice_pdf_key"].isin(matched_xml_keys)].copy()
                omitted_ids = list(dict.fromkeys(omitted["invoice_id"].dropna().map(str).tolist()))
                if omitted_ids:
                    preview = ", ".join(omitted_ids[:10])
                    extra = "" if len(omitted_ids) <= 10 else f" y {len(omitted_ids) - 10} más"
                    self._record_warning(
                        f"Se omitieron {len(omitted_ids)} facturas del invoices file porque no tienen XML coincidente: "
                        f"{preview}{extra}. No se generaron CSV para esas facturas."
                    )

                invoices = invoices[invoices["invoice_pdf_key"].isin(matched_xml_keys)].copy()
                duplicate_subset = ["invoice_pdf_key", "vendor_key", "store_key", "ceco_key"]
                duplicate_count = int(invoices.duplicated(subset=duplicate_subset, keep="first").sum())
                if duplicate_count:
                    invoices = invoices.drop_duplicates(subset=duplicate_subset, keep="first").copy()
                    self._record_warning(
                        f"Se omitieron {duplicate_count} filas duplicadas del invoices file para evitar generar CSV repetidos."
                    )
                invoice_supports = xml_supports

            self._warn_support_store_mismatches(invoices, invoice_supports)
            invoice_supports = self._apply_support_split_factors(invoice_supports)
            support_best = invoice_supports.sort_values(["invoice_key", "flags"], na_position="first").drop_duplicates("invoice_key")
            support_best = support_best.rename(
                columns={
                    "source_file": "support_source_file",
                    "document_type": "support_ubl_document_type",
                    "invoice_date": "support_invoice_date",
                    "due_date": "support_due_date",
                    "subtotal": "support_subtotal",
                    "total": "support_total",
                    "detected_iva": "support_detected_iva",
                    "withholding_tax": "support_withholding_tax",
                    "payment_terms_text": "support_payment_terms_text",
                    "item_count_hint": "support_item_count_hint",
                    "line_items": "support_line_items",
                    "discounts": "support_discounts",
                    "payable_rounding": "support_payable_rounding",
                    "split_factor": "support_split_factor",
                    "split_weights": "support_split_weights",
                    "flags": "support_flags",
                }
            )
            invoices = invoices.merge(
                support_best[
                    [
                        "invoice_key",
                        "support_source_file",
                        "support_ubl_document_type",
                        "support_invoice_date",
                        "support_due_date",
                        "support_subtotal",
                        "support_total",
                        "support_detected_iva",
                        "support_withholding_tax",
                        "support_payment_terms_text",
                        "support_item_count_hint",
                        "support_line_items",
                        "support_discounts",
                        "support_payable_rounding",
                        "support_split_factor",
                        "support_split_weights",
                        "support_flags",
                    ]
                ],
                left_on="invoice_pdf_key",
                right_on="invoice_key",
                how="left",
            )
            invoices["invoice_date"] = invoices["invoice_date"].fillna(invoices["support_invoice_date"])
            invoices["due_date"] = invoices["due_date"].fillna(invoices["support_due_date"])
            invoices["vat"] = invoices["vat"].where(
                invoices["vat"].fillna(0) != 0,
                invoices["support_detected_iva"],
            ).fillna(0)
            invoices["total"] = invoices["total"].where(invoices["total"].fillna(0) != 0, invoices["support_total"])
            invoices["withholding_tax"] = invoices["withholding_tax"].combine_first(invoices["support_withholding_tax"])
            invoices["payment_terms_text"] = invoices["payment_terms_text"].combine_first(invoices["support_payment_terms_text"])
            invoices["item_count_hint"] = invoices["item_count_hint"].combine_first(invoices["support_item_count_hint"])
            invoices["ubl_document_type"] = invoices["ubl_document_type"].combine_first(
                invoices["support_ubl_document_type"]
            ).fillna("Invoice")
            invoices["invoice_line_items"] = invoices["invoice_line_items"].where(
                invoices["invoice_line_items"].map(lambda value: isinstance(value, list) and len(value) > 0),
                invoices["support_line_items"],
            )
            invoices["invoice_discounts"] = invoices["invoice_discounts"].where(
                invoices["invoice_discounts"].map(lambda value: isinstance(value, list) and len(value) > 0),
                invoices["support_discounts"],
            )
            invoices["payable_rounding"] = invoices["payable_rounding"].combine_first(
                invoices["support_payable_rounding"]
            )
            invoices["amount"] = invoices.apply(self._gross_invoice_amount, axis=1)
            invoices = self._apply_beneficiary_distribution_selections(
                invoices,
                beneficiary_selections or {},
            )
            invoices = self._expand_split_invoice_rows(invoices)
        else:
            if table_input:
                raise PipelineError(
                    "El invoices file requiere al menos una factura XML coincidente. "
                    "Las filas sin XML no generan archivos CSV."
                )
            invoices["support_flags"] = None
            invoices["support_source_file"] = None
            invoices["ubl_document_type"] = invoices["ubl_document_type"].fillna("Invoice")
            invoices["support_split_factor"] = 1
            invoices["support_split_index"] = 1

        invoices["amount"] = invoices.apply(self._gross_invoice_amount, axis=1)

        control = cleaned["control"]
        control["vendor_key"] = control["vendor_code"].fillna(control["vendor"]).map(self._text_key)
        control["store_key"] = control["store"].map(self._text_key)
        control["ceco_key"] = control["ceco"].map(self._ceco_key)
        control["ceco_is_valid"] = control["ceco"].map(self._is_valid_ceco_value)
        control = control.sort_values(["ceco_is_valid", "store_key", "vendor_key"], ascending=[False, True, True])

        contracts = cleaned["contracts"]
        contracts["ceco_key"] = contracts["ceco"].map(self._ceco_key)
        contracts["end_of_term"] = pd.to_datetime(contracts["end_of_term"], errors="coerce")
        contracts["rent_min"] = pd.to_numeric(contracts["rent_min"], errors="coerce").fillna(0)
        contracts["sell_media"] = pd.to_numeric(contracts["sell_media"], errors="coerce").fillna(0)
        if "sell_cam" not in contracts.columns:
            contracts["sell_cam"] = 0
        contracts["sell_cam"] = pd.to_numeric(contracts["sell_cam"], errors="coerce").fillna(0)

        prorateo = cleaned["prorateo"]
        if "vendor_code" not in prorateo.columns:
            prorateo["vendor_code"] = None
        if "vendor" not in prorateo.columns:
            prorateo["vendor"] = None
        prorateo["vendor_key"] = prorateo["vendor_code"].fillna(prorateo["vendor"]).map(self._text_key)
        prorateo["store_key"] = prorateo["store"].map(self._text_key)
        prorateo["ceco_key"] = prorateo["ceco"].map(self._ceco_key)
        prorateo["vw_percent"] = pd.to_numeric(prorateo["vw_percent"], errors="coerce")
        prorateo["vw_percent"] = prorateo["vw_percent"].where(prorateo["vw_percent"] <= 1, prorateo["vw_percent"] / 100.0)
        if "vq_percent" not in prorateo.columns:
            prorateo["vq_percent"] = None
        prorateo["vq_percent"] = pd.to_numeric(prorateo["vq_percent"], errors="coerce")
        prorateo["vq_percent"] = prorateo["vq_percent"].where(prorateo["vq_percent"] <= 1, prorateo["vq_percent"] / 100.0)
        prorateo["vq_percent"] = prorateo["vq_percent"].combine_first((1 - prorateo["vw_percent"]).clip(lower=0, upper=1))

        distribution = cleaned["distribution"]
        distribution["source_vendor_key"] = distribution["source_vendor"].map(self._text_key)
        distribution["target_vendor"] = distribution["target_vendor"].map(self._clean_numeric_code)
        distribution["split_percent"] = pd.to_numeric(distribution["split_percent"], errors="coerce")
        distribution["split_percent"] = distribution["split_percent"].where(
            distribution["split_percent"] <= 1,
            distribution["split_percent"] / 100.0,
        )
        self.distribution_rules = distribution

        macro_database = cleaned["macro_database"]
        if not macro_database.empty:
            macro_database["vendor"] = macro_database["vendor"].map(self._clean_numeric_code)
            macro_database["store"] = macro_database["store"].map(self._clean_text)
            macro_database["ceco"] = macro_database["ceco"].map(self._clean_text)
            macro_database["profit_center"] = macro_database["profit_center"].map(self._clean_text)
            macro_database["account"] = macro_database["account"].map(self._clean_numeric_code)
            macro_database["amount"] = pd.to_numeric(macro_database["amount"], errors="coerce")
            macro_database["posting_key"] = macro_database["posting_key"].map(self._clean_text)
            macro_database["tax_code"] = macro_database["tax_code"].map(self._clean_text)
            macro_database["text"] = macro_database["text"].map(self._clean_text)
            macro_database["vendor_key"] = macro_database["vendor"].map(self._text_key)
            macro_database["store_key"] = macro_database["store"].map(self._text_key)
            macro_database["ceco_key"] = macro_database["ceco"].map(self._ceco_key)
            macro_database["end_of_term"] = pd.to_datetime(macro_database["end_of_term"], errors="coerce")
            macro_database["vw_percent"] = pd.to_numeric(macro_database["vw_percent"], errors="coerce")
            macro_database["vw_percent"] = macro_database["vw_percent"].where(
                macro_database["vw_percent"] <= 1,
                macro_database["vw_percent"] / 100.0,
            )
            if "vq_percent" not in macro_database.columns:
                macro_database["vq_percent"] = None
            macro_database["vq_percent"] = pd.to_numeric(macro_database["vq_percent"], errors="coerce")
            macro_database["vq_percent"] = macro_database["vq_percent"].where(
                macro_database["vq_percent"] <= 1,
                macro_database["vq_percent"] / 100.0,
            )
            macro_database["vq_percent"] = macro_database["vq_percent"].combine_first(
                (1 - macro_database["vw_percent"]).clip(lower=0, upper=1)
            )
            macro_database["row_rank"] = macro_database.groupby(["vendor_key", "store_key"]).cumcount() + 1

        self._log(f"Standardized {len(invoices)} invoices.")
        for idx, row in invoices.iterrows():
            support_log = ""
            if "support_total" in row and pd.notna(row["support_total"]):
                support_log = (
                    f" | Matched PDF/XML support: Date={row.get('support_invoice_date')}, "
                    f"Total={row.get('support_total')}, IVA={row.get('support_detected_iva')}, "
                    f"Withholding={row.get('support_withholding_tax')}, "
                    f"PaymentTerms={row.get('support_payment_terms_text')}"
                )
            self._log(
                f"  Invoice [{row['invoice_id']}]: Vendor={row['vendor']}, Store={row['store']}, "
                f"Date={row['invoice_date']}, Total={row['total']}, VAT={row['vat']}{support_log}"
            )

        cleaned["invoices"] = invoices
        cleaned["control"] = control
        cleaned["contracts"] = contracts
        cleaned["prorateo"] = prorateo
        cleaned["distribution"] = distribution
        cleaned["invoice_supports"] = invoice_supports
        cleaned["macro_database"] = macro_database
        return cleaned

    def _exclude_manual_discount_invoices(
        self,
        data: dict[str, pd.DataFrame],
        selections: dict[str, str],
    ) -> dict[str, pd.DataFrame]:
        skip_keys = {key for key, action in selections.items() if action == "skip"}
        invoices = data.get("invoices", pd.DataFrame())
        if not skip_keys or invoices.empty:
            return data

        invoice_keys = invoices["invoice_id"].map(self._invoice_key)
        skipped = invoices[invoice_keys.isin(skip_keys)].copy()
        for _, row in skipped.drop_duplicates(subset=["invoice_id"]).iterrows():
            invoice_id = self._clean_text(row.get("invoice_id")) or "SIN_NUMERO"
            source_file = self._clean_text(row.get("support_source_file")) or "XML sin nombre"
            self._record_warning(
                f"Factura {invoice_id} ({source_file}): omitida por decision del usuario para procesamiento manual. "
                "No se genero ningun CSV para esta factura."
            )

        data["invoices"] = invoices[~invoice_keys.isin(skip_keys)].copy()
        if data["invoices"].empty:
            raise PipelineError(
                "Todas las facturas del lote se marcaron para procesamiento manual. No se generaron archivos CSV."
            )
        return data

    def merge_data(
        self,
        data: dict[str, pd.DataFrame],
        contract_selections: dict[str, str] | None = None,
    ) -> pd.DataFrame:
        self._log("Merging invoices with reference tables.")
        invoices = data["invoices"].copy()
        control = data["control"]
        contracts = data["contracts"].copy()
        if "sell_cam" not in contracts.columns:
            contracts["sell_cam"] = 0
        macro_database = data.get("macro_database", pd.DataFrame())

        vendor_map = control.drop_duplicates(subset=["vendor_key"])[["vendor_key", "store", "ceco", "vendor_code"]].rename(
            columns={"store": "control_store_vendor", "ceco": "control_ceco_vendor", "vendor_code": "control_vendor_code_vendor"}
        )
        store_map = control.drop_duplicates(subset=["store_key"])[["store_key", "store", "ceco", "vendor_code"]].rename(
            columns={"store": "control_store_store", "ceco": "control_ceco_store", "vendor_code": "control_vendor_code_store"}
        )

        merged = invoices.merge(vendor_map, on="vendor_key", how="left")
        merged = merged.merge(store_map, on="store_key", how="left")

        prorateo = data["prorateo"]
        generic_prorateo = prorateo[
            prorateo["vendor_key"].notna() & prorateo["store_key"].isna() & prorateo["ceco_key"].isna()
        ]
        prorateo_vendor = generic_prorateo.drop_duplicates(subset=["vendor_key"])[["vendor_key", "vw_percent", "vq_percent", "ceco"]].rename(
            columns={"vw_percent": "vw_percent_vendor", "vq_percent": "vq_percent_vendor", "ceco": "prorateo_ceco_vendor"}
        )
        prorateo_store = prorateo.drop_duplicates(subset=["store_key"])[["store_key", "vw_percent", "vq_percent", "ceco"]].rename(
            columns={"vw_percent": "vw_percent_store", "vq_percent": "vq_percent_store", "ceco": "prorateo_ceco_store"}
        )
        prorateo_ceco = prorateo.drop_duplicates(subset=["ceco_key"])[["ceco_key", "vw_percent", "vq_percent"]].rename(
            columns={"vw_percent": "vw_percent_ceco", "vq_percent": "vq_percent_ceco"}
        )
        merged = merged.merge(prorateo_vendor, on="vendor_key", how="left")
        merged["prorateo_store_key"] = (
            merged[["store", "control_store_store", "control_store_vendor"]]
            .bfill(axis=1).iloc[:, 0].map(self._text_key)
        )
        merged = merged.merge(
            prorateo_store[prorateo_store["store_key"].notna()],
            left_on="prorateo_store_key", right_on="store_key", how="left", suffixes=("", "_prorateo"),
        )

        if not macro_database.empty:
            macro_header = (
                macro_database.sort_values(["vendor_key", "store_key", "row_rank"])
                .drop_duplicates(subset=["vendor_key", "store_key"])[
                    [
                        "vendor_key",
                        "store_key",
                        "ceco",
                        "profit_center",
                        "end_of_term",
                        "status_en_rem",
                        "rent_min",
                        "vw_percent",
                        "vq_percent",
                    ]
                ]
                .rename(
                    columns={
                        "ceco": "macro_ceco",
                        "profit_center": "macro_profit_center",
                        "end_of_term": "macro_end_of_term",
                        "status_en_rem": "macro_status_en_rem",
                        "rent_min": "macro_rent_min",
                        "vw_percent": "macro_vw_percent",
                        "vq_percent": "macro_vq_percent",
                    }
                )
            )
            merged = merged.merge(macro_header, on=["vendor_key", "store_key"], how="left")
        else:
            merged["macro_ceco"] = None
            merged["macro_profit_center"] = None
            merged["macro_end_of_term"] = pd.NaT
            merged["macro_status_en_rem"] = None
            merged["macro_rent_min"] = np.nan
            merged["macro_vw_percent"] = np.nan
            merged["macro_vq_percent"] = np.nan

        merged["mapped_vendor"] = (
            merged["control_vendor_code_store"]
            .combine_first(merged["control_vendor_code_vendor"])
            .combine_first(merged["vendor"])
            .map(self._clean_numeric_code)
        )
        merged["mapped_store"] = merged[["store", "control_store_store", "control_store_vendor"]].bfill(axis=1).iloc[:, 0]
        merged["mapped_ceco"] = merged[
            ["macro_ceco", "control_ceco_store", "control_ceco_vendor", "ceco", "prorateo_ceco_store", "prorateo_ceco_vendor"]
        ].bfill(axis=1).iloc[:, 0]
        merged["mapped_ceco"] = merged["mapped_ceco"].where(merged["mapped_ceco"].map(self._is_valid_ceco_value))
        merged["mapped_ceco"] = merged["mapped_ceco"].combine_first(merged["macro_ceco"])
        merged["mapped_ceco_key"] = merged["mapped_ceco"].map(self._ceco_key)

        for key, mapped_key, source in [
            ("store_key", "prorateo_store_key", "vendor_store"),
            ("ceco_key", "mapped_ceco_key", "vendor_ceco"),
        ]:
            specific_prorateo = (
                prorateo[prorateo["vendor_key"].notna() & prorateo[key].notna()]
                .drop_duplicates(subset=["vendor_key", key])[["vendor_key", key, "vw_percent", "vq_percent"]]
                .rename(columns={
                    key: mapped_key,
                    "vw_percent": f"vw_percent_{source}",
                    "vq_percent": f"vq_percent_{source}",
                })
            )
            merged = merged.merge(specific_prorateo, on=["vendor_key", mapped_key], how="left")
        merged = merged.merge(
            prorateo_ceco[prorateo_ceco["ceco_key"].notna()].rename(columns={"ceco_key": "mapped_ceco_key"}),
            on="mapped_ceco_key", how="left",
        )

        contract_context = {}
        for ceco_key, rows in merged.groupby("mapped_ceco_key", dropna=True):
            contract_context[str(ceco_key)] = {
                "stores": list(dict.fromkeys(rows["mapped_store"].dropna().map(str).tolist())),
                "invoice_ids": list(dict.fromkeys(rows["invoice_id"].dropna().map(str).tolist())),
            }

        contracts = self._resolve_contract_duplicates(
            contracts,
            needed_ceco_keys=set(merged["mapped_ceco_key"].dropna()),
            selections=contract_selections or {},
            context_by_ceco=contract_context,
        )

        merged = merged.merge(
            contracts[["ceco_key", "end_of_term", "rent_min", "sell_media", "sell_cam", "status_en_rem"]],
            left_on="mapped_ceco_key",
            right_on="ceco_key",
            how="left",
        )
        merged["end_of_term"] = merged["end_of_term"].combine_first(merged["macro_end_of_term"])
        merged["rent_min"] = merged["rent_min"].combine_first(merged["macro_rent_min"])
        merged["status_en_rem"] = merged["status_en_rem"].combine_first(merged["macro_status_en_rem"])
        merged["vw_percent"] = np.nan
        merged["vq_percent"] = np.nan
        merged["prorateo_source"] = None
        for source, vw_column, vq_column in [
            ("vendor_store", "vw_percent_vendor_store", "vq_percent_vendor_store"),
            ("vendor_ceco", "vw_percent_vendor_ceco", "vq_percent_vendor_ceco"),
            ("store", "vw_percent_store", "vq_percent_store"),
            ("ceco", "vw_percent_ceco", "vq_percent_ceco"),
            ("vendor", "vw_percent_vendor", "vq_percent_vendor"),
            ("macro", "macro_vw_percent", "macro_vq_percent"),
        ]:
            use_source = merged["vw_percent"].isna() & merged[vw_column].notna()
            merged.loc[use_source, "vw_percent"] = merged.loc[use_source, vw_column]
            merged.loc[use_source, "vq_percent"] = merged.loc[use_source, vq_column]
            merged.loc[use_source, "prorateo_source"] = source
        merged["vq_percent"] = merged["vq_percent"].combine_first((1 - merged["vw_percent"]).clip(lower=0, upper=1))
        merged["profit_center"] = merged["macro_profit_center"].combine_first(merged["mapped_ceco"])
        self._log(f"Merging completed. Matched variables for {len(merged)} invoice rows:")
        for idx, row in merged.iterrows():
            self._log(
                f"  Invoice [{row['invoice_id']}]: Vendor={row['vendor']} -> "
                f"Store={row['store']} (Mapped Store={row['mapped_store']}), "
                f"CECO={row['ceco']} (Mapped CECO={row['mapped_ceco']}), "
                f"Contract End={row['end_of_term']} (Rent Min={row['rent_min']}, "
                f"Sell Media={row['sell_media']}, Sell CAM={row['sell_cam']}, "
                f"Status={row['status_en_rem']}), "
                f"Prorateo VW%={row['vw_percent']}, VQ%={row['vq_percent']} "
                f"(Source={row['prorateo_source']}), Profit Center={row['profit_center']}"
            )
        return merged

    def apply_rules(
        self,
        merged: pd.DataFrame,
        period: str | None = None,
        discount_selections: dict[str, str] | None = None,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        self._log("Applying business rules and creating output rows.")
        df = merged.copy()
        df["contract_active"] = df["status_en_rem"].map(self._is_active_contract_status)
        df["contract_status_blocked"] = df["status_en_rem"].map(self._is_blocked_contract_status)
        df["rent_type"] = np.where(df["contract_active"], "RF", "RV")
        df["account"] = np.where(df["contract_active"], self.FIXED_ACCOUNT, self.VARIABLE_ACCOUNT)
        office_mask = df["store"].map(self._is_office_store)
        if office_mask.any():
            office_rent_account = self._office_account_for_concept("RF")
            if not office_rent_account:
                raise PipelineError(
                    "Falta la cuenta de arriendos para Oficina en Diccionario_Conceptos_Simple.xlsx "
                    "(concepto RF OFICINA)."
                )
            df.loc[office_mask, "account"] = office_rent_account
        df["gross_amount"] = pd.to_numeric(df["amount"], errors="coerce").map(
            lambda value: round_cop(value, 0)
        )
        registration_date = pd.Timestamp.now().normalize()
        selections = discount_selections or {}
        discount_evaluations = df.apply(
            lambda row: self._evaluate_invoice_discounts(
                row.get("invoice_discounts"),
                row.get("invoice_date"),
                registration_date,
                selections.get(self._invoice_key(row.get("invoice_id"))),
            ),
            axis=1,
        )
        df["detected_invoice_discounts"] = df["invoice_discounts"]
        df["invoice_discounts"] = discount_evaluations.map(lambda result: result["applied_discounts"])
        df["discount_decisions"] = discount_evaluations.map(lambda result: result["decisions"])
        df["estimated_payment_date"] = discount_evaluations.map(lambda result: result["estimated_payment_date"])
        review_keys = set()
        for (_, row), evaluation in zip(df.iterrows(), discount_evaluations):
            invoice_id = row.get("invoice_id")
            for decision in evaluation["decisions"]:
                if decision.get("source") != "xml_note":
                    continue
                if decision.get("decision_reason") == "manual_review_required":
                    review_key = self._invoice_key(invoice_id)
                    if review_key not in review_keys:
                        self.discount_review_items.append(
                            self._discount_review_option(row, decision, review_key)
                        )
                        review_keys.add(review_key)
                    continue
                self._record_warning(self._discount_decision_message(invoice_id, decision))
        df["discount_total"] = df["invoice_discounts"].map(self._invoice_discount_total)
        df["payable_rounding"] = pd.to_numeric(df["payable_rounding"], errors="coerce").map(
            lambda value: round_cop(value, 0)
        )
        df["posting_discount_total"] = np.where(
            df["discount_total"] > 0,
            (df["discount_total"] - df["payable_rounding"]).clip(lower=0),
            0,
        )
        df["posting_discount_total"] = pd.to_numeric(df["posting_discount_total"], errors="coerce").map(
            lambda value: round_cop(value, 0)
        )
        df["amount"] = (df["gross_amount"] - df["posting_discount_total"]).map(
            lambda value: round_cop(value, 0)
        )
        df["vat"] = pd.to_numeric(df["vat"], errors="coerce").map(lambda value: round_cop(value, 0))
        df["vat_total"] = df["vat"]
        df["vat_vw"] = (df["vat"] * df["vw_percent"].fillna(0)).map(lambda value: round_cop(value, 0))
        df["vat_vq"] = (df["vat"] - df["vat_vw"]).map(lambda value: round_cop(value, 0))
        df["withholding_tax_amount"] = pd.to_numeric(df["withholding_tax"], errors="coerce").map(round_cop)
        df["invoice_total"] = (
            pd.to_numeric(df["amount"], errors="coerce").fillna(0) + df["vat_total"].fillna(0)
        ).map(lambda value: round_cop(value, 0))

        period_dates = self._resolve_period(df["invoice_date"], period)
        df["concept"] = [
            rent_type if self._is_office_store(store) else self._resolve_account_concept(account, rent_type)
            for account, rent_type, store in zip(df["account"], df["rent_type"], df["store"])
        ]
        df["text"] = [
            self._build_text(date_value, concept, store, ceco)
            for date_value, concept, store, ceco in zip(period_dates, df["concept"], df["mapped_store"], df["mapped_ceco"])
        ]

        self._log(f"Evaluated business rules for {len(df)} invoices:")
        for idx, row in df.iterrows():
            self._log(
                f"  Invoice [{row['invoice_id']}]: Contract Status={row['status_en_rem']} "
                f"-> ContractActive={row['contract_active']} -> Class={row['rent_type']} -> Account={row['account']}. "
                f"VAT={row['vat_total']} -> VW={row['vat_vw']} (using VW%={row['vw_percent']}), "
                f"VQ={row['vat_vq']}, VQ reversal={row['vat_vw']}. "
                f"Discount={row['discount_total']} (posting reduction={row['posting_discount_total']}), "
                f"Withholding={row['withholding_tax_amount']}, Text='{row['text']}'"
            )

        validation_df = self._build_validations(df)
        expanded = self._apply_distribution(df)
        output_df = expanded[
            [
                "invoice_id",
                "vendor",
                "store",
                "ceco",
                "profit_center",
                "account",
                "concept",
                "rent_type",
                "sell_media",
                "sell_cam",
                "amount",
                "gross_amount",
                "discount_total",
                "posting_discount_total",
                "payable_rounding",
                "invoice_total",
                "vat_total",
                "vat_vw",
                "vat_vq",
                "vw_percent",
                "vq_percent",
                "withholding_tax_amount",
                "text",
                "invoice_date",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
                "ubl_document_type",
                "invoice_line_items",
                "invoice_discounts",
                "detected_invoice_discounts",
                "discount_decisions",
                "estimated_payment_date",
                "support_source_file",
                "support_split_factor",
                "support_split_index",
                "contract_status_blocked",
                "support_flags",
            ]
        ].copy()
        return output_df, validation_df

    def generate_output(
        self,
        output_df: pd.DataFrame,
        validation_df: pd.DataFrame,
        macro_database_df: pd.DataFrame,
        macro_template_path: Path | None = None,
    ) -> PipelineResult:
        pending_concepts = self.build_concept_review_items(output_df)
        if pending_concepts:
            bad_invoice_ids = set()
            for _, row in output_df.iterrows():
                for item in self._invoice_line_items(row.get("invoice_line_items")):
                    description = self._clean_text(item.get("description"))
                    amount = pd.to_numeric(pd.Series([item.get("amount")]), errors="coerce").iloc[0]
                    if not description or pd.isna(amount) or abs(float(amount)) == 0:
                        continue
                    if not self._match_learning_dictionary(description, row):
                        bad_invoice_ids.add(row.get("invoice_id"))

            if bad_invoice_ids:
                preview = ", ".join(item["description"] for item in pending_concepts[:5])
                extra = "" if len(pending_concepts) <= 5 else f" y {len(pending_concepts) - 5} mas"
                self._record_warning(
                    f"Se omitieron {len(bad_invoice_ids)} facturas porque contienen conceptos no aprendidos "
                    f"en el diccionario. Asigna sus conceptos para poder procesarlas: {preview}{extra}"
                )
                for inv_id in bad_invoice_ids:
                    self.preflight_processing_issues.append({
                        "invoice_id": str(inv_id),
                        "source_file": "Diccionario",
                        "store": None,
                        "location": "Diccionario",
                        "problem": "La factura contiene items sin concepto aprendido en Diccionario_Conceptos_Simple.xlsx.",
                        "possible_solution": "Agrega las frases correspondientes al diccionario de conceptos.",
                        "issue_type": "unmapped_concept",
                    })
                output_df = output_df[~output_df["invoice_id"].isin(bad_invoice_ids)].copy()

        if not validation_df.empty and "invoice_id" in validation_df.columns:
            invalid_invoice_ids = set(validation_df["invoice_id"].dropna())
            if invalid_invoice_ids:
                self._log(f"Skipping CSV generation for invoices with validation errors: {', '.join(map(str, invalid_invoice_ids))}")
                output_df = output_df[~output_df["invoice_id"].isin(invalid_invoice_ids)].copy()

        if output_df.empty:
            raise PipelineError("Todas las facturas del lote fallaron la validacion o no tienen conceptos. Corrige los errores antes de generar los archivos.")

        self._log("Writing per-invoice Allocated Costs CSV ZIP and individual invoice ZIP bundles.")
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        allocated_bundle_name = f"allocated_costs_csv_bundle_{timestamp}.zip"
        invoice_bundle_name = f"output_invoice_csv_bundle_{timestamp}.zip"
        manual_style_bundle_name = f"output_manual_style_csv_bundle_{timestamp}.zip"

        summary_df, headers_df, lines_df = self._build_lucy_views(output_df)
        blocked_invoice_ids = set(
            summary_df.loc[summary_df.get("contract_status_blocked", False).fillna(False), "invoice_id"]
            if "contract_status_blocked" in summary_df.columns
            else []
        )
        if blocked_invoice_ids:
            self._log(
                "Skipping CSV generation for invoices with REEMPLAZO contract status: "
                + ", ".join(sorted(str(invoice_id) for invoice_id in blocked_invoice_ids))
            )
            summary_df = summary_df[~summary_df["invoice_id"].isin(blocked_invoice_ids)].copy()
            headers_df = headers_df[~headers_df["invoice_id"].isin(blocked_invoice_ids)].copy()
            lines_df = lines_df[~lines_df["invoice_id"].isin(blocked_invoice_ids)].copy()
        self._log(f"Writing {len(summary_df)} summary records to Lucy Excel/CSV:")
        for idx, row in summary_df.iterrows():
            self._log(
                f"  Record [{row.get('invoice_id')}]: Vendor={row.get('vendor')}, "
                f"CECO={row.get('ceco')}, Account={row.get('account')}, "
                f"Amount={row.get('amount')}, VAT={row.get('vat_total')} (VW={row.get('vat_vw')}, VQ={row.get('vat_vq')}), "
                f"Text='{row.get('text')}'"
            )
        header_export_df = self._build_header_export(summary_df)
        lucy_export_df = self._build_macro_lucy_export(
            summary_df,
            lines_df,
            macro_database_df=macro_database_df,
            macro_template_path=macro_template_path,
        )
        self._write_allocated_costs_csv_bundle(
            summary_df=summary_df,
            lucy_export_df=lucy_export_df,
            bundle_name=allocated_bundle_name,
            timestamp=timestamp,
        )
        self._write_invoice_csv_bundle(
            summary_df=summary_df,
            header_export_df=header_export_df,
            lucy_export_df=lucy_export_df,
            bundle_name=invoice_bundle_name,
            timestamp=timestamp,
        )
        self._write_manual_style_csv_bundle(
            summary_df=summary_df,
            header_export_df=header_export_df,
            lucy_export_df=lucy_export_df,
            bundle_name=manual_style_bundle_name,
            timestamp=timestamp,
        )
        # Write execution logs to a physical log file inside self.output_dir
        log_file_path = self.output_dir / "pipeline.log"
        try:
            self.logs.append(f"Writing complete execution log to: {log_file_path.name}")
            with open(log_file_path, "w", encoding="utf-8") as f:
                f.write("\n".join(self.logs))
        except Exception as exc:
            LOGGER.error(f"Failed to write pipeline log file: {exc}")

        run_folder_name = self.output_dir.name
        return PipelineResult(
            output_csv=f"{run_folder_name}/{allocated_bundle_name}",
            invoice_csv_bundle=f"{run_folder_name}/{invoice_bundle_name}",
            manual_style_csv_bundle=f"{run_folder_name}/{manual_style_bundle_name}",
            output_rows=len(summary_df),
            header_rows=len(header_export_df),
            validation_rows=len(validation_df),
            processing_issues=[
                *self.preflight_processing_issues,
                *self._build_processing_issues(validation_df),
            ],
            logs=self.logs.copy(),
            warnings=self.warnings.copy(),
        )

    def _load_invoice_supports(
        self,
        support_paths: list[Path],
        support_split_factors: dict[str, int] | None = None,
        support_split_weights: dict[str, list[str | float]] | None = None,
    ) -> pd.DataFrame:
        if not support_paths:
            return pd.DataFrame()
        self._log(f"Parsing {len(support_paths)} invoice support files (PDF/XML).")
        df = InvoiceSupportReader().parse_many(support_paths)
        split_factor_by_name = {
            Path(path).name: max(int(factor or 1), 1)
            for path, factor in (support_split_factors or {}).items()
        }
        split_weights_by_name = {
            Path(path).name: weights
            for path, weights in (support_split_weights or {}).items()
        }
        df["split_factor"] = df["source_file"].map(split_factor_by_name).fillna(1).astype(int)
        df["split_weights"] = df.apply(
            lambda row: self._valid_split_weights(
                split_weights_by_name.get(row["source_file"]),
                int(row["split_factor"]),
            ),
            axis=1,
        )
        for _, row in df.iterrows():
            split_note = ""
            if int(row.get("split_factor") or 1) > 1:
                percentages = "/".join(f"{float(value):g}%" for value in row["split_weights"])
                split_note = f", Split: {row['split_factor']} tiendas ({percentages})"
            self._log(
                f"  Support File parsed: {row['source_file']} ({row['source_type'].upper()}) -> "
                f"Invoice_ID: {row['invoice_id']}, Date: {row['invoice_date']}, "
                f"Subtotal: {row['subtotal']}, Total: {row['total']}, VAT: {row['detected_iva']}, "
                f"Withholding: {row['withholding_tax']}{split_note}, Flags: {row['flags']}"
            )
        return df

    @staticmethod
    def _valid_split_weights(weights, split_factor: int) -> list[float]:
        if not isinstance(weights, (list, tuple)) or len(weights) != split_factor:
            return [1.0] * split_factor
        parsed = pd.to_numeric(pd.Series(weights), errors="coerce")
        if parsed.isna().any() or (parsed <= 0).any() or float(parsed.sum()) <= 0:
            return [1.0] * split_factor
        return parsed.astype(float).tolist()

    def _apply_support_split_factors(self, supports: pd.DataFrame) -> pd.DataFrame:
        df = supports.copy()
        if "split_factor" not in df.columns:
            df["split_factor"] = 1
        df["split_factor"] = pd.to_numeric(df["split_factor"], errors="coerce").fillna(1).clip(lower=1).astype(int)
        return df

    def _split_invoice_line_items(self, line_items, split_factor: int) -> list[list[dict]]:
        return self._allocate_invoice_line_items(line_items, [1] * split_factor, split_factor - 1)

    def _allocate_invoice_line_items(self, line_items, weights, residual_index: int) -> list[list[dict]]:
        results = [[] for _ in weights]
        for item in self._invoice_line_items(line_items):
            allocations = {}
            for key in ["amount", "net_amount", "allowance_amount", "charge_amount", "tax_amount"]:
                value = pd.to_numeric(pd.Series([item.get(key)]), errors="coerce").iloc[0]
                if pd.notna(value):
                    allocations[key] = allocate_cop(value, weights, residual_index=residual_index)
            for split_index in range(len(weights)):
                scaled = dict(item)
                for key, values in allocations.items():
                    scaled[key] = values[split_index]
                results[split_index].append(scaled)
        return results

    def _split_invoice_discounts(self, discounts, split_factor: int) -> list[list[dict]]:
        return self._allocate_invoice_discounts(discounts, [1] * split_factor, split_factor - 1)

    def _allocate_invoice_discounts(self, discounts, weights, residual_index: int) -> list[list[dict]]:
        results = [[] for _ in weights]
        for discount in self._invoice_discounts(discounts):
            allocations = {}
            for key in ["amount", "base_amount"]:
                value = pd.to_numeric(pd.Series([discount.get(key)]), errors="coerce").iloc[0]
                if pd.notna(value):
                    allocations[key] = allocate_cop(value, weights, residual_index=residual_index)
            for split_index in range(len(weights)):
                scaled = dict(discount)
                for key, values in allocations.items():
                    scaled[key] = values[split_index]
                results[split_index].append(scaled)
        return results

    def _expand_split_invoice_rows(self, invoices: pd.DataFrame) -> pd.DataFrame:
        if "support_split_factor" not in invoices.columns:
            invoices["support_split_factor"] = 1
        invoices["support_split_factor"] = pd.to_numeric(invoices["support_split_factor"], errors="coerce").fillna(1).clip(lower=1)
        expanded_rows = []
        for _, row in invoices.iterrows():
            split_factor = int(row.get("support_split_factor") or 1)
            split_weights = self._valid_split_weights(
                row.get("support_split_weights"),
                split_factor,
            )
            residual_index = split_factor - 1
            monetary_allocations = {}
            for col in [
                "total",
                "vat",
                "withholding_tax",
                "payable_rounding",
                "amount",
                "support_subtotal",
                "support_total",
                "support_detected_iva",
                "support_withholding_tax",
                "support_payable_rounding",
            ]:
                if col not in row:
                    continue
                value = pd.to_numeric(pd.Series([row.get(col)]), errors="coerce").iloc[0]
                if pd.notna(value):
                    monetary_allocations[col] = allocate_cop(
                        value,
                        split_weights,
                        residual_index=residual_index,
                    )
            split_items = self._allocate_invoice_line_items(
                row.get("invoice_line_items"),
                split_weights,
                residual_index,
            )
            split_discounts = self._allocate_invoice_discounts(
                row.get("invoice_discounts"),
                split_weights,
                residual_index,
            )
            for split_index in range(1, split_factor + 1):
                new_row = row.copy()
                new_row["support_split_index"] = split_index
                new_row["support_split_percent"] = (
                    float(split_weights[split_index - 1]) / sum(split_weights) * 100
                )
                for col, values in monetary_allocations.items():
                    new_row[col] = values[split_index - 1]
                new_row["invoice_line_items"] = split_items[split_index - 1]
                new_row["invoice_discounts"] = split_discounts[split_index - 1]
                expanded_rows.append(new_row)
            if split_factor > 1:
                percentages = "/".join(f"{value:g}%" for value in split_weights)
                self._log(
                    f"  Invoice [{row.get('invoice_id')}]: split into {split_factor} files "
                    f"using {percentages}."
                )
        return pd.DataFrame(expanded_rows).reset_index(drop=True)

    def _load_invoices(self, path: Path) -> pd.DataFrame:
        if path.suffix.lower() == ".xml":
            df = self._load_invoices_from_xml(path)
            df["invoice_input_type"] = "xml"
            return df

        df = self._safe_read_table(path)
        original_columns = [self._clean_text(col) or str(col) for col in df.columns]
        df = df.rename(
            columns=self._best_effort_rename(
                df.columns,
                {
                    "invoice_id": [
                        "invoice_id",
                        "invoice",
                        "factura",
                        "numero_factura",
                        "numero_de_factura",
                        "numero_de_documento",
                        "documento",
                    ],
                    "vendor": [
                        "vendor",
                        "sap_vendor_code",
                        "vendor_code",
                        "acreedor",
                        "acreedor_administracion",
                    ],
                    "supplier_nit": ["supplier_nit", "nit", "tax_id", "identificacion_fiscal"],
                    "store": ["store", "tienda", "local"],
                    "reference": ["reference", "referencia", "centro"],
                    "total": [
                        "total",
                        "amount",
                        "valor",
                        "valor_total",
                        "valor_factura",
                        "valor_de_factura",
                        "valor_facturado",
                        "valor_facturado_total",
                        "canon",
                    ],
                    "vat": ["vat", "iva", "valor_iva", "tax"],
                    "invoice_date": ["invoice_date", "fecha_factura", "fecha", "date"],
                    "ceco": ["ceco", "cost_center", "cc", "c_c"],
                },
            )
        )
        required = {"invoice_id", "vendor"}
        missing = required - set(df.columns)
        if missing:
            raise PipelineError(
                "Invoice file is missing required columns. "
                f"Expected at least {sorted(required)} and could not find {sorted(missing)}. "
                f"Detected columns: {original_columns}."
            )
        df["invoice_input_type"] = "table"
        return df

    def _load_invoices_from_xml(self, path: Path) -> pd.DataFrame:
        parsed = InvoiceSupportReader().parse_file(path)
        if parsed.source_type != "xml":
            raise PipelineError("Only XML files can be used directly as invoice input.")

        store = self._extract_store_from_filename(path)
        data = {
            "invoice_id": parsed.invoice_id,
            "vendor": parsed.supplier_id,
            "supplier_nit": parsed.supplier_id,
            "ubl_document_type": parsed.document_type,
            "vendor_name": parsed.supplier_name,
            "store": store,
            "reference": parsed.invoice_id,
            "total": self._calculated_invoice_total(parsed.subtotal, parsed.detected_iva, parsed.total),
            "vat": parsed.detected_iva,
            "invoice_date": parsed.invoice_date,
            "due_date": parsed.due_date,
            "withholding_tax": parsed.withholding_tax,
            "payment_terms_text": parsed.payment_terms_text,
            "item_count_hint": parsed.item_count_hint,
            "invoice_line_items": parsed.line_items,
            "invoice_discounts": parsed.discounts,
            "payable_rounding": parsed.payable_rounding,
        }
        return pd.DataFrame([data])

    def _load_control(self, path: Path) -> pd.DataFrame:
        df = pd.read_excel(path, sheet_name="CONTROL ADM", header=4)
        df = df.rename(
            columns=self._best_effort_rename(
                df.columns,
                {
                    "store": ["tienda"],
                    "store_name": ["centro"],
                    "ceco": ["c_c", "cc", "ceco"],
                    "vendor_code": ["acreedor_arriendo"],
                    "vendor": ["proveedor_arriendo"],
                },
            )
        )
        if "store" not in df.columns and "store_name" in df.columns:
            df["store"] = df["store_name"]
        keep = [col for col in ["store", "store_name", "ceco", "vendor_code", "vendor"] if col in df.columns]
        result = df[keep].copy()
        if "vendor_code" not in result.columns or "ceco" not in result.columns:
            raise PipelineError("CONTROL ARRI ADMON file could not be parsed into store/vendor/CECO fields.")
        result["vendor_code"] = result["vendor_code"].map(self._extract_primary_code)
        return result.dropna(how="all")

    def _load_contracts(self, path: Path) -> pd.DataFrame:
        xl = pd.ExcelFile(path)
        frames = []
        for sheet_name in xl.sheet_names:
            if sheet_name.strip().upper() in {"CONDICIONES", "CONTABILIZACION", "CUENTAS"}:
                continue
            preview = xl.parse(sheet_name, header=None, nrows=30)
            header_row = self._find_contract_header_row(preview)
            if header_row is None:
                if self._concept_key(sheet_name) == "COLOMBIA":
                    raise PipelineError(
                        "No se pudo localizar la fila de encabezados en la hoja COLOMBIA de "
                        "Contratos_con_condiciones. Revisa que existan las columnas CeCo, End of Term y Status en REM."
                    )
                self._log(f"Skipped contracts sheet '{sheet_name}': header row was not recognized.")
                continue
            frame = xl.parse(sheet_name, header=header_row)
            frame = frame.rename(
                columns=self._best_effort_rename(
                    frame.columns,
                    {
                        "ceco": ["ceco"],
                        "contract_id": ["contract", "contrato"],
                        "contract_name": ["contract_name", "nombre_contrato"],
                        "end_of_term": ["end_of_term", "end_of_term_en_virtual_contract"],
                        "rent_min": ["rent_min_rent", "rent_min"],
                        "sell_media": ["sell_media"],
                        "sell_cam": ["sell_cam"],
                        "status_en_rem": ["status_en_rem"],
                    },
                )
            )
            if {"ceco", "end_of_term"} - set(frame.columns):
                if self._concept_key(sheet_name) == "COLOMBIA":
                    raise PipelineError(
                        "La hoja COLOMBIA de Contratos_con_condiciones no contiene las columnas requeridas "
                        "CeCo y End of Term."
                    )
                continue
            frame["country_sheet"] = sheet_name
            frame["source_row"] = frame.index + header_row + 2
            keep = [
                "ceco",
                "contract_id",
                "contract_name",
                "end_of_term",
                "rent_min",
                "sell_media",
                "sell_cam",
                "status_en_rem",
                "country_sheet",
                "source_row",
            ]
            frame = frame[[col for col in keep if col in frame.columns]].copy()
            for column in keep:
                if column not in frame.columns:
                    frame[column] = None
            frame["contract_key"] = frame.apply(self._contract_row_key, axis=1)
            frames.append(frame[keep + ["contract_key"]])
        if not frames:
            raise PipelineError("Contracts workbook could not be parsed.")
        return pd.concat(frames, ignore_index=True)

    def _find_contract_header_row(self, preview: pd.DataFrame) -> int | None:
        best_row = None
        best_score = 0
        for row_index, row in preview.iterrows():
            columns = {self._normalize_column_name(value) for value in row.tolist() if self._clean_text(value)}
            has_ceco = "ceco" in columns
            has_end = bool(columns & {"end_of_term", "end_of_term_en_virtual_contract"})
            has_status = bool(columns & {"status_en_rem", "status_en_virtual_contract"})
            has_contract = bool(columns & {"contract", "contrato"})
            score = sum([has_ceco, has_end, has_status, has_contract])
            if has_ceco and has_end and score > best_score:
                best_row = int(row_index)
                best_score = score
        return best_row if best_score >= 3 else None

    def _contract_row_key(self, row) -> str:
        parts = [
            self._clean_text(row.get("country_sheet")) or "SHEET",
            self._normalize_invoice_id(row.get("contract_id")) or "SIN_CONTRATO",
            self._clean_text(row.get("source_row")) or "SIN_FILA",
        ]
        return "|".join(self._concept_key(part) for part in parts)

    def _resolve_contract_duplicates(
        self,
        contracts: pd.DataFrame,
        needed_ceco_keys: set[str],
        selections: dict[str, str],
        context_by_ceco: dict[str, dict] | None = None,
    ) -> pd.DataFrame:
        if contracts.empty:
            return contracts

        frame = contracts.copy()
        colombia_mask = frame["country_sheet"].map(self._concept_key).eq("COLOMBIA")
        if colombia_mask.any():
            frame = frame[colombia_mask].copy()
        frame = frame[frame["ceco_key"].isin(needed_ceco_keys)].copy()
        if frame.empty:
            return contracts.iloc[0:0].copy()

        frame["end_of_term"] = pd.to_datetime(frame["end_of_term"], errors="coerce")
        resolved_rows = []
        context_by_ceco = context_by_ceco or {}

        for ceco_key, group in frame.groupby("ceco_key", sort=False):
            group = group.sort_values("end_of_term", ascending=False, na_position="last").copy()
            if len(group) == 1:
                resolved_rows.append(group.iloc[0])
                continue

            valid_group = group[~group["status_en_rem"].map(self._is_blocked_contract_status)].copy()
            if not valid_group.empty and len(valid_group) < len(group):
                ignored = len(group) - len(valid_group)
                self._record_warning(
                    f"CECO {ceco_key}: se ignoraron {ignored} contrato(s) con estado REEMPLAZO porque existe "
                    "otro contrato utilizable para el mismo centro de costo."
                )
                group = valid_group

            if len(group) == 1:
                resolved_rows.append(group.iloc[0])
                continue

            selected_key = self._clean_text(selections.get(str(ceco_key)))
            selected = group[group["contract_key"].map(self._clean_text).eq(selected_key)] if selected_key else group.iloc[0:0]
            if not selected.empty:
                chosen = selected.iloc[0]
                resolved_rows.append(chosen)
                self._record_warning(
                    f"CECO {ceco_key}: se utilizo manualmente el contrato "
                    f"{self._normalize_invoice_id(chosen.get('contract_id')) or 'SIN_NUMERO'} con estado "
                    f"{self._clean_text(chosen.get('status_en_rem')) or 'SIN_ESTADO'}."
                )
                continue

            status_groups = {self._contract_status_group(value) for value in group["status_en_rem"]}
            if len(status_groups) == 1 and "REVISION" not in status_groups:
                chosen = group.iloc[0]
                resolved_rows.append(chosen)
                statuses = ", ".join(
                    dict.fromkeys(
                        self._clean_text(value) or "SIN_ESTADO"
                        for value in group["status_en_rem"].tolist()
                    )
                )
                classification = next(iter(status_groups))
                self._record_warning(
                    f"CECO {ceco_key}: se encontraron {len(group)} contratos duplicados con estados equivalentes "
                    f"({statuses}). El proceso continuo como {classification} usando el contrato "
                    f"{self._normalize_invoice_id(chosen.get('contract_id')) or 'SIN_NUMERO'}."
                )
                continue

            context = context_by_ceco.get(str(ceco_key), {})
            options = [self._contract_review_option(row) for _, row in group.iterrows()]
            self.contract_review_items.append(
                {
                    "ceco": str(ceco_key),
                    "stores": context.get("stores", []),
                    "invoice_ids": context.get("invoice_ids", []),
                    "options": options,
                }
            )
            resolved_rows.append(group.iloc[0])

        if not resolved_rows:
            return frame.iloc[0:0].copy()
        return pd.DataFrame(resolved_rows).reset_index(drop=True)

    def _contract_status_group(self, value) -> str:
        status = self._concept_key(value)
        if status in {"VIGENTE", "POR VENCER"}:
            return "RF"
        if status.startswith("CESADO") or status.startswith("CERRADA"):
            return "RV"
        if status.startswith("REEMPLAZO") or status.startswith("REEMPLAZADO"):
            return "REEMPLAZO"
        return "REVISION"

    def _contract_review_option(self, row) -> dict:
        end_of_term = pd.to_datetime(row.get("end_of_term"), errors="coerce")
        return {
            "contract_key": self._clean_text(row.get("contract_key")),
            "contract_id": self._normalize_invoice_id(row.get("contract_id")) or "SIN_NUMERO",
            "contract_name": self._clean_text(row.get("contract_name")) or "Sin nombre de contrato",
            "status": self._clean_text(row.get("status_en_rem")) or "SIN_ESTADO",
            "classification": self._contract_status_group(row.get("status_en_rem")),
            "end_of_term": end_of_term.strftime("%d-%m-%Y") if pd.notna(end_of_term) else "Sin fecha",
            "sheet": self._clean_text(row.get("country_sheet")) or "Sin hoja",
            "source_row": self._clean_text(row.get("source_row")) or "Sin fila",
        }

    def _load_prorateo(self, path: Path) -> pd.DataFrame:
        df = pd.read_excel(path, sheet_name="CONSOLIDADO ARRIENDOS PROY IVA", header=0)
        df = df.rename(
            columns=self._best_effort_rename(
                df.columns,
                {
                    "vendor_code": ["acreedor_arriendo"],
                    "vendor": ["proveedor_arriendo"],
                    "store": ["tienda"],
                    "ceco": ["ceco"],
                    "vw_percent": ["porcentaje_iva_vw", "vw", "iva_deducible_vw"],
                    "vq_percent": ["porcentaje_iva_vq", "vq", "iva_vq"],
                },
            )
        )
        if "vendor_code" in df.columns:
            df["vendor_code"] = df["vendor_code"].map(self._clean_numeric_code)
        result = df[[col for col in ["vendor_code", "vendor", "store", "ceco", "vw_percent", "vq_percent"] if col in df.columns]].copy()
        if "vendor" not in result.columns:
            result["vendor"] = None
        if "vq_percent" not in result.columns:
            result["vq_percent"] = None
        return result

    def _load_learning_dictionary(self, path: Path | None = None) -> pd.DataFrame:
        target = Path(path) if path else self.learning_dictionary_path
        columns = ["concept", "account", "phrases", "text"]
        if not target.exists():
            return pd.DataFrame(columns=columns)
        df = pd.read_excel(target, sheet_name="Diccionario", dtype=str).fillna("")
        df = df.rename(
            columns=self._best_effort_rename(
                df.columns,
                {
                    "concept": ["concepto"],
                    "account": ["gl_account", "gl account", "cuenta", "cuenta_contable"],
                    "phrases": ["frases_palabras", "frases/palabras", "frase_palabra", "palabras", "frases"],
                    "text": ["texto_corto_lucy", "texto corto lucy", "texto", "text", "sigla"],
                },
            )
        )
        if "concept" not in df.columns or "phrases" not in df.columns:
            return pd.DataFrame(columns=columns)
        if "account" not in df.columns:
            df["account"] = ""
        if "text" not in df.columns:
            df["text"] = df["concept"]
        df["concept"] = df["concept"].map(self._clean_text)
        df["account"] = df["account"].map(self._clean_text)
        df["phrases"] = df["phrases"].map(self._clean_text)
        df["text"] = df["text"].map(self._clean_text)
        return df[[col for col in columns if col in df.columns]].dropna(how="all")

    def _load_distribution(self, path: Path) -> pd.DataFrame:
        df = pd.read_excel(path, sheet_name="Distrubución", header=1)
        records = []
        current_source = None
        for _, row in df.iterrows():
            contract_name = self._clean_text(row.iloc[0]) if len(row) > 0 else None
            distribution_label = self._clean_text(row.iloc[2]) if len(row) > 2 else None
            target_vendor = row.iloc[3] if len(row) > 3 else None
            source_vendor = row.iloc[5] if len(row) > 5 else None

            if contract_name:
                current_source = source_vendor

            percent = self._extract_percent(distribution_label)
            if percent is None or pd.isna(target_vendor):
                continue

            records.append(
                {
                    "source_vendor": self._clean_numeric_code(current_source),
                    "target_vendor": self._clean_numeric_code(target_vendor),
                    "split_percent": percent,
                    "distribution_label": distribution_label,
                    "contract_name": contract_name,
                }
            )

        if not records:
            self._log("Distribution workbook did not produce split rules; invoices will remain unsplit.")
            return pd.DataFrame(columns=["source_vendor", "target_vendor", "split_percent"])
        return pd.DataFrame(records)

    def _load_macro_database(self, path: Path | None) -> pd.DataFrame:
        if path is None:
            return pd.DataFrame()
        try:
            todos = pd.read_excel(path, sheet_name="TODOS", header=None, engine="openpyxl")
        except Exception as exc:
            self._log(f"Could not parse macro database workbook. Detail: {exc}")
            return pd.DataFrame()

        records = self._extract_macro_database_rows(todos)
        return pd.DataFrame(records)

    def _apply_distribution(self, df: pd.DataFrame) -> pd.DataFrame:
        if self.distribution_rules.empty:
            result = df.copy()
            result["vendor"] = result["mapped_vendor"]
            result["ceco"] = result["mapped_ceco"]
            return result

        expanded_rows = []
        for _, row in df.iterrows():
            matches = self.distribution_rules[
                self.distribution_rules["source_vendor_key"] == self._text_key(row["mapped_vendor"])
            ]
            if matches.empty:
                new_row = row.copy()
                new_row["vendor"] = row["mapped_vendor"]
                new_row["ceco"] = row["mapped_ceco"]
                expanded_rows.append(new_row)
                continue

            valid_matches = matches[matches["target_vendor"].notna()].copy()
            if valid_matches.empty:
                new_row = row.copy()
                new_row["vendor"] = row["mapped_vendor"]
                new_row["ceco"] = row["mapped_ceco"]
                expanded_rows.append(new_row)
                continue

            total_pct = valid_matches["split_percent"].sum()
            if total_pct <= 0:
                total_pct = 1.0
            weights = [float(value) for value in valid_matches["split_percent"].fillna(0).tolist()]
            residual_index = max(range(len(weights)), key=weights.__getitem__)
            monetary_columns = [
                "amount",
                "gross_amount",
                "discount_total",
                "posting_discount_total",
                "payable_rounding",
                "invoice_total",
                "vat_total",
                "vat_vw",
                "vat_vq",
                "withholding_tax_amount",
            ]
            allocations = {
                col: allocate_cop(row.get(col) or 0, weights, residual_index=residual_index)
                for col in monetary_columns
            }
            allocated_items = self._allocate_invoice_line_items(
                row.get("invoice_line_items"),
                weights,
                residual_index,
            )
            allocated_discounts = self._allocate_invoice_discounts(
                row.get("invoice_discounts"),
                weights,
                residual_index,
            )

            self._log(
                f"  Invoice [{row['invoice_id']}]: Vendor {row['mapped_vendor']} matched distribution rule. "
                f"Splitting original amount {row['amount']} into {len(valid_matches)} targets:"
            )

            for split_index, (_, split) in enumerate(valid_matches.iterrows()):
                new_row = row.copy()
                new_row["vendor"] = split["target_vendor"]
                new_row["ceco"] = row["mapped_ceco"]
                for col, values in allocations.items():
                    new_row[col] = values[split_index]
                new_row["invoice_line_items"] = allocated_items[split_index]
                new_row["invoice_discounts"] = allocated_discounts[split_index]
                expanded_rows.append(new_row)
                self._log(
                    f"    -> Split Target Vendor: {new_row['vendor']} | Share: {split['split_percent']*100}% | "
                    f"Allocated Amount: {new_row['amount']} | Allocated VAT: {new_row['vat_total']}"
                )

        return pd.DataFrame(expanded_rows)

    def _build_validations(self, df: pd.DataFrame) -> pd.DataFrame:
        validation_rows = []
        for _, row in df.iterrows():
            missing_ceco = pd.isna(row.get("mapped_ceco"))
            missing_contract = pd.isna(row.get("end_of_term"))
            missing_prorrateo = pd.isna(row.get("vw_percent"))
            amount_value = pd.to_numeric(pd.Series([row.get("amount")]), errors="coerce").iloc[0]
            missing_amount = pd.isna(row.get("total")) and (pd.isna(amount_value) or float(amount_value) == 0)
            blocked_replacement = self._is_blocked_contract_status(row.get("status_en_rem"))
            unresolved_discount_note = any(
                decision.get("decision_reason") == "manual_review_required"
                for decision in self._invoice_discounts(row.get("discount_decisions"))
            )
            discount_total = self._invoice_discount_total(row.get("invoice_discounts"))
            gross_amount = pd.to_numeric(pd.Series([row.get("gross_amount")]), errors="coerce").iloc[0]
            invalid_discount = pd.notna(gross_amount) and discount_total > float(gross_amount) + 0.01
            if (
                missing_ceco
                or missing_contract
                or missing_prorrateo
                or missing_amount
                or blocked_replacement
                or unresolved_discount_note
                or invalid_discount
            ):
                validation_rows.append(
                    {
                        "invoice_id": row.get("invoice_id"),
                        "source_file": row.get("support_source_file"),
                        "vendor": row.get("mapped_vendor"),
                        "store": row.get("mapped_store"),
                        "ceco": row.get("mapped_ceco"),
                        "contract_status": row.get("status_en_rem"),
                        "missing_ceco": bool(missing_ceco),
                        "missing_contract": bool(missing_contract),
                        "missing_prorrateo": bool(missing_prorrateo),
                        "missing_amount": bool(missing_amount),
                        "blocked_replacement_status": bool(blocked_replacement),
                        "unresolved_discount_note": bool(unresolved_discount_note),
                        "invalid_discount": bool(invalid_discount),
                        "message": (
                            f"CECO {row.get('mapped_ceco')} tiene status '{row.get('status_en_rem')}' en Contratos_con_condiciones; "
                            "no se genera CSV para esta factura. Revisar manualmente el contrato/reemplazo antes de procesar."
                            if blocked_replacement
                            else (
                                "El XML menciona un descuento o incentivo que no pudo calcularse automaticamente; "
                                "revisar la nota antes de procesar."
                                if unresolved_discount_note
                                else (
                                    "El descuento detectado supera el valor bruto de la factura; revisar importes."
                                    if invalid_discount
                                    else None
                                )
                            )
                        ),
                    }
                )

        return pd.DataFrame(validation_rows)

    def _build_processing_issues(self, validation_df: pd.DataFrame) -> list[dict]:
        issues = []
        if validation_df.empty:
            return issues

        for _, row in validation_df.iterrows():
            descriptions = []
            solutions = []
            ceco = self._clean_text(row.get("ceco"))
            if bool(row.get("missing_ceco")):
                descriptions.append("No se encontro un CeCo valido")
                solutions.append("Completa o corrige la tienda y el CeCo en el invoices file y en la tabla de control")
            if bool(row.get("missing_contract")):
                descriptions.append(
                    f"No se encontro contrato para el CeCo {ceco}"
                    if ceco
                    else "No se encontro contrato"
                )
                solutions.append("Agrega o corrige el CeCo y su estado en Contratos_con_condiciones")
            if bool(row.get("missing_prorrateo")):
                descriptions.append("No se encontro prorrateo VW/VQ")
                solutions.append("Agrega el Vendor, tienda o CeCo y sus porcentajes VQ/VW en el archivo PRORATEO")
            if bool(row.get("missing_amount")):
                descriptions.append("No se pudo determinar el importe de la factura")
                solutions.append("Revisa que el XML incluya items, subtotal, cargos y descuentos con importes validos")
            if bool(row.get("blocked_replacement_status")):
                descriptions.append("El contrato tiene estado REEMPLAZO")
                solutions.append("Revisa el contrato y actualiza su estado, o procesa la factura manualmente")
            if bool(row.get("unresolved_discount_note")):
                descriptions.append("El descuento requiere una decision manual")
                solutions.append("Revisa la condicion del descuento y decide si debe aplicarse antes de reprocesar")
            if bool(row.get("invalid_discount")):
                descriptions.append("El descuento supera el valor bruto de la factura")
                solutions.append("Verifica los importes del XML o procesa la factura manualmente")
            if not descriptions:
                descriptions.append(self._clean_text(row.get("message")) or "Requiere revision manual")
            if not solutions:
                solutions.append("Revisa la factura y sus datos en los archivos maestros antes de reprocesarla")

            source_file = self._clean_text(row.get("source_file"))
            store = self._clean_text(row.get("store"))
            issues.append(
                {
                    "invoice_id": self._clean_text(row.get("invoice_id")) or "Sin numero",
                    "source_file": source_file,
                    "store": store,
                    "location": " / ".join(value for value in [source_file, store] if value) or "Sin archivo o tienda",
                    "problem": "; ".join(descriptions),
                    "possible_solution": "; ".join(dict.fromkeys(solutions)),
                    "issue_type": "invoice_validation",
                }
            )
        return issues

    def _extract_macro_database_rows(self, todos_values: pd.DataFrame) -> list[dict]:
        records: list[dict] = []
        current: dict = {}

        for idx, series in todos_values.iterrows():
            row_idx = idx + 1
            if row_idx < 9:
                continue

            values = series.reindex(range(25)).tolist()
            category = self._clean_text(values[0])
            vendor = self._clean_numeric_code(values[1])
            account_or_name = self._clean_text(values[2])
            amount = values[3]
            tax_code = self._clean_text(values[4])
            posting_key = self._clean_text(values[5])
            text = self._clean_text(values[6])
            store = self._clean_text(values[6])
            group_ceco = self._clean_text(values[7])
            line_ceco = self._clean_text(values[8])

            if category in {"ADMINISTRACION", "ARRIENDOS"}:
                current = {
                    "category": category,
                    "vendor": vendor,
                    "vendor_name": account_or_name,
                    "store": store,
                    "group_ceco": group_ceco,
                    "end_of_term": values[9],
                    "status_en_rem": self._clean_text(values[10]),
                    "rent_min": values[11],
                    "current_canon": values[14],
                    "vat_total": values[15],
                    "vat_vw": values[16],
                    "vat_vq": values[17],
                    "vw_percent": values[18],
                    "vq_percent": values[19],
                }
                continue

            if not current or account_or_name == "TOTAL":
                continue

            if not self._is_macro_line_row(values[1], amount, posting_key, text):
                continue

            ceco = line_ceco if self._clean_text(line_ceco) else current.get("group_ceco")
            records.append(
                {
                    "category": current.get("category"),
                    "vendor": current.get("vendor"),
                    "vendor_name": current.get("vendor_name"),
                    "store": current.get("store"),
                    "ceco": ceco,
                    "profit_center": ceco,
                    "account": self._clean_numeric_code(values[1]),
                    "account_name": account_or_name,
                    "amount": amount,
                    "tax_code": tax_code,
                    "posting_key": posting_key,
                    "text": text,
                    "end_of_term": current.get("end_of_term"),
                    "status_en_rem": current.get("status_en_rem"),
                    "rent_min": current.get("rent_min"),
                    "current_canon": current.get("current_canon"),
                    "vat_total": current.get("vat_total"),
                    "vat_vw": current.get("vat_vw"),
                    "vat_vq": current.get("vat_vq"),
                    "vw_percent": current.get("vw_percent"),
                    "vq_percent": current.get("vq_percent"),
                    "source_row": row_idx,
                }
            )

        return records

    def _build_macro_lucy_export(
        self,
        summary_df: pd.DataFrame,
        lines_df: pd.DataFrame,
        macro_database_df: pd.DataFrame,
        macro_template_path: Path | None = None,
    ) -> pd.DataFrame:
        template_columns = self._resolve_macro_lucy_columns(macro_template_path)
        export_rows = []
        for _, row in lines_df.iterrows():
            export_rows.append(
                {
                    "_posting_index": row.get("posting_index"),
                    "D/A$S/H": row.get("posting_key"),
                    "Imp$Amount": row.get("amount"),
                    "Cod_Iva$Vat Code": row.get("tax_code"),
                    "Conto$GL Account": row.get("account"),
                    "CdC$Cost Center": None,
                    "ProfitCenter$Profit Center": row.get("profit_center"),
                    "Text$Text": row.get("text"),
                    "Ordine$Internal Order": None,
                    "Wbs$WBS": None,
                    "Attribuzione$Assignment": row.get("reference"),
                    "Network$Network": None,
                    "OpNetwork$OpNetwork": None,
                    "MaterialNumber$Material Number": None,
                    "Quantity$Quantity": None,
                    "TransactionType$Transaction Type": None,
                    "POR$POR#": None,
                    "CompanyGL$CompanyGL": None,
                    "Brand$Brand": row.get("store"),
                    "TradingPartner$Trading Partner": None,
                    "PosOrd$Order Position": None,
                    "OdaOdv$PO/SR": None,
                }
            )

        export_df = pd.DataFrame(export_rows)
        if export_df.empty:
            export_df = pd.DataFrame(columns=self.MACRO_LUCY_COLUMNS)

        for column in template_columns:
            if column not in export_df.columns:
                export_df[column] = None

        export_df = export_df[["_posting_index"] + template_columns].copy()
        return export_df

    def _build_header_export(self, summary_df: pd.DataFrame) -> pd.DataFrame:
        rows = []
        for _, row in summary_df.iterrows():
            invoice_date = pd.to_datetime(row.get("invoice_date"), errors="coerce")
            due_date = pd.to_datetime(row.get("due_date"), errors="coerce")
            invoice_date_text = invoice_date.strftime("%d-%m-%Y") if pd.notna(invoice_date) else None
            posting_date = invoice_date.strftime("%d-%m-%Y") if pd.notna(invoice_date) else None
            due_date_text = due_date.strftime("%d-%m-%Y") if pd.notna(due_date) else None
            period = invoice_date.strftime("%m") if pd.notna(invoice_date) else None
            vat_total = pd.to_numeric(pd.Series([row.get("vat_total")]), errors="coerce").fillna(0).iloc[0]
            amount = row.get("amount")
            invoice_total = row.get("invoice_total")
            withholding_tax_amount = row.get("withholding_tax_amount")
            vat_code = "V0" if float(vat_total) == 0 else "I1"

            data = {
                "_posting_index": row.get("posting_index"),
                "checkSkipSubjectChannel": False,
                "Legal entity": "8140 - GMO Colombia",
                "Original Invoice Number": row.get("invoice_id"),
                "Assignment 1": row.get("text"),
                "Post date": posting_date,
                "System": "Lucy",
                "Period": period,
                "Invoice Date": invoice_date_text,
                "Original Invoice Ref Number": row.get("invoice_id"),
                "Invoice Number": row.get("invoice_id"),
                "Document Type Sap": "KR",
                "Currency": "COP",
                "Vat Code": vat_code,
                "Invoice Total Amount": invoice_total,
                "Total Impuestos Retenidos": withholding_tax_amount,
                "Total Taxable Amount": amount,
                "Total VAT Amount": vat_total,
                "Total Vat Amount": vat_total,
                "Withholding Tax Amount": withholding_tax_amount,
                "Allocated Costs": row.get("invoice_id"),
                "Balance": True,
                "Accounting Counterpart": amount,
                "Calculate Tax": "X" if pd.notna(vat_total) and float(vat_total) > 0 else None,
                "SAP_KOSTL": row.get("ceco"),
                "SAP_PROJK": row.get("profit_center"),
                "Due Date": due_date_text,
                "Payment Terms": row.get("payment_terms_text"),
                "Group Company": "8140",
            }
            rows.append(data)

        header_df = pd.DataFrame(rows)
        if header_df.empty:
            header_df = pd.DataFrame(columns=self.HEADER_EXPORT_COLUMNS)

        return header_df.reindex(columns=["_posting_index"] + self.HEADER_EXPORT_COLUMNS)

    def _build_header_matrix(self) -> pd.DataFrame:
        rows = [
            {
                "column_name": "Invoice Date",
                "fill_status": "fillable",
                "source": "PDF or invoice file",
                "logic": "Uses invoice_date resolved from PDF first, then invoice file.",
            },
            {
                "column_name": "Invoice Number",
                "fill_status": "fillable",
                "source": "Invoices file / PDF",
                "logic": "Uses normalized invoice_id.",
            },
            {
                "column_name": "Original Invoice Ref Number",
                "fill_status": "fillable",
                "source": "Invoices file / PDF",
                "logic": "Uses invoice_id as reference until business defines another field.",
            },
            {
                "column_name": "Original Invoice Number",
                "fill_status": "fillable",
                "source": "Invoices file / PDF",
                "logic": "Uses invoice_id.",
            },
            {
                "column_name": "Document Type Sap",
                "fill_status": "fillable",
                "source": "Business rule",
                "logic": "Defaults to KR for supplier invoice flow.",
            },
            {
                "column_name": "Currency",
                "fill_status": "fillable_with_assumption",
                "source": "Business rule",
                "logic": "Defaults to COP for this Colombia lease process.",
            },
            {
                "column_name": "Invoice Total Amount",
                "fill_status": "fillable",
                "source": "Invoices file + PDF enrichment",
                "logic": "Uses resolved invoice_total.",
            },
            {
                "column_name": "Total Taxable Amount",
                "fill_status": "fillable",
                "source": "Invoices file + PDF enrichment",
                "logic": "Uses amount without VAT.",
            },
            {
                "column_name": "Total VAT Amount",
                "fill_status": "fillable",
                "source": "XML / invoice file",
                "logic": "Uses detected VAT; XML invoices without a VAT value are treated as zero.",
            },
            {
                "column_name": "Withholding Tax Amount",
                "fill_status": "partial",
                "source": "PDF / invoice file",
                "logic": "Filled when the PDF exposes retentions or withholding values explicitly.",
            },
            {
                "column_name": "Total Impuestos Retenidos",
                "fill_status": "partial",
                "source": "PDF / invoice file",
                "logic": "Mirrors Withholding Tax Amount when the source invoice exposes retained tax.",
            },
            {
                "column_name": "Vat Code",
                "fill_status": "fillable",
                "source": "Derived rule",
                "logic": "Uses V0 when VAT is zero or absent from XML, and I1 when VAT is positive.",
            },
            {
                "column_name": "Post date",
                "fill_status": "fillable",
                "source": "Invoice Date",
                "logic": "Uses invoice_date as the posting date in the current flow.",
            },
            {
                "column_name": "Period",
                "fill_status": "fillable",
                "source": "Invoice Date",
                "logic": "Uses the invoice month to populate the period field.",
            },
            {
                "column_name": "Legal entity",
                "fill_status": "fillable_with_assumption",
                "source": "Business rule",
                "logic": "Defaults to 8140 - GMO Colombia for this process.",
            },
            {
                "column_name": "SAP_KOSTL",
                "fill_status": "fillable",
                "source": "Macro DB / control / contracts / prorateo",
                "logic": "Uses resolved CECO.",
            },
            {
                "column_name": "SAP_PROJK",
                "fill_status": "fillable",
                "source": "Macro DB / resolved CECO",
                "logic": "Uses resolved profit center, usually same as CECO in current flow.",
            },
            {
                "column_name": "Payment Terms",
                "fill_status": "partial",
                "source": "PDF text when available",
                "logic": "Captures human-readable payment term text from the PDF, not SAP-coded terms yet.",
            },
            {
                "column_name": "Payment Mode",
                "fill_status": "not_filled_yet",
                "source": "Missing master data",
                "logic": "Not present in current PDFs or Excel masters.",
            },
            {
                "column_name": "Due Date",
                "fill_status": "partial",
                "source": "PDF",
                "logic": "Filled when the invoice PDF exposes due date clearly.",
            },
            {
                "column_name": "Allocated Costs",
                "fill_status": "helper_only",
                "source": "Generated reference",
                "logic": "Uses invoice_id as a link to the separate allocated costs export.",
            },
            {
                "column_name": "Accounting Counterpart",
                "fill_status": "fillable",
                "source": "Derived from amount",
                "logic": "Uses the taxable/base amount for the current posting row.",
            },
            {
                "column_name": "Calculate Tax",
                "fill_status": "partial",
                "source": "Derived rule",
                "logic": "Marked with X only when a positive VAT value is resolved.",
            },
            {
                "column_name": "Balance",
                "fill_status": "fillable_with_assumption",
                "source": "Business rule",
                "logic": "Defaults to True in the generated header export.",
            },
            {
                "column_name": "Group Company",
                "fill_status": "fillable_with_assumption",
                "source": "Business rule",
                "logic": "Defaults to 8140 for the current Colombia legal entity flow.",
            },
            {
                "column_name": "Other workflow/system columns",
                "fill_status": "left_blank",
                "source": "System generated",
                "logic": "Kept blank because they belong to OCR, workflow, SAP response, or later automation stages.",
            },
        ]
        return pd.DataFrame(rows)

    def _write_allocated_costs_csv_bundle(
        self,
        summary_df: pd.DataFrame,
        lucy_export_df: pd.DataFrame,
        bundle_name: str,
        timestamp: str,
    ) -> None:
        bundle_dir = self.output_dir / f"allocated_costs_csvs_{timestamp}"
        bundle_dir.mkdir(parents=True, exist_ok=True)

        for _, summary_row in summary_df.dropna(subset=["invoice_id"]).iterrows():
            safe_invoice_id = self._posting_file_stem(summary_row)
            allocated_slice = self._lucy_slice_for_posting(lucy_export_df, summary_row)
            if allocated_slice.empty:
                visible_columns = lucy_export_df.drop(
                    columns=["_posting_index"],
                    errors="ignore",
                ).columns
                allocated_slice = pd.DataFrame(columns=visible_columns)

            file_path = bundle_dir / f"{safe_invoice_id}_allocated_costs.csv"
            allocated_slice.to_csv(file_path, index=False, encoding="utf-8-sig")

        zip_path = self.output_dir / bundle_name
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in bundle_dir.glob("*.csv"):
                zip_file.write(file_path, arcname=file_path.name)

    def _write_invoice_csv_bundle(
        self,
        summary_df: pd.DataFrame,
        header_export_df: pd.DataFrame,
        lucy_export_df: pd.DataFrame,
        bundle_name: str,
        timestamp: str,
    ) -> None:
        bundle_dir = self.output_dir / f"invoice_csvs_{timestamp}"
        bundle_dir.mkdir(parents=True, exist_ok=True)

        for _, summary_row in summary_df.dropna(subset=["invoice_id"]).iterrows():
            invoice_id = str(summary_row["invoice_id"])
            safe_invoice_id = self._posting_file_stem(summary_row)
            invoice_dir = bundle_dir / safe_invoice_id
            invoice_dir.mkdir(parents=True, exist_ok=True)

            header_slice = self._header_slice_for_posting(header_export_df, summary_row)
            allocated_slice = self._lucy_slice_for_posting(lucy_export_df, summary_row)

            if header_slice.empty:
                header_slice = pd.DataFrame(columns=self.HEADER_EXPORT_COLUMNS)
            if allocated_slice.empty:
                allocated_slice = pd.DataFrame(columns=lucy_export_df.columns)

            header_path = invoice_dir / f"{safe_invoice_id}_header.csv"
            allocated_path = invoice_dir / f"{safe_invoice_id}_allocated_costs.csv"
            header_slice.to_csv(header_path, index=False, encoding="utf-8-sig")
            allocated_slice.to_csv(allocated_path, index=False, encoding="utf-8-sig")

        zip_path = self.output_dir / bundle_name
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in bundle_dir.rglob("*"):
                if file_path.is_file():
                    zip_file.write(file_path, arcname=file_path.relative_to(bundle_dir.parent))

    def _write_manual_style_csv_bundle(
        self,
        summary_df: pd.DataFrame,
        header_export_df: pd.DataFrame,
        lucy_export_df: pd.DataFrame,
        bundle_name: str,
        timestamp: str,
    ) -> None:
        bundle_dir = self.output_dir / f"invoice_manual_style_csvs_{timestamp}"
        bundle_dir.mkdir(parents=True, exist_ok=True)

        for _, summary_row in summary_df.dropna(subset=["invoice_id"]).iterrows():
            safe_invoice_id = self._posting_file_stem(summary_row)
            file_path = bundle_dir / f"{safe_invoice_id}_manual_style.csv"

            header_slice = self._header_slice_for_posting(header_export_df, summary_row)
            allocated_slice = self._lucy_slice_for_posting(lucy_export_df, summary_row)
            summary_slice = pd.DataFrame([summary_row])

            self._write_manual_style_csv(
                path=file_path,
                header_slice=header_slice,
                allocated_slice=allocated_slice,
                summary_slice=summary_slice,
            )

        zip_path = self.output_dir / bundle_name
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in bundle_dir.rglob("*"):
                if file_path.is_file():
                    zip_file.write(file_path, arcname=file_path.relative_to(bundle_dir.parent))

    def _posting_file_stem(self, summary_row) -> str:
        invoice_id = self._safe_filename(str(summary_row.get("invoice_id") or "SIN_FACTURA"))
        posting_count = pd.to_numeric(pd.Series([summary_row.get("posting_count")]), errors="coerce").iloc[0]
        posting_index = pd.to_numeric(pd.Series([summary_row.get("posting_index")]), errors="coerce").iloc[0]
        if pd.notna(posting_count) and int(posting_count) > 1 and pd.notna(posting_index):
            return f"{invoice_id}_{int(posting_index)}"
        return invoice_id

    def _header_slice_for_posting(self, header_export_df: pd.DataFrame, summary_row) -> pd.DataFrame:
        invoice_id = str(summary_row.get("invoice_id") or "")
        result = header_export_df[header_export_df["Invoice Number"].astype(str) == invoice_id].copy()
        posting_index = pd.to_numeric(pd.Series([summary_row.get("posting_index")]), errors="coerce").iloc[0]
        if "_posting_index" in result.columns and pd.notna(posting_index):
            narrowed = result[pd.to_numeric(result["_posting_index"], errors="coerce") == int(posting_index)]
            if not narrowed.empty:
                result = narrowed.copy()
        ceco = self._clean_text(summary_row.get("ceco"))
        if ceco and "SAP_KOSTL" in result.columns:
            narrowed = result[result["SAP_KOSTL"].astype(str).map(self._clean_text) == ceco]
            if not narrowed.empty:
                result = narrowed.copy()
        return result.drop(columns=["_posting_index"], errors="ignore")

    def _lucy_slice_for_posting(self, lucy_export_df: pd.DataFrame, summary_row) -> pd.DataFrame:
        invoice_id = str(summary_row.get("invoice_id") or "")
        result = lucy_export_df[lucy_export_df["Attribuzione$Assignment"].astype(str) == invoice_id].copy()
        posting_index = pd.to_numeric(pd.Series([summary_row.get("posting_index")]), errors="coerce").iloc[0]
        if "_posting_index" in result.columns and pd.notna(posting_index):
            narrowed = result[pd.to_numeric(result["_posting_index"], errors="coerce") == int(posting_index)]
            if not narrowed.empty:
                result = narrowed.copy()
        ceco = self._clean_text(summary_row.get("ceco"))
        profit_center = self._clean_text(summary_row.get("profit_center"))
        if ceco and "CdC$Cost Center" in result.columns:
            narrowed = result[result["CdC$Cost Center"].astype(str).map(self._clean_text) == ceco]
            if not narrowed.empty:
                result = narrowed.copy()
        if profit_center and "ProfitCenter$Profit Center" in result.columns:
            narrowed = result[result["ProfitCenter$Profit Center"].astype(str).map(self._clean_text) == profit_center]
            if not narrowed.empty:
                result = narrowed.copy()
        if "CdC$Cost Center" in result.columns:
            result["CdC$Cost Center"] = None
        return result.drop(columns=["_posting_index"], errors="ignore")

    def _write_manual_style_csv(
        self,
        path: Path,
        header_slice: pd.DataFrame,
        allocated_slice: pd.DataFrame,
        summary_slice: pd.DataFrame,
    ) -> None:
        header_columns = [
            "checkSkipSubjectChannel",
            "Invoice Date",
            "Invoice Number",
            "Document Type Sap",
            "Currency",
            "Vat Code",
            "Invoice Total Amount",
            "Total Taxable Amount",
            "Total VAT Amount",
            "Payment Terms",
            "Payment Mode",
            "Allocated Costs",
            "WHT",
            "Partner Bank",
            "Action",
            "Balance",
            "Accounting Counterpart",
            "FISCAL_CODE",
            "Result",
            "Calculate Tax",
            "Last Sap Execution",
            "Due Date",
            "Archive Link Done",
            "SAP_ARCHIVE",
        ]
        allocated_columns = ["S/H", "Amount", "GL Account", "Text", "Vat Code", "Profit Center"]
        wht_columns = ["Wht Description", "Wht Code", "Wht Type", "Wht Base Amount", "Wht Amount"]

        header_row = {}
        if not header_slice.empty:
            raw = header_slice.iloc[0]
            allocated_text = self._coalesce_nonblank(raw.get("Assignment 1"), raw.get("Allocated Costs"))
            header_row = {
                "checkSkipSubjectChannel": self._format_csv_scalar(raw.get("checkSkipSubjectChannel")),
                "Invoice Date": raw.get("Invoice Date"),
                "Invoice Number": raw.get("Invoice Number"),
                "Document Type Sap": raw.get("Document Type Sap"),
                "Currency": raw.get("Currency"),
                "Vat Code": self._resolve_manual_header_vat_code(raw),
                "Invoice Total Amount": self._format_csv_scalar(raw.get("Invoice Total Amount")),
                "Total Taxable Amount": self._format_csv_scalar(raw.get("Total Taxable Amount")),
                "Total VAT Amount": self._format_csv_scalar(raw.get("Total VAT Amount")),
                "Payment Terms": self._resolve_manual_payment_terms(raw.get("Payment Terms")),
                "Payment Mode": self._resolve_manual_payment_mode(raw.get("Payment Mode")),
                "Allocated Costs": allocated_text,
                "WHT": None,
                "Partner Bank": self.PARTNER_BANK_DEFAULT,
                "Action": self._coalesce_nonblank(raw.get("Action"), "EASYAUT"),
                "Balance": self._format_csv_scalar(raw.get("Balance")),
                "Accounting Counterpart": self._format_csv_scalar(raw.get("Accounting Counterpart")),
                "FISCAL_CODE": self._coalesce_nonblank(raw.get("FISCAL_CODE"), "01"),
                "Result": self._format_csv_scalar(raw.get("Result") if pd.notna(raw.get("Result")) else 0),
                "Calculate Tax": raw.get("Calculate Tax"),
                "Last Sap Execution": self._coalesce_nonblank(raw.get("Last Sap Execution")),
                "Due Date": raw.get("Due Date"),
                "Archive Link Done": self._coalesce_nonblank(raw.get("Archive Link Done")),
                "SAP_ARCHIVE": self._format_csv_scalar(raw.get("SAP_ARCHIVE") if pd.notna(raw.get("SAP_ARCHIVE")) else 1),
            }

        allocated_rows = []
        if not allocated_slice.empty:
            for _, row in allocated_slice.iterrows():
                allocated_rows.append(
                    {
                        "S/H": row.get("D/A$S/H"),
                        "Amount": self._format_csv_scalar(row.get("Imp$Amount")),
                        "GL Account": row.get("Conto$GL Account"),
                        "Text": row.get("Text$Text"),
                        "Vat Code": row.get("Cod_Iva$Vat Code"),
                        "Profit Center": row.get("ProfitCenter$Profit Center"),
                    }
                )

        wht_rows = self._build_manual_wht_rows(summary_slice)

        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, quoting=csv.QUOTE_ALL)
            writer.writerow(header_columns)
            writer.writerow([header_row.get(column) for column in header_columns])
            writer.writerow([])
            writer.writerow(["CostiRipartiti"])
            writer.writerow(allocated_columns)
            for row in allocated_rows:
                writer.writerow([row.get(column) for column in allocated_columns])
            writer.writerow([])
            writer.writerow(["Wht"])
            writer.writerow(wht_columns)
            for row in wht_rows:
                writer.writerow([row.get(column) for column in wht_columns])

    def _build_manual_wht_rows(self, summary_slice: pd.DataFrame) -> list[dict]:
        if summary_slice.empty:
            return [self._empty_manual_wht_row()]

        total_withholding = pd.to_numeric(summary_slice["withholding_tax_amount"], errors="coerce").fillna(0).sum()
        taxable_amount = pd.to_numeric(summary_slice["amount"], errors="coerce").fillna(0).sum()
        if total_withholding <= 0:
            return [self._empty_manual_wht_row()]

        return [
            {
                "Wht Description": "Withholding income tax on fees",
                "Wht Code": "A2",
                "Wht Type": "YA",
                "Wht Base Amount": self._format_csv_scalar(round_cop(taxable_amount, 0)),
                "Wht Amount": self._format_csv_scalar(round_cop(total_withholding, 0)),
            }
        ]

    def _empty_manual_wht_row(self) -> dict:
        return {
            "Wht Description": None,
            "Wht Code": None,
            "Wht Type": None,
            "Wht Base Amount": None,
            "Wht Amount": None,
        }

    def _resolve_manual_header_vat_code(self, raw: pd.Series) -> str | None:
        total_vat = pd.to_numeric(pd.Series([raw.get("Total VAT Amount")]), errors="coerce").fillna(0).iloc[0]
        if float(total_vat) > 0:
            return "VQ"
        return "V0"

    def _resolve_manual_payment_terms(self, value) -> str | None:
        return self.PAYMENT_TERMS_DEFAULT

    def _resolve_manual_payment_mode(self, value) -> str | None:
        return self.PAYMENT_MODE_DEFAULT

    def _format_csv_scalar(self, value):
        if pd.isna(value):
            return None
        if isinstance(value, (bool, np.bool_)):
            return "true" if value else "false"
        if isinstance(value, (np.integer, int)):
            return str(int(value))
        if isinstance(value, (np.floating, float)):
            if float(value).is_integer():
                return str(int(value))
            return f"{float(value):.2f}".rstrip("0").rstrip(".")
        return str(value)

    def _coalesce_nonblank(self, *values):
        for value in values:
            if pd.isna(value) or value is None:
                continue
            text = str(value).strip()
            if not text or text.lower() == "nan":
                continue
            return text
        return None

    def _build_macro_export_from_database(self, summary_df: pd.DataFrame, macro_database_df: pd.DataFrame) -> pd.DataFrame:
        macro_lines = []
        for _, summary_row in summary_df.iterrows():
            vendor_key = self._text_key(summary_row.get("vendor"))
            store_key = self._text_key(summary_row.get("store"))
            matches = macro_database_df[
                (macro_database_df["vendor_key"] == vendor_key) & (macro_database_df["store_key"] == store_key)
            ].copy()

            if matches.empty:
                continue

            for _, macro_row in matches.iterrows():
                macro_lines.append(
                    {
                        "D/A$S/H": macro_row.get("posting_key"),
                        "Imp$Amount": round_cop(macro_row.get("amount"), 0),
                        "Cod_Iva$Vat Code": macro_row.get("tax_code"),
                        "Conto$GL Account": macro_row.get("account"),
                        "CdC$Cost Center": None,
                        "ProfitCenter$Profit Center": macro_row.get("profit_center") or macro_row.get("ceco"),
                        "Text$Text": macro_row.get("text"),
                        "Ordine$Internal Order": None,
                        "Wbs$WBS": None,
                        "Attribuzione$Assignment": summary_row.get("invoice_id"),
                        "Network$Network": None,
                        "OpNetwork$OpNetwork": None,
                        "MaterialNumber$Material Number": None,
                        "Quantity$Quantity": None,
                        "TransactionType$Transaction Type": None,
                        "POR$POR#": None,
                        "CompanyGL$CompanyGL": None,
                        "Brand$Brand": summary_row.get("store"),
                        "TradingPartner$Trading Partner": None,
                        "PosOrd$Order Position": None,
                        "OdaOdv$PO/SR": None,
                    }
                )

        return pd.DataFrame(macro_lines)

    def _resolve_macro_lucy_columns(self, macro_template_path: Path | None) -> list[str]:
        if not macro_template_path:
            return self.MACRO_LUCY_COLUMNS.copy()

        try:
            macro_preview = pd.read_excel(macro_template_path, sheet_name="lucy", nrows=1)
        except Exception as exc:  # pragma: no cover
            self._log(f"Could not read macro workbook template; using default Lucy export columns. Detail: {exc}")
            return self.MACRO_LUCY_COLUMNS.copy()

        columns = [self._clean_text(col) for col in macro_preview.columns]
        columns = [col for col in columns if col]
        if not columns:
            return self.MACRO_LUCY_COLUMNS.copy()
        return columns

    def _build_lucy_views(self, output_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        summary_df = output_df.copy()
        for col in [
            "amount",
            "gross_amount",
            "discount_total",
            "posting_discount_total",
            "payable_rounding",
            "invoice_total",
            "vat_total",
            "vat_vw",
            "vat_vq",
            "withholding_tax_amount",
        ]:
            if col in summary_df.columns:
                summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce").map(round_cop)
        if "sell_cam" not in summary_df.columns:
            summary_df["sell_cam"] = 0
        summary_df["invoice_date"] = pd.to_datetime(summary_df["invoice_date"], errors="coerce")
        summary_df["posting_key"] = summary_df["invoice_id"].fillna("") + "|" + summary_df["vendor"].fillna("")
        summary_df["posting_index"] = summary_df.groupby("invoice_id").cumcount() + 1
        summary_df["posting_count"] = summary_df.groupby("invoice_id")["invoice_id"].transform("size")
        if "support_split_index" not in summary_df.columns:
            summary_df["support_split_index"] = summary_df["posting_index"]
        if "support_split_factor" not in summary_df.columns:
            summary_df["support_split_factor"] = summary_df["posting_count"]
        summary_df["is_credit_note"] = (
            summary_df.get("ubl_document_type", pd.Series(index=summary_df.index, dtype=object))
            .fillna("Invoice")
            .astype(str)
            .str.casefold()
            .eq("creditnote")
        )
        summary_df["document_type"] = np.where(
            summary_df["is_credit_note"],
            "Supplier Credit Note",
            "Supplier Invoice",
        )
        summary_df["movement_type"] = "5-FI"
        summary_df["currency"] = "COP"
        summary_df["invoice_total"] = (
            summary_df["amount"].fillna(0) + summary_df["vat_total"].fillna(0)
        ).map(lambda value: round_cop(value, 0))
        summary_df["taxable_total"] = summary_df["amount"]
        summary_df["profit_center"] = summary_df["profit_center"].combine_first(summary_df["ceco"])
        summary_df["reference"] = summary_df["invoice_id"]
        summary_df["vat_status"] = np.select(
            [
                summary_df["vat_total"].notna() & summary_df["vat_total"].eq(0),
                summary_df["vat_total"].notna(),
            ],
            [
                "resolved_zero",
                "resolved_value",
            ],
            default="unresolved",
        )

        headers_df = summary_df[
            [
                "invoice_id",
                "vendor",
                "posting_index",
                "posting_count",
                "support_split_index",
                "support_split_factor",
                "invoice_date",
                "document_type",
                "movement_type",
                "currency",
                "invoice_total",
                "taxable_total",
                "gross_amount",
                "discount_total",
                "posting_discount_total",
                "withholding_tax_amount",
                "store",
                "ceco",
                "profit_center",
                "reference",
                "concept",
                "text",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
                "invoice_line_items",
                "vat_status",
                "is_credit_note",
            ]
        ].copy()

        line_rows: list[dict] = []
        for _, row in summary_df.iterrows():
            base = {
                "invoice_id": row["invoice_id"],
                "vendor": row["vendor"],
                "posting_index": row["posting_index"],
                "posting_count": row["posting_count"],
                "support_split_index": row.get("support_split_index"),
                "support_split_factor": row.get("support_split_factor"),
                "invoice_date": row["invoice_date"],
                "store": row["store"],
                "ceco": row["ceco"],
                "profit_center": row["profit_center"],
                "reference": row["reference"],
                "concept": row["concept"],
                "text": row["text"],
                "is_credit_note": row["is_credit_note"],
            }

            item_rows = []
            for item in self._invoice_line_items(row.get("invoice_line_items")):
                item_amount = pd.to_numeric(pd.Series([item.get("amount")]), errors="coerce").iloc[0]
                if pd.isna(item_amount) or abs(float(item_amount)) == 0:
                    continue
                item_concept, item_account = self._item_line_classification(item, row)
                item_text = self._build_text(row["invoice_date"], item_concept, row["store"], row["ceco"])
                credit_note = bool(row.get("is_credit_note", False))
                item_amount = -abs(float(item_amount)) if credit_note else abs(float(item_amount))
                item_rows.append(
                    {
                        **base,
                        "line_type": "invoice_item",
                        "posting_key": "40",
                        "account": item_account,
                        "concept": item_concept,
                        "tax_code": self._expense_tax_code(row),
                        "amount": round_cop(item_amount, 0),
                        "text": item_text,
                    }
                )

            if item_rows:
                expense_rows = item_rows
            else:
                expense_rows = [
                    {
                        **base,
                        "line_type": "expense",
                        "posting_key": "40",
                        "account": row["account"],
                        "concept": row["concept"],
                        "tax_code": self._expense_tax_code(row),
                        "amount": row["gross_amount"],
                    }
                ]
            line_rows.extend(expense_rows)

            posting_discount_total = pd.to_numeric(
                pd.Series([row.get("posting_discount_total")]),
                errors="coerce",
            ).iloc[0]
            if pd.notna(posting_discount_total) and float(posting_discount_total) > 0:
                allocations = self._allocate_proportionally(
                    [abs(float(expense_row.get("amount") or 0)) for expense_row in expense_rows],
                    float(posting_discount_total),
                )
                discount_text = self._build_text(
                    row["invoice_date"],
                    "INCENTIVO GC",
                    row["store"],
                    row["ceco"],
                )
                for expense_row, allocated_discount in zip(expense_rows, allocations):
                    if allocated_discount <= 0:
                        continue
                    line_rows.append(
                        {
                            **base,
                            "line_type": "discount",
                            "posting_key": "50",
                            "account": expense_row.get("account"),
                            "concept": expense_row.get("concept"),
                            "tax_code": expense_row.get("tax_code"),
                            "amount": allocated_discount,
                            "text": discount_text,
                        }
                    )

            if self._is_full_vw_prorate(row.get("vw_percent")) or self._is_full_vq_prorate(row.get("vq_percent")):
                continue

            if pd.notna(row["vat_vw"]) and abs(float(row["vat_vw"])) > 0:
                line_rows.append(
                    {
                        **base,
                        "line_type": "vat_vw",
                        "posting_key": "40",
                        "account": self.VAT_ACCOUNT,
                        "concept": row["concept"],
                        "tax_code": "VW",
                        "amount": row["vat_vw"],
                    }
                )

            if pd.notna(row["vat_vq"]) and abs(float(row["vat_vq"])) > 0:
                line_rows.append(
                    {
                        **base,
                        "line_type": "vat_vq",
                        "posting_key": "50",
                        "account": self.VAT_ACCOUNT,
                        "concept": row["concept"],
                        "tax_code": "VQ",
                        "amount": row["vat_vq"],
                    }
                )

            if pd.notna(row["vat_vw"]) and abs(float(row["vat_vw"])) > 0:
                line_rows.append(
                    {
                        **base,
                        "line_type": "vat_vq_reversal",
                        "posting_key": "50",
                        "account": self.VAT_ACCOUNT,
                        "concept": row["concept"],
                        "tax_code": "VQ",
                        "amount": row["vat_vw"],
                    }
                )

        lines_df = pd.DataFrame(line_rows)
        if lines_df.empty:
            lines_df = pd.DataFrame(
                columns=[
                    "invoice_id",
                    "vendor",
                    "posting_index",
                    "posting_count",
                    "support_split_index",
                    "support_split_factor",
                    "invoice_date",
                    "store",
                    "ceco",
                    "profit_center",
                    "reference",
                    "text",
                    "line_type",
                    "posting_key",
                    "account",
                    "concept",
                    "tax_code",
                    "amount",
                ]
            )
        else:
            lines_df["amount"] = pd.to_numeric(lines_df["amount"], errors="coerce").map(
                lambda value: round_cop(value, 0)
            )
            credit_note_mask = lines_df["is_credit_note"].fillna(False).astype(bool)
            lines_df.loc[credit_note_mask, "posting_key"] = (
                lines_df.loc[credit_note_mask, "posting_key"]
                .astype(str)
                .map({"40": "50", "50": "40"})
                .fillna(lines_df.loc[credit_note_mask, "posting_key"])
            )
            lines_df = lines_df.drop(columns=["is_credit_note"])

        summary_df = summary_df[
            [
                "invoice_id",
                "vendor",
                "store",
                "ceco",
                "account",
                "concept",
                "rent_type",
                "sell_cam",
                "amount",
                "gross_amount",
                "discount_total",
                "posting_discount_total",
                "payable_rounding",
                "vat_total",
                "vat_vw",
                "vat_vq",
                "vw_percent",
                "vq_percent",
                "withholding_tax_amount",
                "text",
                "invoice_date",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
                "invoice_line_items",
                "invoice_discounts",
                "posting_index",
                "posting_count",
                "support_split_index",
                "support_split_factor",
                "invoice_total",
                "vat_status",
                "ubl_document_type",
                "is_credit_note",
            ]
        ].copy()

        return summary_df, headers_df, lines_df

    def _build_lucy_mapping(
        self,
        summary_df: pd.DataFrame,
        headers_df: pd.DataFrame,
        lines_df: pd.DataFrame,
    ) -> pd.DataFrame:
        header_example = headers_df.iloc[0].to_dict() if not headers_df.empty else {}
        line_example = lines_df.iloc[0].to_dict() if not lines_df.empty else {}
        summary_example = summary_df.iloc[0].to_dict() if not summary_df.empty else {}

        rows = [
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Invoice Date",
                "excel_sheet": "lucy_headers",
                "excel_column": "invoice_date",
                "example_value": header_example.get("invoice_date"),
                "how_to_use": "Use this as the invoice date in the left panel of Lucy.",
            },
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Invoice Number / Reference",
                "excel_sheet": "lucy_headers",
                "excel_column": "reference",
                "example_value": header_example.get("reference"),
                "how_to_use": "Use this for the invoice number or document reference field.",
            },
            {
                "lucy_area": "Top Header",
                "lucy_field": "Vendor Code",
                "excel_sheet": "lucy_headers",
                "excel_column": "vendor",
                "example_value": header_example.get("vendor"),
                "how_to_use": "Use this in the vendor code field at the top of Lucy.",
            },
            {
                "lucy_area": "Top Header",
                "lucy_field": "Document Class",
                "excel_sheet": "lucy_headers",
                "excel_column": "document_type",
                "example_value": header_example.get("document_type"),
                "how_to_use": "Guide value for document class. Adjust only if business requires another class.",
            },
            {
                "lucy_area": "Top Header",
                "lucy_field": "Movement Type",
                "excel_sheet": "lucy_headers",
                "excel_column": "movement_type",
                "example_value": header_example.get("movement_type"),
                "how_to_use": "Guide value for the movement type shown in Lucy.",
            },
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Total Invoice Amount",
                "excel_sheet": "lucy_headers",
                "excel_column": "invoice_total",
                "example_value": header_example.get("invoice_total"),
                "how_to_use": "Use this as the total document amount for the posting.",
            },
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Due Date",
                "excel_sheet": "lucy_headers",
                "excel_column": "due_date",
                "example_value": header_example.get("due_date"),
                "how_to_use": "Use this when Lucy exposes a due date field in the invoice header.",
            },
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Total Taxable Amount",
                "excel_sheet": "lucy_headers",
                "excel_column": "taxable_total",
                "example_value": header_example.get("taxable_total"),
                "how_to_use": "Use this as the taxable/base amount when Lucy asks for taxable total.",
            },
            {
                "lucy_area": "Invoice Information",
                "lucy_field": "Withholding Tax Amount",
                "excel_sheet": "lucy_headers",
                "excel_column": "withholding_tax_amount",
                "example_value": header_example.get("withholding_tax_amount"),
                "how_to_use": "Use this only when the invoice exposes retained taxes and Lucy asks for them.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "G/L Account",
                "excel_sheet": "lucy_lines",
                "excel_column": "account",
                "example_value": line_example.get("account"),
                "how_to_use": "Use one row per posting line in the allocated costs section.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Posting Key",
                "excel_sheet": "lucy_lines",
                "excel_column": "posting_key",
                "example_value": line_example.get("posting_key"),
                "how_to_use": "This mirrors the debit or credit behavior seen later in SAP.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Tax Code",
                "excel_sheet": "lucy_lines",
                "excel_column": "tax_code",
                "example_value": line_example.get("tax_code"),
                "how_to_use": "Use this to identify whether the line is VW, VQ, or expense behavior.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Amount",
                "excel_sheet": "lucy_lines",
                "excel_column": "amount",
                "example_value": line_example.get("amount"),
                "how_to_use": "Use the amount for that specific line in the allocated costs grid.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Cost Center",
                "excel_sheet": "lucy_lines",
                "excel_column": "ceco",
                "example_value": line_example.get("ceco"),
                "how_to_use": "Use this in the cost center field for each line.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Profit Center",
                "excel_sheet": "lucy_lines",
                "excel_column": "profit_center",
                "example_value": line_example.get("profit_center"),
                "how_to_use": "Use this only if Lucy or SAP requires a profit center in that case.",
            },
            {
                "lucy_area": "Allocated Costs Grid",
                "lucy_field": "Line Text",
                "excel_sheet": "lucy_lines",
                "excel_column": "text",
                "example_value": line_example.get("text"),
                "how_to_use": "Use this as the text shown on each accounting line.",
            },
            {
                "lucy_area": "Control / Review",
                "lucy_field": "Posting Split",
                "excel_sheet": "summary",
                "excel_column": "posting_index / posting_count",
                "example_value": f"{summary_example.get('posting_index', '')}/{summary_example.get('posting_count', '')}",
                "how_to_use": "Use this to understand whether the invoice becomes one posting or several postings/vendors.",
            },
            {
                "lucy_area": "Control / Review",
                "lucy_field": "VAT Status",
                "excel_sheet": "summary",
                "excel_column": "vat_status",
                "example_value": summary_example.get("vat_status"),
                "how_to_use": "Use this to know whether VAT was resolved, resolved as zero, or still unresolved.",
            },
            {
                "lucy_area": "Control / Review",
                "lucy_field": "Rent Logic",
                "excel_sheet": "summary",
                "excel_column": "rent_type",
                "example_value": summary_example.get("rent_type"),
                "how_to_use": "This explains whether the posting was classified as RF or RV before reaching Lucy.",
            },
        ]

        return pd.DataFrame(rows)

    def _style_output_workbook(self, workbook_path: Path) -> None:
        workbook = load_workbook(workbook_path)

        self._style_mapping_sheet(workbook["lucy_mapping"])
        self._style_mapping_sheet(workbook["lucy_header_matrix"])
        self._style_grouped_sheet(workbook["summary"])
        self._style_grouped_sheet(workbook["lucy_header_export"])
        self._style_grouped_sheet(workbook["lucy_headers"])
        self._style_grouped_sheet(workbook["lucy_lines"])
        self._style_grouped_sheet(workbook["lucy_export"])

        workbook.save(workbook_path)

    def _style_mapping_sheet(self, worksheet) -> None:
        self._style_header_row(worksheet)
        self._autosize_columns(worksheet)

    def _style_grouped_sheet(self, worksheet) -> None:
        self._style_header_row(worksheet)

        invoice_col = self._find_column_index(worksheet, "invoice_id")
        posting_col = self._find_column_index(worksheet, "posting_index")
        if invoice_col is None:
            self._autosize_columns(worksheet)
            return

        invoice_palette: dict[str, tuple[str, list[str]]] = {}
        palette_index = 0

        for row_idx in range(2, worksheet.max_row + 1):
            invoice_id = worksheet.cell(row=row_idx, column=invoice_col).value
            if invoice_id is None:
                continue

            invoice_key = str(invoice_id)
            if invoice_key not in invoice_palette:
                invoice_palette[invoice_key] = self.GROUP_COLORS[palette_index % len(self.GROUP_COLORS)]
                palette_index += 1

            strong_color, soft_colors = invoice_palette[invoice_key]
            posting_index = 1
            if posting_col is not None:
                raw_posting = worksheet.cell(row=row_idx, column=posting_col).value
                try:
                    posting_index = int(raw_posting)
                except (TypeError, ValueError):
                    posting_index = 1

            fill_color = strong_color if posting_index <= 1 else soft_colors[min(posting_index - 2, len(soft_colors) - 1)]
            font_color = "FFFFFF" if posting_index <= 1 else "111827"
            self._fill_row(worksheet, row_idx, fill_color, font_color)

        worksheet.freeze_panes = "A2"
        self._autosize_columns(worksheet)

    def _style_header_row(self, worksheet) -> None:
        for cell in worksheet[1]:
            cell.fill = PatternFill(fill_type="solid", fgColor=self.HEADER_FILL)
            cell.font = Font(color=self.HEADER_FONT, bold=True)

    def _fill_row(self, worksheet, row_idx: int, fill_color: str, font_color: str) -> None:
        fill = PatternFill(fill_type="solid", fgColor=fill_color)
        font = Font(color=font_color, bold=False)
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.fill = fill
            cell.font = font

    def _find_column_index(self, worksheet, column_name: str) -> int | None:
        for idx, cell in enumerate(worksheet[1], start=1):
            if cell.value == column_name:
                return idx
        return None

    def _autosize_columns(self, worksheet) -> None:
        for column_cells in worksheet.columns:
            values = [str(cell.value) for cell in column_cells if cell.value is not None]
            if not values:
                continue
            max_length = min(max(len(value) for value in values) + 2, 40)
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length

    def _safe_read_table(self, path: Path | None) -> pd.DataFrame:
        if path is None:
            return pd.DataFrame()
        if path.suffix.lower() == ".csv":
            try:
                return pd.read_csv(path, sep=None, engine="python", encoding="utf-8-sig")
            except UnicodeDecodeError:
                return pd.read_csv(path, sep=None, engine="python", encoding="latin1")
        return pd.read_excel(path)

    def _best_effort_rename(self, columns: Iterable[str], expected: dict[str, list[str]]) -> dict[str, str]:
        normalized = {self._normalize_column_name(col): col for col in columns}
        rename_map = {}
        for target, options in expected.items():
            for option in options:
                option_key = self._normalize_column_name(option)
                if option_key in normalized:
                    rename_map[normalized[option_key]] = target
                    break
        return rename_map

    def _normalize_column_name(self, value) -> str:
        text = self._clean_text(value) or ""
        text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
        return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")

    def _clean_text(self, value):
        if pd.isna(value):
            return None
        text = str(value).replace("\n", " ").replace("\xa0", " ").strip()
        text = re.sub(r"\s+", " ", text)
        return text or None

    def _text_key(self, value) -> str | None:
        text = self._clean_text(value)
        return re.sub(r"[^A-Z0-9]", "", text.upper()) if text else None

    def _ceco_key(self, value) -> str | None:
        text = self._clean_text(value)
        return re.sub(r"[^A-Z0-9]", "", text.upper()) if text else None

    def _invoice_key(self, value) -> str | None:
        text = self._normalize_invoice_id(value)
        if not text:
            return None
        key = re.sub(r"[^A-Z0-9]", "", text.upper())
        if key.isdigit():
            return key.lstrip("0") or "0"
        return key

    def _normalize_invoice_id(self, value) -> str | None:
        text = self._clean_text(value)
        if not text:
            return None
        text = re.sub(
            r"^\s*NOTA\s+CR[EÉ]DITO\s*[:#-]?\s*",
            "",
            text,
            flags=re.IGNORECASE,
        )
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
        text = text.replace(" ", "")
        return text

    def _clean_numeric_code(self, value) -> str | None:
        text = self._clean_text(value)
        if not text:
            return None
        if re.fullmatch(r"\d+\.0+", text):
            text = text.split(".")[0]
        digits = re.sub(r"[^0-9]", "", text)
        return digits or None

    def _extract_primary_code(self, value) -> str | None:
        text = self._clean_text(value)
        if not text:
            return None
        match = re.search(r"\d+", text)
        return match.group(0) if match else None

    def _resolve_period(self, invoice_dates: pd.Series, period: str | None) -> pd.Series:
        resolved = pd.to_datetime(invoice_dates, errors="coerce")
        if period:
            try:
                base = pd.to_datetime(f"01-{period}", format="%d-%m-%Y")
            except ValueError as exc:
                raise PipelineError("Period must be in MM-YYYY format, for example 03-2026.") from exc
            return resolved.fillna(base)
        return resolved.fillna(pd.Timestamp.today())

    def _calculated_invoice_total(self, subtotal, vat, fallback_total=None):
        subtotal_value = pd.to_numeric(pd.Series([subtotal]), errors="coerce").iloc[0]
        vat_value = pd.to_numeric(pd.Series([vat]), errors="coerce").iloc[0]
        if pd.notna(subtotal_value) and pd.notna(vat_value):
            return round_cop(float(subtotal_value) + float(vat_value), 0)
        fallback_value = pd.to_numeric(pd.Series([fallback_total]), errors="coerce").iloc[0]
        return fallback_value if pd.notna(fallback_value) else None

    def _resolve_account_concept(self, account, source_text=None) -> str:
        account_code = self._clean_numeric_code(account)
        source_key = self._concept_key(source_text)
        if account_code and source_key:
            for rule_account, rule_text, concept in self.CONCEPT_RULES:
                if account_code == rule_account and self._concept_key(rule_text) in source_key:
                    return concept
        return self.ACCOUNT_FALLBACK_CONCEPTS.get(account_code, "RV")

    def _concept_key(self, value) -> str:
        text = self._clean_text(value) or ""
        text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
        return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()

    def _invoice_line_items(self, value) -> list[dict]:
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        return []

    def _beneficiary_distribution_candidate(self, row) -> dict | None:
        items = self._invoice_line_items(row.get("invoice_line_items"))
        if len(items) < 2 or self._invoice_discounts(row.get("invoice_discounts")):
            return None

        descriptions = [
            re.sub(r"\s+", " ", str(item.get("description") or "")).strip()
            for item in items
        ]
        if not descriptions[0] or len(set(descriptions)) != 1:
            return None

        beneficiary_ids = [self._clean_text(item.get("beneficiary_id")) for item in items]
        if any(not beneficiary_id for beneficiary_id in beneficiary_ids):
            return None
        if len(set(beneficiary_ids)) != len(items):
            return None

        signatures = set()
        amounts = []
        beneficiaries = []
        for item, description, beneficiary_id in zip(items, descriptions, beneficiary_ids):
            allowance = pd.to_numeric(pd.Series([item.get("allowance_amount")]), errors="coerce").fillna(0).iloc[0]
            charge = pd.to_numeric(pd.Series([item.get("charge_amount")]), errors="coerce").fillna(0).iloc[0]
            if float(allowance) != 0 or float(charge) != 0:
                return None
            amount = pd.to_numeric(pd.Series([item.get("amount")]), errors="coerce").iloc[0]
            if pd.isna(amount) or float(amount) <= 0:
                return None
            quantity = pd.to_numeric(pd.Series([item.get("quantity")]), errors="coerce").iloc[0]
            tax_percent = pd.to_numeric(pd.Series([item.get("tax_percent")]), errors="coerce").iloc[0]
            signatures.add(
                (
                    description,
                    self._clean_text(item.get("item_code")) or "",
                    None if pd.isna(quantity) else float(quantity),
                    None if pd.isna(tax_percent) else float(tax_percent),
                    self._clean_text(item.get("tax_scheme_id")) or "",
                )
            )
            rounded_amount = round_cop(amount, 0)
            amounts.append(rounded_amount)
            beneficiaries.append(
                {
                    "beneficiary_id": beneficiary_id,
                    "amount": rounded_amount,
                }
            )
        if len(signatures) != 1:
            return None

        items_total = round_cop(sum(amounts), 0)
        subtotal = pd.to_numeric(
            pd.Series(
                [
                    row.get("support_subtotal"),
                    (row.get("total") or 0) - (row.get("vat") or 0),
                    row.get("amount"),
                ]
            ),
            errors="coerce",
        ).dropna()
        if subtotal.empty or round_cop(subtotal.iloc[0], 0) != items_total:
            return None

        for beneficiary in beneficiaries:
            beneficiary["percentage"] = (
                float(beneficiary["amount"]) / float(items_total) * 100 if items_total else 0
            )

        invoice_id = self._clean_text(row.get("invoice_id")) or "SIN_NUMERO"
        review_key = self._clean_text(row.get("invoice_pdf_key")) or self._invoice_key(invoice_id)
        if not review_key:
            return None
        return {
            "review_key": review_key,
            "invoice_id": invoice_id,
            "source_file": self._clean_text(row.get("support_source_file")) or "XML sin nombre",
            "store": self._clean_text(row.get("store")) or "Sin tienda",
            "ceco": self._clean_text(row.get("ceco")) or "Sin CeCo",
            "description": descriptions[0],
            "line_count": len(items),
            "beneficiary_count": len(beneficiary_ids),
            "beneficiaries": beneficiaries,
            "items_total": items_total,
            "split_factor": max(int(row.get("support_split_factor") or 1), 1),
        }

    def _consolidate_beneficiary_line_items(self, line_items) -> list[dict]:
        items = self._invoice_line_items(line_items)
        if not items:
            return []
        consolidated = dict(items[0])
        for key in ["amount", "net_amount", "allowance_amount", "charge_amount", "tax_amount"]:
            values = pd.to_numeric(
                pd.Series([item.get(key) for item in items]),
                errors="coerce",
            ).dropna()
            if not values.empty:
                consolidated[key] = round_cop(values.sum(), 0)
        consolidated["line_id"] = "CONSOLIDATED_BENEFICIARIES"
        consolidated["beneficiary_ids"] = [item.get("beneficiary_id") for item in items]
        consolidated["beneficiary_id"] = None
        consolidated["source"] = "consolidated_beneficiary_distribution"
        return [consolidated]

    def _apply_beneficiary_distribution_selections(
        self,
        invoices: pd.DataFrame,
        selections: dict[str, str],
    ) -> pd.DataFrame:
        if invoices.empty:
            return invoices

        result = invoices.copy()
        skip_indexes = []
        reviewed_keys = set()
        skipped_keys = set()
        for index, row in result.iterrows():
            candidate = self._beneficiary_distribution_candidate(row)
            if not candidate:
                continue
            review_key = candidate["review_key"]
            action = str(selections.get(review_key) or "").lower()
            if action == "consolidate":
                result.at[index, "invoice_line_items"] = self._consolidate_beneficiary_line_items(
                    row.get("invoice_line_items")
                )
                self._record_warning(
                    f"Factura {candidate['invoice_id']}: se consolidaron {candidate['line_count']} lineas "
                    f"de beneficiarios en una sola linea por {candidate['items_total']}."
                )
            elif action == "skip":
                skip_indexes.append(index)
                if review_key not in skipped_keys:
                    self.preflight_processing_issues.append(
                        {
                            "invoice_id": candidate["invoice_id"],
                            "source_file": candidate["source_file"],
                            "store": candidate["store"],
                            "location": f"{candidate['source_file']} / {candidate['store']}",
                            "problem": (
                                f"El XML contiene {candidate['line_count']} lineas iguales asignadas a "
                                f"beneficiarios diferentes por un total de {candidate['items_total']}; "
                                "la factura fue omitida por decision del usuario"
                            ),
                            "possible_solution": (
                                f"Revisa en el PDF si el concepto aparece como una sola linea por "
                                f"{candidate['items_total']}. Luego vuelve a procesar y elige consolidar "
                                "o continuar normalmente."
                            ),
                            "issue_type": "beneficiary_distribution_skipped",
                        }
                    )
                    skipped_keys.add(review_key)
            elif action == "keep":
                continue
            elif review_key not in reviewed_keys:
                self.beneficiary_review_items.append(candidate)
                reviewed_keys.add(review_key)

        if skip_indexes:
            result = result.drop(index=skip_indexes).copy()
            self._record_warning(
                f"Se omitieron {len(skipped_keys)} facturas con distribucion interna por beneficiarios "
                "por decision del usuario. El resto del lote continuo."
            )
        if result.empty and skip_indexes:
            raise PipelineError(
                "Todas las facturas del lote se marcaron para procesamiento manual. No se generaron archivos CSV."
            )
        return result

    def _invoice_discounts(self, value) -> list[dict]:
        if isinstance(value, list):
            return [discount for discount in value if isinstance(discount, dict)]
        return []

    def _evaluate_invoice_discounts(
        self,
        discounts,
        invoice_date,
        registration_date: pd.Timestamp,
        manual_decision: str | None = None,
    ) -> dict:
        estimated_payment = self._estimated_payment_monday(registration_date)
        applied_discounts = []
        decisions = []
        invoice_timestamp = pd.to_datetime(invoice_date, errors="coerce")

        for raw_discount in self._invoice_discounts(discounts):
            discount = dict(raw_discount)
            source = self._clean_text(discount.get("source")) or "unknown"
            condition_type = self._clean_text(discount.get("condition_type")) or "formal"
            applied = True
            deadline = None
            decision_reason = "formal_xml_discount"

            if source == "xml_note" and condition_type == "explicit_deadline":
                deadline = pd.to_datetime(discount.get("deadline_date"), errors="coerce")
                inclusive = bool(discount.get("deadline_inclusive", True))
                applied = bool(
                    pd.notna(deadline)
                    and (estimated_payment <= deadline if inclusive else estimated_payment < deadline)
                )
                decision_reason = "deadline_met" if applied else "deadline_missed"
            elif source == "xml_note" and condition_type == "relative_period":
                deadline = self._relative_discount_deadline(
                    invoice_timestamp,
                    discount.get("relative_days"),
                    bool(discount.get("business_days")),
                )
                applied = bool(pd.notna(deadline) and estimated_payment <= deadline)
                decision_reason = "relative_deadline_met" if applied else "relative_deadline_missed"
            elif source == "xml_note":
                if manual_decision == "apply":
                    applied = True
                    decision_reason = "manual_apply"
                elif manual_decision == "omit":
                    applied = False
                    decision_reason = "manual_omit"
                else:
                    decision_reason = "manual_review_required"

            discount["applied"] = applied
            discount["registration_date"] = registration_date.strftime("%Y-%m-%d")
            discount["estimated_payment_date"] = estimated_payment.strftime("%Y-%m-%d")
            discount["calculated_deadline"] = deadline.strftime("%Y-%m-%d") if pd.notna(deadline) else None
            discount["decision_reason"] = decision_reason
            decisions.append(discount)
            if applied:
                applied_discounts.append(discount)

        return {
            "applied_discounts": applied_discounts,
            "decisions": decisions,
            "estimated_payment_date": estimated_payment,
        }

    def _discount_review_option(self, row, decision: dict, review_key: str) -> dict:
        amount = pd.to_numeric(pd.Series([decision.get("amount")]), errors="coerce").fillna(0).iloc[0]
        percentage = pd.to_numeric(pd.Series([decision.get("percentage")]), errors="coerce").iloc[0]
        return {
            "review_key": review_key,
            "invoice_id": self._clean_text(row.get("invoice_id")) or "SIN_NUMERO",
            "source_file": self._clean_text(row.get("support_source_file")) or "XML sin nombre",
            "store": self._clean_text(row.get("mapped_store")) or self._clean_text(row.get("store")) or "Sin tienda",
            "ceco": self._clean_text(row.get("mapped_ceco")) or self._clean_text(row.get("ceco")) or "Sin CeCo",
            "amount": float(amount),
            "percentage": float(percentage) if pd.notna(percentage) else None,
            "condition": self._clean_text(decision.get("reason")) or "Condicion no especificada",
        }

    def _estimated_payment_monday(self, registration_date) -> pd.Timestamp:
        registration = pd.Timestamp(registration_date).normalize()
        days_to_wednesday = (2 - registration.weekday()) % 7
        proposal_wednesday = registration + pd.Timedelta(days=days_to_wednesday)
        return proposal_wednesday + pd.Timedelta(days=5)

    def _relative_discount_deadline(self, invoice_date, day_count, business_days: bool):
        if pd.isna(invoice_date):
            return pd.NaT
        count = pd.to_numeric(pd.Series([day_count]), errors="coerce").iloc[0]
        if pd.isna(count) or int(count) <= 0:
            return pd.NaT

        month_start = pd.Timestamp(invoice_date).to_period("M").start_time.normalize()
        if not business_days:
            return month_start + pd.Timedelta(days=int(count) - 1)

        colombia_holidays = holidays.country_holidays("CO", years=[month_start.year])
        current = month_start
        elapsed = 0
        while current.month == month_start.month:
            if current.weekday() < 5 and current.date() not in colombia_holidays:
                elapsed += 1
                if elapsed == int(count):
                    return current
            current += pd.Timedelta(days=1)
        return pd.NaT

    def _discount_decision_message(self, invoice_id, decision: dict) -> str:
        amount = pd.to_numeric(pd.Series([decision.get("amount")]), errors="coerce").fillna(0).iloc[0]
        payment = self._format_audit_date(decision.get("estimated_payment_date"))
        deadline = self._format_audit_date(decision.get("calculated_deadline"))
        reason = decision.get("decision_reason")
        if reason == "manual_apply":
            return (
                f"Factura {invoice_id}: descuento no calculable de {round_cop(amount, 0)} aplicado por decision del usuario."
            )
        if reason == "manual_omit":
            return f"Factura {invoice_id}: descuento no calculable de {round_cop(amount, 0)} omitido por decision del usuario."
        action = "aplicado" if decision.get("applied") else "omitido"
        return (
            f"Factura {invoice_id}: incentivo por pronto pago de {round_cop(amount, 0)} {action}. "
            f"Pago estimado {payment}; fecha limite {deadline}."
        )

    def _format_audit_date(self, value) -> str:
        timestamp = pd.to_datetime(value, errors="coerce")
        return timestamp.strftime("%d-%m-%Y") if pd.notna(timestamp) else "NO_CALCULABLE"

    def _invoice_discount_total(self, value) -> float:
        total = 0.0
        for discount in self._invoice_discounts(value):
            amount = pd.to_numeric(pd.Series([discount.get("amount")]), errors="coerce").iloc[0]
            if pd.notna(amount) and float(amount) > 0:
                total += float(amount)
        return round_cop(total, 0)

    def _gross_invoice_amount(self, row) -> float | None:
        item_amounts = []
        for item in self._invoice_line_items(row.get("invoice_line_items")):
            amount = pd.to_numeric(pd.Series([item.get("amount")]), errors="coerce").iloc[0]
            if pd.notna(amount):
                item_amounts.append(float(amount))
        if item_amounts:
            return round_cop(sum(item_amounts), 0)

        for candidate in [
            row.get("support_subtotal"),
            row.get("amount"),
            (row.get("total") or 0) - (row.get("vat") or 0),
        ]:
            amount = pd.to_numeric(pd.Series([candidate]), errors="coerce").iloc[0]
            if pd.notna(amount):
                return round_cop(amount, 0)
        return None

    def _allocate_proportionally(self, amounts, total: float) -> list[float]:
        numeric_amounts = [
            float(value) if pd.notna(value) else 0.0
            for value in pd.to_numeric(pd.Series(list(amounts)), errors="coerce").fillna(0)
        ]
        amount_sum = sum(max(amount, 0) for amount in numeric_amounts)
        if amount_sum <= 0 or total <= 0:
            return [0.0 for _ in numeric_amounts]

        largest_index = max(range(len(numeric_amounts)), key=numeric_amounts.__getitem__)
        return allocate_cop(total, numeric_amounts, residual_index=largest_index)

    def _multiply_invoice_line_items(self, line_items, factor: float) -> list[dict]:
        scaled_items = []
        for item in self._invoice_line_items(line_items):
            scaled = dict(item)
            for key in ["amount", "net_amount", "allowance_amount", "charge_amount", "tax_amount"]:
                value = pd.to_numeric(pd.Series([scaled.get(key)]), errors="coerce").iloc[0]
                if pd.notna(value):
                    scaled[key] = round_cop(float(value) * float(factor), 0)
            scaled_items.append(scaled)
        return scaled_items

    def _multiply_invoice_discounts(self, discounts, factor: float) -> list[dict]:
        scaled_discounts = []
        for discount in self._invoice_discounts(discounts):
            scaled = dict(discount)
            for key in ["amount", "base_amount"]:
                value = pd.to_numeric(pd.Series([scaled.get(key)]), errors="coerce").iloc[0]
                if pd.notna(value):
                    scaled[key] = round_cop(float(value) * float(factor), 0)
            scaled_discounts.append(scaled)
        return scaled_discounts

    def _item_line_concept(self, item: dict, row) -> str:
        description = self._clean_text(item.get("description"))
        account_concept = self._resolve_account_concept(row.get("account"), description)
        if description and account_concept == self.ACCOUNT_FALLBACK_CONCEPTS.get(self._clean_numeric_code(row.get("account")), "RV"):
            return self._compact_item_concept(description)
        return account_concept

    def _item_line_classification(self, item: dict, row) -> tuple[str, str]:
        description = self._clean_text(item.get("description"))
        dictionary_match = self._match_learning_dictionary(description, row)
        if dictionary_match:
            concept, account = self._resolve_gc_item_classification(*dictionary_match, row)
            if self._is_office_store(row.get("store")) or self._is_office_store(row.get("mapped_store")):
                account = self._office_account_for_concept(concept) or account
            return concept, account
        return "SIN CLASIFICAR", ""

    def _resolve_gc_item_classification(self, concept: str, account: str, row) -> tuple[str, str]:
        normalized_concept = self._concept_key(concept)
        if normalized_concept == "GC V":
            sell_cam = pd.to_numeric(pd.Series([row.get("sell_cam")]), errors="coerce").iloc[0]
            if pd.notna(sell_cam) and float(sell_cam) > 0:
                return "GF", (
                    self.ADMIN_SELL_CAM_ACCOUNT
                    or self.FIXED_ACCOUNT
                    or account
                )
            return "GV", (
                self.ADMIN_VARIABLE_ACCOUNT
                or self._account_code_for_concept("GC VARIABLE")
                or self._account_code_for_concept("GC")
                or account
            )
        if normalized_concept not in {"GC", "GC VARIABLE"}:
            return concept, account
        sell_media = pd.to_numeric(pd.Series([row.get("sell_media")]), errors="coerce").iloc[0]
        if pd.notna(sell_media) and float(sell_media) > 0:
            return "GC", self.FIXED_ACCOUNT
        return "GC V", self._account_code_for_concept("GC VARIABLE") or self._account_code_for_concept("GC")

    def _match_learning_dictionary(self, description: str | None, row) -> tuple[str, str] | None:
        description_key = self._concept_key(normalize_learning_phrase(description))
        if not description_key or self.learning_dictionary.empty:
            return None

        is_office_invoice = self._is_office_store(row.get("store")) or self._is_office_store(row.get("mapped_store"))
        candidates = []
        for _, rule in self.learning_dictionary.iterrows():
            concept = self._clean_text(rule.get("concept"))
            account = self._clean_text(rule.get("account"))
            sigla = self._clean_text(rule.get("text"))
            if self._is_office_dictionary_concept(concept) and not is_office_invoice:
                continue
            phrases = self._split_dictionary_phrases(rule.get("phrases"))
            if not concept or not phrases:
                continue
            for phrase in phrases:
                phrase_key = self._concept_key(normalize_learning_phrase(phrase))
                if phrase_key and phrase_key in description_key:
                    candidates.append((len(phrase_key), concept, account, phrase_key, sigla))

        if not candidates:
            return None

        _, concept, account, _, sigla = sorted(candidates, key=lambda item: item[0], reverse=True)[0]
        if self._is_office_dictionary_concept(concept):
            office_concept = self._concept_key(concept)
            if office_concept == "RF OFICINA":
                concept = self._clean_text(row.get("rent_type")) or "RV"
            else:
                concept = sigla or self._office_concept_counterpart(concept)
            account = self._clean_numeric_code(account) or account
            return concept, account

        if account == "SEGUN CONTRATO" or account == "SEGUN_CONTRATO" or concept == "RENTA":
            account = self._clean_numeric_code(row.get("account")) or self.VARIABLE_ACCOUNT
            concept = self._resolve_account_concept(account, row.get("rent_type"))
            if is_office_invoice:
                concept = self._clean_text(row.get("rent_type")) or concept
                account = self._office_account_for_concept("RF") or account
        elif not account:
            return None
        else:
            account = self._clean_numeric_code(account) or account
            concept = sigla if sigla else concept
        return concept, account

    def _is_office_store(self, value) -> bool:
        return "OFICINA" in self._concept_key(value)

    def _is_office_dictionary_concept(self, concept: str | None) -> bool:
        return " OFICINA" in f" {self._concept_key(concept)}"

    def _office_concept_counterpart(self, concept: str | None) -> str:
        concept_key = self._concept_key(concept)
        if concept_key == "GC OFICINA":
            return "GC V"
        if concept_key == "ACUED OFICINA":
            return "ACUED"
        if concept_key == "ENERGIA OFICINA":
            return "ENERGIA"
        if concept_key == "RF OFICINA":
            return "RENTA"
        return self._clean_text(concept) or ""

    def _office_account_for_concept(self, concept: str | None) -> str | None:
        concept_key = self._concept_key(concept)
        if concept_key in {"RF", "RV", "RENTA", "RF OFICINA"}:
            office_concept = "RF OFICINA"
        elif concept_key in {"ENERGIA", "ENERGIA OFICINA"}:
            office_concept = "ENERGIA OFICINA"
        elif concept_key in {"ACUED", "ACUED OFICINA"}:
            office_concept = "ACUED OFICINA"
        elif concept_key in {"GC", "GC V", "GC VARIABLE", "GV", "GF", "GC OFICINA"}:
            office_concept = "GC OFICINA"
        else:
            return None

        if self.learning_dictionary.empty:
            raise PipelineError(
                f"Falta el concepto {office_concept} en Diccionario_Conceptos_Simple.xlsx."
            )
        matches = self.learning_dictionary[
            self.learning_dictionary["concept"].map(self._concept_key).eq(office_concept)
        ]
        accounts = {
            account
            for account in matches["account"].map(self._clean_numeric_code)
            if account
        }
        if len(accounts) > 1:
            raise PipelineError(
                f"El concepto {office_concept} tiene mas de una cuenta en Diccionario_Conceptos_Simple.xlsx."
            )
        if not accounts:
            raise PipelineError(
                f"Falta la cuenta para {office_concept} en Diccionario_Conceptos_Simple.xlsx."
            )
        return next(iter(accounts))

    def _split_dictionary_phrases(self, value) -> list[str]:
        text = self._clean_text(value)
        if not text:
            return []
        return [part.strip() for part in str(text).split(",") if part.strip()]

    def build_concept_review_items(self, output_df: pd.DataFrame) -> list[dict]:
        review_items: dict[str, dict] = {}
        for _, row in output_df.iterrows():
            for item in self._invoice_line_items(row.get("invoice_line_items")):
                description = self._clean_text(item.get("description"))
                amount = pd.to_numeric(pd.Series([item.get("amount")]), errors="coerce").iloc[0]
                if not description or pd.isna(amount) or abs(float(amount)) == 0:
                    continue
                if self._match_learning_dictionary(description, row):
                    continue

                suggested_phrase = normalize_learning_phrase(description) or description
                key = self._concept_key(suggested_phrase)
                if key not in review_items:
                    review_items[key] = {
                        "key": key,
                        "description": description,
                        "suggested_phrase": suggested_phrase,
                        "count": 0,
                        "total_amount": 0.0,
                        "examples": [],
                    }
                review_items[key]["count"] += 1
                review_items[key]["total_amount"] += float(amount)
                if len(review_items[key]["examples"]) < 3:
                    review_items[key]["examples"].append(
                        {
                            "invoice_id": row.get("invoice_id"),
                            "store": row.get("store"),
                            "ceco": row.get("ceco"),
                            "amount": round_cop(amount, 0),
                        }
                    )
        return sorted(review_items.values(), key=lambda item: item["description"])

    def learning_concepts(self) -> list[dict]:
        if self.learning_dictionary.empty:
            return []
        concepts = []
        seen = set()
        for _, row in self.learning_dictionary.iterrows():
            concept = self._clean_text(row.get("concept"))
            account = self._clean_text(row.get("account"))
            if not concept or concept in seen:
                continue
            seen.add(concept)
            concepts.append({"concept": concept, "account": account})
        return concepts

    def _compact_item_concept(self, description: str) -> str:
        concept = self._concept_key(description)
        concept = re.sub(r"\b(MES|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE|ENERO|FEBRERO|MARZO|ABRIL)\b", "", concept)
        concept = re.sub(r"\b20\d{2}\b", "", concept)
        concept = re.sub(r"\s+", " ", concept).strip()
        return concept[:45] if concept else "ITEM"

    def _build_text(self, period_date, rent_type, store, ceco) -> str:
        month_year = pd.to_datetime(period_date).strftime("%m-%Y") if pd.notna(period_date) else "00-0000"
        return self._limit_text(f"{month_year} {rent_type or 'RV'} {store or 'NO_STORE'} {ceco or 'NO_CECO'}")

    def _limit_text(self, value: str) -> str:
        text = re.sub(r"\s+", " ", str(value or "")).strip()
        if len(text) <= self.TEXT_MAX_LENGTH:
            return text

        words = text.split()
        compacted = []
        current = ""
        for word in words:
            candidate = f"{current} {word}".strip()
            if len(candidate) > self.TEXT_MAX_LENGTH:
                break
            compacted.append(word)
            current = candidate
        if compacted:
            return " ".join(compacted)
        return text[: self.TEXT_MAX_LENGTH].rstrip()

    def _is_active_contract_status(self, value) -> bool:
        status = self._concept_key(value)
        if not status:
            return False
        return status in {"VIGENTE", "POR VENCER"}

    def _is_blocked_contract_status(self, value) -> bool:
        status = self._concept_key(value)
        return status.startswith("REEMPLAZO") or status.startswith("REEMPLAZADO")

    def _extract_percent(self, label) -> float | None:
        text = self._clean_text(label)
        if not text:
            return None
        match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", text)
        return float(match.group(1).replace(",", ".")) / 100.0 if match else None

    def _is_full_vw_prorate(self, value) -> bool:
        percent = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
        if pd.isna(percent):
            return False
        if percent > 1:
            percent = percent / 100.0
        return float(percent) >= 0.999999

    def _is_full_vq_prorate(self, value) -> bool:
        percent = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
        if pd.isna(percent):
            return False
        if percent > 1:
            percent = percent / 100.0
        return float(percent) >= 0.999999

    def _expense_tax_code(self, row) -> str:
        if self._is_full_vw_prorate(row.get("vw_percent")):
            return "VW"
        return "VQ"

    def _is_valid_ceco_value(self, value) -> bool:
        text = self._clean_text(value)
        if not text:
            return False
        blocked = {
            "VARIABLE",
            "VIGENTE",
            "CESADO",
            "OK",
            "NO",
            "N/A",
            "NA",
            "TOTAL",
        }
        if text.upper() in blocked:
            return False
        return bool(re.fullmatch(r"[A-Z]?\d{2,8}", text.upper()))

    def _is_macro_line_row(self, account, amount, posting_key, text) -> bool:
        if pd.isna(account) or pd.isna(amount) or pd.isna(posting_key):
            return False
        account_code = self._clean_numeric_code(account)
        if not account_code or self._clean_text(text) == "TOTAL":
            return False
        try:
            float(str(amount).replace(",", "").replace(" ", ""))
        except ValueError:
            return False
        return True

    def _looks_numeric(self, value) -> bool:
        try:
            float(value)
            return True
        except (TypeError, ValueError):
            return False

    def _safe_filename(self, value: str) -> str:
        safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value).strip())
        return safe or "invoice"

    def _extract_store_from_filename(self, path: Path) -> str | None:
        stem = path.stem.upper()
        match = re.search(r"(^|[^A-Z0-9])([A-Z]{1,5}\d{2,5})(?=$|[^A-Z0-9])", stem)
        return match.group(2) if match else None

    def _warn_support_store_mismatches(self, invoices: pd.DataFrame, supports: pd.DataFrame) -> None:
        if invoices.empty or supports.empty:
            return

        for _, support in supports.iterrows():
            source_file = self._clean_text(support.get("source_file"))
            filename_store = self._extract_store_from_filename(Path(source_file)) if source_file else None
            invoice_key = self._clean_text(support.get("invoice_key"))
            if not filename_store or not invoice_key:
                continue

            matches = invoices[invoices["invoice_pdf_key"] == invoice_key]
            expected_stores = sorted(
                {
                    store
                    for store in matches.get("store", pd.Series(dtype=object)).map(self._clean_text).tolist()
                    if store
                }
            )
            if not expected_stores or self._text_key(filename_store) in {self._text_key(store) for store in expected_stores}:
                continue

            invoice_id = self._clean_text(support.get("invoice_id")) or "SIN_NUMERO"
            expected_label = ", ".join(expected_stores)
            warning = (
                f"El archivo '{source_file}' indica la tienda {filename_store} en su nombre, pero la factura "
                f"{invoice_id} pertenece a la tienda {expected_label} en el invoices file. "
                "Revisa si el XML fue nombrado o enviado incorrectamente; el proceso continuó sin modificarlo."
            )
            self._record_warning(warning)

    def _record_warning(self, warning: str) -> None:
        if warning not in self.warnings:
            self.warnings.append(warning)
            self._log(f"WARNING: {warning}")

    def _log(self, message: str) -> None:
        LOGGER.info(message)
        self.logs.append(message)
