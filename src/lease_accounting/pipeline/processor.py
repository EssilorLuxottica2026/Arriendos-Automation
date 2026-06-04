import csv
import logging
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from .pdf_reader import InvoiceSupportReader


LOGGER = logging.getLogger(__name__)


class PipelineError(Exception):
    """Controlled error for pipeline and upload issues."""


@dataclass
class PipelineResult:
    output_excel: str
    output_excel_fixed: str
    output_csv: str
    header_csv: str
    invoice_csv_bundle: str
    manual_style_csv_bundle: str
    validation_excel: str
    normalized_invoices_excel: str
    output_rows: int
    header_rows: int
    validation_rows: int
    logs: list[str]

    def to_dict(self) -> dict:
        return {
            "output_excel": self.output_excel,
            "output_excel_fixed": self.output_excel_fixed,
            "output_csv": self.output_csv,
            "header_csv": self.header_csv,
            "invoice_csv_bundle": self.invoice_csv_bundle,
            "manual_style_csv_bundle": self.manual_style_csv_bundle,
            "validation_excel": self.validation_excel,
            "normalized_invoices_excel": self.normalized_invoices_excel,
            "output_rows": self.output_rows,
            "header_rows": self.header_rows,
            "validation_rows": self.validation_rows,
            "logs": self.logs,
        }


class LeaseAccountingPipeline:
    FIXED_ACCOUNT = "1245150017"
    VARIABLE_ACCOUNT = "1537210004"
    VAT_ACCOUNT = "1149120011"
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

    def __init__(self, output_dir: Path):
        self.base_output_dir = Path(output_dir)
        self.base_output_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir = self.base_output_dir
        self.logs: list[str] = []
        self.distribution_rules = pd.DataFrame()

    def run(
        self,
        invoices_path: Path,
        control_path: Path,
        contracts_path: Path,
        prorateo_path: Path,
        distribution_path: Path,
        history_path: Path | None = None,
        support_paths: list[Path] | None = None,
        macro_template_path: Path | None = None,
        period: str | None = None,
    ) -> dict:
        self.logs = []
        run_folder_name = pd.Timestamp.now().strftime("%Y.%m.%d_%H.%M")
        self.output_dir = self.base_output_dir / run_folder_name
        self.output_dir.mkdir(parents=True, exist_ok=True)
        data = self.load_data(
            invoices_path=invoices_path,
            control_path=control_path,
            contracts_path=contracts_path,
            prorateo_path=prorateo_path,
            distribution_path=distribution_path,
            history_path=history_path,
            support_paths=support_paths,
            macro_template_path=macro_template_path,
        )
        data = self.clean_data(data)
        merged = self.merge_data(data)
        output_df, validation_df = self.apply_rules(merged, period=period)
        return self.generate_output(
            output_df,
            validation_df,
            data["invoices"],
            macro_database_df=data.get("macro_database", pd.DataFrame()),
            macro_template_path=macro_template_path,
        ).to_dict()

    def load_data(
        self,
        invoices_path: Path,
        control_path: Path,
        contracts_path: Path,
        prorateo_path: Path,
        distribution_path: Path,
        history_path: Path | None = None,
        support_paths: list[Path] | None = None,
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
            "invoice_supports": self._load_invoice_supports(support_paths or []),
            "macro_database": self._load_macro_database(macro_template_path) if macro_template_path else pd.DataFrame(),
        }

    def clean_data(self, data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
        self._log("Cleaning and standardizing data.")
        cleaned: dict[str, pd.DataFrame] = {}
        for name, frame in data.items():
            if frame.empty:
                cleaned[name] = frame.copy()
                continue
            df = frame.copy()
            df.columns = [self._normalize_column_name(col) for col in df.columns]
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].map(self._clean_text)
            cleaned[name] = df

        invoices = cleaned["invoices"]
        for col in [
            "invoice_id",
            "vendor",
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
        ]:
            if col not in invoices.columns:
                invoices[col] = None
        invoices["invoice_date"] = pd.to_datetime(invoices["invoice_date"], errors="coerce")
        invoices["due_date"] = pd.to_datetime(invoices["due_date"], errors="coerce")
        invoices["total"] = pd.to_numeric(invoices["total"], errors="coerce")
        invoices["vat"] = pd.to_numeric(invoices["vat"], errors="coerce").fillna(0)
        invoices["withholding_tax"] = pd.to_numeric(invoices["withholding_tax"], errors="coerce")
        invoices["item_count_hint"] = pd.to_numeric(invoices["item_count_hint"], errors="coerce")
        invoices["invoice_id"] = invoices["invoice_id"].map(self._normalize_invoice_id)
        invoices["vendor"] = invoices["vendor"].map(self._clean_numeric_code)
        invoices["amount"] = (invoices["total"].fillna(0) - invoices["vat"].fillna(0)).round(2)
        invoices["vendor_key"] = invoices["vendor"].map(self._text_key)
        invoices["store_key"] = invoices["store"].map(self._text_key)
        invoices["ceco_key"] = invoices["ceco"].map(self._ceco_key)
        invoices["invoice_vendor_key"] = invoices["invoice_id"].map(self._clean_text).fillna("") + "|" + invoices["vendor_key"].fillna("")
        invoices["invoice_pdf_key"] = invoices["invoice_id"].map(self._invoice_key)

        invoice_supports = cleaned["invoice_supports"]
        if not invoice_supports.empty:
            invoice_supports["invoice_key"] = invoice_supports["invoice_id"].map(self._invoice_key)
            support_best = invoice_supports.sort_values(["invoice_key", "flags"], na_position="first").drop_duplicates("invoice_key")
            support_best = support_best.rename(
                columns={
                    "invoice_date": "support_invoice_date",
                    "due_date": "support_due_date",
                    "subtotal": "support_subtotal",
                    "total": "support_total",
                    "detected_iva": "support_detected_iva",
                    "withholding_tax": "support_withholding_tax",
                    "payment_terms_text": "support_payment_terms_text",
                    "item_count_hint": "support_item_count_hint",
                    "flags": "support_flags",
                }
            )
            invoices = invoices.merge(
                support_best[
                    [
                        "invoice_key",
                        "support_invoice_date",
                        "support_due_date",
                        "support_subtotal",
                        "support_total",
                        "support_detected_iva",
                        "support_withholding_tax",
                        "support_payment_terms_text",
                        "support_item_count_hint",
                        "support_flags",
                    ]
                ],
                left_on="invoice_pdf_key",
                right_on="invoice_key",
                how="left",
            )
            invoices["invoice_date"] = invoices["invoice_date"].fillna(invoices["support_invoice_date"])
            invoices["due_date"] = invoices["due_date"].fillna(invoices["support_due_date"])
            invoices["vat"] = invoices["vat"].where(invoices["vat"].fillna(0) != 0, invoices["support_detected_iva"])
            invoices["total"] = invoices["total"].where(invoices["total"].fillna(0) != 0, invoices["support_total"])
            invoices["withholding_tax"] = invoices["withholding_tax"].combine_first(invoices["support_withholding_tax"])
            invoices["payment_terms_text"] = invoices["payment_terms_text"].combine_first(invoices["support_payment_terms_text"])
            invoices["item_count_hint"] = invoices["item_count_hint"].combine_first(invoices["support_item_count_hint"])
            invoices["amount"] = invoices["support_subtotal"].combine_first(invoices["total"].fillna(0) - invoices["vat"].fillna(0)).round(2)
        else:
            invoices["support_flags"] = None

        control = cleaned["control"]
        control["vendor_key"] = control["vendor_code"].fillna(control["vendor"]).map(self._text_key)
        control["store_key"] = control["store"].map(self._text_key)
        control["ceco_key"] = control["ceco"].map(self._ceco_key)
        control["ceco_is_valid"] = control["ceco_key"].notna()
        control = control.sort_values(["ceco_is_valid", "store_key", "vendor_key"], ascending=[False, True, True])

        contracts = cleaned["contracts"]
        contracts["ceco_key"] = contracts["ceco"].map(self._ceco_key)
        contracts["end_of_term"] = pd.to_datetime(contracts["end_of_term"], errors="coerce")
        contracts["rent_min"] = pd.to_numeric(contracts["rent_min"], errors="coerce").fillna(0)

        prorateo = cleaned["prorateo"]
        prorateo["vendor_key"] = prorateo["vendor_code"].fillna(prorateo["vendor"]).map(self._text_key)
        prorateo["store_key"] = prorateo["store"].map(self._text_key)
        prorateo["ceco_key"] = prorateo["ceco"].map(self._ceco_key)
        prorateo["vw_percent"] = pd.to_numeric(prorateo["vw_percent"], errors="coerce")
        prorateo["vw_percent"] = prorateo["vw_percent"].where(prorateo["vw_percent"] <= 1, prorateo["vw_percent"] / 100.0)

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

    def merge_data(self, data: dict[str, pd.DataFrame]) -> pd.DataFrame:
        self._log("Merging invoices with reference tables.")
        invoices = data["invoices"].copy()
        control = data["control"]
        contracts = data["contracts"].sort_values(["ceco_key", "end_of_term"], ascending=[True, False]).drop_duplicates("ceco_key")
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
        prorateo_vendor = prorateo.drop_duplicates(subset=["vendor_key"])[["vendor_key", "vw_percent", "ceco"]].rename(
            columns={"vw_percent": "vw_percent_vendor", "ceco": "prorateo_ceco_vendor"}
        )
        prorateo_store = prorateo.drop_duplicates(subset=["store_key"])[["store_key", "vw_percent", "ceco"]].rename(
            columns={"vw_percent": "vw_percent_store", "ceco": "prorateo_ceco_store"}
        )
        merged = merged.merge(prorateo_vendor, on="vendor_key", how="left")
        merged = merged.merge(prorateo_store, on="store_key", how="left")

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

        merged["mapped_vendor"] = merged["vendor"].map(self._clean_numeric_code)
        merged["mapped_store"] = merged[["store", "control_store_store", "control_store_vendor"]].bfill(axis=1).iloc[:, 0]
        merged["mapped_ceco"] = merged[
            ["macro_ceco", "control_ceco_store", "control_ceco_vendor", "ceco", "prorateo_ceco_store", "prorateo_ceco_vendor"]
        ].bfill(axis=1).iloc[:, 0]
        merged["mapped_ceco"] = merged["mapped_ceco"].where(merged["mapped_ceco"].map(self._is_valid_ceco_value))
        merged["mapped_ceco"] = merged["mapped_ceco"].combine_first(merged["macro_ceco"])
        merged["mapped_ceco_key"] = merged["mapped_ceco"].map(self._ceco_key)

        merged = merged.merge(
            contracts[["ceco_key", "end_of_term", "rent_min", "status_en_rem"]],
            left_on="mapped_ceco_key",
            right_on="ceco_key",
            how="left",
        )
        merged["end_of_term"] = merged["end_of_term"].combine_first(merged["macro_end_of_term"])
        merged["rent_min"] = merged["rent_min"].combine_first(merged["macro_rent_min"])
        merged["status_en_rem"] = merged["status_en_rem"].combine_first(merged["macro_status_en_rem"])
        merged["vw_percent"] = merged["vw_percent_vendor"].combine_first(merged["vw_percent_store"]).combine_first(merged["macro_vw_percent"])
        merged["profit_center"] = merged["macro_profit_center"].combine_first(merged["mapped_ceco"])
        self._log(f"Merging completed. Matched variables for {len(merged)} invoice rows:")
        for idx, row in merged.iterrows():
            self._log(
                f"  Invoice [{row['invoice_id']}]: Vendor={row['vendor']} -> "
                f"Store={row['store']} (Mapped Store={row['mapped_store']}), "
                f"CECO={row['ceco']} (Mapped CECO={row['mapped_ceco']}), "
                f"Contract End={row['end_of_term']} (Rent Min={row['rent_min']}, Status={row['status_en_rem']}), "
                f"Prorateo VW%={row['vw_percent']}, Profit Center={row['profit_center']}"
            )
        return merged

    def apply_rules(self, merged: pd.DataFrame, period: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
        self._log("Applying business rules and creating output rows.")
        df = merged.copy()
        effective_dates = pd.to_datetime(df["invoice_date"], errors="coerce")
        df["contract_active"] = df["end_of_term"].ge(effective_dates)
        df["rent_type"] = np.where(df["contract_active"], "RF", "RV")
        df["account"] = np.where(df["contract_active"], self.FIXED_ACCOUNT, self.VARIABLE_ACCOUNT)
        df["vat_total"] = df["vat"].round(2)
        df["vat_vw"] = (df["vat"] * df["vw_percent"].fillna(0)).round(2)
        df["vat_vq"] = (df["vat"] - df["vat_vw"]).round(2)
        df["withholding_tax_amount"] = pd.to_numeric(df["withholding_tax"], errors="coerce").round(2)
        df["invoice_total"] = pd.to_numeric(df["total"], errors="coerce").round(2)

        period_dates = self._resolve_period(df["invoice_date"], period)
        df["text"] = [
            self._build_text(date_value, rent_type, store, ceco)
            for date_value, rent_type, store, ceco in zip(period_dates, df["rent_type"], df["mapped_store"], df["mapped_ceco"])
        ]

        self._log(f"Evaluated business rules for {len(df)} invoices:")
        for idx, row in df.iterrows():
            self._log(
                f"  Invoice [{row['invoice_id']}]: Date {row['invoice_date']} vs EndOfTerm {row['end_of_term']} "
                f"-> ContractActive={row['contract_active']} -> Class={row['rent_type']} -> Account={row['account']}. "
                f"VAT={row['vat_total']} -> VW={row['vat_vw']} (using VW%={row['vw_percent']}), VQ={row['vat_vq']}. "
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
                "rent_type",
                "amount",
                "invoice_total",
                "vat_total",
                "vat_vw",
                "vat_vq",
                "withholding_tax_amount",
                "text",
                "invoice_date",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
            ]
        ].copy()
        return output_df, validation_df

    def generate_output(
        self,
        output_df: pd.DataFrame,
        validation_df: pd.DataFrame,
        normalized_invoices_df: pd.DataFrame,
        macro_database_df: pd.DataFrame,
        macro_template_path: Path | None = None,
    ) -> PipelineResult:
        self._log("Writing Excel, CSV, and validation outputs.")
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        excel_name = f"output_lucy_ready_{timestamp}.xlsx"
        excel_fixed_name = "output_lucy_ready.xlsx"
        csv_name = f"output_lucy_ready_{timestamp}.csv"
        header_csv_name = f"output_lucy_headers_{timestamp}.csv"
        invoice_bundle_name = f"output_invoice_csv_bundle_{timestamp}.zip"
        manual_style_bundle_name = f"output_manual_style_csv_bundle_{timestamp}.zip"
        validation_name = f"output_validations_{timestamp}.xlsx"
        normalized_name = f"normalized_invoices_{timestamp}.xlsx"

        summary_df, headers_df, lines_df = self._build_lucy_views(output_df)
        self._log(f"Writing {len(summary_df)} summary records to Lucy Excel/CSV:")
        for idx, row in summary_df.iterrows():
            self._log(
                f"  Record [{row.get('invoice_id')}]: Vendor={row.get('vendor')}, "
                f"CECO={row.get('ceco')}, Account={row.get('account')}, "
                f"Amount={row.get('amount')}, VAT={row.get('vat_total')} (VW={row.get('vat_vw')}, VQ={row.get('vat_vq')}), "
                f"Text='{row.get('text')}'"
            )
        mapping_df = self._build_lucy_mapping(summary_df, headers_df, lines_df)
        header_export_df = self._build_header_export(summary_df)
        header_matrix_df = self._build_header_matrix()
        lucy_export_df = self._build_macro_lucy_export(
            summary_df,
            lines_df,
            macro_database_df=macro_database_df,
            macro_template_path=macro_template_path,
        )

        with pd.ExcelWriter(self.output_dir / excel_name, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="summary", index=False)
            header_export_df.to_excel(writer, sheet_name="lucy_header_export", index=False)
            headers_df.to_excel(writer, sheet_name="lucy_headers", index=False)
            lines_df.to_excel(writer, sheet_name="lucy_lines", index=False)
            lucy_export_df.to_excel(writer, sheet_name="lucy_export", index=False)
            mapping_df.to_excel(writer, sheet_name="lucy_mapping", index=False)
            header_matrix_df.to_excel(writer, sheet_name="lucy_header_matrix", index=False)
        self._style_output_workbook(self.output_dir / excel_name)

        try:
            with pd.ExcelWriter(self.output_dir / excel_fixed_name, engine="openpyxl") as writer:
                summary_df.to_excel(writer, sheet_name="summary", index=False)
                header_export_df.to_excel(writer, sheet_name="lucy_header_export", index=False)
                headers_df.to_excel(writer, sheet_name="lucy_headers", index=False)
                lines_df.to_excel(writer, sheet_name="lucy_lines", index=False)
                lucy_export_df.to_excel(writer, sheet_name="lucy_export", index=False)
                mapping_df.to_excel(writer, sheet_name="lucy_mapping", index=False)
                header_matrix_df.to_excel(writer, sheet_name="lucy_header_matrix", index=False)
            self._style_output_workbook(self.output_dir / excel_fixed_name)
        except PermissionError:
            self._log(f"Could not overwrite {excel_fixed_name} because the file is open.")
        lucy_export_df.to_csv(self.output_dir / csv_name, index=False, encoding="utf-8-sig")
        header_export_df.to_csv(self.output_dir / header_csv_name, index=False, encoding="utf-8-sig")
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
        validation_df.to_excel(self.output_dir / validation_name, index=False)
        normalized_snapshot = normalized_invoices_df[
            [
                "invoice_id",
                "vendor",
                "store",
                "ceco",
                "invoice_date",
                "due_date",
                "total",
                "vat",
                "withholding_tax",
                "amount",
                "payment_terms_text",
                "item_count_hint",
                "support_flags",
            ]
        ].copy()
        normalized_snapshot.to_excel(self.output_dir / normalized_name, index=False)

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
            output_excel=f"{run_folder_name}/{excel_name}",
            output_excel_fixed=f"{run_folder_name}/{excel_fixed_name}",
            output_csv=f"{run_folder_name}/{csv_name}",
            header_csv=f"{run_folder_name}/{header_csv_name}",
            invoice_csv_bundle=f"{run_folder_name}/{invoice_bundle_name}",
            manual_style_csv_bundle=f"{run_folder_name}/{manual_style_bundle_name}",
            validation_excel=f"{run_folder_name}/{validation_name}",
            normalized_invoices_excel=f"{run_folder_name}/{normalized_name}",
            output_rows=len(summary_df),
            header_rows=len(header_export_df),
            validation_rows=len(validation_df),
            logs=self.logs.copy(),
        )

    def _load_invoice_supports(self, support_paths: list[Path]) -> pd.DataFrame:
        if not support_paths:
            return pd.DataFrame()
        self._log(f"Parsing {len(support_paths)} invoice support files (PDF/XML).")
        df = InvoiceSupportReader().parse_many(support_paths)
        for _, row in df.iterrows():
            self._log(
                f"  Support File parsed: {row['source_file']} ({row['source_type'].upper()}) -> "
                f"Invoice_ID: {row['invoice_id']}, Date: {row['invoice_date']}, "
                f"Subtotal: {row['subtotal']}, Total: {row['total']}, VAT: {row['detected_iva']}, "
                f"Withholding: {row['withholding_tax']}, Flags: {row['flags']}"
            )
        return df

    def _load_invoices(self, path: Path) -> pd.DataFrame:
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
                    "vendor": ["vendor", "nit", "proveedor", "sap_vendor_code", "vendor_code", "acreedor"],
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
        return df

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
            frame = xl.parse(sheet_name, header=4)
            frame = frame.rename(
                columns=self._best_effort_rename(
                    frame.columns,
                    {
                        "ceco": ["ceco"],
                        "end_of_term": ["end_of_term", "end_of_term_en_virtual_contract"],
                        "rent_min": ["rent_min_rent", "rent_min"],
                        "status_en_rem": ["status_en_rem"],
                    },
                )
            )
            if {"ceco", "end_of_term"} - set(frame.columns):
                continue
            frame["country_sheet"] = sheet_name
            frames.append(frame[[col for col in ["ceco", "end_of_term", "rent_min", "status_en_rem", "country_sheet"] if col in frame.columns]])
        if not frames:
            raise PipelineError("Contracts workbook could not be parsed.")
        return pd.concat(frames, ignore_index=True)

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
                },
            )
        )
        if "vendor_code" in df.columns:
            df["vendor_code"] = df["vendor_code"].map(self._clean_numeric_code)
        return df[[col for col in ["vendor_code", "vendor", "store", "ceco", "vw_percent"] if col in df.columns]].copy()

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

            self._log(
                f"  Invoice [{row['invoice_id']}]: Vendor {row['mapped_vendor']} matched distribution rule. "
                f"Splitting original amount {row['amount']} into {len(valid_matches)} targets:"
            )

            for _, split in valid_matches.iterrows():
                factor = split["split_percent"] / total_pct
                new_row = row.copy()
                new_row["vendor"] = split["target_vendor"]
                new_row["ceco"] = row["mapped_ceco"]
                new_row["amount"] = round((row["amount"] or 0) * factor, 2)
                new_row["invoice_total"] = round((row.get("invoice_total") or 0) * factor, 2)
                new_row["vat_total"] = round((row["vat_total"] or 0) * factor, 2)
                new_row["vat_vw"] = round((row["vat_vw"] or 0) * factor, 2)
                new_row["vat_vq"] = round((row["vat_vq"] or 0) * factor, 2)
                new_row["withholding_tax_amount"] = round((row.get("withholding_tax_amount") or 0) * factor, 2)
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
            if missing_ceco or missing_contract or missing_prorrateo or missing_amount:
                validation_rows.append(
                    {
                        "invoice_id": row.get("invoice_id"),
                        "vendor": row.get("mapped_vendor"),
                        "store": row.get("mapped_store"),
                        "ceco": row.get("mapped_ceco"),
                        "missing_ceco": bool(missing_ceco),
                        "missing_contract": bool(missing_contract),
                        "missing_prorrateo": bool(missing_prorrateo),
                        "missing_amount": bool(missing_amount),
                    }
                )

        return pd.DataFrame(validation_rows)

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
        if not macro_database_df.empty:
            macro_export = self._build_macro_export_from_database(summary_df, macro_database_df)
            if not macro_export.empty:
                export_rows = macro_export.to_dict(orient="records")

        if not export_rows:
            for _, row in lines_df.iterrows():
                export_rows.append(
                    {
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

        export_df = export_df[template_columns].copy()
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
            vat_total = row.get("vat_total")
            amount = row.get("amount")
            invoice_total = row.get("invoice_total")
            withholding_tax_amount = row.get("withholding_tax_amount")
            vat_code = None
            if pd.notna(vat_total):
                vat_code = "V0" if float(vat_total) == 0 else "I1"

            data = {
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

        return header_df.reindex(columns=self.HEADER_EXPORT_COLUMNS)

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
                "fill_status": "partial",
                "source": "PDF / invoice file",
                "logic": "Filled when VAT is detected or explicitly zero; blank if unresolved.",
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
                "fill_status": "partial",
                "source": "Derived rule",
                "logic": "Uses V0 when VAT is zero, I1 when VAT is positive, blank when unresolved.",
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

        invoice_ids = [invoice_id for invoice_id in summary_df["invoice_id"].dropna().astype(str).unique().tolist() if invoice_id]
        for invoice_id in invoice_ids:
            safe_invoice_id = self._safe_filename(invoice_id)
            invoice_dir = bundle_dir / safe_invoice_id
            invoice_dir.mkdir(parents=True, exist_ok=True)

            header_slice = header_export_df[header_export_df["Invoice Number"].astype(str) == str(invoice_id)].copy()
            allocated_slice = lucy_export_df[
                lucy_export_df["Attribuzione$Assignment"].astype(str) == str(invoice_id)
            ].copy()

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

        invoice_ids = [invoice_id for invoice_id in summary_df["invoice_id"].dropna().astype(str).unique().tolist() if invoice_id]
        for invoice_id in invoice_ids:
            safe_invoice_id = self._safe_filename(invoice_id)
            file_path = bundle_dir / f"{safe_invoice_id}_manual_style.csv"

            header_slice = header_export_df[header_export_df["Invoice Number"].astype(str) == str(invoice_id)].copy()
            allocated_slice = lucy_export_df[lucy_export_df["Attribuzione$Assignment"].astype(str) == str(invoice_id)].copy()
            summary_slice = summary_df[summary_df["invoice_id"].astype(str) == str(invoice_id)].copy()

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
                "Partner Bank": self._coalesce_nonblank(raw.get("Partner Bank"), "_"),
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
            return [{"Wht Description": None, "Wht Code": None, "Wht Type": None, "Wht Base Amount": None, "Wht Amount": None}]

        total_withholding = pd.to_numeric(summary_slice["withholding_tax_amount"], errors="coerce").fillna(0).sum()
        taxable_amount = pd.to_numeric(summary_slice["amount"], errors="coerce").fillna(0).sum()
        if total_withholding <= 0:
            return [{"Wht Description": None, "Wht Code": None, "Wht Type": None, "Wht Base Amount": None, "Wht Amount": None}]

        return [
            {
                "Wht Description": "Withholding tax",
                "Wht Code": None,
                "Wht Type": None,
                "Wht Base Amount": self._format_csv_scalar(round(taxable_amount, 2)),
                "Wht Amount": self._format_csv_scalar(round(total_withholding, 2)),
            }
        ]

    def _resolve_manual_header_vat_code(self, raw: pd.Series) -> str | None:
        header_vat_code = raw.get("Vat Code")
        total_vat = raw.get("Total VAT Amount")
        if pd.notna(header_vat_code):
            return header_vat_code
        if pd.notna(total_vat) and float(total_vat) > 0:
            return "VQ"
        if pd.notna(total_vat) and float(total_vat) == 0:
            return "V0"
        return None

    def _resolve_manual_payment_terms(self, value) -> str | None:
        if pd.isna(value) or value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text.isdigit():
            return text.zfill(4)

        normalized = self._clean_text(text)
        term_map = {
            "CREDITO": "0001",
            "CRÉDITO": "0001",
            "CONTADO": "0001",
        }
        return term_map.get(normalized, text)

    def _resolve_manual_payment_mode(self, value) -> str | None:
        if pd.notna(value) and value not in {None, ""}:
            return value
        return "B"

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
                        "Imp$Amount": macro_row.get("amount"),
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
        summary_df["invoice_date"] = pd.to_datetime(summary_df["invoice_date"], errors="coerce")
        summary_df["posting_key"] = summary_df["invoice_id"].fillna("") + "|" + summary_df["vendor"].fillna("")
        summary_df["posting_index"] = summary_df.groupby("invoice_id").cumcount() + 1
        summary_df["posting_count"] = summary_df.groupby("invoice_id")["invoice_id"].transform("size")
        summary_df["document_type"] = "Supplier Invoice"
        summary_df["movement_type"] = "5-FI"
        summary_df["currency"] = "COP"
        summary_df["invoice_total"] = summary_df["invoice_total"].combine_first(summary_df["amount"].fillna(0) + summary_df["vat_total"].fillna(0))
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
                "invoice_date",
                "document_type",
                "movement_type",
                "currency",
                "invoice_total",
                "taxable_total",
                "withholding_tax_amount",
                "store",
                "ceco",
                "profit_center",
                "reference",
                "text",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
                "vat_status",
            ]
        ].copy()

        line_rows: list[dict] = []
        for _, row in summary_df.iterrows():
            base = {
                "invoice_id": row["invoice_id"],
                "vendor": row["vendor"],
                "posting_index": row["posting_index"],
                "posting_count": row["posting_count"],
                "invoice_date": row["invoice_date"],
                "store": row["store"],
                "ceco": row["ceco"],
                "profit_center": row["profit_center"],
                "reference": row["reference"],
                "text": row["text"],
            }

            line_rows.append(
                {
                    **base,
                    "line_type": "expense",
                    "posting_key": "40",
                    "account": row["account"],
                    "tax_code": "VQ",
                    "amount": row["amount"],
                }
            )

            if pd.notna(row["vat_vw"]) and abs(float(row["vat_vw"])) > 0:
                line_rows.append(
                    {
                        **base,
                        "line_type": "vat_vw",
                        "posting_key": "40",
                        "account": self.VAT_ACCOUNT,
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
                        "tax_code": "VQ",
                        "amount": row["vat_vq"],
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
                    "invoice_date",
                    "store",
                    "ceco",
                    "profit_center",
                    "reference",
                    "text",
                    "line_type",
                    "posting_key",
                    "account",
                    "tax_code",
                    "amount",
                ]
            )

        summary_df = summary_df[
            [
                "invoice_id",
                "vendor",
                "store",
                "ceco",
                "account",
                "rent_type",
                "amount",
                "vat_total",
                "vat_vw",
                "vat_vq",
                "withholding_tax_amount",
                "text",
                "invoice_date",
                "due_date",
                "payment_terms_text",
                "item_count_hint",
                "posting_index",
                "posting_count",
                "invoice_total",
                "vat_status",
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
        return re.sub(r"[^A-Z0-9]", "", text.upper()) if text else None

    def _normalize_invoice_id(self, value) -> str | None:
        text = self._clean_text(value)
        if not text:
            return None
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
        text = text.replace(" ", "")
        return text

    def _clean_numeric_code(self, value) -> str | None:
        text = self._clean_text(value)
        if not text:
            return None
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

    def _build_text(self, period_date, rent_type, store, ceco) -> str:
        month_year = pd.to_datetime(period_date).strftime("%m-%Y") if pd.notna(period_date) else "00-0000"
        return f"{month_year} {rent_type or 'RV'} {store or 'NO_STORE'} {ceco or 'NO_CECO'}"

    def _extract_percent(self, label) -> float | None:
        text = self._clean_text(label)
        if not text:
            return None
        match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", text)
        return float(match.group(1).replace(",", ".")) / 100.0 if match else None

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

    def _log(self, message: str) -> None:
        LOGGER.info(message)
        self.logs.append(message)
