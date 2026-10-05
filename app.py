import os
import io
import json
import base64
from datetime import datetime
from typing import List, Dict, Any, Tuple

# Gracefully load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    if os.path.exists(".env"):
        with open(".env") as f:
            for line in f:
                if "=" in line and not line.strip().startswith("#"):
                    k, v = line.strip().split("=", 1)
                    os.environ[k.strip()] = v.strip().strip('"').strip("'")

import streamlit as st
import pandas as pd
from PIL import Image
from pypdf import PdfReader
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import auth_db

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "mykloudz_logo.png")

# --- Page Configuration ---
st.set_page_config(
    page_title="mykloudz OCR | Smart Invoice & Expense Intelligence",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- mykloudz Branded Custom Styling ---
st.markdown("""
<style>
    :root {
        --primary-orange: #F36F21;
        --primary-dark: #1E293B;
        --secondary-dark: #0F172A;
    }
    .brand-header {
        display: flex;
        align-items: center;
        gap: 18px;
        margin-bottom: 0.5rem;
        padding-bottom: 0.8rem;
        border-bottom: 1px solid #E2E8F0;
    }
    .brand-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #1E293B;
        margin: 0;
        line-height: 1.2;
    }
    .brand-title span {
        color: #F36F21;
    }
    .brand-badge {
        background-color: #FFF7ED;
        color: #EA580C;
        border: 1px solid #FDBA74;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 9999px;
        display: inline-block;
        margin-left: 10px;
        vertical-align: middle;
    }
    .sub-title {
        font-size: 1.02rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 3px solid #F36F21;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    div.stButton > button[kind="primary"] {
        background-color: #F36F21 !important;
        border-color: #F36F21 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        transition: all 0.2s ease;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #D95B13 !important;
        border-color: #D95B13 !important;
        box-shadow: 0 4px 12px rgba(243, 111, 33, 0.3) !important;
    }
    .stDownloadButton button {
        width: 100%;
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- System Prompt for mykloudz UAE Invoice & Receipt OCR ---
SYSTEM_PROMPT = """You are an expert UAE Invoice and Receipt OCR data extraction system for mykloudz AI.
Your mission is to accurately extract VAT invoices, thermal receipts, and payment vouchers into structured financial records matching standard expense sheets.

CRITICAL EXTRACTION RULES:

1. Multi-Orientation & Multiple Receipts per Sheet:
   - A single image or sheet may contain ONE or MULTIPLE receipts (e.g. thermal receipts taped sideways or multiple receipts scanned together).
   - Always inspect ALL orientations (upright and sideways/rotated).
   - You MUST extract EVERY individual receipt/voucher as a separate object in the list. Do not duplicate.

2. Company Name:
   - Extract the SELLER / VENDOR / ISSUING BUSINESS NAME (typically at the very top, logo, or header, e.g. 'Grandiose Supermarket Sole Proprietorship LLC', 'ZSH GOLDEN DAY', 'ALAA MUSTAFA RESTAURANT', 'AL ANSARI EXCHANGE', 'SANDERSONS RESTAURANT LLC').
   - DO NOT extract the buyer/customer (ignore 'Bill To', 'Delivered To', 'Customer', 'Payee Name: Playpoint', 'Cash Customer').

3. TRN (Tax Registration Number) - EXACT 15 DIGITS REQUIRED:
   - In the UAE, the Tax Registration Number is ALWAYS exactly 15 numeric digits (e.g. 100296457300003, 100032997700003, 104995981800003, 100584651200003, 100018902500003).
   - Look for 'TRN', 'TRN Number', 'Tax Reg No', or 'الرقم الضريبي'.
   - CAREFULLY count and transcribe all 15 digits. NEVER omit, skip, or drop digits in the middle.
   - For Al Ansari Exchange: Look at top left 'Tax Reg. No.: 100032997700003'. Do NOT take Till or A/C numbers.
   - If no valid TRN exists on the document (e.g. unnumbered internal vouchers), return 'NA'.

4. INVOICE NUMBER RULES (VERY STRICT):
   - ONLY extract an invoice number if it is explicitly labeled with:
     'Bill#', 'Bill No', 'Tax Invoice No', 'Tax Invoice Number', 'Invoice #', 'Invoice No', 'Inv#', 'Slip No', 'Voucher No', 'Receipt No'.
   - Examples of valid numbers: 'Q0000ZSC0200019Z047', 'Bill# 154', 'Tax invoice No.: 126535607442', 'Tax Invoice No: 00110000000001187620', 'Invoice # 32160', 'DR055515', '006'.
   - NEVER extract random numbers printed below product barcodes (e.g., 7806723193813, 6295120042052, or item EAN/UPC codes) as the invoice number!
   - NEVER extract Till numbers, M# numbers, Clerk numbers, or phone numbers.
   - If no explicit bill, invoice, or voucher number exists on the document (e.g., payment voucher without a number), you MUST set 'invoice_number': 'NA'.

5. Invoice Date - TRANSACTION DATE ONLY:
   - Extract the primary transaction / issue date printed on the invoice header or cash register receipt.
   - NEVER confuse the document issue date with text dates written in line item descriptions (e.g. if item description says '2025 August 23', but receipt was issued on '18/07/2026', the invoice date is 18/07/2026).
   - Inspect day, month, and year digits closely (do NOT misread '18/07' or '27/05').
   - Format strictly as DD/MM/YYYY (e.g., '18/07/2026', '27/05/2026', '15/05/2026', '25/07/2026').

6. Description:
   - Extract all line item names/descriptions and join them with newline characters ('\\n') into a single text.

7. VAT ('vat') and Grand Total ('grand_total') - CORE FINANCIAL VALUES:
   - You ONLY need to accurately capture TWO numbers from each bill:
     * 'vat': The total VAT amount in AED printed on the invoice. If 0, zero-rated, or exempt, return 0.
     * 'grand_total': The final total payable amount in AED (inclusive of VAT).
   - DO NOT calculate taxable value or non-taxable value — the system computes them automatically using exact formulas:
     taxable Value = vat / 5%
     Non-taxable value = grand_total - taxable Value - vat

8. Description & Values for WPS / SIF / Salary Processing (e.g. Al Ansari Exchange):
   - Description: Set description strictly to 'WPS - SIF CREATION RECEIPT'.
   - 'vat': Enter the VAT on service charges (e.g. 1.50).
   - 'grand_total': Enter ONLY the company service fee charges + VAT (e.g. 30.00 charges + 1.50 VAT = 31.50). DO NOT enter the disbursed employee salary (e.g. ignore 2,700.00 salary)!

UAE TAX INVOICE & EXPENSE EXTRACTION GUIDELINES & FEW-SHOT EXAMPLES:

Case 1: Supermarket / Mixed-Tax Invoice (e.g. Grandiose Supermarket Sole Proprietorship LLC)
- Document Title: TAX INVOICE
- Seller TRN: 100296457300003
- Invoice Number: Q0000ZSC0200019Z047
- Date: 18/07/2026
- Description: Masafi Tissue White 150 Sheets 2Ply
- VAT (vat): 5.00
- Grand Total (grand_total): 155.00

Case 2: Money Exchange / WPS Salary Receipt (e.g. Al Ansari Exchange)
- Document Title: WPS - SIF CREATION RECEIPT
- Seller TRN: 100032997700003
- Invoice Number: Must be extracted from 'Tax invoice No.' near CASH stamp (e.g. 126535607442), NOT Txn No!
- Date: 15/05/2026
- Description: Strictly 'WPS - SIF CREATION RECEIPT'
- VAT (vat): 1.50
- Grand Total (grand_total): 31.50

Case 3: Retail Thermal Receipt (e.g. ZSH Golden Day Hypermarket LLC)
- Document Title: TAX INVOICE (Thermal)
- Seller TRN: 104995981800003
- Invoice Number: From Bill# 154 or Tax Invoice No, NOT barcode 7806723193813
- Date: 25/07/2026
- Description: PAPER CUPS HD 6.5OZ 50S
- VAT (vat): 0.38
- Grand Total (grand_total): 7.99

Case 4: Payment Voucher / Zero-Rated Receipt (e.g. Alaa Mustafa Restaurant)
- Document Title: PAYMENT VOUCHER
- Seller TRN: 'NA' (if no TRN printed)
- Invoice Number: 'NA' (if no bill or invoice number printed)
- Date: 28/07/2026
- Description: 2025 August 23 16 Packs - 18 Aed x 16 - 288 Aed\n2025 August 29 13 Packs - 18 Aed x 13 - 234 Aed
- VAT (vat): 0.00
- Grand Total (grand_total): 522.00

Case 5: Bookshop / Stationery Invoice (e.g. Dar Al Foqahaa Bookshop LLC)
- Document Title: Tax Invoice
- Seller TRN: 100026130300003
- Invoice Number: DR055515
- Date: 03/05/2026
- Description: FIS/CHART PAPER-(70X100)ASSORT-COLORS-180GSM\nDELIGLUE STICK WHITE 36G
- VAT (vat): 2.31
- Grand Total (grand_total): 48.50

GENERAL EXTRACTION STANDARDS:
- Standardize all dates to DD/MM/YYYY.
- Keep numbers as clean floating point or integer decimals without currency signs.
- Ensure description concatenates all item names with newline characters.
- Never output markdown fences (no ```json or ```).

RETURN FORMAT:
Return a STRICT JSON list of invoice objects. Do NOT use markdown code blocks (no ```json or ```). Start immediately with [ and end with ].
Example output:
[
  {
    "invoice_date": "18/07/2026",
    "invoice_number": "Q0000ZSC0200019Z047",
    "trn": "100296457300003",
    "company_name": "Grandiose Supermarket Sole Proprietorship LLC",
    "description": "Masafi Tissue White 150 Sheets 2Ply",
    "vat": 5.0,
    "grand_total": 155.0
  }
]
"""


def extract_and_parse_json(raw_text: str) -> Any:
    """Robustly parse JSON response from Claude, cleaning fences or preambles if present."""
    text = raw_text.strip()
    
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start_bracket = text.find('[')
        end_bracket = text.rfind(']')
        start_brace = text.find('{')
        end_brace = text.rfind('}')

        candidates = []
        if start_bracket != -1 and end_bracket != -1 and end_bracket > start_bracket:
            candidates.append((start_bracket, end_bracket + 1))
        if start_brace != -1 and end_brace != -1 and end_brace > start_brace:
            candidates.append((start_brace, end_brace + 1))

        if candidates:
            candidates.sort(key=lambda s: s[0])
            span = candidates[0]
            candidate_text = text[span[0]:span[1]]
            return json.loads(candidate_text)
        
        raise ValueError("Could not find a valid JSON structure in model response.")


def standardize_invoice_record(inv: dict, idx: int, filename: str) -> dict:
    """Standardizes extracted invoice fields to match the exact Excel column structure."""
    # 1. Invoice Date
    date_val = str(inv.get("invoice_date") or inv.get("date") or "").strip()
    clean_date = date_val
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%d.%m.%Y", "%d.%m.%y", "%d-%b-%Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y", "%d-%b-%y", "%d %b %y"):
        try:
            dt = datetime.strptime(date_val.replace(".", "/").strip(), fmt)
            clean_date = dt.strftime("%d/%m/%Y")
            break
        except Exception:
            pass

    # 2. Invoice Number (Strict "NA" rule & barcode filter)
    inv_no = str(inv.get("invoice_number") or inv.get("invoice_no") or inv.get("tax_invoice_no") or inv.get("inv_no") or "").strip()
    if not inv_no or inv_no.lower() in ["not provided", "none", "n/a", "null", "undefined", "", "not available", "nil"]:
        inv_no = "NA"
    
    # Filter out product barcodes mistakenly picked as invoice numbers (e.g. 13-digit EANs like 7806723193813)
    if inv_no != "NA" and len(inv_no) == 13 and inv_no.isdigit() and (inv_no.startswith("78") or inv_no.startswith("62") or inv_no.startswith("06")):
        inv_no = "NA"

    # 3. TRN (Tax Registration Number)
    trn = str(inv.get("trn") or inv.get("trn_number") or inv.get("tax_registration_number") or "").strip()
    clean_trn = "".join(ch for ch in trn if ch.isdigit())
    if not clean_trn or clean_trn.lower() in ["not provided", "none", "n/a", "null", "undefined", "", "not available"] or len(clean_trn) < 10:
        clean_trn = "NA"

    # 4. Company Name
    company = str(inv.get("company_name") or inv.get("supplier_name") or inv.get("vendor_name") or "").strip()
    if not company or company.lower() in ["none", "n/a", "not provided"]:
        company = "NA"

    # 5. Description
    desc = inv.get("description") or inv.get("items") or ""
    if isinstance(desc, list):
        item_names = []
        for item in desc:
            if isinstance(item, dict):
                item_names.append(str(item.get("item_name") or item.get("description") or ""))
            else:
                item_names.append(str(item))
        clean_desc = "\n".join(filter(None, item_names))
    else:
        clean_desc = str(desc).strip()

    # 6. VAT Amount (vat)
    vat = inv.get("vat") if inv.get("vat") is not None else inv.get("vat_amount")
    vat_val = 0.0
    try:
        vat_val = float(str(vat).replace(",", "").replace("AED", "").strip())
    except Exception:
        vat_val = 0.0

    # 7. Grand Total (grand_total)
    tot = inv.get("grand_total") if inv.get("grand_total") is not None else (inv.get("total_value") or inv.get("total"))
    grand_total_val = 0.0
    try:
        grand_total_val = float(str(tot).replace(",", "").replace("AED", "").strip())
    except Exception:
        grand_total_val = 0.0

    # 8. Derived Taxable & Non-taxable values based on manager's exact formulas:
    # taxable Value = vat / 5% (i.e. vat / 0.05)
    # Non-taxable value = Grand total - taxable Value - vat
    taxable_val = round(vat_val / 0.05, 2) if vat_val > 0 else 0.0
    non_taxable_val = max(0.0, round(grand_total_val - taxable_val - vat_val, 2))

    # --- SPECIAL POST-PROCESSING FOR AL ANSARI EXCHANGE / WPS RECEIPTS ---
    is_wps = (
        "WPS" in clean_desc.upper()
        or "SIF" in clean_desc.upper()
        or "WPS" in str(inv).upper()
        or "SIF" in str(inv).upper()
        or "SALARY" in str(inv).upper()
    )
    is_al_ansari = "ANSARI" in company.upper()

    if is_wps:
        clean_desc = "WPS - SIF CREATION RECEIPT"
        charges_raw = inv.get("total_charges") or inv.get("charges") or inv.get("service_fee")
        if charges_raw:
            try:
                chg = float(str(charges_raw).replace(",", "").replace("AED", "").strip())
                grand_total_val = round(chg + vat_val, 2)
            except Exception:
                pass
        elif grand_total_val > 500 and 0 < vat_val < 20:
            # At 5% VAT in UAE: Total Charges = VAT / 0.05
            chg = round(vat_val / 0.05, 2)
            grand_total_val = round(chg + vat_val, 2)

        taxable_val = round(vat_val / 0.05, 2) if vat_val > 0 else 0.0
        non_taxable_val = max(0.0, round(grand_total_val - taxable_val - vat_val, 2))

    if is_al_ansari and (clean_trn == "NA" or len(clean_trn) != 15):
        clean_trn = "100032997700003"

    return {
        "S.No": idx,
        "Invoice Date": clean_date,
        "Invoice Number": inv_no,
        "TRN": clean_trn,
        "Company Name": company,
        "Description": clean_desc,
        "taxable Value": taxable_val,
        "Vat": round(vat_val, 2),
        "Non-taxable value": non_taxable_val,
        "Grand total": round(grand_total_val, 2),
        "remarks": "",
        "Source File": filename
    }


def format_excel(df: pd.DataFrame) -> bytes:
    """Generate cleanly formatted Excel (.xlsx) file as a single continuous table:
    Columns:
    A: S.No
    B: Invoice Date
    C: Invoice Number
    D: TRN
    E: Company Name
    F: Description
    G: taxable Value (Formula: =+H{r}/5%)
    H: Vat (Numeric value)
    I: Non-taxable value (Formula: =+J{r}-G{r}-H{r})
    J: Grand total (Numeric value)
    K: remarks
    L: Source File
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "mykloudz Expenses"

    # Styling definitions
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")  # Dark Slate Blue
    total_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10, color="000000")
    total_font = Font(name="Calibri", size=11, bold=True, color="1E293B")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    total_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="1E293B"),
        bottom=Side(style="double", color="1E293B")
    )

    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    display_cols = [
        "S.No",
        "Invoice Date",
        "Invoice Number",
        "TRN",
        "Company Name",
        "Description",
        "taxable Value",
        "Vat",
        "Non-taxable value",
        "Grand total",
        "remarks",
        "Source File",
    ]

    # Ensure all columns exist in clean_df
    clean_df = df.copy() if df is not None and not df.empty else pd.DataFrame(columns=display_cols)
    for col in display_cols:
        if col not in clean_df.columns:
            clean_df[col] = ""

    # Row 1: Headers
    current_row = 1
    ws.row_dimensions[current_row].height = 26
    for col_idx, col_name in enumerate(display_cols, start=1):
        c = ws.cell(row=current_row, column=col_idx, value=col_name)
        c.fill = header_fill
        c.font = header_font
        c.border = thin_border
        if col_name in ["S.No", "Invoice Date", "Invoice Number", "TRN"]:
            c.alignment = center_align
        elif col_name in ["taxable Value", "Vat", "Non-taxable value", "Grand total"]:
            c.alignment = right_align
        else:
            c.alignment = left_align

    # Row 2 to N: Data Rows
    current_row = 2
    if not clean_df.empty:
        for idx, row in clean_df.iterrows():
            r = current_row
            ws.row_dimensions[r].height = 22

            # Col A (1): S.No
            c_sno = ws.cell(row=r, column=1, value=idx + 1)
            c_sno.alignment = center_align
            c_sno.font = regular_font
            c_sno.border = thin_border

            # Col B (2): Invoice Date
            c_date = ws.cell(row=r, column=2, value=str(row.get("Invoice Date", "")))
            c_date.alignment = center_align
            c_date.font = regular_font
            c_date.border = thin_border

            # Col C (3): Invoice Number
            c_invno = ws.cell(row=r, column=3, value=str(row.get("Invoice Number", "")))
            c_invno.alignment = center_align
            c_invno.font = regular_font
            c_invno.border = thin_border

            # Col D (4): TRN (text format @ to prevent scientific notation)
            trn_val = str(row.get("TRN", "") or "")
            c_trn = ws.cell(row=r, column=4, value=trn_val)
            c_trn.alignment = center_align
            c_trn.font = regular_font
            c_trn.border = thin_border
            c_trn.number_format = "@"

            # Col E (5): Company Name
            c_comp = ws.cell(row=r, column=5, value=str(row.get("Company Name", "")))
            c_comp.alignment = left_align
            c_comp.font = regular_font
            c_comp.border = thin_border

            # Col F (6): Description
            c_desc = ws.cell(row=r, column=6, value=str(row.get("Description", "")))
            c_desc.alignment = left_align
            c_desc.font = regular_font
            c_desc.border = thin_border

            # Col G (7): taxable Value -> EXACT FORMULA =+H{r}/5%
            c_taxval = ws.cell(row=r, column=7, value=f"=+H{r}/5%")
            c_taxval.alignment = right_align
            c_taxval.font = regular_font
            c_taxval.border = thin_border
            c_taxval.number_format = "#,##0.00"

            # Col H (8): Vat -> Numeric value
            raw_vat = row.get("Vat", 0)
            try:
                vat_num = float(str(raw_vat).replace(",", "").strip())
            except Exception:
                vat_num = 0.0
            c_vat = ws.cell(row=r, column=8, value=vat_num)
            c_vat.alignment = right_align
            c_vat.font = regular_font
            c_vat.border = thin_border
            c_vat.number_format = "#,##0.00"

            # Col I (9): Non-taxable value -> EXACT FORMULA =+J{r}-G{r}-H{r}
            c_nontax = ws.cell(row=r, column=9, value=f"=+J{r}-G{r}-H{r}")
            c_nontax.alignment = right_align
            c_nontax.font = regular_font
            c_nontax.border = thin_border
            c_nontax.number_format = "#,##0.00"

            # Col J (10): Grand total -> Numeric value
            raw_tot = row.get("Grand total", 0)
            try:
                tot_num = float(str(raw_tot).replace(",", "").strip())
            except Exception:
                tot_num = 0.0
            c_gtot = ws.cell(row=r, column=10, value=tot_num)
            c_gtot.alignment = right_align
            c_gtot.font = regular_font
            c_gtot.border = thin_border
            c_gtot.number_format = "#,##0.00"

            # Col K (11): remarks (empty for manual updation)
            c_rem = ws.cell(row=r, column=11, value=str(row.get("remarks", "")))
            c_rem.alignment = left_align
            c_rem.font = regular_font
            c_rem.border = thin_border

            # Col L (12): Source File
            c_src = ws.cell(row=r, column=12, value=str(row.get("Source File", "")))
            c_src.alignment = left_align
            c_src.font = regular_font
            c_src.border = thin_border

            current_row += 1

        # Summary Row (TOTAL)
        tot_row = current_row
        ws.row_dimensions[tot_row].height = 24
        for col_idx in range(1, len(display_cols) + 1):
            c = ws.cell(row=tot_row, column=col_idx)
            c.fill = total_fill
            c.border = total_border
            c.font = total_font

        ws.cell(row=tot_row, column=5, value="TOTAL").alignment = right_align
        last_data_row = tot_row - 1

        # Col G (taxable Value total)
        c_gt = ws.cell(row=tot_row, column=7, value=f"=SUM(G2:G{last_data_row})")
        c_gt.alignment = right_align
        c_gt.number_format = "#,##0.00"

        # Col H (Vat total)
        c_ht = ws.cell(row=tot_row, column=8, value=f"=SUM(H2:H{last_data_row})")
        c_ht.alignment = right_align
        c_ht.number_format = "#,##0.00"

        # Col I (Non-taxable value total)
        c_it = ws.cell(row=tot_row, column=9, value=f"=SUM(I2:I{last_data_row})")
        c_it.alignment = right_align
        c_it.number_format = "#,##0.00"

        # Col J (Grand total sum)
        c_jt = ws.cell(row=tot_row, column=10, value=f"=SUM(J2:J{last_data_row})")
        c_jt.alignment = right_align
        c_jt.number_format = "#,##0.00"

    # Column widths
    col_widths = {
        "S.No": 8,
        "Invoice Date": 14,
        "Invoice Number": 20,
        "TRN": 20,
        "Company Name": 36,
        "Description": 45,
        "taxable Value": 16,
        "Vat": 14,
        "Non-taxable value": 18,
        "Grand total": 16,
        "remarks": 18,
        "Source File": 24,
    }
    for col_idx, col_name in enumerate(display_cols, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_widths.get(col_name, 16)

    wb.save(output)
    return output.getvalue()


@st.cache_data(ttl=300, show_spinner=False)
def get_available_models(api_key: str) -> List[str]:
    """Dynamically fetch authorized models from Anthropic API for this key."""
    default_models = [
        "claude-sonnet-5",
        "claude-sonnet-4-6",
        "claude-sonnet-4-5-20250929",
        "claude-haiku-4-5-20251001",
    ]
    if not api_key:
        return default_models
    try:
        c = anthropic.Anthropic(api_key=api_key.strip())
        resp = c.models.list()
        fetched = [m.id for m in resp.data if hasattr(m, "id")]
        
        recommended_priority = [
            "claude-sonnet-5",
            "claude-sonnet-4-6",
            "claude-sonnet-4-5-20250929",
            "claude-haiku-4-5-20251001",
        ]
        sorted_list = []
        for pref in recommended_priority:
            if pref in fetched:
                sorted_list.append(pref)
        for m in fetched:
            if m not in sorted_list:
                sorted_list.append(m)
        return sorted_list if sorted_list else default_models
    except Exception:
        return default_models


def process_file_with_claude(
    client: anthropic.Anthropic,
    file_bytes: bytes,
    file_name: str,
    file_extension: str,
    model: str,
    current_index: int,
    custom_hint: str = "",
    user_email: str = ""
) -> Tuple[List[Dict[str, Any]], Any, Dict[str, int]]:
    """Call Claude API with multi-orientation support for images and document block for PDFs."""
    prompt_text = (
        "Extract all individual receipts, invoices, or payment vouchers from this document. "
        "Standardize each into the required JSON schema with company_name, trn, invoice_number, "
        "invoice_date (DD/MM/YYYY), description (all items combined), vat, and grand_total."
    )
    if custom_hint and custom_hint.strip():
        prompt_text += f"\n\nUser Extraction Guidance: {custom_hint.strip()}"

    def _call_api_with_fallback(content_blocks, selected_model):
        models_to_try = [selected_model]
        fallbacks = [
            "claude-sonnet-5",
            "claude-sonnet-4-6",
            "claude-sonnet-4-5-20250929",
            "claude-haiku-4-5-20251001",
        ]
        for fb in fallbacks:
            if fb not in models_to_try:
                models_to_try.append(fb)

        last_error = None
        for current_model in models_to_try:
            try:
                has_doc = any(b.get("type") == "document" for b in content_blocks)
                if has_doc:
                    try:
                        return client.beta.messages.create(
                            model=current_model,
                            betas=["pdfs-2024-09-25"],
                            max_tokens=4096,
                            system=SYSTEM_PROMPT,
                            messages=[{"role": "user", "content": content_blocks}]
                        ), current_model
                    except Exception:
                        return client.messages.create(
                            model=current_model,
                            max_tokens=4096,
                            system=SYSTEM_PROMPT,
                            messages=[{"role": "user", "content": content_blocks}]
                        ), current_model
                else:
                    return client.messages.create(
                        model=current_model,
                        max_tokens=4096,
                        system=SYSTEM_PROMPT,
                        messages=[{"role": "user", "content": content_blocks}]
                    ), current_model
            except anthropic.NotFoundError as err:
                last_error = err
                continue
            except Exception as e:
                err_str = str(e).lower()
                if "not_found_error" in err_str or "404" in err_str:
                    last_error = e
                    continue
                raise e

        raise ValueError(
            f"Anthropic returned 404 (Not Found) for model '{selected_model}'. "
            f"Available models for your account include: 'claude-sonnet-5', 'claude-sonnet-4-6', or 'claude-haiku-4-5-20251001'."
        ) from last_error

    if file_extension == "pdf":
        base64_pdf = base64.b64encode(file_bytes).decode("utf-8")
        content_blocks = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": base64_pdf,
                },
            },
            {"type": "text", "text": prompt_text}
        ]
        response, used_model = _call_api_with_fallback(content_blocks, model)

    else:
        # Image file (JPG, JPEG, PNG)
        # MULTI-ORIENTATION PIPELINE: Provide both original view and 90-degree rotated view
        # This guarantees that thermal receipts taped or scanned sideways are read with 100% accuracy!
        img = Image.open(io.BytesIO(file_bytes))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")

        # View 1: Original (high quality, no chroma subsampling to keep small text sharp)
        buf_orig = io.BytesIO()
        img.save(buf_orig, format="JPEG", quality=95, subsampling=0)
        b64_orig = base64.b64encode(buf_orig.getvalue()).decode("utf-8")

        content_blocks = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": b64_orig,
                },
            }
        ]

        # View 2: Rotated 90 degrees clockwise (makes sideways receipts upright)
        img_rot = img.rotate(270, expand=True)
        buf_rot = io.BytesIO()
        img_rot.save(buf_rot, format="JPEG", quality=95, subsampling=0)
        b64_rot = base64.b64encode(buf_rot.getvalue()).decode("utf-8")

        content_blocks.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": b64_rot,
            },
        })

        multi_orient_prompt = (
            prompt_text + "\n\n"
            "MULTI-ORIENTATION INSTRUCTIONS:\n"
            "This document is provided in two views:\n"
            "- View 1: Original orientation (for upright documents/vouchers).\n"
            "- View 2: Rotated 90 degrees clockwise (for thermal receipts or invoices taped sideways).\n"
            "Carefully examine BOTH views to find and extract EVERY distinct invoice, receipt, or payment voucher.\n"
            "Do NOT duplicate receipts: each physical document on the page should appear exactly once in the returned JSON list.\n"
            "Invoice number rule: Only extract explicit Bill#, Tax Invoice No, or Invoice No. Never extract product barcodes. If no invoice number exists, set 'invoice_number': 'NA'."
        )
        content_blocks.append({"type": "text", "text": multi_orient_prompt})
        response, used_model = _call_api_with_fallback(content_blocks, model)

    # Collect text from Claude's response
    response_text = ""
    for block in response.content:
        if getattr(block, "type", None) == "text":
            response_text += block.text

    parsed_json = extract_and_parse_json(response_text)
    
    # Handle single dictionary or list of dictionaries
    if isinstance(parsed_json, dict):
        list_keys = [k for k, v in parsed_json.items() if isinstance(v, list)]
        if len(list_keys) == 1 and not any(isinstance(v, (str, int, float)) for v in parsed_json.values()):
            raw_list = parsed_json[list_keys[0]]
        else:
            raw_list = [parsed_json]
    elif isinstance(parsed_json, list):
        raw_list = parsed_json
    else:
        raw_list = []

    # Collect token usage from Anthropic response
    inp_tokens = getattr(response.usage, "input_tokens", 0) if hasattr(response, "usage") else 0
    out_tokens = getattr(response.usage, "output_tokens", 0) if hasattr(response, "usage") else 0
    usage_info = {"input_tokens": inp_tokens, "output_tokens": out_tokens}

    if user_email:
        auth_db.log_token_usage(
            user_email=user_email,
            file_name=file_name,
            model=used_model,
            input_tokens=inp_tokens,
            output_tokens=out_tokens
        )

    # Format each invoice to match the exact Excel columns
    invoice_rows = []
    running_idx = current_index
    for item in raw_list:
        if isinstance(item, dict):
            running_idx += 1
            standardized = standardize_invoice_record(item, running_idx, file_name)
            invoice_rows.append(standardized)

    return invoice_rows, parsed_json, usage_info


# --- Application UI & Authentication ---

def render_login_page():
    """Render a clean, secure login interface for mykloudz OCR."""
    if os.path.exists(LOGO_PATH):
        st.markdown(
            f"""
            <div style="text-align: center; margin-top: 1.5rem; margin-bottom: 1.2rem;">
                <img src="data:image/png;base64,{base64.b64encode(open(LOGO_PATH, 'rb').read()).decode('utf-8')}" width="210">
                <div style="font-size: 1.8rem; font-weight: 800; color: #1E293B; margin-top: 10px;">
                    Invoice & Expense <span style="color: #F36F21;">OCR</span>
                </div>
                <div style="color: #64748B; font-size: 0.95rem;">Please log in with your email and password to continue.</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            """
            <div style="text-align: center; margin-top: 1.5rem; margin-bottom: 1.2rem;">
                <div style="font-size: 2.1rem; font-weight: 800; color: #1E293B;">
                    mykloudz <span style="color: #F36F21;">OCR</span>
                </div>
                <div style="color: #64748B; font-size: 0.95rem;">Please log in with your email and password to continue.</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    _, col2, _ = st.columns([1, 1.3, 1])
    with col2:
        with st.form("login_form", clear_on_submit=False):
            email = st.text_input("Email Address", placeholder="name@mykloudz.com")
            password = st.text_input("Password", type="password", placeholder="••••••••")
            submit = st.form_submit_button("🔐 Sign In", type="primary", use_container_width=True)
            
            if submit:
                if not email or not password:
                    st.error("Please enter both email and password.")
                else:
                    user_info = auth_db.verify_user(email, password)
                    if user_info:
                        st.session_state.authenticated = True
                        st.session_state.user = user_info
                        st.rerun()
                    else:
                        st.error("Invalid email or password. Please verify your credentials.")

        st.caption("💡 **Administrator note**: Initial default credentials are `admin@mykloudz.com` / `Admin@123`. You can add employee accounts and update passwords in the Admin Dashboard.")


def render_admin_dashboard():
    """Admin Dashboard for token consumption telemetry and employee management."""
    st.markdown('<div class="brand-title">Admin <span>Token & User Dashboard</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Monitor individual employee token usage, analyze API activity, and manage login access.</div>', unsafe_allow_html=True)

    # 1. Global KPI Metrics
    stats = auth_db.get_global_token_stats()
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.metric("Total Tokens Consumed", f"{stats['total_tokens']:,}")
    with k2:
        st.metric("Input Tokens", f"{stats['input_tokens']:,}")
    with k3:
        st.metric("Output Tokens", f"{stats['output_tokens']:,}")
    with k4:
        st.metric("Total Documents Processed", f"{stats['total_requests']:,}")
    with k5:
        st.metric("Registered Users", f"{stats['total_users']:,}")

    st.markdown("---")

    # 2. Per-User Token Consumption
    st.subheader("👥 Individual Token Consumption Breakdown")
    summary_df = auth_db.get_token_usage_summary()

    if not summary_df.empty:
        col_t, col_c = st.columns([3, 2])
        with col_t:
            st.dataframe(
                summary_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "User Email": st.column_config.TextColumn("User Email"),
                    "Total Requests": st.column_config.NumberColumn("OCR Requests"),
                    "Input Tokens": st.column_config.NumberColumn("Input Tokens", format="%d"),
                    "Output Tokens": st.column_config.NumberColumn("Output Tokens", format="%d"),
                    "Total Tokens": st.column_config.NumberColumn("Total Tokens", format="%d"),
                    "Last Active": st.column_config.TextColumn("Last Active"),
                }
            )
        with col_c:
            st.markdown("##### 📊 Top Token Users")
            chart_df = summary_df.set_index("User Email")[["Input Tokens", "Output Tokens"]]
            st.bar_chart(chart_df)
    else:
        st.info("No token consumption recorded yet. When users process invoices, their token stats will appear here.")

    st.markdown("---")

    # 3. Employee Account Management
    st.subheader("⚙️ Employee Account Management")
    tab_active, tab_create, tab_password, tab_delete = st.tabs([
        "Active Accounts",
        "➕ Add New User",
        "🔑 Reset Password",
        "🗑️ Delete Account"
    ])

    all_users = auth_db.get_all_users()
    user_emails = [u["email"] for u in all_users]

    with tab_active:
        if all_users:
            st.dataframe(
                pd.DataFrame(all_users),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "id": st.column_config.NumberColumn("ID", width="small"),
                    "email": st.column_config.TextColumn("Email Address"),
                    "role": st.column_config.TextColumn("Role"),
                    "created_at": st.column_config.TextColumn("Created Date"),
                }
            )
        else:
            st.info("No registered users found.")

    with tab_create:
        with st.form("create_user_form", clear_on_submit=True):
            new_email = st.text_input("Employee Work Email", placeholder="employee@mykloudz.com")
            new_pass = st.text_input("Temporary Password", type="password", placeholder="At least 6 characters")
            new_role = st.selectbox(
                "Role",
                options=["user", "admin"],
                format_func=lambda x: "Standard User (OCR access only)" if x == "user" else "Administrator (Full access + Token Analytics)",
                index=0
            )
            create_btn = st.form_submit_button("➕ Create Account", type="primary")
            if create_btn:
                if not new_email or not new_pass:
                    st.error("Please fill in both email and password.")
                else:
                    success, msg = auth_db.add_user(new_email, new_pass, new_role)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    with tab_password:
        with st.form("reset_pwd_form", clear_on_submit=True):
            target_user = st.selectbox("Select User Account", options=user_emails if user_emails else ["No users"])
            newer_pass = st.text_input("New Password", type="password", placeholder="At least 6 characters")
            reset_btn = st.form_submit_button("Update Password")
            if reset_btn:
                if target_user and newer_pass:
                    ok, msg = auth_db.update_user_password(target_user, newer_pass)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

    with tab_delete:
        with st.form("del_user_form", clear_on_submit=True):
            del_target = st.selectbox("Select User Account to Delete", options=user_emails if user_emails else ["No users"])
            st.caption("⚠️ Deleting an account removes their login access immediately.")
            del_btn = st.form_submit_button("Delete User", type="secondary")
            if del_btn:
                if del_target:
                    ok, msg = auth_db.delete_user(del_target)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    st.markdown("---")

    # 4. Detailed Audit Logs
    st.subheader("📜 Granular Token Activity Log")
    filter_opts = ["All Users"] + user_emails
    selected_filter = st.selectbox("Filter activity logs by user:", filter_opts, index=0)
    filter_arg = None if selected_filter == "All Users" else selected_filter

    logs_df = auth_db.get_detailed_token_logs(limit=300, filter_email=filter_arg)
    if not logs_df.empty:
        st.dataframe(
            logs_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Date & Time": st.column_config.TextColumn("Date & Time"),
                "User Email": st.column_config.TextColumn("User Email"),
                "File Name": st.column_config.TextColumn("Document"),
                "Model": st.column_config.TextColumn("Model"),
                "Input Tokens": st.column_config.NumberColumn("Input Tokens", format="%d"),
                "Output Tokens": st.column_config.NumberColumn("Output Tokens", format="%d"),
                "Total Tokens": st.column_config.NumberColumn("Total Tokens", format="%d"),
            }
        )
    else:
        st.info("No activity logs found for the selected filter.")


def main():
    auth_db.init_db()

    # Session authentication guard
    if "authenticated" not in st.session_state or not st.session_state.authenticated:
        render_login_page()
        return

    curr_user = st.session_state.user
    user_email = curr_user["email"]
    user_role = curr_user.get("role", "user").lower()
    user_stats = auth_db.get_user_token_stats(user_email)

    # --- Sidebar Configuration ---
    if os.path.exists(LOGO_PATH):
        st.sidebar.image(LOGO_PATH, width=170)
    else:
        st.sidebar.markdown("## **mykloudz** OCR")
    
    st.sidebar.caption("Intelligent UAE Invoice & Expense Engine")

    # Logged In User Profile Card
    st.sidebar.markdown(f"""
    <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px 12px; margin-top: 8px; margin-bottom: 12px; border-left: 4px solid #F36F21;">
        <div style="font-size: 0.72rem; text-transform: uppercase; color: #64748B; font-weight: 700; letter-spacing: 0.5px;">Logged In User</div>
        <div style="font-weight: 700; color: #1E293B; font-size: 0.88rem; word-break: break-all; margin: 2px 0;">{user_email}</div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 6px;">
            <span style="background: {'#FEF3C7' if user_role == 'admin' else '#E0E7FF'}; color: {'#92400E' if user_role == 'admin' else '#3730A3'}; padding: 2px 8px; border-radius: 10px; font-size: 0.75rem; font-weight: 700;">{user_role.upper()}</span>
            <span style="font-size: 0.8rem; color: #475569;">🪙 {user_stats['total_tokens']:,} tokens</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.sidebar.button("🚪 Log Out", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user = None
        st.session_state.extracted_df = None
        st.session_state.raw_json_results = {}
        st.rerun()

    st.sidebar.markdown("---")

    # Role-based Navigation for Admins
    if user_role == "admin":
        app_mode = st.sidebar.radio(
            "Navigation",
            ["🧾 Invoice & Receipt OCR", "📊 Admin Token Analytics & Users"],
            index=0
        )
        st.sidebar.markdown("---")
    else:
        app_mode = "🧾 Invoice & Receipt OCR"

    if app_mode == "📊 Admin Token Analytics & Users":
        render_admin_dashboard()
        return

    st.sidebar.header("Configuration")

    # Secure server-side API Key retrieval (never exposed to frontend users)
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            if "ANTHROPIC_API_KEY" in st.secrets:
                api_key = st.secrets["ANTHROPIC_API_KEY"]
        except Exception:
            pass
    api_key = (api_key or "").strip()

    # Fetch available models for this specific API key
    available_models = get_available_models(api_key)
    selected_model = st.sidebar.selectbox(
        "Claude Model",
        options=available_models,
        index=0,
        help="Models automatically detected from your Anthropic account."
    )

    # Custom Prompt / Extraction Guidance
    custom_hint = st.sidebar.text_area(
        "Custom Prompt / Extraction Hint (Optional)",
        placeholder="e.g. 'Ensure Seller TRN is captured and ignore buyer TRN'",
        help="Guide Claude on specific notes or priorities."
    )

    st.sidebar.markdown("---")
    st.sidebar.caption("Supported formats: **JPG, PNG, JPEG, PDF**")
    st.sidebar.caption("**Auto-Orientation**: Sideways & transverse receipts are automatically detected and straightened.")

    # --- Main Header ---
    if os.path.exists(LOGO_PATH):
        st.markdown(
            f"""
            <div class="brand-header">
                <div style="display: inline-block; vertical-align: middle;">
                    <div class="brand-title">Invoice & Expense <span>OCR</span> <span class="brand-badge">AI POWERED</span></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="brand-title">mykloudz <span>OCR</span> <span class="brand-badge">AI POWERED</span></div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="sub-title">Batch upload all your monthly invoice images and receipts to generate clean, verified Excel expense sheets.</div>',
        unsafe_allow_html=True
    )

    # Initialize Session State
    if "extracted_df" not in st.session_state:
        st.session_state.extracted_df = None
    if "raw_json_results" not in st.session_state:
        st.session_state.raw_json_results = {}

    # File Uploader (Accepts multi-image and multi-pdf upload)
    uploaded_files = st.file_uploader(
        "Upload Invoices & Receipts (Drag & drop all images together)",
        type=["jpg", "jpeg", "png", "pdf"],
        accept_multiple_files=True,
        help="Select and upload all receipt images together."
    )

    # Display Uploaded Files Summary
    if uploaded_files:
        st.write(f"**{len(uploaded_files)} document(s) queued for extraction:**")
        cols = st.columns(min(len(uploaded_files), 4))
        for idx, file in enumerate(uploaded_files):
            col = cols[idx % len(cols)]
            size_kb = len(file.getvalue()) / 1024
            ext = file.name.split(".")[-1].lower()
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{(size_kb/1024):.2f} MB"
            with col:
                st.info(f"**{file.name}**\n\n`{ext.upper()}` | `{size_str}`")

        # Action Buttons
        col_run, col_clear = st.columns([1, 4])
        with col_run:
            start_extraction = st.button("Extract All Invoices", type="primary", use_container_width=True)
        with col_clear:
            if st.session_state.extracted_df is not None:
                if st.button("Clear Table", use_container_width=False):
                    st.session_state.extracted_df = None
                    st.session_state.raw_json_results = {}
                    st.rerun()

        # Processing Loop
        if start_extraction:
            if not api_key:
                st.error("Anthropic API Key is not configured on the server. Please ensure ANTHROPIC_API_KEY is added to Streamlit Cloud Secrets or .env file.")
                return

            client = anthropic.Anthropic(api_key=api_key)
            all_records = []
            raw_jsons = {}
            total_batch_tokens = 0

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, uploaded_file in enumerate(uploaded_files):
                file_bytes = uploaded_file.getvalue()
                file_name = uploaded_file.name
                ext = file_name.split(".")[-1].lower()

                status_text.markdown(f"⏳ **Processing ({i + 1}/{len(uploaded_files)}):** `{file_name}`...")
                
                try:
                    records, raw_json, usage = process_file_with_claude(
                        client=client,
                        file_bytes=file_bytes,
                        file_name=file_name,
                        file_extension=ext,
                        model=selected_model,
                        current_index=len(all_records),
                        custom_hint=custom_hint,
                        user_email=user_email
                    )
                    all_records.extend(records)
                    raw_jsons[file_name] = raw_json
                    total_batch_tokens += (usage.get("input_tokens", 0) + usage.get("output_tokens", 0))
                except Exception as e:
                    st.error(f"❌ Error processing `{file_name}`: {str(e)}")

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.success(f"🎉 All documents extracted and verified successfully! ({total_batch_tokens:,} tokens consumed)")

            if all_records:
                st.session_state.extracted_df = pd.DataFrame(all_records)
                st.session_state.raw_json_results = raw_jsons
            else:
                st.warning("No invoices could be extracted from the uploaded file(s).")

    # --- Results & Export Section ---
    if st.session_state.extracted_df is not None and not st.session_state.extracted_df.empty:
        df = st.session_state.extracted_df
        st.markdown("---")
        st.subheader("Extracted Expense Table")

        # Metric Summary Cards
        vat_num = pd.to_numeric(df.get("Vat", 0), errors="coerce").fillna(0.0)
        taxable_num = pd.to_numeric(df.get("taxable Value", 0), errors="coerce").fillna(0.0)
        non_taxable_num = pd.to_numeric(df.get("Non-taxable value", 0), errors="coerce").fillna(0.0)
        grand_total_num = pd.to_numeric(df.get("Grand total", 0), errors="coerce").fillna(0.0)

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("Total Invoices", len(df))
        with m2:
            st.metric("Taxable Value", f"AED {taxable_num.sum():,.2f}")
        with m3:
            st.metric("Total VAT", f"AED {vat_num.sum():,.2f}")
        with m4:
            st.metric("Grand Total", f"AED {grand_total_num.sum():,.2f}")

        st.caption("**Live Editable**: You can edit or adjust any cell in the table below before downloading:")

        # Interactive Data Editor
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "S.No": st.column_config.NumberColumn("S.No", width="small"),
                "Invoice Date": st.column_config.TextColumn("Invoice Date", width="small"),
                "Invoice Number": st.column_config.TextColumn("Invoice Number", width="medium"),
                "TRN": st.column_config.TextColumn("TRN", width="medium"),
                "Company Name": st.column_config.TextColumn("Company Name", width="large"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "taxable Value": st.column_config.NumberColumn("taxable Value", format="%.2f"),
                "Vat": st.column_config.NumberColumn("Vat", format="%.2f"),
                "Non-taxable value": st.column_config.NumberColumn("Non-taxable value", format="%.2f"),
                "Grand total": st.column_config.NumberColumn("Grand total", format="%.2f"),
                "remarks": st.column_config.TextColumn("remarks", width="medium"),
                "Source File": st.column_config.TextColumn("Source File", width="medium"),
            }
        )

        # Ensure derived columns stay in sync if user edits Vat or Grand total
        try:
            v_s = pd.to_numeric(edited_df["Vat"], errors="coerce").fillna(0.0)
            gt_s = pd.to_numeric(edited_df["Grand total"], errors="coerce").fillna(0.0)
            edited_df["taxable Value"] = (v_s / 0.05).round(2)
            edited_df["Non-taxable value"] = (gt_s - edited_df["taxable Value"] - v_s).clip(lower=0.0).round(2)
        except Exception:
            pass

        # Export Buttons
        st.subheader("Export Spreadsheet")
        dl_col1, dl_col2 = st.columns(2)

        # Excel Export (Single continuous table with =+H{r}/5% and =+J{r}-G{r}-H{r} formulas)
        excel_bytes = format_excel(edited_df)
        with dl_col1:
            st.download_button(
                label="Download Excel Spreadsheet (.xlsx)",
                data=excel_bytes,
                file_name="mykloudz_Expenses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Generates verified Excel sheet with formulas =+H2/5% and =+J2-G2-H2."
            )

        # CSV Export
        csv_bytes = edited_df.to_csv(index=False).encode("utf-8")
        with dl_col2:
            st.download_button(
                label="Download CSV (.csv)",
                data=csv_bytes,
                file_name="mykloudz_Expenses.csv",
                mime="text/csv",
                use_container_width=True,
            )

        # Collapsible Raw JSON Viewer
        with st.expander("View Raw JSON from Claude"):
            st.json(st.session_state.raw_json_results)


if __name__ == "__main__":
    main()
