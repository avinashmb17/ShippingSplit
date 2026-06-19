import streamlit as st
import pandas as pd
import pdfplumber
import re
import os

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from io import BytesIO


# =====================================================
# HELPERS
# =====================================================

def clean(x):
    return str(x).replace(".0", "").strip()


def safe_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))


# =====================================================
# PDF TEXT CLEAN
# =====================================================

def clean_pdf_text(text):
    text = text.upper()
    text = re.sub(r'CHB NAME.*', '', text)
    text = re.sub(r'PAGE\s*\d+.*', '', text)
    text = re.sub(r'\n\s*\n', '\n', text)
    return text


# =====================================================
# COMPANY DETECTION
# =====================================================

def detect_company(text):

    if not text:
        return "UNKNOWN"

    text = text.upper()
    text = re.sub(r'\s+', ' ', text)

    if re.search(r'BOM\s*:?\s*\d+', text):
        return "MALCA"

    if re.search(r'IGM/S\.?B\.?\s*NO\.?\s*:?\s*\d+', text):
        return "JK"

    
    if (
    "BVC BRINK'S" in text
    or "BVC BRINKS" in text
    or "BVC BRINK'S DIAMOND" in text
    or "BVC BRINKS DIAMOND" in text
    ):
        return "BVC"
    
    if "BHARAT DIAMOND BOURSE" in text:
        return "BDB"
    
    #if "BVC BRINK'S" in text or "BVC BRINKS" in text:
    #    return "BVC"

    return "UNKNOWN"


# =====================================================
# NUMBER EXTRACTION
# =====================================================

def extract_sb_number(text):

    if not text:
        return None

    text = text.upper()
    text = re.sub(r'\s+', ' ', text)

    patterns = [
        r'BOM\s*:?\s*(\d+)',
        r'IGM/S\.?B\.?\s*NO\.?\s*:?\s*(\d+)',
        r'SB\s*NO\.?\s*:?\s*(\d+)'
    ]

    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1)

    return None

def extract_haw_no(text):

    if not text:
        return None

    text = text.upper()
    text = re.sub(r'\s+', ' ', text)

    patterns = [
        r'BOM\s*:?\s*(\d+)',
        r'HAWB\s*:?\s*(\d+)',
        r'HAWB\s*NO\.?\s*:?\s*([A-Z0-9\-\/]+)',
        r'HAW[A]?\s*NO\.?\s*:?\s*([A-Z0-9\-\/]+)'
    ]

    for pattern in patterns:
        m = re.search(pattern, text)

        if m:
            return m.group(1)

    return None

def extract_invoice_no(text):

    if not text:
        return None

    text = text.upper()

    patterns = [
        r'INVOICE\s*NO\.?\s*:?\s*([A-Z0-9\-\/]+)',
        r'INV\s*NO\.?\s*:?\s*([A-Z0-9\-\/]+)'
    ]

    for pattern in patterns:
        m = re.search(pattern, text)

        if m:
            return m.group(1)

    return None
# =====================================================
# COMMODITY
# =====================================================

def normalize(text):

    if not text:
        return ""

    text = str(text).upper().strip()
    text = re.sub(r'[^A-Z0-9 ]', ' ', text)
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


def extract_commodity(text):

    if not text:
        return None

    text = text.upper()
    text = re.sub(r'\s+', ' ', text)

    patterns = [
        r'COMMODITY\s*:?\s*([A-Z0-9 \-/&,\.]+)',
        r'SAID[- ]TO[- ]CONTAIN\s*:?\s*([A-Z0-9 \-/&,\.]+)',
        r'CONTAINING\s*:?\s*([A-Z0-9 \-/&,\.]+)'
    ]

    for pattern in patterns:

        m = re.search(pattern, text, re.IGNORECASE)

        if m:

            value = m.group(1)

            value = re.split(
                r'WEIGHT|PCS|PIECES|VALUE|AMOUNT|QTY|QUANTITY|MARKS|PACKAGE|TOTAL',
                value,
                flags=re.IGNORECASE
            )[0].strip()

            return normalize(value)

    return None


# =====================================================
# COST CENTER
# =====================================================

def get_cost_center(df_cc, commodity):

    if not commodity:
        return None

    commodity = normalize(commodity)

    commodity_dict = {
        normalize(str(r.get("Description", ""))):
        str(r.get("Cost Center", "")).strip()
        for _, r in df_cc.iterrows()
    }

    for desc, cc in commodity_dict.items():

        if desc and desc in commodity:
            return cc

    return None


# =====================================================
# PDF HEADER
# =====================================================

def add_cost_center_header(page, cost_center):

    packet = BytesIO()

    width = float(page.mediabox.width)
    height = float(page.mediabox.height)

    can = canvas.Canvas(packet, pagesize=(width, height))

    can.setFont("Helvetica-Bold", 10)

    can.drawString(
        20,
        height - 20,
        f"COST CENTER : {cost_center}"
    )

    can.save()

    packet.seek(0)

    overlay_pdf = PdfReader(packet)
    overlay_page = overlay_pdf.pages[0]

    page.merge_page(overlay_page)

    return page


# =====================================================
# EXCEL LOOKUP
# =====================================================

def get_filename_from_excel(df, value, column):

    value = clean(value)

    df[column] = df[column].astype(str).apply(clean)

    matches = df[
        df[column].astype(str).str.strip() ==
        str(value).strip()
    ]

    if matches.empty:
        return None

    so_values = []
    inv_values = []

    for _, row in matches.iterrows():

        so = str(row.get("SO NO", "")).strip()
        inv = str(row.get("Invoice No", "")).strip()

        if so and so.lower() != "nan":
            so_values.append(so)

        if (
            (not so or so.lower() == "nan")
            and inv
            and inv.lower() != "nan"
        ):
            inv_values.append(inv)

    if so_values:

        base = "SO_" + "-".join(
            sorted(set(so_values))
        )

        if inv_values:
            base += "_" + "-".join(
                sorted(set(inv_values))
            )

        return base

    if inv_values:
        #return "INV_" + "-".join(
        return "SO_" + "-".join(
            sorted(set(inv_values))
        )

    return None


# =====================================================
# MAIN PROCESS
# =====================================================
#report_rows = []

#report_rows = []

def split_pdf_and_search_excel(
    uploaded_files,
    excel_file,
    base_path
):

    global report_rows
    report_rows = []

    output_dir = os.path.join(base_path, "OUTPUT")
    os.makedirs(output_dir, exist_ok=True)

    df = pd.read_excel(excel_file, sheet_name="Sheet1")
    df.columns = df.columns.str.strip()

    df_cc = pd.read_excel(excel_file, sheet_name="Commodity")
    df_cc.columns = df_cc.columns.str.strip()

    for pdf_file in uploaded_files:

        reader = PdfReader(pdf_file)

        with pdfplumber.open(pdf_file) as pdf:

            for page_num, page in enumerate(pdf.pages):

                text = clean_pdf_text(page.extract_text() or "")

                st.write("COMPANY:", detect_company(text))
                st.text(text[:5000])

                company = detect_company(text)

                sb_no = extract_sb_number(text)
                haw_no = extract_haw_no(text)
                inv_no = extract_invoice_no(text)
                commodity = extract_commodity(text)

                base_name = None
                cost_center = None

                # -----------------------------
                # COMPANY LOGIC
                # -----------------------------

                if company == "MALCA":

                    if haw_no:
                        base_name = get_filename_from_excel(df, haw_no, "HAW No")

                elif company == "BVC":

                    if haw_no:
                        base_name = haw_no

                    if not base_name:
                        m = re.search(r'HAWB\s*:?\s*(\d+)', text, re.IGNORECASE)
                        if m:
                            base_name = m.group(1)

                else:

                    if sb_no:
                        base_name = get_filename_from_excel(df, sb_no, "SB No")

                # -----------------------------
                # COST CENTER
                # -----------------------------
                if company != "MALCA":
                    cost_center = get_cost_center(df_cc, commodity)

                # -----------------------------
                # BDB RULE
                # -----------------------------
                if company == "BDB" and not base_name and inv_no:
                    base_name = inv_no

                if not base_name:
                    base_name = f"PAGE_{page_num + 1}"

                filename = safe_filename(f"{base_name}_{company}")

                writer = PdfWriter()
                pdf_page = reader.pages[page_num]

                if company != "MALCA" and cost_center:
                    pdf_page = add_cost_center_header(pdf_page, cost_center)

                writer.add_page(pdf_page)

                output_file = os.path.join(output_dir, f"{filename}.pdf")
                split_output_file = str.split(f"{filename}",sep="_")[1]
                
                st.write(split_output_file)
                counter = 1
                original = output_file
                
                while os.path.exists(output_file):
                    output_file = original.replace(".pdf", f"_{counter}.pdf")
                    counter += 1

                with open(output_file, "wb") as f:
                    writer.write(f)

                st.success(f"Saved: {os.path.basename(output_file)}")

                # =====================================================
                #  EXCEL REPORT (FIXED LOCATION)
                # =====================================================
                charge_rows = extract_charges(company, text)

                if not charge_rows:

                    report_rows.append({
                        "File Name": f"{filename}.pdf",
                        "File Path": output_file,
                        "SO_NO":split_output_file,
                        "Company": company,
                        "SO No": sb_no,
                        "Invoice No": inv_no,
                        "Charge Type": "",
                        "Charges": "",
                        "CGST": "",
                        "SGST": "",
                        "IGST": ""
                    })

                else:

                    for r in charge_rows:

                        report_rows.append({
                            "File Name": f"{filename}.pdf",
                            "File Path": output_file,
                            "SO_NO":split_output_file,
                            "Company": company,
                            "SO No": sb_no,
                            "Invoice No": inv_no,
                            "Charge Type": r["Charge Type"],
                            "Charges": r["Charges"],
                            "CGST": r["CGST"],
                            "SGST": r["SGST"],
                            "IGST": r["IGST"]
                        })

    # =====================================================
    # FINAL EXCEL EXPORT (AFTER LOOP)
    # =====================================================

    report_df = pd.DataFrame(report_rows)

    output_excel = os.path.join(output_dir, "PDF_REPORT.xlsx")

    report_df.to_excel(output_excel, index=False)

    st.success(f"Excel Saved: {output_excel}")
def extract_charges(company, text):

    rows = []

    # -----------------------------
    # BVC (FIXED: no duplicates, clean mapping)
    # -----------------------------
    if company == "BVC":

        rows = []

        # normalize
        t = re.sub(r'\s+', ' ', text.upper())

        # 🔥 STEP 1: extract ALL service rows for 996799 / 998599
        pattern = re.findall(
            r'(AIRLINE DELIVERY CHARGES|DELIVERY CHARGES|ACC CHARGES)\s+(996799|998599)\s+([\d,]+\.\d{2})\s+9%\s+([\d,]+\.\d{2})\s+9%\s+([\d,]+\.\d{2})',
            t
        )

        seen = set()

        for label, code, amount, cgst, sgst in pattern:

            key = (label, amount)

            #  BLOCK DUPLICATES
            if key in seen:
                continue
            seen.add(key)

            # normalize label
            if "AIRLINE" in label:
                final_label = "AIRLINE DELIVERY"
            elif "ACC" in label:
                final_label = "ACC CHARGES"
            else:
                final_label = "DELIVERY"

            rows.append({
                "Charge Type": final_label,
                "Charges": amount.replace(",", ""),
                "CGST": cgst,
                "SGST": sgst,
                "IGST": ""
            })

        return rows
    # -----------------------------
    # JK
    # -----------------------------
    elif company == "JK":

        p = re.search(
            r'TOTAL\s*:\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})',
            text,
            re.IGNORECASE
        )

        if p:
            rows.append({
                "Charge Type": "TOTAL",
                "Charges": p.group(1),
                "CGST": p.group(2),
                "SGST": p.group(3),
                "IGST": p.group(4)
            })

    # -----------------------------
    # MALCA (FIXED + added FREIGHT VALUATION)
    # -----------------------------
    elif company == "MALCA":

        matches = re.findall(
            r'(FREIGHT(?:\s+VALUATION|\s+OTHER)?)\s+996531\s+([\d,]+\.\d{2})\s+18\.00\s+([\d,]+\.\d{2})',
            text,
            re.IGNORECASE
        )

        for m in matches:
            rows.append({
                "Charge Type": m[0],
                "Charges": m[1],
                "CGST": "",
                "SGST": "",
                "IGST": m[2]
            })

    # -----------------------------
    # BDB (FIXED parsing)
    # -----------------------------
    elif company == "BDB":

        p = re.search(
            r'996719\s+([\d,]+\.\d{2})\s+0\.00\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})',
            text,
            re.IGNORECASE
        )

        if p:
            rows.append({
                "Charge Type": "996719",
                "Charges": p.group(1),
                "CGST": p.group(2),
                "SGST": p.group(3),
                "IGST": ""
            })

    return rows

# =====================================================
# STREAMLIT UI
# =====================================================

st.set_page_config(
    page_title="PDF Splitter",
    layout="wide"
)

uploaded_files = st.sidebar.file_uploader(
    "Choose PDF Files",
    type=["pdf"],
    accept_multiple_files=True
)

excel_file = st.sidebar.file_uploader(
    "Choose Excel File",
    type=["xlsx", "xls"]
)

if uploaded_files and excel_file:

    base_path = st.text_input(
        "Enter Base Folder Path (Example: Z:\\PDF\\)"
    )

    if st.button("Split PDF"):

        if not base_path:
            st.error(
                "Please enter a base folder path like Z:\\PDF\\"
            )

        else:

            split_pdf_and_search_excel(
                uploaded_files,
                excel_file,
                base_path
            )

            st.success(
                "Completed Successfully"
            )

else:

    st.info(
        "Upload PDF and Excel file"
    )

