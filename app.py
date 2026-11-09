from datetime import datetime
from io import BytesIO
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. PAGE CONFIGURATION & PROFESSIONAL DEPED CSS
# ==========================================
st.set_page_config(page_title="District Book Inventory", layout="wide")

st.markdown("""
    <style>
    /* Main Background */
    .stApp {
        background-color: #F8FAFC !important;
    }

    /* Dark Navy Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #0F172A !important;
    }
    [data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    /* Main Page Select Box & Inputs Styling */
    div[data-baseweb="select"] > div, 
    div[data-baseweb="input"] > div,
    textarea {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }

    /* Force Selectbox & Input Text to Crisp Navy */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div,
    input, textarea {
        color: #0F172A !important;
        font-weight: 500 !important;
    }

    /* Input & Select Labels */
    .stTextInput label, .stDateInput label, .stTextArea label, .stSelectbox label {
        color: #1E293B !important;
        font-weight: 700 !important;
    }

    /* Metric Cards Styling */
    [data-testid="stMetricValue"] {
        color: #1D4ED8 !important;
        font-weight: 800 !important;
    }
    
    [data-testid="stMetricLabel"] {
        color: #475569 !important;
        font-weight: 600 !important;
    }

    /* Primary Buttons (DepEd Royal Blue) */
    .stButton > button[kind="primary"] {
        background-color: #1D4ED8 !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #1E40AF !important;
    }

    /* Info Notice Box */
    .stAlert {
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CONNECT TO GOOGLE SHEETS & PRESETS
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data(worksheet_name):
    try:
        return conn.read(worksheet=worksheet_name, ttl=10)
    except Exception as e:
        st.warning(f"⚠️ Unable to fetch worksheet '{worksheet_name}'. Please verify the tab name matches your Google Sheet and permissions are set properly.")
        return pd.DataFrame()

SCHOOL_LIST = [
    "Arcaflor Maniapao ES",
    "Balabag ES",
    "Casildo B. Nonol Sr. ES",
    "Colorado ES",
    "Damñas ES",
    "Digos City Central ES",
    "Domingo Abawag ES",
    "Dulangan ES",
    "Federico Alferez ES",
    "Jolencio R. Alberca ES",
    "Lungag ES",
    "Mahayahay ES",
    "Pedro Basalan ES",
    "Ranao ES",
    "Remedios N. Saplala ES",
    "Ruparan ES",
]

# --- SIDEBAR CONTROLS ---
st.sidebar.title("Navigation")
role = st.sidebar.radio("Select View:", ["Principal View", "Custodian View"])

# ==========================================
# 3. HEADER, PROMINENT SCHOOL SELECTOR & METRICS
# ==========================================
st.title("📚 District Book Inventory Tracker")

if role == "Custodian View":
    selected_school = None
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="🏫 Total Schools", value=len(SCHOOL_LIST))
    with col2:
        st.metric(label="📦 Total Pending Dispatches", value="12")
    with col3:
        st.metric(label="✅ Completed Orders", value="148")

else:
    # --- PROMINENT SCHOOL SELECTOR ON MAIN PAGE ---
    selected_school = st.selectbox(
        "🏫 **Select Your School to View Inventory & ICS:**", 
        SCHOOL_LIST,
        index=0
    )

    # --- PRINCIPAL VIEW: Property Accountability Metrics ---
    school_df = load_data("school_inventory")
    master_df = load_data("master_inventory")

    # Filter records for selected school
    if not school_df.empty and "school_name" in school_df.columns:
        school_records = school_df[school_df["school_name"].astype(str).str.strip() == selected_school.strip()]
    else:
        school_records = pd.DataFrame()

    if not school_records.empty and not master_df.empty:
        # Case-insensitive merge for price calculation
        school_records["book_title_clean"] = school_records["book_title"].astype(str).str.strip().str.lower()
        cols_to_merge = ["book_title", "unit_cost"] if "unit_cost" in master_df.columns else ["book_title"]
        master_merge = master_df[cols_to_merge].copy()
        master_merge["book_title_clean"] = master_merge["book_title"].astype(str).str.strip().str.lower()
        master_merge = master_merge.drop(columns=["book_title"])

        merged_records = school_records.merge(master_merge, on="book_title_clean", how="left").drop(columns=["book_title_clean"])
        
        if "unit_cost" not in merged_records.columns:
            merged_records["unit_cost"] = 90.00
        else:
            merged_records["unit_cost"] = merged_records["unit_cost"].fillna(90.00)
    else:
        merged_records = school_records
        if not merged_records.empty:
            merged_records["unit_cost"] = 90.00

    if not merged_records.empty and "status" in merged_records.columns:
        # 1. Pending Dispatches/ICS Batches
        pending_items = merged_records[merged_records["status"].astype(str).str.lower() == "for release"]
        pending_count = len(pending_items)
        
        # 2. Total Accountable Value for pickup (Qty * Price)
        pending_value = (pending_items["quantity_received"] * pending_items["unit_cost"]).sum() if not pending_items.empty else 0.0
        
        # 3. Total Received Value (Completed Items)
        received_items = merged_records[merged_records["status"].astype(str).str.lower() == "received"]
        received_value = (received_items["quantity_received"] * received_items["unit_cost"]).sum() if not received_items.empty else 0.0
    else:
        pending_count = 0
        pending_value = 0.0
        received_value = 0.0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            label="📋 Pending ICS Dispatches", 
            value=f"{pending_count} Batches",
            help="Number of book titles ready for pickup at the District Office"
        )
    with col2:
        st.metric(
            label="💵 Total Accountable Value", 
            value=f"₱{pending_value:,.2f}",
            help="Total monetary value of ready materials requiring signature on the ICS form"
        )
    with col3:
        st.metric(
            label="🏛️ Accountable Value Received", 
            value=f"₱{received_value:,.2f}",
            help="Total value of books previously collected and acknowledged"
        )

st.divider()

# ==========================================
# 4. FORMAL DEPED ICS EXCEL GENERATOR (WITH NO. OF LM, DR & IAR)
# ==========================================
def generate_official_deped_ics_excel(school_name, date_str, df_items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ICS Form"

    ws.views.sheetView[0].showGridLines = True

    font_header = Font(name="Calibri", size=10)
    font_title = Font(name="Calibri", size=14, bold=True)
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border_side = Side(border_style="thin", color="000000")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # 1. Header (Rows 1-5 across 10 columns A to J)
    headers = [
        "Republic of the Philippines",
        "Department of Education",
        "Region XI – Davao Region",
        "City Schools Division of Digos",
        "Digos Occidental District"
    ]
    for r_idx, text in enumerate(headers, start=1):
        ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=10)
        cell = ws.cell(row=r_idx, column=1, value=text)
        cell.font = font_header
        cell.alignment = align_center

    # 2. Section Title (Row 7 across A7:J7)
    ws.merge_cells("A7:J7")
    title_cell = ws.cell(row=7, column=1, value="INVENTORY CUSTODIAN SLIP (ICS)")
    title_cell.font = font_title
    title_cell.alignment = align_center

    # 3. Metadata Section (Rows 11-12)
    ws.cell(row=11, column=1, value=f"Entity Name: {school_name}").font = font_bold
    ws.cell(row=12, column=1, value="INVENTORY CUSTODIAN SLIP (ICS) No.: ____________________").font = font_bold

    # 4. Table Column Headers (Row 14 - 10 Columns)
    table_headers = [
        "Quantity", "Unit", "Unit Cost", "Total Cost", "Description", 
        "No. of LM", "DR No.", "IAR No.", "Inventory Item No.", "Estimated Useful Life"
    ]
    ws.row_dimensions[14].height = 28
    
    for c_idx, h_text in enumerate(table_headers, start=1):
        cell = ws.cell(row=14, column=c_idx, value=h_text)
        cell.font = font_bold
        cell.alignment = align_center
        cell.border = thin_border

    # 5. Populate Data Rows (Starting at Row 15)
    current_row = 15
    for _, row in df_items.iterrows():
        qty = int(row.get("quantity_received", 0))
        unit_cost = float(row.get("unit_cost", 90.00)) if "unit_cost" in row else 90.00
        desc = str(row.get("book_title", ""))
        no_of_lm = str(row.get("no_of_lm", qty)) if "no_of_lm" in row and pd.notna(row.get("no_of_lm")) else qty
        dr_no = str(row.get("dr_no", "")) if "dr_no" in row and pd.notna(row.get("dr_no")) else ""
        iar_no = str(row.get("iar_no", "")) if "iar_no" in row and pd.notna(row.get("iar_no")) else ""
        useful_life = int(row.get("useful_life", 3))

        ws.cell(row=current_row, column=1, value=qty).alignment = align_center
        ws.cell(row=current_row, column=2, value="PCS").alignment = align_center
        
        c3 = ws.cell(row=current_row, column=3, value=unit_cost)
        c3.number_format = '#,##0.00'
        c3.alignment = align_right

        # Calculated Total Cost
        total_cost = qty * unit_cost
        c4 = ws.cell(row=current_row, column=4, value=total_cost)
        c4.number_format = '#,##0.00'
        c4.alignment = align_right

        ws.cell(row=current_row, column=5, value=desc).alignment = align_left
        ws.cell(row=current_row, column=6, value=no_of_lm).alignment = align_center  # No. of LM
        ws.cell(row=current_row, column=7, value=dr_no).alignment = align_center     # DR No.
        ws.cell(row=current_row, column=8, value=iar_no).alignment = align_center    # IAR No.
        ws.cell(row=current_row, column=9, value="").alignment = align_center         # Inventory Item No.
        ws.cell(row=current_row, column=10, value=useful_life).alignment = align_center

        for col in range(1, 11):
            ws.cell(row=current_row, column=col).border = thin_border
            ws.cell(row=current_row, column=col).font = font_regular

        current_row += 1

    # Fill blank padded rows to keep standard sheet length
    target_end_row = max(current_row + 3, 26)
    for r in range(current_row, target_end_row):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c, value="")
            cell.border = thin_border

    # 6. Signatures Block (Spanning 10 Columns: Cols 1-5 for Custodian, Cols 6-10 for Recipient)
    sig_start = target_end_row + 1
    
    ws.merge_cells(start_row=sig_start, start_column=1, end_row=sig_start, end_column=5)
    ws.cell(row=sig_start, column=1, value="Received from:").font = font_bold
    
    ws.merge_cells(start_row=sig_start, start_column=6, end_row=sig_start, end_column=10)
    ws.cell(row=sig_start, column=6, value="Received by:").font = font_bold

    ws.merge_cells(start_row=sig_start+2, start_column=1, end_row=sig_start+2, end_column=5)
    c_cust = ws.cell(row=sig_start+2, column=1, value="HERICK REEL D. SORDILLA")
    c_cust.font = font_bold
    c_cust.alignment = align_center

    ws.merge_cells(start_row=sig_start+3, start_column=1, end_row=sig_start+3, end_column=5)
    ws.cell(row=sig_start+3, column=1, value="District Property Custodian").alignment = align_center

    ws.merge_cells(start_row=sig_start+4, start_column=1, end_row=sig_start+4, end_column=5)
