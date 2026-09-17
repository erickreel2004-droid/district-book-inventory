from datetime import datetime
from io import BytesIO
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. PAGE CONFIGURATION & CUSTOM CSS
# ==========================================
st.set_page_config(page_title="District Book Inventory", layout="wide")

st.markdown("""
    <style>
    /* Dark Slate Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }

    /* Main Page Select Box & Inputs Styling */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
    }
    
    /* Force main page selectbox text to be clear dark slate */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {
        color: #0F172A !important;
    }

    textarea, div[data-baseweb="input"] > div {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
    }

    input, textarea {
        color: #0F172A !important;
    }

    .stTextInput label, .stDateInput label, .stTextArea label, .stSelectbox label {
        color: #0F172A !important;
        font-weight: 600;
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
        # Merge prices from master inventory
        cols_to_merge = ["book_title", "unit_cost"] if "unit_cost" in master_df.columns else ["book_title"]
        merged_records = school_records.merge(master_df[cols_to_merge], on="book_title", how="left")
        
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
# 4. FORMAL DEPED ICS EXCEL GENERATOR (WITH DR & IAR)
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

    # 1. Header (Rows 1-5 across 9 columns A to I)
    headers = [
        "Republic of the Philippines",
        "Department of Education",
        "Region XI – Davao Region",
        "City Schools Division of Digos",
        "Digos Occidental District"
    ]
    for r_idx, text in enumerate(headers, start=1):
        ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=9)
        cell = ws.cell(row=r_idx, column=1, value=text)
        cell.font = font_header
        cell.alignment = align_center

    # 2. Section Title (Row 7 across A7:I7)
    ws.merge_cells("A7:I7")
    title_cell = ws.cell(row=7, column=1, value="INVENTORY CUSTODIAN SLIP (ICS)")
    title_cell.font = font_title
    title_cell.alignment = align_center

    # 3. Metadata Section (Rows 11-12)
    ws.cell(row=11, column=1, value=f"Entity Name: {school_name}").font = font_bold
    ws.cell(row=12, column=1, value="INVENTORY CUSTODIAN SLIP (ICS) No.: ____________________").font = font_bold

    # 4. Table Column Headers (Row 14 - 9 Columns)
    table_headers = [
        "Quantity", "Unit", "Unit Cost", "Total Cost", "Description", 
        "DR No.", "IAR No.", "Inventory Item No.", "Estimated Useful Life"
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
        dr_no = str(row.get("dr_no", "")) if "dr_no" in row and pd.notna(row.get("dr_no")) else ""
        iar_no = str(row.get("iar_no", "")) if "iar_no" in row and pd.notna(row.get("iar_no")) else ""
        useful_life = int(row.get("useful_life", 3))

        ws.cell(row=current_row, column=1, value=qty).alignment = align_center
        ws.cell(row=current_row, column=2, value="PCS").alignment = align_center
        
        c3 = ws.cell(row=current_row, column=3, value=unit_cost)
        c3.number_format = '#,##0.00'
        c3.alignment = align_right

        # Excel formula for Total Cost (= Quantity * Unit Cost)
        c4 = ws.cell(row=current_row, column=4, value=f"=A{current_row}*C{current_row}")
        c4.number_format = '#,##0.00'
        c4.alignment = align_right

        ws.cell(row=current_row, column=5, value=desc).alignment = align_left
        ws.cell(row=current_row, column=6, value=dr_no).alignment = align_center   # DR No.
        ws.cell(row=current_row, column=7, value=iar_no).alignment = align_center  # IAR No.
        ws.cell(row=current_row, column=8, value="").alignment = align_center       # Inventory Item No.
        ws.cell(row=current_row, column=9, value=useful_life).alignment = align_center

        for col in range(1, 10):
            ws.cell(row=current_row, column=col).border = thin_border
            ws.cell(row=current_row, column=col).font = font_regular

        current_row += 1

    # Fill blank padded rows to keep standard sheet length
    target_end_row = max(current_row + 3, 26)
    for r in range(current_row, target_end_row):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c, value="")
            cell.border = thin_border

    # 6. Signatures Block (Spanning 9 Columns: Cols 1-5 for Custodian, Cols 6-9 for Recipient)
    sig_start = target_end_row + 1
    
    # Left Header (Custodian)
    ws.merge_cells(start_row=sig_start, start_column=1, end_row=sig_start, end_column=5)
    ws.cell(row=sig_start, column=1, value="Received from:").font = font_bold
    
    # Right Header (Recipient)
    ws.merge_cells(start_row=sig_start, start_column=6, end_row=sig_start, end_column=9)
    ws.cell(row=sig_start, column=6, value="Received by:").font = font_bold

    # Custodian Name
    ws.merge_cells(start_row=sig_start+2, start_column=1, end_row=sig_start+2, end_column=5)
    c_cust = ws.cell(row=sig_start+2, column=1, value="HERICK REEL D. SORDILLA")
    c_cust.font = font_bold
    c_cust.alignment = align_center

    ws.merge_cells(start_row=sig_start+3, start_column=1, end_row=sig_start+3, end_column=5)
    ws.cell(row=sig_start+3, column=1, value="District Property Custodian").alignment = align_center

    ws.merge_cells(start_row=sig_start+4, start_column=1, end_row=sig_start+4, end_column=5)
    ws.cell(row=sig_start+4, column=1, value="Date: ____________________").alignment = align_center

    # Recipient Line
    ws.merge_cells(start_row=sig_start+2, start_column=6, end_row=sig_start+2, end_column=9)
    ws.cell(row=sig_start+2, column=6, value="__________________________________").alignment = align_center

    ws.merge_cells(start_row=sig_start+3, start_column=6, end_row=sig_start+3, end_column=9)
    ws.cell(row=sig_start+3, column=6, value="Signature over Printed Name of End-User").alignment = align_center

    ws.merge_cells(start_row=sig_start+4, start_column=6, end_row=sig_start+4, end_column=9)
    ws.cell(row=sig_start+4, column=6, value=f"Date: {date_str}").alignment = align_center

    # Signatures Outer Border Outline
    for r in range(sig_start, sig_start+5):
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = thin_border

    # Set Column Widths for clean 9-column layout
    col_widths = {1: 10, 2: 8, 3: 12, 4: 14, 5: 30, 6: 14, 7: 14, 8: 18, 9: 18}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. CUSTODIAN VIEW
# ==========================================
if role == "Custodian View":
    st.header("🔒 Custodian Control Panel")

    password = st.text_input("Enter Custodian Password to Access:", type="password")
    CUSTODIAN_PASSWORD = "admin123"

    if password == CUSTODIAN_PASSWORD:
        st.success("Access Granted!")

        col1, col2 = st.columns([1, 2])

        with col1:
            st.subheader("Add / Update Central Stock")
            with st.form("add_book_form"):
                title = st.text_input("Book Title")
                stock = st.number_input("Quantity to Add", min_value=1, step=1, value=100)
                unit_cost = st.number_input("Unit Cost (₱)", min_value=0.0, step=0.01, value=90.00, format="%.2f")
                dr_no = st.text_input("DR No. (Optional)")
                iar_no = st.text_input("IAR No. (Optional)")
                useful_life = st.number_input("Useful Life (Years)", min_value=1, step=1, value=3)
                submit = st.form_submit_button("Add to Master Stock")

                if submit and title:
                    master_df = load_data("master_inventory")
                    title_clean = title.strip()

                    if not master_df.empty and "book_title" in master_df.columns and title_clean in master_df["book_title"].values:
                        master_df.loc[master_df["book_title"] == title_clean, "central_stock"] += stock
                        master_df.loc[master_df["book_title"] == title_clean, "unit_cost"] = unit_cost
                        master_df.loc[master_df["book_title"] == title_clean, "dr_no"] = dr_no
                        master_df.loc[master_df["book_title"] == title_clean, "iar_no"] = iar_no
                        master_df.loc[master_df["book_title"] == title_clean, "useful_life"] = useful_life
                    else:
                        new_row = pd.DataFrame([{
                            "book_title": title_clean, 
                            "central_stock": stock,
                            "unit_cost": unit_cost,
                            "dr_no": dr_no,
                            "iar_no": iar_no,
                            "useful_life": useful_life
                        }])
                        master_df = pd.concat([master_df, new_row], ignore_index=True)

                    conn.update(worksheet="master_inventory", data=master_df)
                    st.success(f"Added {stock} copies of '{title_clean}'!")
                    st.rerun()

        with col2:
            st.subheader("Dispatch Books to Schools")
            master_df = load_data("master_inventory")

            if master_df.empty or "book_title" not in master_df.columns:
                st.info("Please add books to Central Stock first before dispatching.")
            else:
                master_df = master_df.sort_values(by="book_title", ascending=False)
                book_options = master_df["book_title"].tolist()

                m_col1, m_col2 = st.columns([3, 1])
                selected_book = m_col1.selectbox("📖 Select Book Title to Dispatch:", book_options)
                dispatch_all_btn = m_col2.button("⚡ Batch Dispatch All", type="primary", use_container_width=True)

                current_stock = int(master_df.loc[master_df["book_title"] == selected_book, "central_stock"].values[0])

                st.divider()

                live_total_requested = 0
                dispatch_selections = []

                for idx, school in enumerate(SCHOOL_LIST):
                    c1, c2, c3 = st.columns([4, 2, 2])
                    c1.write(f"**{school}**")

                    input_key = f"dispatch_qty_{selected_book}_{idx}"
                    if input_key not in st.session_state:
                        st.session_state[input_key] = 10

                    qty = c2.number_input(
                        f"Qty for {school}",
                        min_value=1,
                        step=1,
                        key=input_key,
                        label_visibility="collapsed",
                    )
                    live_total_requested += qty
                    dispatch_selections.append({"school": school, "qty": qty})

                    if c3.button("Dispatch", key=f"dispatch_btn_{selected_book}_{idx}"):
                        if qty > current_stock:
                            st.error(f"Not enough stock! Only {current_stock} available.")
                        else:
                            master_df.loc[master_df["book_title"] == selected_book, "central_stock"] -= qty
                            conn.update(worksheet="master_inventory", data=master_df)

                            school_df = load_data("school_inventory")
                            new_dispatch = pd.DataFrame([{
                                "school_name": school,
                                "book_title": selected_book,
                                "quantity_received": qty,
                                "status": "For Release"
                            }])
                            school_df = pd.concat([school_df, new_dispatch], ignore_index=True)
                            conn.update(worksheet="school_inventory", data=school_df)

                            st.success(f"Dispatched {qty} copies of '{selected_book}' to {school}!")
                            st.rerun()

                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("📦 Warehouse Stock Available", current_stock)
                m2.metric("📋 Total Input (All Schools)", live_total_requested)
                m3.metric("🟢 Remaining Stock After Batch", current_stock - live_total_requested)

                if dispatch_all_btn:
                    if live_total_requested > current_stock:
                        st.error(f"Cannot batch dispatch! Required `{live_total_requested}` copies, but only `{current_stock}` available.")
                    else:
                        master_df.loc[master_df["book_title"] == selected_book, "central_stock"] -= live_total_requested
                        conn.update(worksheet="master_inventory", data=master_df)

                        school_df = load_data("school_inventory")
                        new_rows = pd.DataFrame([
                            {
                                "school_name": item["school"],
                                "book_title": selected_book,
                                "quantity_received": item["qty"],
                                "status": "For Release"
                            } for item in dispatch_selections
                        ])
                        school_df = pd.concat([school_df, new_rows], ignore_index=True)
                        conn.update(worksheet="school_inventory", data=school_df)

                        st.success(f"Successfully batch dispatched **{selected_book}** to all schools!")
                        st.rerun()

        st.divider()

        tab1, tab2, tab3 = st.tabs(["📊 Central Warehouse Stock", "🚚 Dispatched Inventory Log", "📅 Scheduled Appointments"])

        with tab1:
            central_df = load_data("master_inventory")
            if not central_df.empty:
                st.dataframe(central_df, use_container_width=True)

        with tab2:
            dispatched_df = load_data("school_inventory")
            if not dispatched_df.empty:
                st.dataframe(dispatched_df, use_container_width=True)

        with tab3:
            appointments_df = load_data("appointments")
            if not appointments_df.empty:
                st.dataframe(appointments_df, use_container_width=True)

# ==========================================
# 6. PRINCIPAL VIEW
# ==========================================
else:
    st.header(f"Principal Portal - {selected_school}")

    st.info(f"ℹ️ **Notice for {selected_school}:** Please be informed that the following books/learning materials assigned to your school are now ready for pickup at the District Office.")

    school_df = load_data("school_inventory")
    master_df = load_data("master_inventory")

    if not school_df.empty and "school_name" in school_df.columns:
        filtered = school_df[school_df["school_name"].astype(str).str.strip() == selected_school.strip()]
        
        if not filtered.empty and "status" in filtered.columns:
            filtered = filtered[filtered["status"].astype(str).str.strip().str.lower() != "received"]
        
        if not filtered.empty:
            summary = filtered.groupby(["book_title", "status"])["quantity_received"].sum().reset_index()
            
            # --- MERGE BOOK DETAILS (PRICES, DR, IAR, USEFUL LIFE) FROM MASTER INVENTORY ---
            if not master_df.empty:
                cols_to_merge = ["book_title"]
                for col_name in ["unit_cost", "useful_life", "dr_no", "iar_no"]:
                    if col_name in master_df.columns:
                        cols_to_merge.append(col_name)
                
                summary = summary.merge(master_df[cols_to_merge], on="book_title", how="left")
            
            if "unit_cost" not in summary.columns:
                summary["unit_cost"] = 90.00
            else:
                summary["unit_cost"] = summary["unit_cost"].fillna(90.00)

            if "useful_life" not in summary.columns:
                summary["useful_life"] = 3
            else:
                summary["useful_life"] = summary["useful_life"].fillna(3)

            summary = summary.sort_values(by="book_title", ascending=False)
            
            # Display readable summary on web interface
            display_cols = ["book_title", "status", "quantity_received", "unit_cost"]
            for col_opt in ["dr_no", "iar_no"]:
                if col_opt in summary.columns:
                    display_cols.append(col_opt)

            st.dataframe(summary[display_cols], use_container_width=True)

            excel_data = generate_official_deped_ics_excel(
                school_name=selected_school,
                date_str=datetime.now().strftime("%B %d, %Y"),
                df_items=summary
            )

            st.download_button(
                label="📊 Download DepEd ICS Form (Excel)",
                data=excel_data,
                file_name=f"ICS_{selected_school.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.success("🎉 No pending dispatches for this school. All books have been received!")
    else:
        st.info("No dispatches logged for this school yet.")
