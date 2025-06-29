import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.units import mm
from reportlab.lib import colors
from collections import Counter
import plotly.express as px
import os

# --- Helper Functions ---

def load_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    dfs = {sheet: xls.parse(sheet) for sheet in sheets}
    return dfs

def get_unique(df, cols):
    if isinstance(cols, str):
        cols = [cols]
    vals = []
    for col in cols:
        if col in df.columns:
            vals += df[col].dropna().astype(str).tolist()
    return sorted(set([v.strip() for v in vals if v.strip()]))

def filter_df(df, filters):
    mask = pd.Series([True] * len(df))
    for col, selected in filters.items():
        if col in df.columns and selected:
            if isinstance(selected, list):
                mask &= df[col].astype(str).isin(selected)
            else:
                mask &= df[col].astype(str).str.contains(selected, case=False, na=False)
    return df[mask]

def multi_col_name_search(df, search, cols):
    if not search.strip():
        return pd.Series([True] * len(df))
    mask = pd.Series([False] * len(df))
    for col in cols:
        if col in df.columns:
            mask |= df[col].astype(str).str.contains(search, case=False, na=False)
    return mask

def val(row, col):
    return str(row.get(col, "")).strip() if pd.notnull(row.get(col, "")) else ""

def location_presence(row, locations):
    result = []
    for loc in locations:
        for suffix in ['CO', 'Branch', 'Plant']:
            col_name = f"{loc}"
            if col_name in row and isinstance(row[col_name], str) and suffix in row[col_name]:
                result.append(f"{suffix} in {loc}")
    return ", ".join(result)

def combine_vals(row, cols, sep=", "):
    return sep.join([val(row, c) for c in cols if val(row, c)])

def card_html(lines):
    content = ""
    for label, value in lines:
        if value:
            content += f"<strong style='color:#6366f1'>{label}:</strong> <span style='color:#334155'>{value}</span><br>"
    return f"""
    <div style="border:2px solid #6366f1;padding:16px 20px 14px 20px;border-radius:12px;margin-bottom:20px;background:linear-gradient(135deg,#f1f5f9 0%,#e0e7ff 100%);color:#000;min-height:120px;box-shadow:0 2px 8px #6366f122;">
        {content}
    </div>
    """

def render_cards(df, card_func, columns_per_row=3):
    cards = []
    for _, row in df.iterrows():
        cards.append(card_func(row))
    for i in range(0, len(cards), columns_per_row):
        cols = st.columns(columns_per_row)
        for j in range(columns_per_row):
            if i + j < len(cards):
                with cols[j]:
                    st.markdown(cards[i + j], unsafe_allow_html=True)

# --- PDF Export with ReportLab ---

def download_pdf_reportlab(filtered_df, card_func, filename="contacts.pdf"):
    if filtered_df.empty:
        return BytesIO()
    first_row = filtered_df.iloc[0]
    card_lines = card_func(first_row, as_lines=True)
    card_fields = [label for label, value in card_lines]

    # Decide orientation based on number of columns
    if len(card_fields) > 8:
        page_size = landscape(A4)
        total_width_mm = 297 - 20 - 20  # A4 landscape width - margins
    else:
        page_size = A4
        total_width_mm = 210 - 20 - 20  # A4 portrait width - margins

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle(
        'cell',
        fontSize=7,
        leading=9,
        wordWrap='CJK',  # enables wrapping for long words
        alignment=0,     # left
        spaceAfter=2,
    )
    header_style = ParagraphStyle(
        'header',
        fontSize=7.5,
        leading=10,
        alignment=0,
        spaceAfter=2,
        fontName='Helvetica-Bold'
    )
    story = []

    # --- Add logo at the top ---
    logo_path = "logo.png"  # Change to your logo file path
    if os.path.exists(logo_path):
        img = Image(logo_path, width=120, height=40)  # Adjust size as needed
        story.append(img)
        story.append(Spacer(1, 12))  # Add some space after the logo

    # Prepare table data: header + rows (with Paragraph for wrapping)
    table_data = [[Paragraph(str(field), header_style) for field in card_fields]]
    for _, row in filtered_df.iterrows():
        lines = dict(card_func(row, as_lines=True))
        row_data = [Paragraph(str(lines.get(field, "")), cell_style) for field in card_fields]
        table_data.append(row_data)

    # Calculate column widths to fit page width
    col_width = total_width_mm / len(card_fields)
    col_widths = [col_width * mm for _ in card_fields]

    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(table)

    doc.build(story)
    buffer.seek(0)
    return buffer

# --- Card Layouts for Each Tab ---

def listed_companies_card(row, as_lines=False):
    locations = ["Agra", "Mumbai", "NCR", "Chennai", "Vadodara", "Bangalore", "Pune", "Kolkata", "Hyderabad", "Ahmedabad"]
    ceo_designation = combine_vals(row, ['CEO Name ', 'Designation'], ", ")
    company_bloomberg = combine_vals(row, ['Corporate Name', 'Bloomberg Code'], ", ")
    sector_subsector = combine_vals(row, ['Sector', 'Sub Sector'], ", ")
    analyst_team = combine_vals(row, ['Relevant Analyst Team (Sector Wise)'], ", ")
    analyst_location = combine_vals(row, ['Head Office'], ", ")
    analyst_team_and_loc = f"{analyst_team}, {analyst_location}".strip()
    cfo_designation = combine_vals(row, ['CFO Connects', 'Designation.1'], ", ")
    


    loc_presence = location_presence(row, locations)
    lines = [
        ("CEO & Designation", ceo_designation),
        ("CEO City", val(row, 'CEO City')),
        ("Corporate & Bl.Code", company_bloomberg),
        ("Sector & Subsector", sector_subsector),
        ("Analyst & Head Office", analyst_team_and_loc),
        ("CFO Connect & Designation", cfo_designation),
        ("Location Presence", loc_presence),
        
    ]
    if as_lines:
        return lines
    return card_html(lines)

def expert_confirmed_card(row, as_lines=False):
    name = val(row, "Name")
    designation = val(row, "Designation")
    sector_segment = combine_vals(row, ['Sector', 'Segments'], " ")
    company_desc = combine_vals(row, ['Company', 'Description'], ", ")
    city = val(row, "Location")
    
    lines = [
        ("Name", name),
        ("Designation", designation),
        ("Sector & Segment", sector_segment),
        ("Company & Description", company_desc),
        ("Location", city),
        
    ]
    if as_lines:
        return lines
    return card_html(lines)

def expert_potential_card(row, as_lines=False):
    name = val(row, "Name")
    designation = val(row, "Designation")
    sector_segment = combine_vals(row, ['Sector', 'Segment'], " ")
    company_desc = combine_vals(row, ['Company', 'Description'], ", ")
    city = val(row, "Location")
    
    lines = [
        ("Name", name),
        ("Designation", designation),
        ("Sector & Segment", sector_segment),
        ("Company & Description", company_desc),
        ("Location", city),
        
    ]
    if as_lines:
        return lines
    return card_html(lines)

def channel_checks_card(row, as_lines=False):
    name = val(row, "Name")
    sector_subsector = combine_vals(row, ['Sector', 'Sub Sector'], " - ")
    state_city = combine_vals(row, ['State', 'Location'], " ")
    designation = val(row, "Designation / Area of Expertise")
    
    lines = [
        ("Name", name),
        ("Sector - Sub Sector", sector_subsector),
        ("State & City", state_city),
        ("Designation / Area of Expertise", designation),
        
    ]
    if as_lines:
        return lines
    return card_html(lines)

def ir_data_card(row, as_lines=False):
    bloomberg_code = val(row, "Bloomberg Code")
    full_name = val(row, "Full Name")
    sector_subsector = combine_vals(row, ['Sector', 'Sub Sector'], ", ")
    ir_agency = val(row, "IR Agency")
    
    lines = [
        ("Bloomberg Code", bloomberg_code),
        ("Full Name", full_name),
        ("Sector, Sub Sector", sector_subsector),
        ("IR Agency", ir_agency),
    ]
    if as_lines:
        return lines
    return card_html(lines)

def ministry_contacts_card(row, as_lines=False):
    name = val(row, "Name")
    designation = val(row, "Designation")
    sector = val(row, "Sector")
    department = val(row, "Department")
    address = val(row, "Address")
    email_phone = combine_vals(row, ['Email', 'Phone Number'], ", ")
    
    lines = [
        ("Name", name),
        ("Designation", designation),
        ("Sector", sector),
        ("Department", department),
        ("Address", address),
        ("Email", email_phone),  # <-- Add comma here
        
    ]
    if as_lines:
        return lines
    return card_html(lines)

def generic_card(row, columns, as_lines=False):
    lines = []
    for col in columns:
        v = val(row, col)
        if v:
            lines.append((col, v))
    email_phone = combine_vals(row, ['Email', 'Phone Number'], ", ")
    if email_phone:
        lines.append(("Email, Phone", email_phone))
    if as_lines:
        return lines
    return card_html(lines)

# --- All Sheet Dashboard ---

def all_dashboard(dfs):
    st.header("All Sheets Dashboard")
    total_contacts = sum([len(df) for df in dfs.values()])
    st.metric("Total Contacts", total_contacts)

    # Add counts for specific sheets
    sheet_counts = {sheet.lower(): len(df) for sheet, df in dfs.items()}
    st.metric("🌱 Channel Checks", sheet_counts.get('channel checks', 0))
    st.metric("✅ Expert Confirmed", sheet_counts.get('expert confirmed', 0))
    st.metric("⭐ Expert Potential", sheet_counts.get('expert potential', 0))

    # Concatenate all data
    all_df = pd.concat([df.assign(_sheet=sheet) for sheet, df in dfs.items()], ignore_index=True)
    sheet_names = list(dfs.keys())

    # Sidebar filters
    with st.sidebar:
        st.subheader("All Tab Filters")
        st.markdown("**Tip:** All filters are combined (AND). Only rows matching all selected filters will be shown.")
        selected_sheets = st.multiselect("Select Sheet(s)", sheet_names, default=sheet_names)
        name_cols = [c for c in all_df.columns if "name" in c.lower()]
        sector_col = "Sector" if "Sector" in all_df.columns else all_df.columns[0]
        market_cap_col = None
        for c in all_df.columns:
            if "market cap" in c.lower():
                market_cap_col = c
                break
        name_suggestions = get_unique(all_df, name_cols)
        sector_suggestions = get_unique(all_df, sector_col)
        market_cap_suggestions = get_unique(all_df, market_cap_col) if market_cap_col else []
        name_search = st.selectbox("Name Search", [""] + name_suggestions, key="all_name_search")
        sector_search = st.multiselect("Sector", sector_suggestions, key="all_sector_search")
        if market_cap_col:
            market_cap_search = st.multiselect("Market Cap", market_cap_suggestions, key="all_marketcap_search")
        else:
            market_cap_search = []
        # Location-wise filter (searches all location-like columns)
        location_cols = [c for c in all_df.columns if any(x in c.lower() for x in ["location", "city", "state", "address"])]
        location_suggestions = get_unique(all_df, location_cols)
        location_search = st.selectbox("Location Search (All Tabs)", [""] + location_suggestions, key="all_location_search")

    # Filter by selected sheets
    if selected_sheets:
        dfs = {sheet: dfs[sheet] for sheet in selected_sheets}
        all_df = all_df[all_df["_sheet"].isin(selected_sheets)].reset_index(drop=True)

    # If any filter is used, filter each sheet and show results grouped by sheet
    if name_search or sector_search or (market_cap_col and market_cap_search) or location_search:
        st.write("Showing filtered contacts grouped by sheet:")
        filtered_dfs = {}
        for sheet, df in dfs.items():
            mask = pd.Series([True] * len(df))
            filter_applied = False  # Track if any filter is applied to this sheet

            # Name filter
            if name_search:
                name_cols_sheet = [c for c in df.columns if "name" in c.lower()]
                if name_cols_sheet:
                    mask &= multi_col_name_search(df, name_search, name_cols_sheet)
                    filter_applied = True

            # Sector filter
            if sector_search and sector_col in df.columns:
                mask &= df[sector_col].astype(str).isin(sector_search)
                filter_applied = True

            # Market Cap filter
            if market_cap_col and market_cap_col in df.columns and market_cap_search:
                mask &= df[market_cap_col].astype(str).isin(market_cap_search)
                filter_applied = True

            # Location filter
            if location_search:
                loc_cols = [c for c in df.columns if any(x in c.lower() for x in ["location", "city", "state", "address"])]
                if loc_cols:
                    mask_loc = pd.Series([False] * len(df))
                    for col in loc_cols:
                        mask_loc |= df[col].astype(str).str.contains(location_search, case=False, na=False)
                    mask &= mask_loc
                    filter_applied = True

            filtered_df = df[mask]
            # Only show if a filter was actually applied to this sheet and results are not empty
            if filter_applied and not filtered_df.empty:
                st.subheader(f"{sheet} ({len(filtered_df)})")
                st.dataframe(filtered_df)
                filtered_dfs[sheet] = filtered_df

        # Download options
        if filtered_dfs:
            download_option = st.radio(
                "Download filtered results from:",
                ["All Selected Sheets (Combined)", "Choose Sheet(s)"],
                horizontal=True,
                key="download_option"
            )
            if download_option == "All Selected Sheets (Combined)":
                # Excel download (separate sheets)
                excel_bytes = BytesIO()
                with pd.ExcelWriter(excel_bytes, engine="xlsxwriter") as writer:
                    for sheet, df in filtered_dfs.items():
                        df.to_excel(writer, sheet_name=sheet[:31], index=False)
                excel_bytes.seek(0)
                st.download_button(
                    "Download All Filtered as Excel (Separate Sheets)",
                    data=excel_bytes,
                    file_name="filtered_contacts_by_sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # PDF download (all sheets, one after another)
                pdf_buffers = []
                for sheet, df in filtered_dfs.items():
                    # Choose the right card function for each sheet if needed
                    card_func = generic_card
                    if "listed companies" in sheet.lower():
                        card_func = listed_companies_card
                    elif "expert confirmed" in sheet.lower():
                        card_func = expert_confirmed_card
                    elif "expert potential" in sheet.lower():
                        card_func = expert_potential_card
                    elif "channel checks" in sheet.lower():
                        card_func = channel_checks_card
                    elif "ir data" in sheet.lower():
                        card_func = ir_data_card
                    elif "ministry contacts" in sheet.lower():
                        card_func = ministry_contacts_card

                    pdf_buffer = download_pdf_reportlab(df, card_func)
                    pdf_buffers.append((sheet, pdf_buffer.read()))
                # Combine PDFs (simple concatenation, works for most viewers)
                from PyPDF2 import PdfMerger
                merger = PdfMerger()
                for sheet, pdf_bytes in pdf_buffers:
                    merger.append(BytesIO(pdf_bytes))
                merged_pdf = BytesIO()
                merger.write(merged_pdf)
                merger.close()
                merged_pdf.seek(0)
                st.download_button(
                    "Download All Filtered as PDF (Combined)",
                    data=merged_pdf,
                    file_name="filtered_contacts_by_sheet.pdf",
                    mime="application/pdf"
                )

            else:
                sheet_to_download = st.multiselect(
                    "Select sheet(s) to download",
                    list(filtered_dfs.keys()),
                    default=list(filtered_dfs.keys()),
                    key="sheet_download_select"
                )
                if sheet_to_download:
                    # Excel download (selected sheets)
                    excel_bytes = BytesIO()
                    with pd.ExcelWriter(excel_bytes, engine="xlsxwriter") as writer:
                        for sheet in sheet_to_download:
                            filtered_dfs[sheet].to_excel(writer, sheet_name=sheet[:31], index=False)
                    excel_bytes.seek(0)
                    st.download_button(
                        "Download Selected Sheets as Excel (Separate Sheets)",
                        data=excel_bytes,
                        file_name="filtered_contacts_selected_sheets.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # PDF download (selected sheets)
                    pdf_buffers = []
                    for sheet in sheet_to_download:
                        df = filtered_dfs[sheet]
                        card_func = generic_card
                        if "listed companies" in sheet.lower():
                            card_func = listed_companies_card
                        elif "expert confirmed" in sheet.lower():
                            card_func = expert_confirmed_card
                        elif "expert potential" in sheet.lower():
                            card_func = expert_potential_card
                        elif "channel checks" in sheet.lower():
                            card_func = channel_checks_card
                        elif "ir data" in sheet.lower():
                            card_func = ir_data_card
                        elif "ministry contacts" in sheet.lower():
                            card_func = ministry_contacts_card

                        pdf_buffer = download_pdf_reportlab(df, card_func)
                        pdf_buffers.append((sheet, pdf_buffer.read()))
                    from PyPDF2 import PdfMerger
                    merger = PdfMerger()
                    for sheet, pdf_bytes in pdf_buffers:
                        merger.append(BytesIO(pdf_bytes))
                    merged_pdf = BytesIO()
                    merger.write(merged_pdf)
                    merger.close()
                    merged_pdf.seek(0)
                    st.download_button(
                        "Download Selected Sheets as PDF (Combined)",
                        data=merged_pdf,
                        file_name="filtered_contacts_selected_sheets.pdf",
                        mime="application/pdf"
                    )
        return  # Stop further processing if any filter is used

    # If no filter, show all
    st.write(f"Showing {len(all_df)} contacts")
    st.dataframe(all_df)
    excel_bytes = BytesIO()
    all_df.to_excel(excel_bytes, index=False)
    excel_bytes.seek(0)
    st.download_button("Download All as Excel", data=excel_bytes, file_name="all_contacts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Streamlit App ---

st.set_page_config(page_title="Dealer Directory", layout="wide")

# Add your logo here (adjust width as needed)
st.image("logo.png", width=180)  # <-- Place your logo file in the same directory

st.markdown("<h2 style='color:#000;'>Dealer Directory</h2>", unsafe_allow_html=True)

st.success("Welcome! Upload your Excel file to get started.")

# Custom CSS
st.markdown("""
    <style>
    /* App background: soft gradient */
    .stApp {
        background: linear-gradient(135deg, #f0f4fd 0%, #e0e7ff 100%);
        color: #000;
    }
    /* Sidebar: slightly darker gradient */
    .stSidebar {
        background: linear-gradient(135deg, #e0e7ff 0%, #6366f1 100%);
        color: #000;
    }
    /* Headings */
    .stHeader, .stSubheader, h1, h2, h3, h4 {
        color: #000 !important;
        letter-spacing: 0.5px;
    }
    /* Card look for DataFrames and widgets */
    .stDataFrame, .stSelectbox, .stMultiSelect, .stTextInput, .stTextArea, .stNumberInput {
        background: #fff !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 8px #6366f122 !important;
        color: #1e293b !important;
    }
    /* Card HTML (custom cards) */
    div[style*="border:2px solid"] {
        background: linear-gradient(135deg, #6366f1 0%, #f1f5f9 100%) !important;
        color: #fff !important;
        border-radius: 14px !important;
        box-shadow: 0 4px 16px #6366f133 !important;
        border: none !important;
    }
    div[style*="border:2px solid"] strong {
        color: #fbbf24 !important;
    }
    div[style*="border:2px solid"] span {
        color: #fff !important;
    }
    /* Buttons */
    .stButton>button, .stDownloadButton>button {
        color: #fff !important;
        background: linear-gradient(90deg, #6366f1 0%, #818cf8 100%) !important;
        border: none;
        border-radius: 8px;
        padding: 0.5em 1.2em;
        font-weight: 600;
        box-shadow: 0 2px 8px #6366f122;
        transition: background 0.2s;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background: linear-gradient(90deg, #818cf8 0%, #6366f1 100%) !important;
    }
    /* Metrics */
    .stMetric {
        background: #fff;
        border-radius: 10px;
        color: #6366f1 !important;
        padding: 0.5em 1em;
        box-shadow: 0 2px 8px #6366f122;
        border: 1px solid #e0e7ff;
    }
    /* Success/info boxes */
    .stAlert-success {
        background: linear-gradient(90deg, #bbf7d0 0%, #22d3ee 100%) !important;
        color: #134e4a !important;
        border-radius: 10px;
        font-weight: 500;
    }
    .stAlert-info {
        background: linear-gradient(90deg, #e0e7ff 0%, #f1f5f9 100%) !important;
        color: #334155 !important;
        border-radius: 10px;
        font-weight: 500;
    }
    /* Scrollbar for dark sidebar */
    .stSidebar [data-testid="stVerticalBlock"]::-webkit-scrollbar-thumb {
        background: #6366f1 !important;
    }
    </style>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    dfs = load_excel(uploaded_file)
    sheet_names = list(dfs.keys())
    st.sidebar.header("Sheet Selection")
    select_options = ["All"] + sheet_names
    selected_sheet = st.sidebar.selectbox("Select sheet", select_options)

    pdf_bytes = None

    if selected_sheet == "All":
        all_dashboard(dfs)
    else:
        df = dfs[selected_sheet]
        tab = selected_sheet.lower()
        st.sidebar.header("Filters")
        filters = {}

        if "listed companies" in tab:
            name_cols = ['CEO Name ', 'CFO Connects', 'Relevant Analyst Team Sector Wise']
            name_suggestions = get_unique(df, name_cols)
            name_search = st.sidebar.selectbox("Name Search", [""] + name_suggestions, key="lc_name_search")
            filter_fields = {
                "Corporate Name": "Corporate Name",
                "Bloomberg Code": "Bloomberg Code",
                "Sector": "Sector",
                "Coverage": "Coverage",
                "SEBI Classification": "SEBI Classification",
                "Relation lead": "Relation lead (Research / Corp Access / IB / IR Agency / Parent / Sales)",
                "Head Office": "Head Office",
            }
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"lc_{col}_filter_{idx}")
            mask = pd.Series([True] * len(df))
            if name_search:
                mask &= multi_col_name_search(df, name_search, name_cols)
            for col, selected in filters.items():
                if col in df.columns and selected:
                    mask &= df[col].astype(str).isin(selected)
            filtered_df = df[mask]
            card_func = listed_companies_card

        elif "expert confirmed" in tab:
            name_suggestions = get_unique(df, "Name")
            name_search = st.sidebar.selectbox("Name", [""] + name_suggestions, key="ec_name_search")
            designation_suggestions = get_unique(df, "Designation")
            designation_search = st.sidebar.multiselect("Designation", designation_suggestions, key="ec_designation")
            city_suggestions = get_unique(df, "Location")
            city_search = st.sidebar.multiselect("Location", city_suggestions, key="ec_city")
            filter_fields = {
                "Sector": "Sector",
                "Segment": "Segments",
                "Company": "Company",
                "Description": "Description"
            }
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"ec_{col}_filter_{idx}")
            mask = pd.Series([True] * len(df))
            if name_search:
                mask &= df["Name"].astype(str) == name_search
            if designation_search and "Designation" in df.columns:
                mask &= df["Designation"].astype(str).isin(designation_search)
            if city_search and "City" in df.columns:
                mask &= df["City"].astype(str).isin(city_search)
            for col, selected in filters.items():
                if col in df.columns and selected:
                    mask &= df[col].astype(str).isin(selected)
            filtered_df = df[mask]
            card_func = expert_confirmed_card

        elif "expert potential" in tab:
            name_suggestions = get_unique(df, "Name")
            name_search = st.sidebar.selectbox("Name", [""] + name_suggestions, key="ep_name_search")
            designation_suggestions = get_unique(df, "Designation")
            designation_search = st.sidebar.multiselect("Designation", designation_suggestions, key="ep_designation")
            city_suggestions = get_unique(df, "Location")
            city_search = st.sidebar.multiselect("Location", city_suggestions, key="ep_city")
            filter_fields = {
                "Sector": "Sector",
                "Segment": "Segment",
                "Company": "Company",
                "Description": "Description"
            }
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"ep_{col}_filter_{idx}")
            mask = pd.Series([True] * len(df))
            if name_search:
                mask &= df["Name"].astype(str) == name_search
            if designation_search and "Designation" in df.columns:
                mask &= df["Designation"].astype(str).isin(designation_search)
            if city_search and "Location" in df.columns:
                mask &= df["Location"].astype(str).isin(city_search)
            for col, selected in filters.items():
                if col in df.columns and selected:
                    mask &= df[col].astype(str).isin(selected)
            filtered_df = df[mask]
            card_func = expert_potential_card

        elif "channel checks" in tab:
            name_suggestions = get_unique(df, "Name")
            name_search = st.sidebar.selectbox("Name", [""] + name_suggestions, key="cc_name_search")
            sector_suggestions = get_unique(df, "Sector")
            sector_search = st.sidebar.multiselect("Sector", sector_suggestions, key="cc_sector")
            sub_sector_suggestions = get_unique(df, "Sub Sector")
            sub_sector_search = st.sidebar.multiselect("Sub Sector", sub_sector_suggestions, key="cc_sub_sector")
            state_suggestions = get_unique(df, "State")
            state_search = st.sidebar.multiselect("State", state_suggestions, key="cc_state")
            city_suggestions = get_unique(df, "Location")
            city_search = st.sidebar.multiselect("Location", city_suggestions, key="cc_city")
            designation_col = "Designation / Area of Expertise" if "Designation / Area of Expertise" in df.columns else "Designation"
            designation_suggestions = get_unique(df, designation_col)
            designation_search = st.sidebar.multiselect(designation_col, designation_suggestions, key="cc_designation")
            mask = pd.Series([True] * len(df))
            if name_search:
                mask &= df["Name"].astype(str) == name_search
            if sector_search and "Sector" in df.columns:
                mask &= df["Sector"].astype(str).isin(sector_search)
            if sub_sector_search and "Sub Sector" in df.columns:
                mask &= df["Sub Sector"].astype(str).isin(sub_sector_search)
            if state_search and "State" in df.columns:
                mask &= df["State"].astype(str).isin(state_search)
            if city_search and "Location" in df.columns:
                mask &= df["Location"].astype(str).isin(city_search)
            if designation_search and designation_col in df.columns:
                mask &= df[designation_col].astype(str).isin(designation_search)
            filtered_df = df[mask]
            card_func = channel_checks_card

        elif "ir data" in tab:
            filter_fields = {col: col for col in df.columns}
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"ir_{col}_filter_{idx}")
            mask = pd.Series([True] * len(df))
            for col, selected in filters.items():
                if col in df.columns and selected:
                    mask &= df[col].astype(str).isin(selected)
            filtered_df = df[mask]
            card_func = ir_data_card

        elif "ministry contacts" in tab:
            filter_fields = {col: col for col in df.columns}
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"mc_{col}_filter_{idx}")
            mask = pd.Series([True] * len(df))
            for col, selected in filters.items():
                if col in df.columns and selected:
                    mask &= df[col].astype(str).isin(selected)
            filtered_df = df[mask]
            card_func = ministry_contacts_card

        else:
            filter_fields = {col: col for col in df.columns}
            for idx, (label, col) in enumerate(filter_fields.items()):
                options = get_unique(df, col)
                if len(options) > 1 and len(options) < 100:
                    filters[col] = st.sidebar.multiselect(label, options, default=[], key=f"gen_{col}_filter_{idx}")
                else:
                    filters[col] = st.sidebar.text_input(label, key=f"gen_{col}_input_{idx}")
            mask = pd.Series([True] * len(df))
            for col, selected in filters.items():
                if col in df.columns and selected:
                    if isinstance(selected, list):
                        mask &= df[col].astype(str).isin(selected)
                    else:
                        mask &= df[col].astype(str).str.contains(selected, case=False, na=False)
            filtered_df = df[mask]
            card_func = lambda row, as_lines=False: generic_card(row, df.columns[:8], as_lines=as_lines)

        pdf_bytes = download_pdf_reportlab(filtered_df, card_func)
        st.download_button("Download results as PDF", data=pdf_bytes, file_name="contacts.pdf", mime="application/pdf")

        # Excel download button for the filtered data
        excel_bytes = BytesIO()
        filtered_df.to_excel(excel_bytes, index=False)
        excel_bytes.seek(0)
        st.download_button(
            "Download results as Excel",
            data=excel_bytes,
            file_name="contacts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader(f"Contacts ({len(filtered_df)})")
        render_cards(filtered_df, card_func, columns_per_row=3)
else:
    st.info("Please upload an Excel file to get started.")
