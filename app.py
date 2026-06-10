import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import io
from datetime import datetime, timedelta
import math

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Automated Multi-Site RAAWA Generator", 
    page_icon="📄", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS FOR PROFESSIONAL DESIGN ---
st.markdown("""
<style>
    /* Main header styling */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
    }
    
    /* Card styling */
    .custom-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Feature box styling */
    .feature-box {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        border-left: 4px solid #667eea;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    /* Advantage box styling */
    .advantage-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        color: white;
    }
    
    /* Sidebar styling */
    .sidebar-content {
        padding: 1rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        border-radius: 5px;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* Success message styling */
    .stSuccess {
        background: linear-gradient(135deg, #84fab0 0%, #8fd3f4 100%);
        padding: 1rem;
        border-radius: 10px;
    }
    
    /* Info box styling */
    .info-box {
        background: #e3f2fd;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #2196f3;
    }
    
    /* Date input styling */
    .stDateInput {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

def format_contact_number(contact):
    """Format contact number to start with 0 and have 11 digits"""
    if not contact or pd.isna(contact) or contact == 'N/A':
        return "N/A"
    
    # Convert to string and strip
    contact_str = str(contact).strip()
    
    # Remove any non-digit characters
    contact_str = ''.join(filter(str.isdigit, contact_str))
    
    # Check if it's a valid number
    if len(contact_str) == 10:
        # Starts with 9, add leading 0
        if contact_str.startswith('9'):
            return '0' + contact_str
        else:
            return contact_str
    elif len(contact_str) == 11:
        # Already has leading 0 or starts with 63
        if contact_str.startswith('0'):
            return contact_str
        elif contact_str.startswith('63'):
            # Convert from 63xxx to 0xxx
            return '0' + contact_str[2:]
        else:
            return contact_str
    elif len(contact_str) == 12 and contact_str.startswith('639'):
        # 639xxx format, convert to 09xxx
        return '0' + contact_str[2:]
    else:
        # Return as-is if can't format
        return contact_str

def format_id_number(id_num):
    """Format ID number to remove decimal places"""
    if not id_num or pd.isna(id_num) or id_num == 'N/A':
        return "N/A"
    
    # Convert to string
    id_str = str(id_num).strip()
    
    # If it's a float with .0, convert to integer
    try:
        # Check if it's a number with decimal
        if '.' in id_str:
            # Check if it's like "7796.0"
            float_val = float(id_str)
            if float_val.is_integer():
                return str(int(float_val))
            else:
                # Keep as string if not integer
                return id_str
        else:
            return id_str
    except:
        return id_str

# --- DATABASE LOADING ---
@st.cache_data
def load_databases():
    try:
        # Load Main Site Database
        df_sites = pd.read_excel("Globe FO Engr Conatct_Vendor.xlsx", sheet_name="MIN")
        df_sites['PLAID'] = df_sites['PLAID'].astype(str).str.strip()
        # Normalize: Remove the word 'Territory' and whitespace to just get the number
        df_sites['TERRITORY'] = df_sites['TERRITORY'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        # Load Requisitioner Database - FIXED: Use dtype=str to avoid float conversion issues
        df_req = pd.read_excel("Requisitioner.xlsx", header=1, dtype=str)
        
        # Clean up column names
        df_req.columns = df_req.columns.str.strip()
        
        # Normalize: Remove the word 'Territory' and whitespace to just get the number
        if 'Territory no.' in df_req.columns:
            df_req['Territory no.'] = df_req['Territory no.'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        # Format requisitioner data - handle potential NaN values
        for idx, row in df_req.iterrows():
            # Format ID number (remove decimal)
            if 'ID #' in df_req.columns:
                id_value = row.get('ID #', 'N/A')
                if pd.isna(id_value):
                    id_value = 'N/A'
                df_req.at[idx, 'ID #'] = format_id_number(id_value)
            
            # Format contact number (add leading 0)
            if 'Contact No.' in df_req.columns:
                contact_value = row.get('Contact No.', 'N/A')
                if pd.isna(contact_value):
                    contact_value = 'N/A'
                df_req.at[idx, 'Contact No.'] = format_contact_number(contact_value)
        
        # Load Engineer/Technician Database
        try:
            # Skip the first row (title row) and use the second row as header, read all as string
            df_engr_tech = pd.read_excel("EngrTech.xlsx", header=1, dtype=str)
            
            # Standardize column names
            df_engr_tech.columns = df_engr_tech.columns.str.strip()
            
            # Check for expected columns and map them
            if 'Name' in df_engr_tech.columns:
                pass  # Already has correct column name
            elif 'NAME' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'NAME': 'Name'}, inplace=True)
                
            if 'Company' in df_engr_tech.columns:
                pass  # Already has correct column name
            elif 'COMPANY' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'COMPANY': 'Company'}, inplace=True)
                
            # Map SEC ID to ID No
            if 'SEC ID' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'SEC ID': 'ID No'}, inplace=True)
            elif 'ID_NO' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'ID_NO': 'ID No'}, inplace=True)
            elif 'ID No' in df_engr_tech.columns:
                pass  # Already has correct column name
            elif 'ID NUMBER' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'ID NUMBER': 'ID No'}, inplace=True)
            
            # Format engineer/technician data
            if 'ID No' in df_engr_tech.columns:
                for idx, row in df_engr_tech.iterrows():
                    id_value = row.get('ID No', '')
                    if pd.isna(id_value):
                        id_value = ''
                    df_engr_tech.at[idx, 'ID No'] = format_id_number(id_value)
            
            # Drop any rows where Name is NaN or empty
            if 'Name' in df_engr_tech.columns:
                df_engr_tech = df_engr_tech.dropna(subset=['Name'], how='all')
                df_engr_tech = df_engr_tech[df_engr_tech['Name'].notna()]
                df_engr_tech = df_engr_tech[df_engr_tech['Name'].astype(str).str.strip() != '']
            
            # Fill NaN values in Company and ID No with empty strings
            if 'Company' in df_engr_tech.columns:
                df_engr_tech['Company'] = df_engr_tech['Company'].fillna('').astype(str)
            else:
                df_engr_tech['Company'] = ''
                
            if 'ID No' in df_engr_tech.columns:
                df_engr_tech['ID No'] = df_engr_tech['ID No'].fillna('').astype(str)
            else:
                df_engr_tech['ID No'] = ''
            
        except Exception as e:
            st.warning(f"EngrTech.xlsx not found or error loading: {e}")
            df_engr_tech = pd.DataFrame(columns=['Name', 'Company', 'ID No'])
        
        return df_sites, df_req, df_engr_tech
    except Exception as e:
        st.error(f"Error loading database files: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None, None, None

def create_raawa_file(matching_sites, personnel_list, scope_of_work, start_date, end_date, req_profile, facility_manager, batch_num=1, total_batches=1):
    """Helper function to create a single RAAWA file with dynamic font sizing"""
    template_file = "MIN591__MANUAL RAAWA_APPLICATION_June8,2026.xlsx"
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active 
    
    # --- WRITE REQUISITIONER DETAILS with formatting ---
    ws["D3"].value = req_profile["name"]
    ws["D4"].value = req_profile["dept"]
    
    # Format ID as integer without decimal
    id_value = format_id_number(req_profile["id"])
    ws["G4"].value = id_value
    
    # Format contact number with leading 0
    contact_value = format_contact_number(req_profile["contact"])
    ws["J4"].value = contact_value
    
    # --- WRITE SITES LINE-BY-LINE ---
    base_site_row = 6
    num_sites = len(matching_sites)
    for idx, (_, row) in enumerate(matching_sites.iterrows()):
        curr_row = base_site_row + idx
        ws.cell(row=curr_row, column=1, value=f"{row.get('PLAID', '')} - {row.get('SITE', '')}")
        ws.cell(row=curr_row, column=4, value=row.get("SITE_ADD", "N/A"))
    
    # Hide unused site rows
    for r in range(base_site_row + num_sites, 16):
        ws.row_dimensions[r].hidden = True
    
    # Set validity dates
    ws["D17"].value = start_date.strftime("%Y-%m-%d")
    ws["E17"].value = end_date.strftime("%Y-%m-%d")
    
    # --- WRITE PERSONNEL WITH DYNAMIC FONT SIZES ---
    start_personnel_row = 19
    
    # Function to calculate font size based on text length
    def get_font_size(text, min_size=6, max_size=10):
        """Calculate appropriate font size based on text length"""
        if not text or text == '':
            return max_size
        text_length = len(str(text))
        if text_length <= 12:
            return max_size  # 10
        elif text_length <= 18:
            return max_size - 1  # 9
        elif text_length <= 25:
            return max_size - 2  # 8
        elif text_length <= 32:
            return max_size - 3  # 7
        elif text_length <= 40:
            return max_size - 4  # 6
        else:
            return min_size  # 6
    
    for idx, person in enumerate(personnel_list):
        row_index = start_personnel_row + (idx // 2)
        col_offset = 0 if idx % 2 == 0 else 5
        
        # Format ID number for personnel
        formatted_id = format_id_number(person["id_no"])
        
        # Calculate font sizes for each cell
        name_font_size = get_font_size(person["name"])
        company_font_size = get_font_size(person["company"])
        id_font_size = get_font_size(formatted_id)
        
        # Name cell
        name_cell = ws.cell(row=row_index, column=1+col_offset)
        name_cell.value = person["name"]
        name_cell.font = Font(name="Calibri", size=name_font_size)
        
        # Company cell
        company_cell = ws.cell(row=row_index, column=4+col_offset)
        company_cell.value = person["company"]
        company_cell.font = Font(name="Calibri", size=company_font_size)
        
        # ID cell (formatted)
        id_cell = ws.cell(row=row_index, column=5+col_offset)
        id_cell.value = formatted_id
        id_cell.font = Font(name="Calibri", size=id_font_size)
    
    # Hide unused personnel rows
    for r in range(start_personnel_row + (len(personnel_list)//2 + 1), 39):
        ws.row_dimensions[r].hidden = True
    
    # Scope of work
    if total_batches > 1:
        ws["A41"].value = f"{scope_of_work}\n\n(Page {batch_num} of {total_batches} for this location group)"
    else:
        ws["A41"].value = scope_of_work
    
    # --- SIGNATORY REPLACEMENT ---
    # Replace the Facility Manager signatory (NEW ENGINEER_AH) - preserve original formatting
    original_signatory = ws["A48"].value
    if original_signatory:
        ws["A48"].value = str(original_signatory).replace("NEW ENGINEER_AH", facility_manager)
    else:
        ws["A48"].value = f"{facility_manager}\nSignature Over Printed Name / Date"
    
    # --- FINAL CLEANUP - Apply Calibri 6 to all personnel entries ---
    # Apply Calibri 6 font to all personnel cells to ensure consistency
    for row in range(start_personnel_row, start_personnel_row + (len(personnel_list)//2 + 1)):
        for col in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]:
            cell = ws.cell(row=row, column=col)
            if cell.value and row >= start_personnel_row:
                # Only change font size if it's currently larger than 6, otherwise keep as is
                if cell.font and cell.font.size and cell.font.size > 6:
                    cell.font = Font(name="Calibri", size=6)
                elif not cell.font:
                    cell.font = Font(name="Calibri", size=6)
    
    # Ensure header row (row 1) has correct formatting
    header_font = Font(name="Calibri", size=6, bold=False, italic=False)
    for col_idx in range(1, 12):
        ws.cell(row=1, column=col_idx).font = header_font
    
    # Ensure signatory has underline but keep its size
    sig_font = Font(name="Calibri", size=6, underline="single")
    ws["A48"].font = sig_font
    ws["A50"].font = sig_font
    
    # Clean up any random formatting in personnel area
    for row in range(start_personnel_row, 39):
        for col in [1, 4, 5, 6, 7, 8, 9, 10, 11]:
            cell = ws.cell(row=row, column=col)
            if cell.value and row >= start_personnel_row:
                if not cell.font or cell.font.size != 6:
                    cell.font = Font(name="Calibri", size=6)
    
    # Save to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def split_sites_by_territory_and_fm(matching_sites):
    """Split sites by unique combinations of Territory and Facility Manager, then further split if > 10 sites per group"""
    unique_combos = get_unique_combinations(matching_sites)
    
    final_groups = []
    
    for key, combo in unique_combos.items():
        sites_list = combo['sites']
        territory = combo['territory']
        facility_manager = combo['facility_manager']
        
        # Convert sites list to DataFrame
        sites_df = pd.DataFrame(sites_list)
        
        # If more than 10 sites, split into batches of 10
        if len(sites_df) > 10:
            num_batches = math.ceil(len(sites_df) / 10)
            for batch_num in range(num_batches):
                start_idx = batch_num * 10
                end_idx = min((batch_num + 1) * 10, len(sites_df))
                batch_sites = sites_df.iloc[start_idx:end_idx]
                
                final_groups.append({
                    'dataframe': batch_sites,
                    'territory': territory,
                    'facility_manager': facility_manager,
                    'batch_num': batch_num + 1,
                    'total_batches': num_batches,
                    'original_count': len(sites_df)
                })
        else:
            final_groups.append({
                'dataframe': sites_df,
                'territory': territory,
                'facility_manager': facility_manager,
                'batch_num': 1,
                'total_batches': 1,
                'original_count': len(sites_df)
            })
    
    return final_groups

def get_unique_combinations(matching_sites):
    """Get unique combinations of Territory and NEW ENGINEER_AH that must be separated"""
    combinations = {}
    
    for _, row in matching_sites.iterrows():
        territory = str(row.get("TERRITORY", "")).strip()
        facility_manager = str(row.get("NEW ENGINEER_AH", "")).strip()
        
        # Clean up N/A or nan values
        if facility_manager in ['N/A', 'nan', '']:
            facility_manager = "Unassigned FM"
        
        # Create a unique key combining territory and FM
        key = f"{territory}_{facility_manager}"
        
        if key not in combinations:
            combinations[key] = {
                'territory': territory,
                'facility_manager': facility_manager,
                'sites': []
            }
        
        combinations[key]['sites'].append(row)
    
    return combinations

def check_conflicts(matching_sites):
    """Check for conflicts in territories and facility managers"""
    unique_combos = get_unique_combinations(matching_sites)
    territories = set([combo['territory'] for combo in unique_combos.values()])
    facility_managers = set([combo['facility_manager'] for combo in unique_combos.values()])
    
    territory_conflict = '7' in territories and '8' in territories
    fm_conflict = len(facility_managers) > 1
    
    return {
        'territory_conflict': territory_conflict,
        'fm_conflict': fm_conflict,
        'unique_territories': territories,
        'unique_facility_managers': facility_managers,
        'num_combinations': len(unique_combos)
    }

# --- LOAD DATABASES FIRST ---
df_db, df_req_db, df_engr_tech_db = load_databases()

# --- REQUISITIONER FUNCTION (depends on loaded database) ---
def get_requisitioner_for_territory(territory):
    """Get requisitioner profile for a specific territory with formatted values"""
    if df_req_db is not None and 'Territory no.' in df_req_db.columns:
        matching_reqs = df_req_db[df_req_db['Territory no.'] == territory]
        if not matching_reqs.empty:
            req_row = matching_reqs.iloc[0]
            return {
                "name": str(req_row.get("Name", "N/A")),
                "dept": str(req_row.get("Dept./Group", "N/A")),
                "id": format_id_number(req_row.get("ID #", "N/A")),
                "contact": format_contact_number(req_row.get("Contact No.", "N/A"))
            }
    
    # Fallback for territory without requisitioner
    return {
        "name": f"Territory {territory} Engineer",
        "dept": f"TERRITORY {territory}",
        "id": "N/A",
        "contact": "N/A"
    }

# --- SIDEBAR NAVIGATION ---
st.sidebar.markdown("""
<div class="sidebar-content">
    <h2 style="text-align: center; color: #667eea;">📋 RAAWA Generator</h2>
    <hr>
</div>
""", unsafe_allow_html=True)

page = st.sidebar.radio(
    "Navigate",
    ["🏠 Main Generator", "ℹ️ About & Developer"],
    format_func=lambda x: x
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style="text-align: center; padding: 1rem;">
    <small>Version 2.2.0</small><br>
    <small>© 2026 RAAWA Generator</small><br>
    <small>✨ Unlimited Sites Support</small>
</div>
""", unsafe_allow_html=True)

# --- ABOUT PAGE ---
if page == "ℹ️ About & Developer":
    st.markdown("""
    <div class="main-header">
        <h1>📄 About RAAWA Generator</h1>
        <p>Professional Multi-Site RAAWA Document Automation System</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Developer Profile
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("""
        <div class="custom-card" style="text-align: center;">
            <h2>👨‍💻 Developer</h2>
            <hr>
            <h3>John Carlo Rabanes</h3>
            <p><strong>OLT Rollout Engineer</strong></p>
            <p>📞 09669343065</p>
            <p>📧 rabanes.johncarlo4@gmail.com</p>
            <p>🏢 Nokia Shanghai Bell</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="custom-card">
            <h2>🎯 Mission</h2>
            <p>To streamline and automate the RAAWA (Request for Authority to Access Work Area) document generation process, reducing manual effort and eliminating errors in multi-site telecommunications infrastructure projects.</p>
            <br>
            <h2>💡 Vision</h2>
            <p>To become the standard tool for telecommunications field operations, enabling engineers to generate compliant documentation in minutes instead of hours.</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Advantages Section
    st.markdown("## 🌟 Key Advantages")
    col1, col2, col3 = st.columns(3)
    
    advantages = [
        ("⏱️ Time Saving", "Generate complete RAAWA forms in seconds instead of 30+ minutes"),
        ("🎯 Zero Errors", "Eliminate manual data entry mistakes and formatting issues"),
        ("📊 Unlimited Sites", "Handle any number of sites with automatic batching")
    ]
    
    for idx, (title, desc) in enumerate(advantages):
        with [col1, col2, col3][idx]:
            st.markdown(f"""
            <div class="advantage-box">
                <h3>{title}</h3>
                <p>{desc}</p>
            </div>
            """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    advantages2 = [
        ("🔄 Auto-Routing", "Automatic requisitioner mapping based on territory"),
        ("📁 Database Integration", "Load personnel from Excel database with search"),
        ("🔒 Smart Batching", "Auto-splits into multiple RAAWAs when exceeding 10 sites")
    ]
    
    for idx, (title, desc) in enumerate(advantages2):
        with [col1, col2, col3][idx]:
            st.markdown(f"""
            <div class="advantage-box">
                <h3>{title}</h3>
                <p>{desc}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Features Section
    st.markdown("## ⚡ Features")
    
    features = [
        ("🏗️ Unlimited Sites", "Select any number of sites from database with auto-batching", "✅"),
        ("👥 Personnel Management", "Manual, database, or mixed input modes", "✅"),
        ("📝 Smart Conflict Resolution", "Auto-splits by territory and facility manager", "✅"),
        ("📄 Professional Output", "Perfectly formatted Excel with auto font sizing", "✅"),
        ("🔍 Search & Filter", "Quick personnel search from database", "✅"),
        ("💾 Batch Download", "Download multiple RAAWA files at once", "✅"),
        ("🎨 Clean Interface", "User-friendly with professional design", "✅"),
        ("📊 Real-time Preview", "See selected sites and personnel before generation", "✅"),
        ("🔄 Session Management", "Persistent data across navigation", "✅"),
        ("📑 Auto-Batching", "Automatically splits into 10-site batches per signatory group", "✅"),
        ("📅 Flexible Dates", "Choose start date manually or use current date", "✅"),
        ("📞 Auto-Format Contact", "Automatically formats contact numbers to start with 0", "✅"),
        ("🆔 Clean ID Numbers", "Removes decimal places from ID numbers", "✅")
    ]
    
    for feature in features:
        st.markdown(f"""
        <div class="feature-box">
            <h3>{feature[0]} {feature[2]}</h3>
            <p>{feature[1]}</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Tech Stack
    st.markdown("## 🛠️ Technology Stack")
    tech_col1, tech_col2, tech_col3, tech_col4 = st.columns(4)
    
    with tech_col1:
        st.markdown("""
        <div class="feature-box">
            <h3>🐍 Python</h3>
            <p>Core Logic</p>
        </div>
        """, unsafe_allow_html=True)
    
    with tech_col2:
        st.markdown("""
        <div class="feature-box">
            <h3>🎈 Streamlit</h3>
            <p>Web Framework</p>
        </div>
        """, unsafe_allow_html=True)
    
    with tech_col3:
        st.markdown("""
        <div class="feature-box">
            <h3>📊 Pandas</h3>
            <p>Data Processing</p>
        </div>
        """, unsafe_allow_html=True)
    
    with tech_col4:
        st.markdown("""
        <div class="feature-box">
            <h3>📝 OpenPyXL</h3>
            <p>Excel Generation</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Support Section
    st.markdown("## 📞 Support & Contact")
    st.markdown("""
    <div class="info-box">
        <h3>Need Help?</h3>
        <p>📧 Email: <a href="mailto:rabanes.johncarlo4@gmail.com">rabanes.johncarlo4@gmail.com</a></p>
        <p>📱 Phone: 09669343065</p>
        <p>🏢 Company: Nokia Shanghai Bell</p>
        <hr>
        <p><small>For technical support, feature requests, or bug reports, please reach out via email.</small></p>
    </div>
    """, unsafe_allow_html=True)

# --- MAIN GENERATOR PAGE ---
else:
    if df_db is not None and df_req_db is not None:
        # Professional Header
        st.markdown("""
        <div class="main-header">
            <h1>📄 Automated Multi-Site RAAWA Generator</h1>
            <p>Select any number of sites, populate the personnel manifest, and generate perfectly styled RAAWA forms instantly.<br>
            <strong>✨ Unlimited sites supported - Auto-batches into groups of 10 per signatory</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        # --- STEP 1: MULTI-SITE SELECTION INPUT ---
        st.markdown("## 📍 Step 1: Site Selection")
        st.markdown("Select the sites for your RAAWA request (unlimited number supported)")
        
        plaid_list = sorted(df_db['PLAID'].unique())
        selected_plaids = st.multiselect(
            "Search or Select Site IDs (PLAID):", 
            options=plaid_list,
            default=[]
        )
            
        matching_sites = df_db[df_db['PLAID'].isin(selected_plaids)]
        
        if not matching_sites.empty:
            # Check for conflicts using unique combinations
            conflicts = check_conflicts(matching_sites)
            
            if conflicts['territory_conflict']:
                st.warning("⚠️ **Territory Limitation Detected:** You have selected sites from both Territory 7 and Territory 8. These cannot be combined in a single RAAWA.")
            
            if conflicts['fm_conflict']:
                st.warning(f"⚠️ **Facility Manager Conflict Detected:** Multiple Facility Managers found ({', '.join(conflicts['unique_facility_managers'])}). Each unique Territory + Facility Manager combination will get its own RAAWA.")
            
            st.markdown("### 📋 Selected Sites Preview")
            st.dataframe(
                matching_sites[['PLAID', 'SITE', 'REGION', 'TERRITORY', 'NEW ENGINEER_AH', 'CONTACT NUMBER', 'SITE_ADD']],
                hide_index=True,
                use_container_width=True
            )
            
            # Show detailed breakdown of how files will be split
            unique_combos = get_unique_combinations(matching_sites)
            
            # Calculate total files including batching
            total_files = 0
            batch_details = []
            for combo in unique_combos.values():
                num_sites = len(combo['sites'])
                num_files = math.ceil(num_sites / 10)
                total_files += num_files
                batch_details.append({
                    'territory': combo['territory'],
                    'fm': combo['facility_manager'],
                    'sites': num_sites,
                    'files': num_files
                })
            
            st.info(f"📋 **Generation Plan:** The system will create **{total_files} separate RAAWA files**")
            
            # Display detailed breakdown
            with st.expander("📊 View Detailed RAAWA Breakdown"):
                for detail in batch_details:
                    if detail['files'] > 1:
                        st.markdown(f"""
                        **Territory {detail['territory']} - {detail['fm']}**:
                        - {detail['sites']} sites → {detail['files']} RAAWA files (batches of 10)
                        """)
                    else:
                        st.markdown(f"""
                        **Territory {detail['territory']} - {detail['fm']}**:
                        - {detail['sites']} sites → {detail['files']} RAAWA file
                        """)
                    st.markdown("---")
                
        else:
            st.info("💡 Please select at least one site to begin")
            st.stop()

        st.markdown("---")

        # --- STEP 2: ACCESS SCOPE & DEPLOYMENTS ---
        st.markdown("## 📝 Step 2: Access Scope & Deployment Details")
        
        scope_col, date_col = st.columns(2)
        with scope_col:
            scope_of_work = st.text_area(
                "Nature of Access / Detailed Scope of Work:",
                value="SITE SURVEY, INSTALLATION, INTEGRATION, AND ACCEPTANCE TESTING.",
                height=100
            )
        with date_col:
            # Date selection options
            date_option = st.radio(
                "Date Selection Method:",
                ["Use Current Date", "Manual Date Entry"],
                horizontal=True
            )
            
            if date_option == "Use Current Date":
                start_date = datetime.now()
                st.info(f"📅 Start Date: **{start_date.strftime('%Y-%m-%d')}** (Current Date)")
            else:
                start_date = st.date_input(
                    "Select Start Date:",
                    value=datetime.now().date(),
                    min_value=datetime.now().date(),
                    help="Choose the effective start date for the authorization"
                )
                start_date = datetime.combine(start_date, datetime.min.time())
                st.caption(f"Selected Start Date: **{start_date.strftime('%Y-%m-%d')}**")
            
            # Validity period
            validity_days = st.number_input(
                "Authorization Validity (Days):", 
                min_value=1, 
                max_value=365, 
                value=30,
                help="Number of days the authorization will be valid from the start date"
            )
            
            # Calculate end date
            end_date = start_date + timedelta(days=int(validity_days))
            
            # Display coverage
            st.success(f"📅 Form window coverage: **{start_date.strftime('%Y-%m-%d')}** to **{end_date.strftime('%Y-%m-%d')}**")
            st.caption(f"Total validity: {validity_days} days")

        st.markdown("## 👥 Step 3: Team Manifest")
        st.markdown("Add the engineers and technicians who will perform the work")
        
        # --- PERSONNEL SELECTION METHOD ---
        col_method, col_empty = st.columns([2, 3])
        with col_method:
            selection_method = st.radio(
                "Personnel Selection Method:",
                ["Manual Input", "Load from Database", "Mixed (Database + Manual)"],
                horizontal=True
            )
        
        # Initialize personnel list in session state if not exists
        if 'personnel_list' not in st.session_state:
            st.session_state.personnel_list = []
        
        if selection_method == "Load from Database":
            if df_engr_tech_db is not None and not df_engr_tech_db.empty:
                st.success(f"📋 Found {len(df_engr_tech_db)} personnel records in database")
                
                # Display database records for selection
                with st.expander("📁 Select Personnel from Database", expanded=True):
                    # Search/filter functionality
                    search_term = st.text_input("🔍 Search by Name, Company, or ID:", key="personnel_search")
                    
                    filtered_df = df_engr_tech_db
                    if search_term:
                        filtered_df = df_engr_tech_db[
                            df_engr_tech_db['Name'].astype(str).str.contains(search_term, case=False, na=False) |
                            df_engr_tech_db['Company'].astype(str).str.contains(search_term, case=False, na=False) |
                            df_engr_tech_db['ID No'].astype(str).str.contains(search_term, case=False, na=False)
                        ]
                    
                    # Multi-select for personnel
                    selected_indices = st.multiselect(
                        "Select Engineers/Technicians to add:",
                        options=filtered_df.index.tolist(),
                        format_func=lambda x: f"{filtered_df.loc[x, 'Name']} - {filtered_df.loc[x, 'Company']} (ID: {filtered_df.loc[x, 'ID No']})",
                        key="selected_personnel_db"
                    )
                    
                    if st.button("➕ Add Selected Personnel to Manifest"):
                        for idx in selected_indices:
                            person = filtered_df.loc[idx]
                            # Check for duplicates
                            if not any(p['name'] == person['Name'] for p in st.session_state.personnel_list):
                                st.session_state.personnel_list.append({
                                    "name": person['Name'],
                                    "company": person['Company'],
                                    "id_no": person['ID No']
                                })
                        st.success(f"Added {len(selected_indices)} personnel to manifest!")
                        st.rerun()
                
                # Manual addition option within database mode
                st.markdown("---")
                st.subheader("Or Add Manually")
                
                # Manual input fields in database mode
                col1, col2, col3 = st.columns(3)
                with col1:
                    manual_name = st.text_input("Name:", key="manual_name_db")
                with col2:
                    manual_company = st.text_input("Company:", key="manual_company_db")
                with col3:
                    manual_id = st.text_input("ID No.:", key="manual_id_db")
                
                if st.button("➕ Add Manual Entry", key="add_manual_db"):
                    if manual_name:
                        st.session_state.personnel_list.append({
                            "name": manual_name,
                            "company": manual_company,
                            "id_no": manual_id
                        })
                        st.success(f"Added {manual_name} to manifest!")
                        st.rerun()
                
            else:
                st.warning("No personnel data found in EngrTech.xlsx. Please use Manual Input mode.")
                selection_method = "Manual Input"
        
        if selection_method == "Manual Input":
            # Track visibility dynamically up to 40 entries
            for i in range(1, 41):
                if f"p_show_{i}" not in st.session_state:
                    st.session_state[f"p_show_{i}"] = True if i <= 4 else False
                    
            visible_count = sum(1 for i in range(1, 41) if st.session_state[f"p_show_{i}"])
            
            for i in range(1, 41):
                if st.session_state[f"p_show_{i}"]:
                    p_col1, p_col2, p_col3 = st.columns([2, 2, 2])
                    with p_col1:
                        p_name = st.text_input(f"Personnel Name {i}", key=f"name_{i}")
                    with p_col2:
                        p_comp = st.text_input(f"Company/Vendor {i}", key=f"comp_{i}")
                    with p_col3:
                        p_id = st.text_input(f"Security ID No. {i}", key=f"id_{i}")
                        
                    if p_name: 
                        # Check if already added to avoid duplicates
                        if not any(p['name'] == p_name for p in st.session_state.personnel_list):
                            st.session_state.personnel_list.append({"name": p_name, "company": p_comp, "id_no": p_id})

            if visible_count < 40:
                if st.button("➕ Add More Personnel Fields", key="add_more_fields"):
                    for i in range(1, 41):
                        if not st.session_state[f"p_show_{i}"]:
                            st.session_state[f"p_show_{i}"] = True
                            st.rerun()
        
        if selection_method == "Mixed (Database + Manual)":
            st.subheader("Database Selection")
            if df_engr_tech_db is not None and not df_engr_tech_db.empty:
                # Database selection section
                with st.expander("📁 Select from Database", expanded=True):
                    search_term = st.text_input("🔍 Search:", key="mixed_search")
                    
                    filtered_df = df_engr_tech_db
                    if search_term:
                        filtered_df = df_engr_tech_db[
                            df_engr_tech_db['Name'].astype(str).str.contains(search_term, case=False, na=False) |
                            df_engr_tech_db['Company'].astype(str).str.contains(search_term, case=False, na=False) |
                            df_engr_tech_db['ID No'].astype(str).str.contains(search_term, case=False, na=False)
                        ]
                    
                    selected_indices = st.multiselect(
                        "Select personnel:",
                        options=filtered_df.index.tolist(),
                        format_func=lambda x: f"{filtered_df.loc[x, 'Name']} - {filtered_df.loc[x, 'Company']}",
                        key="mixed_selection"
                    )
                    
                    if st.button("➕ Add Selected", key="add_mixed"):
                        for idx in selected_indices:
                            person = filtered_df.loc[idx]
                            if not any(p['name'] == person['Name'] for p in st.session_state.personnel_list):
                                st.session_state.personnel_list.append({
                                    "name": person['Name'],
                                    "company": person['Company'],
                                    "id_no": person['ID No']
                                })
                        st.success(f"Added {len(selected_indices)} personnel!")
                        st.rerun()
            
            st.markdown("---")
            st.subheader("Manual Addition")
            
            # Manual input fields
            col1, col2, col3 = st.columns(3)
            with col1:
                manual_name = st.text_input("Name:", key="manual_name_mixed")
            with col2:
                manual_company = st.text_input("Company:", key="manual_company_mixed")
            with col3:
                manual_id = st.text_input("ID No.:", key="manual_id_mixed")
            
            if st.button("➕ Add Manual Entry", key="add_manual_mixed"):
                if manual_name:
                    st.session_state.personnel_list.append({
                        "name": manual_name,
                        "company": manual_company,
                        "id_no": manual_id
                    })
                    st.success(f"Added {manual_name} to manifest!")
                    st.rerun()
        
        # Display current personnel manifest
        if st.session_state.personnel_list:
            st.markdown("---")
            st.markdown(f"### 📋 Current Team Manifest ({len(st.session_state.personnel_list)} personnel)")
            
            # Create a dataframe for display with formatted IDs
            manifest_df = pd.DataFrame(st.session_state.personnel_list)
            manifest_df['id_no'] = manifest_df['id_no'].apply(format_id_number)
            st.dataframe(manifest_df, hide_index=True, use_container_width=True)
            
            # Option to clear manifest
            if st.button("🗑️ Clear All Personnel"):
                st.session_state.personnel_list = []
                st.rerun()
        else:
            st.info("No personnel added yet. Use the methods above to add team members.")
        
        st.markdown("---")

        # --- STEP 3: LOGIC EXECUTION & RENDER ---
        if st.button("🚀 Build Multi-Site RAAWA Spreadsheet", type="primary"):
            if not st.session_state.personnel_list:
                st.warning("Please add at least one personnel entry to populate the work manifest.")
            else:
                try:
                    # Split sites by territory, facility manager, and batch size
                    site_groups = split_sites_by_territory_and_fm(matching_sites)
                    
                    total_groups = len(site_groups)
                    st.info(f"📋 Generating **{total_groups} RAAWA files**...")
                    
                    # Create RAAWA files for each group
                    file_details = []
                    
                    for group_info in site_groups:
                        group_sites = group_info['dataframe']
                        territory = group_info['territory']
                        facility_manager = group_info['facility_manager']
                        batch_num = group_info['batch_num']
                        total_batches = group_info['total_batches']
                        
                        # Get requisitioner for the territory
                        req_profile = get_requisitioner_for_territory(territory)
                        
                        # Create the RAAWA file
                        buffer = create_raawa_file(
                            group_sites, 
                            st.session_state.personnel_list, 
                            scope_of_work, 
                            start_date, 
                            end_date, 
                            req_profile,
                            facility_manager,
                            batch_num,
                            total_batches
                        )
                        
                        # Generate clean filename
                        clean_fm = facility_manager.replace(" ", "_").replace("/", "_").replace(",", "")
                        if total_batches > 1:
                            filename = f"RAAWA_Territory{territory}_{clean_fm}_Batch{batch_num}of{total_batches}_{group_sites.iloc[0]['PLAID']}.xlsx"
                            display_name = f"Territory {territory} - {facility_manager[:25]} (Batch {batch_num}/{total_batches})"
                        else:
                            filename = f"RAAWA_Territory{territory}_{clean_fm}_{group_sites.iloc[0]['PLAID']}.xlsx"
                            display_name = f"Territory {territory} - {facility_manager[:30]}"
                        
                        file_details.append({
                            "name": filename,
                            "buffer": buffer,
                            "display_name": display_name,
                            "territory": territory,
                            "facility_manager": facility_manager,
                            "requisitioner": req_profile['name'],
                            "num_sites": len(group_sites),
                            "batch_info": f" (Batch {batch_num}/{total_batches})" if total_batches > 1 else ""
                        })
                        
                        st.success(f"✅ Created RAAWA #{len(file_details)}: Territory {territory} - {facility_manager[:40]} ({len(group_sites)} sites){' - Batch ' + str(batch_num) + '/' + str(total_batches) if total_batches > 1 else ''} | Requisitioner: {req_profile['name']}")
                    
                    # Store in session state for persistent access
                    st.session_state['generated_files'] = file_details
                    st.session_state['files_generated'] = True
                    st.rerun()
                        
                except Exception as ex:
                    st.error(f"Error: {ex}")
        
        # --- DISPLAY DOWNLOAD BUTTONS FOR GENERATED FILES ---
        if st.session_state.get('files_generated', False):
            st.markdown("---")
            st.markdown("## 📥 Download Generated RAAWA Files")
            
            file_details = st.session_state.get('generated_files', [])
            
            if len(file_details) == 1:
                # Single file download
                st.success(f"🎉 Successfully generated 1 RAAWA file!")
                file_info = file_details[0]
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.download_button(
                        label=f"💾 Download {file_info['name']}",
                        data=file_info['buffer'],
                        file_name=file_info['name'],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="single_download",
                        use_container_width=True
                    )
                with col2:
                    st.caption(f"📋 {file_info['num_sites']} site(s){file_info['batch_info']}")
                
                st.markdown(f"""
                **Details:**
                - 🏢 **Facility Manager:** {file_info['facility_manager']}
                - 👤 **Requisitioner:** {file_info['requisitioner']}
                - 📍 **Territory:** {file_info['territory']}
                """)
                
            else:
                # Multiple files download
                st.success(f"🎉 Successfully generated {len(file_details)} RAAWA files!")
                st.markdown("**Click each button below to download the respective RAAWA file:**")
                
                # Create expandable section for better organization
                with st.expander("📁 View All Generated Files", expanded=True):
                    for idx, file_info in enumerate(file_details):
                        st.markdown(f"**File #{idx + 1}: Territory {file_info['territory']}{file_info['batch_info']}**")
                        
                        col1, col2, col3 = st.columns([2, 1.5, 1.5])
                        
                        with col1:
                            st.download_button(
                                label=f"📄 {file_info['display_name'][:50]}",
                                data=file_info['buffer'],
                                file_name=file_info['name'],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{idx}_{file_info['territory']}_{file_info['facility_manager'][:20]}",
                                use_container_width=True
                            )
                        
                        with col2:
                            st.caption(f"🏢 FM: {file_info['facility_manager'][:35]}")
                            st.caption(f"📍 Territory: {file_info['territory']}")
                        
                        with col3:
                            st.caption(f"📋 Req: {file_info['requisitioner'][:30]}")
                            st.caption(f"🏗️ {file_info['num_sites']} site(s)")
                        
                        st.markdown("---")
            
            # Add option to clear and start over
            col_clear, col_spacer = st.columns([1, 3])
            with col_clear:
                if st.button("🔄 Clear Files & Start New RAAWA", use_container_width=True):
                    st.session_state['files_generated'] = False
                    st.session_state['generated_files'] = []
                    st.session_state.personnel_list = []
                    st.rerun()
    else:
        st.error("Failed to load databases. Please check that all required Excel files are present.")