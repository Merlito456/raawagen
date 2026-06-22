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
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
    }
    .custom-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .feature-box {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        border-left: 4px solid #667eea;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .advantage-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        color: white;
    }
    .sidebar-content {
        padding: 1rem;
    }
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
    .stSuccess {
        background: linear-gradient(135deg, #84fab0 0%, #8fd3f4 100%);
        padding: 1rem;
        border-radius: 10px;
    }
    .info-box {
        background: #e3f2fd;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #2196f3;
    }
    .stDateInput {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

def format_contact_number(contact):
    """Format contact number to start with 0 and have 11 digits"""
    if not contact or pd.isna(contact) or contact == 'N/A':
        return "N/A"
    
    contact_str = str(contact).strip()
    contact_str = ''.join(filter(str.isdigit, contact_str))
    
    if len(contact_str) == 10:
        if contact_str.startswith('9'):
            return '0' + contact_str
        else:
            return contact_str
    elif len(contact_str) == 11:
        if contact_str.startswith('0'):
            return contact_str
        elif contact_str.startswith('63'):
            return '0' + contact_str[2:]
        else:
            return contact_str
    elif len(contact_str) == 12 and contact_str.startswith('639'):
        return '0' + contact_str[2:]
    else:
        return contact_str

def format_id_number(id_num):
    """Format ID number to remove decimal places"""
    if not id_num or pd.isna(id_num) or id_num == 'N/A':
        return "N/A"
    
    id_str = str(id_num).strip()
    
    try:
        if '.' in id_str:
            float_val = float(id_str)
            if float_val.is_integer():
                return str(int(float_val))
            else:
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
        df_sites['TERRITORY'] = df_sites['TERRITORY'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        if 'REGION' in df_sites.columns:
            df_sites['REGION'] = df_sites['REGION'].astype(str).str.upper().str.strip()
            df_sites['REGION'] = df_sites['REGION'].apply(lambda x: x if x in ['LUZ', 'VIS', 'MIN'] else 'MIN')
        
        # Load Requisitioner Database
        df_req = pd.read_excel("Requisitioner.xlsx", header=1, dtype=str)
        df_req.columns = df_req.columns.str.strip()
        
        if 'Territory no.' in df_req.columns:
            df_req['Territory no.'] = df_req['Territory no.'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        if 'Region' in df_req.columns:
            df_req['Region'] = df_req['Region'].astype(str).str.upper().str.strip()
            df_req['Region'] = df_req['Region'].apply(lambda x: x if x in ['LUZ', 'VIS', 'MIN'] else 'MIN')
        
        for idx, row in df_req.iterrows():
            if 'ID #' in df_req.columns:
                id_value = row.get('ID #', 'N/A')
                if pd.isna(id_value):
                    id_value = 'N/A'
                df_req.at[idx, 'ID #'] = format_id_number(id_value)
            
            if 'Contact No.' in df_req.columns:
                contact_value = row.get('Contact No.', 'N/A')
                if pd.isna(contact_value):
                    contact_value = 'N/A'
                df_req.at[idx, 'Contact No.'] = format_contact_number(contact_value)
        
        # Load Engineer/Technician Database
        try:
            df_engr_tech = pd.read_excel("EngrTech.xlsx", header=0, dtype=str)
            df_engr_tech.columns = df_engr_tech.columns.str.strip()
            
            # Map columns
            if 'Name' not in df_engr_tech.columns:
                for col in ['NAME', 'name']:
                    if col in df_engr_tech.columns:
                        df_engr_tech.rename(columns={col: 'Name'}, inplace=True)
                        break
            
            if 'Company' not in df_engr_tech.columns:
                for col in ['COMPANY', 'company', 'Vendor']:
                    if col in df_engr_tech.columns:
                        df_engr_tech.rename(columns={col: 'Company'}, inplace=True)
                        break
            
            if 'ID No' not in df_engr_tech.columns:
                for col in ['ID_NO', 'ID NUMBER', 'SEC ID', 'ID']:
                    if col in df_engr_tech.columns:
                        df_engr_tech.rename(columns={col: 'ID No'}, inplace=True)
                        break
            
            if 'Region' not in df_engr_tech.columns:
                for col in ['REGION', 'region']:
                    if col in df_engr_tech.columns:
                        df_engr_tech.rename(columns={col: 'Region'}, inplace=True)
                        break
            
            # Format IDs
            if 'ID No' in df_engr_tech.columns:
                df_engr_tech['ID No'] = df_engr_tech['ID No'].astype(str).apply(format_id_number)
            
            # Drop empty rows
            if 'Name' in df_engr_tech.columns:
                df_engr_tech = df_engr_tech.dropna(subset=['Name'], how='all')
                df_engr_tech = df_engr_tech[df_engr_tech['Name'].notna()]
                df_engr_tech = df_engr_tech[df_engr_tech['Name'].astype(str).str.strip() != '']
            
            # Fill NaN values
            if 'Company' in df_engr_tech.columns:
                df_engr_tech['Company'] = df_engr_tech['Company'].fillna('').astype(str)
            else:
                df_engr_tech['Company'] = ''
                
            if 'ID No' in df_engr_tech.columns:
                df_engr_tech['ID No'] = df_engr_tech['ID No'].fillna('').astype(str)
            else:
                df_engr_tech['ID No'] = ''
            
            # Handle Region column
            if 'Region' in df_engr_tech.columns:
                df_engr_tech['Region'] = df_engr_tech['Region'].fillna('').astype(str).str.upper().str.strip()
                df_engr_tech['Region'] = df_engr_tech['Region'].apply(lambda x: x if x in ['LUZ', 'VIS', 'MIN'] else '')
            else:
                df_engr_tech['Region'] = ''
                
        except Exception as e:
            st.warning(f"EngrTech.xlsx error: {e}")
            df_engr_tech = pd.DataFrame(columns=['Name', 'Company', 'ID No', 'Region'])
        
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
    
    ws["D3"].value = req_profile["name"]
    ws["D4"].value = req_profile["dept"]
    
    id_value = format_id_number(req_profile["id"])
    ws["G4"].value = id_value
    
    contact_value = format_contact_number(req_profile["contact"])
    ws["J4"].value = contact_value
    
    base_site_row = 6
    num_sites = len(matching_sites)
    for idx, (_, row) in enumerate(matching_sites.iterrows()):
        curr_row = base_site_row + idx
        ws.cell(row=curr_row, column=1, value=f"{row.get('PLAID', '')} - {row.get('SITE', '')}")
        ws.cell(row=curr_row, column=4, value=row.get("SITE_ADD", "N/A"))
    
    for r in range(base_site_row + num_sites, 16):
        ws.row_dimensions[r].hidden = True
    
    ws["D17"].value = start_date.strftime("%Y-%m-%d")
    ws["E17"].value = end_date.strftime("%Y-%m-%d")
    
    start_personnel_row = 19
    
    def get_font_size(text, min_size=6, max_size=10):
        if not text or text == '':
            return max_size
        text_length = len(str(text))
        if text_length <= 12:
            return max_size
        elif text_length <= 18:
            return max_size - 1
        elif text_length <= 25:
            return max_size - 2
        elif text_length <= 32:
            return max_size - 3
        elif text_length <= 40:
            return max_size - 4
        else:
            return min_size
    
    for idx, person in enumerate(personnel_list):
        row_index = start_personnel_row + (idx // 2)
        col_offset = 0 if idx % 2 == 0 else 5
        
        formatted_id = format_id_number(person["id_no"])
        
        name_cell = ws.cell(row=row_index, column=1+col_offset)
        name_cell.value = person["name"]
        name_cell.font = Font(name="Calibri", size=get_font_size(person["name"]))
        
        company_cell = ws.cell(row=row_index, column=4+col_offset)
        company_cell.value = person["company"]
        company_cell.font = Font(name="Calibri", size=get_font_size(person["company"]))
        
        id_cell = ws.cell(row=row_index, column=5+col_offset)
        id_cell.value = formatted_id
        id_cell.font = Font(name="Calibri", size=get_font_size(formatted_id))
    
    for r in range(start_personnel_row + (len(personnel_list)//2 + 1), 39):
        ws.row_dimensions[r].hidden = True
    
    if total_batches > 1:
        ws["A41"].value = f"{scope_of_work}\n\n(Page {batch_num} of {total_batches} for this location group)"
    else:
        ws["A41"].value = scope_of_work
    
    original_signatory = ws["A48"].value
    if original_signatory:
        ws["A48"].value = str(original_signatory).replace("NEW ENGINEER_AH", facility_manager)
    else:
        ws["A48"].value = f"{facility_manager}\nSignature Over Printed Name / Date"
    
    for row in range(start_personnel_row, start_personnel_row + (len(personnel_list)//2 + 1)):
        for col in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]:
            cell = ws.cell(row=row, column=col)
            if cell.value and row >= start_personnel_row:
                if cell.font and cell.font.size and cell.font.size > 6:
                    cell.font = Font(name="Calibri", size=6)
                elif not cell.font:
                    cell.font = Font(name="Calibri", size=6)
    
    header_font = Font(name="Calibri", size=6, bold=False, italic=False)
    for col_idx in range(1, 12):
        ws.cell(row=1, column=col_idx).font = header_font
    
    sig_font = Font(name="Calibri", size=6, underline="single")
    ws["A48"].font = sig_font
    ws["A50"].font = sig_font
    
    for row in range(start_personnel_row, 39):
        for col in [1, 4, 5, 6, 7, 8, 9, 10, 11]:
            cell = ws.cell(row=row, column=col)
            if cell.value and row >= start_personnel_row:
                if not cell.font or cell.font.size != 6:
                    cell.font = Font(name="Calibri", size=6)
    
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
        
        sites_df = pd.DataFrame(sites_list)
        
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
        
        if facility_manager in ['N/A', 'nan', '']:
            facility_manager = "Unassigned FM"
        
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

# --- REQUISITIONER FUNCTION with Project and Region support ---
def get_requisitioner_for_territory_and_project(territory, project, region):
    """Get requisitioner profile based on territory, project, and region"""
    if df_req_db is not None and 'Territory no.' in df_req_db.columns:
        matching_reqs = df_req_db[
            (df_req_db['Territory no.'] == territory) &
            (df_req_db['Project'] == project) &
            (df_req_db['Region'] == region)
        ]
        
        if not matching_reqs.empty:
            req_row = matching_reqs.iloc[0]
            return {
                "name": str(req_row.get("Name", "N/A")),
                "dept": str(req_row.get("Dept./Group", "N/A")),
                "id": format_id_number(req_row.get("ID #", "N/A")),
                "contact": format_contact_number(req_row.get("Contact No.", "N/A"))
            }
    
    if df_req_db is not None and 'Territory no.' in df_req_db.columns:
        matching_reqs = df_req_db[
            (df_req_db['Territory no.'] == territory) &
            (df_req_db['Region'] == region)
        ]
        
        if not matching_reqs.empty:
            req_row = matching_reqs.iloc[0]
            return {
                "name": str(req_row.get("Name", "N/A")),
                "dept": str(req_row.get("Dept./Group", "N/A")),
                "id": format_id_number(req_row.get("ID #", "N/A")),
                "contact": format_contact_number(req_row.get("Contact No.", "N/A"))
            }
    
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
    <small>Version 2.3.0</small><br>
    <small>© 2026 RAAWA Generator</small><br>
    <small>✨ Region Aware (LUZ, VIS, MIN)</small>
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
    
    st.markdown("## ⚡ Features")
    
    features = [
        ("🏗️ Unlimited Sites", "Select any number of sites from database with auto-batching", "✅"),
        ("👥 Personnel Management", "Manual, database, or mixed input modes", "✅"),
        ("📝 Project & Region Aware", "Filter requisitioners and teams by project/region (LUZ, VIS, MIN)", "✅"),
        ("🏢 Company Filtering", "Select personnel by company with region awareness", "✅"),
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
        st.markdown("""
        <div class="main-header">
            <h1>📄 Automated Multi-Site RAAWA Generator</h1>
            <p>Select any number of sites, populate the personnel manifest, and generate perfectly styled RAAWA forms instantly.<br>
            <strong>✨ Region Aware (LUZ, VIS, MIN) - Auto-batches into groups of 10 per signatory</strong></p>
        </div>
        """, unsafe_allow_html=True)
        
        # --- STEP 1: PROJECT & REGION SELECTION ---
        st.markdown("## 📋 Step 1: Project & Region Configuration")
        
        project_list = []
        
        if df_req_db is not None and 'Project' in df_req_db.columns:
            project_list = sorted(df_req_db['Project'].dropna().unique())
            project_list = [p for p in project_list if str(p).strip() != '' and str(p).strip() != 'nan']
        else:
            project_list = ['Default']
        
        col_project, col_region = st.columns(2)
        
        with col_project:
            selected_project = st.selectbox(
                "Select Project:",
                options=project_list,
                help="Choose the project to determine the requisitioner"
            )
        
        with col_region:
            if df_db is not None and 'REGION' in df_db.columns:
                site_regions = sorted(df_db['REGION'].dropna().unique())
                site_regions = [r for r in site_regions if str(r).strip() != '']
                site_regions = [r for r in site_regions if r in ['LUZ', 'VIS', 'MIN']]
                
                if not site_regions:
                    site_regions = ['LUZ', 'VIS', 'MIN']
            else:
                site_regions = ['LUZ', 'VIS', 'MIN']
            
            selected_region = st.selectbox(
                "Select Region:",
                options=['All Regions'] + site_regions,
                help="Filter sites by region (LUZ, VIS, MIN)"
            )
        
        st.markdown("---")
        
        # --- STEP 2: MULTI-SITE SELECTION INPUT ---
        st.markdown("## 📍 Step 2: Site Selection")
        st.markdown("Select the sites for your RAAWA request (unlimited number supported)")
        
        if selected_region != 'All Regions' and df_db is not None and 'REGION' in df_db.columns:
            filtered_db = df_db[df_db['REGION'].str.upper() == selected_region.upper()]
        else:
            filtered_db = df_db
        
        plaid_list = sorted(filtered_db['PLAID'].unique())
        selected_plaids = st.multiselect(
            "Search or Select Site IDs (PLAID):", 
            options=plaid_list,
            default=[]
        )
            
        matching_sites = filtered_db[filtered_db['PLAID'].isin(selected_plaids)]
        
        if not matching_sites.empty:
            conflicts = check_conflicts(matching_sites)
            
            if conflicts['territory_conflict']:
                st.warning("⚠️ **Territory Limitation Detected:** You have selected sites from both Territory 7 and Territory 8. These cannot be combined in a single RAAWA.")
            
            if conflicts['fm_conflict']:
                st.warning(f"⚠️ **Facility Manager Conflict Detected:** Multiple Facility Managers found ({', '.join(conflicts['unique_facility_managers'])}). Each unique Territory + Facility Manager combination will get its own RAAWA.")
            
            st.markdown("### 📋 Selected Sites Preview")
            
            preview_cols = ['PLAID', 'SITE', 'REGION', 'TERRITORY', 'NEW ENGINEER_AH', 'CONTACT NUMBER', 'SITE_ADD']
            available_cols = [col for col in preview_cols if col in matching_sites.columns]
            st.dataframe(
                matching_sites[available_cols],
                hide_index=True,
                use_container_width=True
            )
            
            unique_combos = get_unique_combinations(matching_sites)
            
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

        # --- STEP 3: ACCESS SCOPE & DEPLOYMENTS ---
        st.markdown("## 📝 Step 3: Access Scope & Deployment Details")
        
        scope_col, date_col = st.columns(2)
        with scope_col:
            scope_of_work = st.text_area(
                "Nature of Access / Detailed Scope of Work:",
                value="SITE SURVEY, INSTALLATION, INTEGRATION, AND ACCEPTANCE TESTING.",
                height=100
            )
        with date_col:
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
            
            validity_days = st.number_input(
                "Authorization Validity (Days):", 
                min_value=1, 
                max_value=365, 
                value=30,
                help="Number of days the authorization will be valid from the start date"
            )
            
            end_date = start_date + timedelta(days=int(validity_days))
            
            st.success(f"📅 Form window coverage: **{start_date.strftime('%Y-%m-%d')}** to **{end_date.strftime('%Y-%m-%d')}**")
            st.caption(f"Total validity: {validity_days} days")

        st.markdown("## 👥 Step 4: Team Manifest")
        st.markdown("Add the engineers and technicians who will perform the work")
        
        # --- PERSONNEL SELECTION METHOD ---
        col_method, col_empty = st.columns([2, 3])
        with col_method:
            selection_method = st.radio(
                "Personnel Selection Method:",
                ["Manual Input", "Load from Database", "Mixed (Database + Manual)"],
                horizontal=True
            )
        
        if 'personnel_list' not in st.session_state:
            st.session_state.personnel_list = []
        
        if selection_method == "Load from Database":
            if df_engr_tech_db is None or df_engr_tech_db.empty:
                st.warning("⚠️ No personnel records found in EngrTech.xlsx")
                st.info("""
                **Expected format for EngrTech.xlsx:**
                | Name | Company | ID No | Region |
                |------|---------|-------|--------|
                | John Doe | Nokia | 12345 | LUZ |
                | Jane Smith | Huawei | 67890 | VIS |
                """)
            else:
                st.success(f"📋 Found {len(df_engr_tech_db)} personnel records in database")
                
                with st.expander("📁 Select Personnel from Database", expanded=True):
                    # Get unique companies and regions
                    company_list = sorted(df_engr_tech_db['Company'].unique())
                    company_list = [c for c in company_list if str(c).strip() != '']
                    
                    region_list = []
                    if 'Region' in df_engr_tech_db.columns:
                        region_list = sorted(df_engr_tech_db['Region'].unique())
                        region_list = [r for r in region_list if str(r).strip() != '']
                    
                    # Filters in columns
                    col_filters = st.columns(3)
                    
                    with col_filters[0]:
                        if company_list:
                            selected_companies = st.multiselect(
                                "Filter by Company:",
                                options=['All Companies'] + company_list,
                                default=['All Companies'],
                                key="company_filter_main"
                            )
                        else:
                            selected_companies = ['All Companies']
                            st.info("No companies found")
                    
                    with col_filters[1]:
                        if region_list:
                            selected_regions = st.multiselect(
                                "Filter by Region:",
                                options=['All Regions'] + region_list,
                                default=['All Regions'],
                                key="region_filter_main"
                            )
                        else:
                            selected_regions = ['All Regions']
                            st.info("No regions found")
                    
                    with col_filters[2]:
                        search_term = st.text_input("🔍 Search:", key="personnel_search_main")
                    
                    # Apply filters
                    filtered_engr = df_engr_tech_db.copy()
                    
                    # Company filter
                    if selected_companies and 'All Companies' not in selected_companies:
                        filtered_engr = filtered_engr[filtered_engr['Company'].isin(selected_companies)]
                    
                    # Region filter
                    if selected_regions and 'All Regions' not in selected_regions:
                        if 'Region' in filtered_engr.columns:
                            filtered_engr = filtered_engr[filtered_engr['Region'].isin(selected_regions)]
                    
                    # Search filter
                    if search_term:
                        filtered_engr = filtered_engr[
                            filtered_engr['Name'].astype(str).str.contains(search_term, case=False, na=False) |
                            filtered_engr['Company'].astype(str).str.contains(search_term, case=False, na=False) |
                            filtered_engr['ID No'].astype(str).str.contains(search_term, case=False, na=False)
                        ]
                    
                    # Show count
                    st.info(f"Showing {len(filtered_engr)} of {len(df_engr_tech_db)} personnel")
                    
                    # Display filtered data - safe column selection
                    display_cols = []
                    if 'Name' in filtered_engr.columns:
                        display_cols.append('Name')
                    if 'Company' in filtered_engr.columns:
                        display_cols.append('Company')
                    if 'ID No' in filtered_engr.columns:
                        display_cols.append('ID No')
                    if 'Region' in filtered_engr.columns:
                        display_cols.append('Region')
                    
                    if not filtered_engr.empty and display_cols:
                        st.dataframe(
                            filtered_engr[display_cols],
                            hide_index=True,
                            use_container_width=True
                        )
                        
                        # Multi-select with filtered data
                        selected_indices = st.multiselect(
                            "Select Engineers/Technicians to add:",
                            options=filtered_engr.index.tolist(),
                            format_func=lambda x: f"{filtered_engr.loc[x, 'Name']} - {filtered_engr.loc[x, 'Company']} (ID: {filtered_engr.loc[x, 'ID No']})",
                            key="selected_personnel_db_main"
                        )
                        
                        col_add, col_clear = st.columns([1, 4])
                        with col_add:
                            if st.button("➕ Add Selected to Manifest", use_container_width=True):
                                if selected_indices:
                                    added_count = 0
                                    for idx in selected_indices:
                                        person = filtered_engr.loc[idx]
                                        if not any(p['name'] == person['Name'] for p in st.session_state.personnel_list):
                                            st.session_state.personnel_list.append({
                                                "name": person['Name'],
                                                "company": person['Company'],
                                                "id_no": person['ID No']
                                            })
                                            added_count += 1
                                    st.success(f"Added {added_count} personnel to manifest!")
                                    st.rerun()
                                else:
                                    st.warning("Please select personnel first")
                    else:
                        st.warning("No personnel match the current filters")
                
                # Manual entry section
                st.markdown("---")
                st.subheader("Or Add Manually")
                
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
                with st.expander("📁 Select from Database", expanded=True):
                    # Get unique companies and regions
                    company_list = sorted(df_engr_tech_db['Company'].unique())
                    company_list = [c for c in company_list if str(c).strip() != '']
                    
                    region_list = []
                    if 'Region' in df_engr_tech_db.columns:
                        region_list = sorted(df_engr_tech_db['Region'].unique())
                        region_list = [r for r in region_list if str(r).strip() != '']
                    
                    # Filters
                    col_filters = st.columns(3)
                    
                    with col_filters[0]:
                        if company_list:
                            selected_companies = st.multiselect(
                                "Filter by Company:",
                                options=['All Companies'] + company_list,
                                default=['All Companies'],
                                key="company_filter_mixed"
                            )
                        else:
                            selected_companies = ['All Companies']
                    
                    with col_filters[1]:
                        if region_list:
                            selected_regions = st.multiselect(
                                "Filter by Region:",
                                options=['All Regions'] + region_list,
                                default=['All Regions'],
                                key="region_filter_mixed"
                            )
                        else:
                            selected_regions = ['All Regions']
                    
                    with col_filters[2]:
                        search_term = st.text_input("🔍 Search:", key="mixed_search")
                    
                    # Apply filters
                    filtered_engr = df_engr_tech_db.copy()
                    
                    if selected_companies and 'All Companies' not in selected_companies:
                        filtered_engr = filtered_engr[filtered_engr['Company'].isin(selected_companies)]
                    
                    if selected_regions and 'All Regions' not in selected_regions:
                        if 'Region' in filtered_engr.columns:
                            filtered_engr = filtered_engr[filtered_engr['Region'].isin(selected_regions)]
                    
                    if search_term:
                        filtered_engr = filtered_engr[
                            filtered_engr['Name'].astype(str).str.contains(search_term, case=False, na=False) |
                            filtered_engr['Company'].astype(str).str.contains(search_term, case=False, na=False) |
                            filtered_engr['ID No'].astype(str).str.contains(search_term, case=False, na=False)
                        ]
                    
                    if not filtered_engr.empty:
                        # Safe column selection
                        display_cols = []
                        if 'Name' in filtered_engr.columns:
                            display_cols.append('Name')
                        if 'Company' in filtered_engr.columns:
                            display_cols.append('Company')
                        if 'ID No' in filtered_engr.columns:
                            display_cols.append('ID No')
                        if 'Region' in filtered_engr.columns:
                            display_cols.append('Region')
                        
                        if display_cols:
                            st.dataframe(
                                filtered_engr[display_cols],
                                hide_index=True,
                                use_container_width=True
                            )
                        
                        selected_indices = st.multiselect(
                            "Select personnel:",
                            options=filtered_engr.index.tolist(),
                            format_func=lambda x: f"{filtered_engr.loc[x, 'Name']} - {filtered_engr.loc[x, 'Company']}",
                            key="mixed_selection"
                        )
                        
                        if st.button("➕ Add Selected", key="add_mixed"):
                            if selected_indices:
                                added_count = 0
                                for idx in selected_indices:
                                    person = filtered_engr.loc[idx]
                                    if not any(p['name'] == person['Name'] for p in st.session_state.personnel_list):
                                        st.session_state.personnel_list.append({
                                            "name": person['Name'],
                                            "company": person['Company'],
                                            "id_no": person['ID No']
                                        })
                                        added_count += 1
                                st.success(f"Added {added_count} personnel!")
                                st.rerun()
                            else:
                                st.warning("Please select personnel first")
            
            st.markdown("---")
            st.subheader("Manual Addition")
            
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
            
            manifest_df = pd.DataFrame(st.session_state.personnel_list)
            manifest_df['id_no'] = manifest_df['id_no'].apply(format_id_number)
            st.dataframe(manifest_df, hide_index=True, use_container_width=True)
            
            if st.button("🗑️ Clear All Personnel"):
                st.session_state.personnel_list = []
                st.rerun()
        else:
            st.info("No personnel added yet. Use the methods above to add team members.")
        
        st.markdown("---")

        # --- STEP 5: LOGIC EXECUTION & RENDER ---
        if st.button("🚀 Build Multi-Site RAAWA Spreadsheet", type="primary"):
            if not st.session_state.personnel_list:
                st.warning("Please add at least one personnel entry to populate the work manifest.")
            else:
                try:
                    site_groups = split_sites_by_territory_and_fm(matching_sites)
                    
                    total_groups = len(site_groups)
                    st.info(f"📋 Generating **{total_groups} RAAWA files**...")
                    
                    file_details = []
                    
                    for group_info in site_groups:
                        group_sites = group_info['dataframe']
                        territory = group_info['territory']
                        facility_manager = group_info['facility_manager']
                        batch_num = group_info['batch_num']
                        total_batches = group_info['total_batches']
                        
                        group_region = selected_region
                        if selected_region == 'All Regions':
                            group_region = group_sites.iloc[0].get('REGION', 'MIN')
                            group_region = group_region if group_region in ['LUZ', 'VIS', 'MIN'] else 'MIN'
                        
                        req_profile = get_requisitioner_for_territory_and_project(
                            territory, 
                            selected_project, 
                            group_region
                        )
                        
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
                        
                        clean_fm = facility_manager.replace(" ", "_").replace("/", "_").replace(",", "")
                        if total_batches > 1:
                            filename = f"RAAWA_{selected_project}_{group_region}_Territory{territory}_{clean_fm}_Batch{batch_num}of{total_batches}_{group_sites.iloc[0]['PLAID']}.xlsx"
                            display_name = f"{selected_project} - {group_region} - Territory {territory} - {facility_manager[:25]} (Batch {batch_num}/{total_batches})"
                        else:
                            filename = f"RAAWA_{selected_project}_{group_region}_Territory{territory}_{clean_fm}_{group_sites.iloc[0]['PLAID']}.xlsx"
                            display_name = f"{selected_project} - {group_region} - Territory {territory} - {facility_manager[:30]}"
                        
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
                        
                        st.success(f"✅ Created RAAWA #{len(file_details)}: {selected_project} - {group_region} - Territory {territory} - {facility_manager[:40]} ({len(group_sites)} sites){' - Batch ' + str(batch_num) + '/' + str(total_batches) if total_batches > 1 else ''} | Requisitioner: {req_profile['name']}")
                    
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
                - 📋 **Project:** {selected_project}
                - 🌍 **Region:** {selected_region if selected_region != 'All Regions' else file_info.get('region', 'MIN')}
                - 🏢 **Facility Manager:** {file_info['facility_manager']}
                - 👤 **Requisitioner:** {file_info['requisitioner']}
                - 📍 **Territory:** {file_info['territory']}
                """)
                
            else:
                st.success(f"🎉 Successfully generated {len(file_details)} RAAWA files!")
                st.markdown(f"**Project:** {selected_project} | **Region:** {selected_region if selected_region != 'All Regions' else 'Multiple'}")
                
                with st.expander("📁 View All Generated Files", expanded=True):
                    for idx, file_info in enumerate(file_details):
                        st.markdown(f"**File #{idx + 1}: {selected_project} - {selected_region if selected_region != 'All Regions' else 'Multiple'} - Territory {file_info['territory']}{file_info['batch_info']}**")
                        
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
            
            col_clear, col_spacer = st.columns([1, 3])
            with col_clear:
                if st.button("🔄 Clear Files & Start New RAAWA", use_container_width=True):
                    st.session_state['files_generated'] = False
                    st.session_state['generated_files'] = []
                    st.session_state.personnel_list = []
                    st.rerun()
    else:
        st.error("Failed to load databases. Please check that all required Excel files are present.")