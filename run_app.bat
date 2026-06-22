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
</style>
""", unsafe_allow_html=True)

def format_contact_number(contact):
    """Format contact number to start with 0"""
    if not contact or pd.isna(contact) or contact == 'N/A':
        return "N/A"
    contact_str = str(contact).strip()
    contact_str = ''.join(filter(str.isdigit, contact_str))
    if len(contact_str) == 10 and contact_str.startswith('9'):
        return '0' + contact_str
    elif len(contact_str) == 12 and contact_str.startswith('639'):
        return '0' + contact_str[2:]
    elif len(contact_str) == 11 and not contact_str.startswith('0'):
        return '0' + contact_str
    return contact_str

def format_id_number(id_num):
    """Format ID number to remove decimal places"""
    if not id_num or pd.isna(id_num) or id_num == 'N/A':
        return "N/A"
    try:
        if '.' in str(id_num):
            float_val = float(str(id_num))
            if float_val.is_integer():
                return str(int(float_val))
    except:
        pass
    return str(id_num).strip()

# --- DATABASE LOADING ---
@st.cache_data
def load_databases():
    try:
        df_sites = pd.read_excel("Globe FO Engr Conatct_Vendor.xlsx", sheet_name="MIN")
        df_sites['PLAID'] = df_sites['PLAID'].astype(str).str.strip()
        df_sites['TERRITORY'] = df_sites['TERRITORY'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        df_req = pd.read_excel("Requisitioner.xlsx", header=1, dtype=str)
        df_req.columns = df_req.columns.str.strip()
        if 'Territory no.' in df_req.columns:
            df_req['Territory no.'] = df_req['Territory no.'].astype(str).str.replace('Territory', '', case=False).str.strip()
        
        for idx, row in df_req.iterrows():
            if 'ID #' in df_req.columns:
                id_val = row.get('ID #', 'N/A')
                if not pd.isna(id_val):
                    df_req.at[idx, 'ID #'] = format_id_number(id_val)
            if 'Contact No.' in df_req.columns:
                contact_val = row.get('Contact No.', 'N/A')
                if not pd.isna(contact_val):
                    df_req.at[idx, 'Contact No.'] = format_contact_number(contact_val)
        
        try:
            df_engr_tech = pd.read_excel("EngrTech.xlsx", header=1, dtype=str)
            df_engr_tech.columns = df_engr_tech.columns.str.strip()
            
            if 'SEC ID' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'SEC ID': 'ID No'}, inplace=True)
            elif 'ID_NO' in df_engr_tech.columns:
                df_engr_tech.rename(columns={'ID_NO': 'ID No'}, inplace=True)
            
            if 'ID No' in df_engr_tech.columns:
                for idx, row in df_engr_tech.iterrows():
                    id_val = row.get('ID No', '')
                    if not pd.isna(id_val):
                        df_engr_tech.at[idx, 'ID No'] = format_id_number(id_val)
            
            df_engr_tech = df_engr_tech.dropna(subset=['Name'], how='all')
            df_engr_tech['Company'] = df_engr_tech.get('Company', pd.Series()).fillna('').astype(str)
            df_engr_tech['ID No'] = df_engr_tech.get('ID No', pd.Series()).fillna('').astype(str)
        except:
            df_engr_tech = pd.DataFrame(columns=['Name', 'Company', 'ID No'])
        
        return df_sites, df_req, df_engr_tech
    except Exception as e:
        st.error(f"Error loading databases: {e}")
        return None, None, None

def create_raawa_file(matching_sites, personnel_list, scope_of_work, start_date, end_date, req_profile, facility_manager, batch_num=1, total_batches=1):
    template_file = "MIN591__MANUAL RAAWA_APPLICATION_June8,2026.xlsx"
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active 
    
    ws["D3"].value = req_profile["name"]
    ws["D4"].value = req_profile["dept"]
    ws["G4"].value = format_id_number(req_profile["id"])
    ws["J4"].value = format_contact_number(req_profile["contact"])
    
    base_site_row = 6
    for idx, (_, row) in enumerate(matching_sites.iterrows()):
        curr_row = base_site_row + idx
        ws.cell(row=curr_row, column=1, value=f"{row.get('PLAID', '')} - {row.get('SITE', '')}")
        ws.cell(row=curr_row, column=4, value=row.get("SITE_ADD", "N/A"))
    
    for r in range(base_site_row + len(matching_sites), 16):
        ws.row_dimensions[r].hidden = True
    
    ws["D17"].value = start_date.strftime("%Y-%m-%d")
    ws["E17"].value = end_date.strftime("%Y-%m-%d")
    
    start_personnel_row = 19
    
    def get_font_size(text, min_size=6, max_size=10):
        if not text or text == '':
            return max_size
        length = len(str(text))
        if length <= 12: return max_size
        elif length <= 18: return max_size - 1
        elif length <= 25: return max_size - 2
        elif length <= 32: return max_size - 3
        elif length <= 40: return max_size - 4
        else: return min_size
    
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
    
    if total_batches > 1:
        ws["A41"].value = f"{scope_of_work}\n\n(Page {batch_num} of {total_batches})"
    else:
        ws["A41"].value = scope_of_work
    
    original_signatory = ws["A48"].value
    if original_signatory:
        ws["A48"].value = str(original_signatory).replace("NEW ENGINEER_AH", facility_manager)
    else:
        ws["A48"].value = f"{facility_manager}\nSignature Over Printed Name / Date"
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def split_sites_by_territory_and_fm(matching_sites):
    unique_combos = {}
    for _, row in matching_sites.iterrows():
        territory = str(row.get("TERRITORY", "")).strip()
        facility_manager = str(row.get("NEW ENGINEER_AH", "")).strip()
        if facility_manager in ['N/A', 'nan', '']:
            facility_manager = "Unassigned FM"
        key = f"{territory}_{facility_manager}"
        if key not in unique_combos:
            unique_combos[key] = {'territory': territory, 'facility_manager': facility_manager, 'sites': []}
        unique_combos[key]['sites'].append(row.to_dict())
    
    final_groups = []
    for combo in unique_combos.values():
        sites_df = pd.DataFrame(combo['sites'])
        if len(sites_df) > 10:
            num_batches = math.ceil(len(sites_df) / 10)
            for batch_num in range(num_batches):
                start_idx = batch_num * 10
                end_idx = min((batch_num + 1) * 10, len(sites_df))
                final_groups.append({
                    'dataframe': sites_df.iloc[start_idx:end_idx],
                    'territory': combo['territory'],
                    'facility_manager': combo['facility_manager'],
                    'batch_num': batch_num + 1,
                    'total_batches': num_batches
                })
        else:
            final_groups.append({
                'dataframe': sites_df,
                'territory': combo['territory'],
                'facility_manager': combo['facility_manager'],
                'batch_num': 1,
                'total_batches': 1
            })
    return final_groups

# --- MAIN APP ---
df_db, df_req_db, df_engr_tech_db = load_databases()

if df_db is not None and df_req_db is not None:
    st.markdown("""
    <div class="main-header">
        <h1>📄 Automated Multi-Site RAAWA Generator</h1>
        <p>Select any number of sites, populate the personnel manifest, and generate perfectly styled RAAWA forms instantly.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Site Selection
    st.markdown("## 📍 Step 1: Site Selection")
    plaid_list = sorted(df_db['PLAID'].unique())
    selected_plaids = st.multiselect("Search or Select Site IDs (PLAID):", options=plaid_list, default=[])
    matching_sites = df_db[df_db['PLAID'].isin(selected_plaids)]
    
    if not matching_sites.empty:
        st.dataframe(matching_sites[['PLAID', 'SITE', 'REGION', 'TERRITORY', 'NEW ENGINEER_AH', 'SITE_ADD']], hide_index=True, use_container_width=True)
    else:
        st.info("💡 Please select at least one site to begin")
        st.stop()
    
    st.markdown("---")
    
    # Access Scope
    st.markdown("## 📝 Step 2: Access Scope")
    scope_of_work = st.text_area("Scope of Work:", value="SITE SURVEY, INSTALLATION, INTEGRATION, AND ACCEPTANCE TESTING.", height=100)
    
    start_date = st.date_input("Start Date:", value=datetime.now().date())
    start_date = datetime.combine(start_date, datetime.min.time())
    validity_days = st.number_input("Validity (Days):", min_value=1, max_value=365, value=30)
    end_date = start_date + timedelta(days=int(validity_days))
    st.success(f"Valid from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    
    st.markdown("---")
    
    # Personnel
    st.markdown("## 👥 Step 3: Team Manifest")
    
    if 'personnel_list' not in st.session_state:
        st.session_state.personnel_list = []
    
    col1, col2, col3 = st.columns(3)
    with col1:
        name = st.text_input("Name:", key="new_name")
    with col2:
        company = st.text_input("Company:", key="new_company")
    with col3:
        id_no = st.text_input("ID No.:", key="new_id")
    
    if st.button("➕ Add Personnel"):
        if name:
            st.session_state.personnel_list.append({"name": name, "company": company, "id_no": id_no})
            st.rerun()
    
    if st.session_state.personnel_list:
        manifest_df = pd.DataFrame(st.session_state.personnel_list)
        st.dataframe(manifest_df, hide_index=True, use_container_width=True)
        if st.button("🗑️ Clear All"):
            st.session_state.personnel_list = []
            st.rerun()
    
    st.markdown("---")
    
    # Generate
    if st.button("🚀 Generate RAAWA", type="primary"):
        if not st.session_state.personnel_list:
            st.warning("Please add at least one personnel entry.")
        else:
            try:
                site_groups = split_sites_by_territory_and_fm(matching_sites)
                
                def get_requisitioner(territory):
                    matching = df_req_db[df_req_db['Territory no.'] == territory]
                    if not matching.empty:
                        row = matching.iloc[0]
                        return {
                            "name": str(row.get("Name", "N/A")),
                            "dept": str(row.get("Dept./Group", "N/A")),
                            "id": str(row.get("ID #", "N/A")),
                            "contact": str(row.get("Contact No.", "N/A"))
                        }
                    return {"name": f"Territory {territory} Engineer", "dept": f"TERRITORY {territory}", "id": "N/A", "contact": "N/A"}
                
                for group in site_groups:
                    req_profile = get_requisitioner(group['territory'])
                    buffer = create_raawa_file(
                        group['dataframe'], st.session_state.personnel_list, scope_of_work,
                        start_date, end_date, req_profile, group['facility_manager'],
                        group['batch_num'], group['total_batches']
                    )
                    
                    filename = f"RAAWA_Territory{group['territory']}_{group['facility_manager'].replace(' ', '_')}.xlsx"
                    st.download_button(label=f"📥 Download {filename}", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.success(f"✅ Generated {len(site_groups)} RAAWA file(s)!")
            except Exception as e:
                st.error(f"Error: {e}")
else:
    st.error("Failed to load databases. Please check that all required Excel files are present.")