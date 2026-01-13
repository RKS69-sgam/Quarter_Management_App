import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Firebase Setup ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        try:
            if "firebase_config" in st.secrets:
                cred_dict = dict(st.secrets["firebase_config"])
                if isinstance(cred_dict.get('private_key'), str):
                    cred_dict['private_key'] = cred_dict['private_key'].replace('\\n', '\n')
                cred = credentials.Certificate(cred_dict)
            else:
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

db = init_db()

# --- 1. Utility Functions ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def calculate_service_and_age(dob_val, doa_val):
    """Age aur Service Length ka automatic calculation"""
    now = datetime.now()
    age, s_year, s_month = "0", "0", "0"
    try:
        if dob_val and str(dob_val) != 'nan':
            dob = pd.to_datetime(dob_val).to_pydatetime()
            age = str(relativedelta(now, dob).years)
        if doa_val and str(doa_val) != 'nan':
            doa = pd.to_datetime(doa_val).to_pydatetime()
            diff = relativedelta(now, doa)
            s_year, s_month = str(diff.years), str(diff.months)
    except: pass
    return age, s_year, s_month

def replace_text_in_paragraph(paragraph, data):
    """Font style preserve karne ke liye advanced replacement logic"""
    for key, value in data.items():
        placeholder = "{{" + key + "}}"
        if placeholder in paragraph.text:
            # Poore paragraph ke text mein replacement
            combined_text = "".join(run.text for run in paragraph.runs)
            new_text = combined_text.replace(placeholder, str(value))
            
            # Formatting bachane ke liye pehle saare runs clear karein
            for i in range(len(paragraph.runs)):
                paragraph.runs[i].text = ""
            
            # Pehle run mein naya text daalein (baaki formatting waisi rahegi)
            if paragraph.runs:
                paragraph.runs[0].text = new_text

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        # Body Paragraphs check karein
        for p in doc.paragraphs:
            replace_text_in_paragraph(p, data)
        # Table Cells (DOB, DOA template mein yahi hain) check karein
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_text_in_paragraph(p, data)
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Error: {e}"); return None

# --- 2. UI ---
st.set_page_config(layout="wide", page_title="Railway PME System")
df_emp = get_employees()

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME Records", "🛠 Update PME"])

with tab1:
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            # Auto-calculation logic
            age, s_year, s_month = calculate_service_and_age(emp_data.get('Date of Birth'), emp_data.get('Date of Appointment'))
            
            pme_vals = {
                "dob": emp_data.get('Date of Birth', ''),
                "doa": emp_data.get('Date of Appointment', ''),
                "name": emp_data.get('Employee Name', ''),
                "age": age,
                "father_name": emp_data.get('Father Name', ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical Category', ''),
                "current_date": c_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_data.get('Last PME', 'N/A'),
                "last_place": emp_data.get('Last PME Place', 'NKJ'),
                "examiner": "ACMS/NKJ",
                "service_year": s_year,
                "service_month": s_month
            }

            if st.form_submit_button("Generate Memo"):
                out = generate_pme_docx(TEMPLATE_PATH, pme_vals)
                if out:
                    st.session_state.pme_file = out
                    st.success("✅ Memo Taiyar hai!")

        if 'pme_file' in st.session_state and st.session_state.pme_file:
            st.download_button("📥 Download PME Memo", st.session_state.pme_file, f"PME_{h_id}.docx")

with tab3:
    st.header("🛠 Update Employee PME Data")
    if not df_emp.empty:
        t_sel = st.selectbox("Update Employee", emp_list, key="upd_pme_final")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("final_upd_form"):
            col1, col2 = st.columns(2)
            m1 = col1.text_input("Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = col2.text_input("Mark 2", t_row.get('Physical Mark 2', ''))
            lp_date = st.text_input("Last Examination Date (DD/MM/YYYY)", t_row.get('Last PME', ''))
            lp_place = st.text_input("Examination Place", t_row.get('Last PME Place', 'NKJ'))
            
            if st.form_submit_button("Save to Database"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_date, "Last PME Place": lp_place
                })
                st.success("✅ Records Updated!")
                st.rerun()
