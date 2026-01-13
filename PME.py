import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. PATH & FIREBASE SETUP ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "pme memo temp.docx")

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

# --- 1. UTILITIES ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def calculate_service(doa_val):
    """Auto calculate years and months from DOA"""
    try:
        if not doa_val or doa_val == 'N/A': return 0, 0
        doa = pd.to_datetime(doa_val).to_pydatetime()
        diff = relativedelta(datetime.now(), doa)
        return diff.years, diff.months
    except: return 0, 0

def replace_text_preserve_format(doc, data):
    [span_0](start_span)[span_1](start_span)"""Font style barkaraar rakhne ke liye logic[span_0](end_span)[span_1](end_span)"""
    for key, value in data.items():
        placeholder = "{{" + key + "}}"
        # Paragraphs mein search karein
        for p in doc.paragraphs:
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
        # Tables mein search karein
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            for run in p.runs:
                                if placeholder in run.text:
                                    run.text = run.text.replace(placeholder, str(value))

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        replace_text_preserve_format(doc, data)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Doc Error: {e}"); return None

# --- 2. MAIN UI ---
st.set_page_config(layout="wide", page_title="PME Management")

if 'pme_file' not in st.session_state: st.session_state.pme_file = None

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME Records", "🛠 Update PME"])
df_emp = get_employees()

with tab1:
    st.header("PME Memo Generation")
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            s_year, s_month = calculate_service(emp_data.get('Date of Appointment', ''))
            
            # [span_2](start_span)Mapping all fields from your template[span_2](end_span)
            pme_vals = {
                "name": emp_data.get('Employee Name', ''),
                "age": emp_data.get('Age', ''),
                "father_name": emp_data.get('Father Name', ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical Category', ''),
                "dob": emp_data.get('Date of Birth', ''),
                "doa": emp_data.get('Date of Appointment', ''),
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
                    db.collection("pme_history").add({**pme_vals, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                    st.session_state.pme_file = out
                    st.success("✅ Memo Taiyar hai!")

        if st.session_state.pme_file:
            st.download_button("📥 Download PME Memo", st.session_state.pme_file, f"PME_{h_id}.docx")

with tab3:
    st.header("🛠 Update PME & Marks")
    if not df_emp.empty:
        t_sel = st.selectbox("Select Employee to Update", emp_list, key="upd_pme")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("update_pme_form"):
            col1, col2 = st.columns(2)
            m1 = col1.text_input("Physical Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = col2.text_input("Physical Mark 2", t_row.get('Physical Mark 2', ''))
            lp_date = st.text_input("Last PME Date (DD/MM/YYYY)", t_row.get('Last PME', ''))
            lp_place = st.text_input("Last Examination Place", t_row.get('Last PME Place', 'NKJ'))
            
            if st.form_submit_button("Update Records"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_date, "Last PME Place": lp_place
                })
                st.success("✅ Database Updated!")
                st.rerun()
