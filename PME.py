import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Setup ---
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

# --- 1. Utilities ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def calculate_service(doa_val):
    if not doa_val or str(doa_val).lower() == 'nan': return "0", "0"
    try:
        doa = pd.to_datetime(doa_val).to_pydatetime()
        diff = relativedelta(datetime.now(), doa)
        return str(diff.years), str(diff.months)
    except: return "0", "0"

def replace_in_paragraphs(paragraphs, data):
    """Paragraphs ke andar runs ko preserve karke text replace karna"""
    for p in paragraphs:
        for key, value in data.items():
            placeholder = "{{" + key + "}}"
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        # 1. Replace in main body paragraphs
        replace_in_paragraphs(doc.paragraphs, data)
        # 2. Replace in all tables (Template contains tables for DOB/DOA)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs, data)
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Doc Generation Error: {e}"); return None

# --- 2. MAIN UI ---
st.set_page_config(layout="wide", page_title="PME System")

if 'pme_file' not in st.session_state: st.session_state.pme_file = None

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME Records", "🛠 Update PME"])
df_emp = get_employees()

with tab1:
    st.header("PME Memo Generation")
    if not df_emp.empty:
        # [span_1](start_span)HRMS ID based selection[span_1](end_span)
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            s_year, s_month = calculate_service(emp_data.get('DOA'))
            
            # [span_2](start_span)Template ke placeholders ke hisaab se mapping[span_2](end_span)
            pme_vals = {
                "dob": emp_data.get('DOB', ''),
                "doa": emp_data.get('DOA', ''),
                "name": emp_data.get('Employee Name', ''),
                "age": emp_data.get('Age', ''),
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

        if st.session_state.pme_file:
            st.download_button("📥 Download PME Memo", st.session_state.pme_file, f"PME_{h_id}.docx")
    else: st.warning("No data found.")

with tab3:
    st.header("🛠 Update PME & Employee Marks")
    if not df_emp.empty:
        t_sel = st.selectbox("Select Employee", emp_list, key="upd_pme")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("upd_pme_f"):
            m1 = st.text_input("Physical Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = st.text_input("Physical Mark 2", t_row.get('Physical Mark 2', ''))
            lp_date = st.text_input("Last Examination Date (DD/MM/YYYY)", t_row.get('Last PME', ''))
            lp_place = st.text_input("Examination Place", t_row.get('Last PME Place', 'NKJ'))
            
            if st.form_submit_button("Update Records"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_date, "Last PME Place": lp_place
                })
                st.success("✅ Database Updated!")
                st.rerun()
