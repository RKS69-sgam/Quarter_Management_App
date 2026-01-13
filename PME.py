import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# =================================================================
# --- 0. PATH & FIREBASE SETUP ---
# =================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "pme memo temp.docx")

EMP_COLLECTION = "employees"
PME_HISTORY_COLLECTION = "pme_history"

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
            st.error(f"Firebase Error: {e}")
            st.stop()
    return firestore.client()

db = init_db()

# =================================================================
# --- 1. UTILITIES ---
# =================================================================
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_pme_history():
    docs = db.collection(PME_HISTORY_COLLECTION).order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def calculate_service(doa_val):
    """Appointment date se aaj tak ka service period nikalne ke liye"""
    try:
        # Agar DOA string hai toh convert karein
        if isinstance(doa_val, str):
            doa = datetime.strptime(doa_val, "%Y-%m-%d")
        else:
            doa = datetime.combine(doa_val, datetime.min.time())
            
        now = datetime.now()
        diff = relativedelta(now, doa)
        return diff.years, diff.months
    except Exception:
        return 0, 0

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path):
        return None
    try:
        doc = Document(template_path)
        # [span_0](start_span)Template placeholders replacement logic[span_0](end_span)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data.items():
                        placeholder = "{{ " + key + " }}"
                        alt_placeholder = "{{" + key + "}}"
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str(value))
                        if alt_placeholder in cell.text:
                            cell.text = cell.text.replace(alt_placeholder, str(value))
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Word Error: {e}")
        return None

# =================================================================
# --- 2. MAIN UI ---
# =================================================================
st.set_page_config(layout="wide", page_title="PME System")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 PME Management Login")
    with st.form("login"):
        u, p = st.text_input("User"), st.text_input("Pass", type="password")
        if st.form_submit_button("Login") and u == "admin" and p == "Sgam@4321":
            st.session_state.auth = True
            st.rerun()
    st.stop()

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME Records", "🛠 Update Physical Marks"])

df_emp = get_employees()

# --- TAB 1: GENERATE ---
with tab1:
    st.header("PME Memo Generation")
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            col1, col2 = st.columns(2)
            c_date = col1.date_input("Memo Date", value=datetime.now())
            
            # Service and Last Examined Date logic
            doa = emp_data.get('Date of Appointment', '')
            s_year, s_month = calculate_service(doa)
            
            # Fetching Last PME Date directly from 'employees' collection
            last_pme_from_db = emp_data.get('Last PME', 'N/A') 

            pme_data = {
                "name": emp_data.get('Employee Name', ''),
                "age": emp_data.get('Age', ''),
                "father_name": emp_data.get('Father Name', ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical Category', ''),
                "dob": emp_data.get('Date of Birth', ''),
                "doa": doa,
                "current_date": c_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                "last_examined_date": last_pme_from_db, 
                "last_place": "NKJ", 
                "examiner": "ACMS/NKJ",
                "service_year": s_year,
                "service_month": s_month
            }

            if st.form_submit_button("Generate & Save"):
                out = generate_pme_docx(TEMPLATE_PATH, pme_data)
                if out:
                    # History record save karna
                    db.collection(PME_HISTORY_COLLECTION).add({
                        **pme_data, "Timestamp": datetime.now(), "HRMS_ID": h_id
                    })
                    st.success("✅ Memo generated successfully!")
                    st.download_button("📥 Download Memo", out, f"PME_{h_id}.docx")
    else:
        st.warning("Employees database khali hai.")

# --- TAB 2 & 3: Records and Update Logic ---
with tab2:
    st.header("📊 PME History Report")
    df_h = get_pme_history()
    if not df_h.empty:
        st.dataframe(df_h[['Timestamp', 'name', 'designation', 'current_date', 'last_examined_date']], use_container_width=True)

with tab3:
    st.header("🛠 Update Employee Details")
    if not df_emp.empty:
        target_emp = st.selectbox("Select Employee to Update", emp_list, key="upd_pme")
        t_id = target_emp.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("marks_upd_pme"):
            colA, colB = st.columns(2)
            m1 = colA.text_input("Physical Mark 1", value=t_row.get('Physical Mark 1', ''))
            m2 = colB.text_input("Physical Mark 2", value=t_row.get('Physical Mark 2', ''))
            last_pme_input = st.text_input("Last PME Date (DD/MM/YYYY)", value=t_row.get('Last PME', ''))
            
            if st.form_submit_button("Save Updates"):
                db.collection(EMP_COLLECTION).document(t_row['id']).update({
                    "Physical Mark 1": m1,
                    "Physical Mark 2": m2,
                    "Last PME": last_pme_input
                })
                st.success("✅ Database updated!")
                st.rerun()

