import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# =================================================================
# --- 0. PATH & FIREBASE SETUP ---
# =================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# File name corrected as per your request
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "Exam NOC Letter temp.docx")

EMP_COLLECTION = "employees"

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
            st.error(f"Firebase Init Error: {e}")
            st.stop()
    return firestore.client()

db = init_db()

# =================================================================
# --- 1. UTILITIES ---
# =================================================================
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def safe_replace(paragraphs, data):
    # Formatting barkaraar rakhne ke liye runs ka istemal
    for p in paragraphs:
        for key, value in data.items():
            placeholder = f"[{key}]"
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
                # Double check for split placeholders
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, str(value))

def generate_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Template not found at: {template_path}")
        return None
    try:
        doc = Document(template_path)
        # 1. [span_0](start_span)Replace in normal paragraphs[span_0](end_span)
        safe_replace(doc.paragraphs, data)
        # 2. [span_1](start_span)Replace in tables (Header & Date)[span_1](end_span)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    safe_replace(cell.paragraphs, data)
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Word file error: {e}")
        return None

# =================================================================
# --- 2. AUTHENTICATION (Sgam@4321) ---
# =================================================================
st.set_page_config(layout="wide", page_title="Exam NOC System")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Exam NOC Login")
    with st.form("login_form"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Invalid Password")
    st.stop()

# =================================================================
# --- 3. MAIN UI ---
# =================================================================
st.header("📋 Exam NOC Letter Taiyar Karein")
df_emp = get_employees()

if not df_emp.empty:
    emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
    selected = st.selectbox("Karmchari Chunein", emp_list)
    
    h_id = selected.split('(')[-1].strip(')')
    emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

    # Session storage to avoid form errors
    if 'noc_out' not in st.session_state:
        st.session_state.noc_out = None
    if 'last_hid' not in st.session_state:
        st.session_state.last_hid = ""

    with st.form(key=f"noc_form_{h_id}"):
        l_date = st.date_input("Letter Date", value=datetime.now())
        
        # [span_2](start_span)NOC Mapping logic[span_2](end_span)
        memo_data = {
            "LetterDate": l_date.strftime("%d/%m/%Y"),
            "EmployeeName": emp_data.get('Employee Name in Hindi', emp_data.get('Employee Name', '')),
            "Designation": emp_data.get('Designation in Hindi', emp_data.get('Designation', '')),
            "UnitNumber": emp_data.get('UNIT No.', ''),
            "PFNumber": emp_data.get('PF Number', '')
        }
        
        if st.form_submit_button("Generate NOC"):
            result = generate_docx(TEMPLATE_PATH, memo_data)
            if result:
                st.session_state.noc_out = result
                st.session_state.last_hid = h_id
                st.success("✅ NOC taiyar ho gaya!")

    # Download Button (Outside Form)
    if st.session_state.noc_out and st.session_state.last_hid == h_id:
        st.download_button(
            label="📥 Download NOC Letter (DOCX)",
            data=st.session_state.noc_out,
            file_name=f"Exam_NOC_{h_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.warning("Database mein koi data nahi mila.")
