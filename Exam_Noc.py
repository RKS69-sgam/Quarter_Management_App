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
# File name updated to 'Exam NOC Letter temp.docx'
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "Exam NOC Letter temp.docx")

SICK_COLLECTION = "sickemp"
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
    [span_0](start_span)[span_1](start_span)"""Formatting barkaraar rakhne ke liye Runs replacement[span_0](end_span)[span_1](end_span)"""
    for p in paragraphs:
        for key, value in data.items():
            placeholder = f"[{key}]"
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, str(value))

def generate_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Template not found: {template_path}")
        return None
    try:
        doc = Document(template_path)
        # 1. Normal Paragraphs
        safe_replace(doc.paragraphs, data)
        # 2. [span_2](start_span)Tables (Header and Date Table)[span_2](end_span)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    safe_replace(cell.paragraphs, data)
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Word file Error: {e}")
        return None

# =================================================================
# --- 2. AUTHENTICATION (Sgam@4321) ---
# =================================================================
st.set_page_config(layout="wide", page_title="Railway NOC System")

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
                st.error("Invalid Password!")
    st.stop()

# =================================================================
# --- 3. UI LOGIC ---
# =================================================================
st.header("📋 Exam NOC Letter Taiyar Karein")
df_emp = get_employees()

if not df_emp.empty:
    emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
    selected = st.selectbox("Karmchari Chunein", emp_list)
    
    h_id = selected.split('(')[-1].strip(')')
    emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

    # Session storage for download button
    if 'noc_doc' not in st.session_state:
        st.session_state.noc_doc = None
    if 'hid_last' not in st.session_state:
        st.session_state.hid_last = ""

    with st.form(key=f"noc_form_v2_{h_id}"):
        l_date = st.date_input("Letter Date", value=datetime.now())
        
        # [span_3](start_span)[span_4](start_span)NOC Mapping[span_3](end_span)[span_4](end_span)
        memo_data = {
            "LetterDate": l_date.strftime("%d/%m/%Y"),
            "EmployeeName": emp_data.get('Employee Name in Hindi', emp_data.get('Employee Name', '')),
            "Designation": emp_data.get('Designation in Hindi', emp_data.get('Designation', '')),
            "UnitNumber": emp_data.get('UNIT No.', ''),
            "PFNumber": emp_data.get('PF Number', '')
        }
        
        if st.form_submit_button("Generate NOC"):
            out = generate_docx(TEMPLATE_PATH, memo_data)
            if out:
                st.session_state.noc_doc = out
                st.session_state.hid_last = h_id
                st.success("✅ NOC generate ho gaya! Niche se download karein.")

    # Download link outside form
    if st.session_state.noc_doc and st.session_state.hid_last == h_id:
        st.download_button(
            label="📥 Download NOC Letter",
            data=st.session_state.noc_doc,
            file_name=f"NOC_{h_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.warning("No data found.")
