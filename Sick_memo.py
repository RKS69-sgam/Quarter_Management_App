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
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "SICK MEMO temp.docx")

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
# --- 1. UTILITIES (FORMATTING FIX) ---
# =================================================================
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def get_sick_records():
    docs = db.collection(SICK_COLLECTION).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    if data:
        df = pd.DataFrame(data)
        if 'Created' in df.columns:
            df = df.sort_values(by='Created', ascending=False)
        return df
    return pd.DataFrame()

def safe_replace(paragraphs, data):
    """Word formatting barkaraar rakhne ke liye 'runs' ka istemal"""
    for p in paragraphs:
        for key, value in data.items():
            placeholder = f"[{key}]"
            if placeholder in p.text:
                # Sirf un 'runs' mein text badle jahan placeholder hai
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
                # Double check agar placeholder runs ke beech mein split ho gaya ho
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
        
        # 2. Tables
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
st.set_page_config(layout="wide", page_title="Railway Sick Management")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Sick Management Login")
    with st.form("login"):
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
tab1, tab2 = st.tabs(["📝 Sick Memo Generate", "📊 Report & Return Update"])

with tab1:
    st.header("📋 Sick Memo Taiyar Karein")
    df_emp = get_employees()
    
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Karmchari Chunein", emp_list)
        
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        if 'memo_docx' not in st.session_state:
            st.session_state.memo_docx = None
        if 'current_h_id' not in st.session_state:
            st.session_state.current_h_id = ""

        with st.form(key=f"memo_form_{h_id}"):
            c1, c2 = st.columns(2)
            memo_date = c1.date_input("Memo Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "OTHER"])
            
            # HINDI NAME MAPPING
            memo_data = {
                "LetterDate": memo_date.strftime("%d/%m/%Y"),
                "EmployeeName": emp_data.get('Employee Name in Hindi', emp_data.get('Employee Name', '')),
                "Designation": emp_data.get('Designation in Hindi', emp_data.get('Designation', '')),
                "UnitNumber": emp_data.get('UNIT No.', '')
            }
            
            if st.form_submit_button("Generate & Save Record"):
                db.collection(SICK_COLLECTION).add({
                    "HRMS_ID": h_id,
                    "Name": memo_data["EmployeeName"],
                    "Designation": memo_data["Designation"],
                    "SickDate": str(memo_date),
                    "Hospital": hospital,
                    "Status": "SICK",
                    "ReturnDate": None,
                    "Created": datetime.now()
                })
                
                docx_out = generate_docx(TEMPLATE_PATH, memo_data)
                if docx_out:
                    st.session_state.memo_docx = docx_out
                    st.session_state.current_h_id = h_id
                    st.success("✅ Data Save ho gaya! Niche se download karein.")

        if st.session_state.memo_docx and st.session_state.current_h_id == h_id:
            st.download_button(
                label="📥 Download Sick Memo (DOCX)",
                data=st.session_state.memo_docx,
                file_name=f"Sick_Memo_{h_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Database mein koi data nahi mila.")

with tab2:
    st.header("📊 Sick Reports & History")
    df_sick = get_sick_records()
    
    if not df_sick.empty:
        st.subheader("🔄 Update Return (Mark FIT)")
        active_sick = df_sick[df_sick['Status'] == 'SICK']
        
        if not active_sick.empty:
            sick_options = active_sick.apply(lambda r: f"{r['Name']} (Sick: {r['SickDate']})", axis=1).tolist()
            returning = st.selectbox("Select Employee", sick_options)
            ret_date = st.date_input("Return (FIT) Date", value=datetime.now())
            
            if st.button("Confirm Return"):
                idx = sick_options.index(returning)
                doc_id = active_sick.iloc[idx]['id']
                db.collection(SICK_COLLECTION).document(doc_id).update({
                    "Status": "FIT",
                    "ReturnDate": str(ret_date)
                })
                st.success("Status Updated.")
                st.rerun()
        
        st.divider()
        st.subheader("📑 History")
        st.dataframe(df_sick.drop(columns=['id', 'Created'], errors='ignore'), use_container_width=True)
    else:
        st.info("No sick records yet.")
