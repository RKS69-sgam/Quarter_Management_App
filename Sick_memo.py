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
# Assets folder ka sahi path set karein
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
                # Local testing ke liye (agar file root mein hai)
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

def get_sick_records():
    docs = db.collection(SICK_COLLECTION).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def generate_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Error: Template file nahi mili is path par: {template_path}")
        return None
    try:
        doc = Document(template_path)
        # Paragraphs mein placeholder replace karein
        for p in doc.paragraphs:
            for key, value in data.items():
                target = f"[{key}]"
                if target in p.text:
                    p.text = p.text.replace(target, str(value))
        
        # Tables mein placeholder replace karein
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data.items():
                        target = f"[{key}]"
                        if target in cell.text:
                            cell.text = cell.text.replace(target, str(value))
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Word file generate karne mein error: {e}")
        return None

# =================================================================
# --- 2. AUTHENTICATION (Sgam@4321) ---
# =================================================================
st.set_page_config(layout="wide", page_title="Sick Management System")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Login")
    with st.form("login_form"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Galt Password!")
    st.stop()

# =================================================================
# --- 3. MAIN UI ---
# =================================================================
tab1, tab2 = st.tabs(["📝 Sick Memo Generate", "📊 Report & Return Update"])

# --- TAB 1: GENERATE ---
with tab1:
    st.header("📋 Sick Memo Taiyar Karein")
    df_emp = get_employees()
    
    if not df_emp.empty:
        # Selection Box
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Karmchari Chunein", emp_list)
        
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
        
        with st.form(key=f"memo_form_{h_id}"):
            c1, c2 = st.columns(2)
            memo_date = c1.date_input("Memo Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "OTHER"])
            
            # Data Mapping for Template
            memo_data = {
                "LetterDate": memo_date.strftime("%d/%m/%Y"),
                "EmployeeName": emp_data.get('Employee Name', ''),
                "Designation": emp_data.get('Designation', ''),
                "UnitNumber": emp_data.get('UNIT No.', '')
            }
            
            if st.form_submit_button("Generate & Firebase mein Save karein"):
                # 1. Firebase Logging
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
                
                # 2. Document Creation
                docx_out = generate_docx(TEMPLATE_PATH, memo_data)
                if docx_out:
                    st.success("Data Save ho gaya!")
                    st.download_button("📥 Download Sick Memo", docx_out, f"Sick_Memo_{h_id}.docx")
    else:
        st.warning("Database mein koi karmchari nahi mila.")

# --- TAB 2: REPORTS & RETURN ---
with tab2:
    st.header("📊 Sick Reports")
    df_sick = get_sick_records()
    
    if not df_sick.empty:
        # Return Update Section
        st.subheader("🔄 Return Entry (Mark FIT)")
        active_sick = df_sick[df_sick['Status'] == 'SICK']
        
        if not active_sick.empty:
            sick_names = active_sick.apply(lambda r: f"{r['Name']} (Sick: {r['SickDate']})", axis=1).tolist()
            returning = st.selectbox("Kaun laut chuka hai?", sick_names)
            ret_date = st.date_input("Return Date")
            
            if st.button("Confirm Return"):
                idx = sick_names.index(returning)
                doc_id = active_sick.iloc[idx]['id']
                db.collection(SICK_COLLECTION).document(doc_id).update({
                    "Status": "FIT",
                    "ReturnDate": str(ret_date)
                })
                st.success("Status FIT kar diya gaya hai.")
                st.rerun()
        
        st.divider()
        st.subheader("Sabhi Records")
        st.dataframe(df_sick.drop(columns=['id', 'Created'], errors='ignore'), use_container_width=True)
    else:
        st.info("Abhi tak koi sick record nahi hai.")
