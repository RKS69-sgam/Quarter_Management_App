import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import firebase_admin
from firebase_admin import credentials, firestore

# =================================================================
# --- 0. FIREBASE SETUP ---
# =================================================================
SERVICE_ACCOUNT_FILE = 'sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json' 
SICK_COLLECTION = "sickemp"
EMP_COLLECTION = "employees"

@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        try:
            if st.secrets.get("firebase_config"):
                cred_dict = dict(st.secrets["firebase_config"])
                if isinstance(cred_dict.get('private_key'), str):
                    cred_dict['private_key'] = cred_dict['private_key'].replace('\\n', '\n')
                cred = credentials.Certificate(cred_dict)
            else:
                cred = credentials.Certificate(SERVICE_ACCOUNT_FILE)
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}")
    return firestore.client()

db = init_db()

# =================================================================
# --- 1. UTILITIES ---
# =================================================================
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_sick_records():
    docs = db.collection(SICK_COLLECTION).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def generate_docx(template_path, data):
    doc = Document(template_path)
    # Paragraphs mein placeholders badalna
    for p in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"[{key}]"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
    
    # Tables ke andar bhi placeholders check karein
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in data.items():
                    placeholder = f"[{key}]"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(value))
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =================================================================
# --- 2. AUTHENTICATION ---
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

# --- TAB 1: SICK MEMO GENERATION ---
with tab1:
    st.header("📋 Generate New Sick Memo")
    df_emp = get_employees()
    
    if not df_emp.empty:
        # Search & Selection
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Search Employee", emp_list)
        
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
        
        with st.form("sick_memo_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            memo_date = c1.date_input("Memo Date", value=datetime.now())
            hospital = c2.selectbox("Hospital Name", ["BEOHARI", "NEW KATNI", "OTHER"])
            
            # Placeholders from Word File
            memo_data = {
                "LetterDate": memo_date.strftime("%d/%m/%Y"),
                "EmployeeName": emp_data.get('Employee Name', 'N/A'),
                "Designation": emp_data.get('Designation', 'N/A'),
                "UnitNumber": emp_data.get('UNIT No.', 'N/A')
            }
            
            if st.form_submit_button("Generate & Save Record"):
                try:
                    # 1. Save to Firebase sickemp collection
                    db.collection(SICK_COLLECTION).add({
                        "HRMS_ID": h_id,
                        "Name": memo_data["EmployeeName"],
                        "Designation": memo_data["Designation"],
                        "SickDate": str(memo_date),
                        "Hospital": hospital,
                        "Status": "SICK",
                        "ReturnDate": None,
                        "Timestamp": datetime.now()
                    })
                    
                    # 2. Generate Word File
                    docx_out = generate_docx("SICK MEMO temp.docx", memo_data)
                    st.success(f"Record for {memo_data['EmployeeName']} saved successfully!")
                    st.download_button("📥 Download Filled Sick Memo", docx_out, f"Sick_Memo_{h_id}.docx")
                except Exception as e:
                    st.error(f"Error: {e}")
    else:
        st.warning("No employees found in master database.")

# --- TAB 2: REPORTS & RETURN UPDATES ---
with tab2:
    st.header("📊 Sick Employee Dashboard")
    df_sick = get_sick_records()
    
    if not df_sick.empty:
        # 1. RETURN TO DUTY SECTION
        st.subheader("🔄 Update Return from Sick (Entry)")
        # Filter only currently SICK employees
        active_sick = df_sick[df_sick['Status'] == 'SICK']
        
        if not active_sick.empty:
            sick_options = active_sick.apply(lambda r: f"{r['Name']} (Since: {r['SickDate']})", axis=1).tolist()
            returning_emp = st.selectbox("Select Employee returning to Duty", sick_options)
            
            idx = sick_options.index(returning_emp)
            doc_to_update = active_sick.iloc[idx]
            
            ret_date = st.date_input("Return (FIT) Date", value=datetime.now())
            
            if st.button("Confirm Return & Mark FIT"):
                db.collection(SICK_COLLECTION).document(doc_to_update['id']).update({
                    "Status": "FIT",
                    "ReturnDate": str(ret_date)
                })
                st.success(f"{doc_to_update['Name']} is now marked as FIT.")
                st.rerun()
        else:
            st.info("No employees are currently on SICK status.")

        # 2. FULL REPORT TABLE
        st.divider()
        st.subheader("📑 All Sick Records History")
        # Formatting for display
        display_df = df_sick.drop(columns=['id', 'Timestamp'], errors='ignore')
        st.dataframe(display_df, use_container_width=True)
        
        # Download Report CSV
        report_csv = display_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button("📥 Download History Report", report_csv, "Sick_History_Report.csv", "text/csv")
    else:
        st.info("No sick records found in 'sickemp' collection.")

