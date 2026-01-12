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
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "Exam NOC Letter temp.docx")

EMP_COLLECTION = "employees"
NOC_HISTORY_COLLECTION = "noc_history"

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

def get_noc_history():
    docs = db.collection(NOC_HISTORY_COLLECTION).order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def get_noc_count_this_year(pf_number):
    current_year = datetime.now().year
    docs = db.collection(NOC_HISTORY_COLLECTION)\
             .where("PFNumber", "==", pf_number)\
             .where("Year", "==", current_year).stream()
    return len(list(docs))

def create_noc_table(doc, emp_data_list):
    # [PFNumber] placeholder dhundh kar table insert karna
    for p in doc.paragraphs:
        if "[PFNumber]" in p.text:
            p.text = p.text.replace("[PFNumber]", "")
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            headers = ['Sr. No.', 'PF Number', 'Employee Name', 'Designation', "Exam's Name", 'Term of NOC']
            for i, h_text in enumerate(headers):
                hdr[i].text = h_text
            
            for idx, emp in enumerate(emp_data_list):
                row = table.add_row().cells
                row[0].text = str(idx + 1)
                row[1].text = str(emp['PFNumber'])
                row[2].text = str(emp['Name'])
                row[3].text = str(emp['Desig'])
                row[4].text = str(emp['ExamName'])
                row[5].text = str(emp['Term'])
            break

def generate_multi_noc(template_path, l_date, emp_data_list):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    
    # 1. Date Replacement
    formatted_date = l_date.strftime("%d-%m-%Y")
    for p in doc.paragraphs:
        if "[LetterDate]" in p.text:
            p.text = p.text.replace("[LetterDate]", formatted_date)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "[LetterDate]" in cell.text:
                    cell.text = cell.text.replace("[LetterDate]", formatted_date)

    # 2. Add Dynamic Table
    create_noc_table(doc, emp_data_list)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =================================================================
# --- 2. MAIN UI ---
# =================================================================
st.set_page_config(layout="wide", page_title="Exam NOC System")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Exam NOC Login")
    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Invalid Credentials")
    st.stop()

tab1, tab2 = st.tabs(["📝 Generate NOC", "📊 NOC Records Report"])

# --- TAB 1: GENERATE ---
with tab1:
    st.header("Exam NOC Taiyar Karein")
    df_emp = get_employees()
    if not df_emp.empty:
        emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('PF Number')})", axis=1).tolist()
        selected_names = st.multiselect("Karmchari Chunein (Multiple Select Kar Sakte Hain)", emp_options)
        
        final_list = []
        if selected_names:
            st.subheader("Exam Details Bharein")
            for name in selected_names:
                pf = name.split('(')[-1].strip(')')
                row = df_emp[df_emp['PF Number'] == pf].iloc[0]
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    exam = st.text_input(f"Exam Name for {row['Employee Name']}", key=f"ex_{pf}")
                with col2:
                    count = get_noc_count_this_year(pf)
                    terms = ["First", "Second", "Third", "Fourth"]
                    if count < 4:
                        term = terms[count]
                        st.info(f"Auto Term: {term}")
                        final_list.append({
                            "PFNumber": pf,
                            "Name": row.get('Employee Name in Hindi', row['Employee Name']),
                            "Desig": row.get('Designation in Hindi', row['Designation']),
                            "ExamName": exam,
                            "Term": term,
                            "Year": datetime.now().year,
                            "Timestamp": datetime.now()
                        })
                    else:
                        st.error("Limit Reached (4 NOCs Already Taken)")

            if st.button("Generate & Save NOC"):
                if final_list and all(item['ExamName'] for item in final_list):
                    out = generate_multi_noc(TEMPLATE_PATH, datetime.now(), final_list)
                    if out:
                        for item in final_list:
                            db.collection(NOC_HISTORY_COLLECTION).add(item)
                        st.success("✅ Records Saved!")
                        st.download_button("📥 Download NOC Letter", out, "Exam_NOC.docx")
                else:
                    st.warning("Sabhi Exam Names bharna zaroori hai.")

# --- TAB 2: REPORT ---
with tab2:
    st.header("📊 Exam NOC History Report")
    df_history = get_noc_history()
    
    if not df_history.empty:
        # Summary Metrics
        c1, c2 = st.columns(2)
        total_noc = len(df_history)
        unique_emp = df_history['PFNumber'].nunique()
        c1.metric("Total NOCs Issued", total_noc)
        c2.metric("Unique Employees", unique_emp)
        
        # Filter by PF Number
        search_pf = st.text_input("PF Number se Search Karein")
        if search_pf:
            df_history = df_history[df_history['PFNumber'].str.contains(search_pf)]
        
        # Display Table
        st.dataframe(
            df_history[['Timestamp', 'PFNumber', 'Name', 'Designation', 'ExamName', 'Term', 'Year']],
            use_container_width=True
        )
        
        # Export Option
        csv = df_history.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Download Full Report (CSV)", csv, "NOC_Report.csv", "text/csv")
    else:
        st.info("Abhi tak koi NOC record nahi hai.")
