import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document

# =================================================================
# --- 0. PATH & FIREBASE SETUP ---
# =================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SICK_TEMP = os.path.join(BASE_DIR, "assets", "SICK MEMO temp.docx")
IOD_TEMP = os.path.join(BASE_DIR, "assets", "IOD_temp.docx")

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
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

db = init_db()

# =================================================================
# --- 1. UTILITIES ---
# =================================================================
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    data = []
    for d in docs:
        emp = d.to_dict()
        raw_pf = emp.get('PF Number', '')
        emp['PF_Clean'] = str(raw_pf).split('.')[0].strip() if raw_pf else ""
        data.append(emp)
    return pd.DataFrame(data) if data else pd.DataFrame()

def get_sick_records():
    docs = db.collection(SICK_COLLECTION).stream()
    data = []
    for d in docs:
        item = d.to_dict()
        item['id'] = d.id
        # Ensure PF_Number exists in dictionary for DataFrame
        if 'PF_Number' not in item: item['PF_Number'] = ""
        if 'ReturnDate' not in item: item['ReturnDate'] = "N/A"
        data.append(item)
    
    # Error Prevention: Define columns if data is empty
    cols = ['Name', 'PF_Number', 'MemoType', 'StartDate', 'Status', 'ReturnDate']
    if not data:
        return pd.DataFrame(columns=cols)
    
    df = pd.DataFrame(data)
    # Ensure all required columns exist in the DF to avoid KeyError
    for c in cols:
        if c not in df.columns: df[c] = ""
        
    if 'Created' in df.columns:
        df = df.sort_values(by='Created', ascending=False)
    return df

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        all_paras = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paras.extend(list(cell.paragraphs))
        
        for p in all_paras:
            for key, value in data.items():
                placeholder = f"[{key}]"
                if placeholder in p.text:
                    for run in p.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, str(value))
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Docx Error: {e}"); return None

# =================================================================
# --- 2. AUTHENTICATION ---
# =================================================================
st.set_page_config(layout="wide", page_title="Railway Health MS")
if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login"):
        u, p = st.text_input("User"), st.text_input("Pass", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True; st.rerun()
            else: st.error("Wrong credentials")
    st.stop()

# =================================================================
# --- 3. MAIN UI ---
# =================================================================
tab1, tab2 = st.tabs(["📝 Generate Memo", "📊 Dashboard & History"])

with tab1:
    st.header("Memo Generation")
    df_emp = get_employees()
    if not df_emp.empty:
        memo_type = st.radio("Choose Memo Type:", ["SICK", "IOD"], horizontal=True)
        emp_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        
        selected_p = st.selectbox("Select Employee (Patient)", emp_options)
        p_pf = selected_p.split('(')[-1].strip(')')
        p_data = df_emp[df_emp['PF_Clean'] == p_pf].iloc[0]

        with st.form("memo_form"):
            c1, c2 = st.columns(2)
            memo_date = c1.date_input("Letter Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "SHAHDOL", "JABALPUR"])
            
            witness_info = {}
            if memo_type == "IOD":
                st.divider()
                st.subheader("Injury & Witness Details")
                ic1, ic2 = st.columns(2)
                injury_date = ic1.date_input("Injury Date")
                injury_time = ic2.text_input("Injury Time")
                injury_reason = st.text_area("Reason of Injury")
                wc1, wc2 = st.columns(2)
                w1_sel = wc1.selectbox("First Witness", ["Select"] + emp_options)
                w2_sel = wc2.selectbox("Second Witness", ["Select"] + emp_options)
                
                if w1_sel != "Select":
                    w1_row = df_emp[df_emp['PF_Clean'] == w1_sel.split('(')[-1].strip(')')].iloc[0]
                    witness_info.update({"WITNESS1_NAME": w1_row['Employee Name'], "WITNESS1_DESIG": w1_row['Designation']})
                if w2_sel != "Select":
                    w2_row = df_emp[df_emp['PF_Clean'] == w2_sel.split('(')[-1].strip(')')].iloc[0]
                    witness_info.update({"WITNESS2_NAME": w2_row['Employee Name'], "WITNESS2_DESIG": w2_row['Designation']})

            if st.form_submit_button("Save & Generate"):
                word_data = {
                    "LETTER DATE": memo_date.strftime("%d/%m/%Y"),
                    "EMPLOYEE NAME": p_data.get('Employee Name in Hindi', p_data.get('Employee Name', '')),
                    "PF NUMBER": p_pf,
                    "DESIGNATION": p_data.get('Designation in Hindi', p_data.get('Designation', '')),
                    "UNIT": p_data.get('UNIT No.', ''), **witness_info
                }
                if memo_type == "IOD":
                    word_data.update({"INJURY DATE": injury_date.strftime("%d/%m/%Y"), "TIME": injury_time, "INJURY REASON": injury_reason})

                db.collection(SICK_COLLECTION).add({
                    "HRMS_ID": p_data.get('HRMS ID', ''),
                    "PF_Number": p_pf,
                    "Name": p_data.get('Employee Name', ''),
                    "MemoType": memo_type,
                    "StartDate": str(memo_date),
                    "Status": "SICK" if memo_type == "SICK" else "IOD_ACTIVE",
                    "Created": datetime.now()
                })
                t_path = IOD_TEMP if memo_type == "IOD" else SICK_TEMP
                st.session_state.memo_bytes = generate_docx(t_path, word_data)
                st.session_state.last_memo_name = f"{memo_type}_{p_pf}.docx"
                st.success("✅ Saved!"); st.rerun()

        if 'memo_bytes' in st.session_state and st.session_state.memo_bytes:
            st.download_button("📥 Download Memo", st.session_state.memo_bytes, st.session_state.last_memo_name)

with tab2:
    st.header("📊 Health Reports")
    df_sick = get_sick_records()
    
    # Counting logic
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Sick", len(df_sick[df_sick['MemoType']=='SICK']))
    m2.metric("Total IOD", len(df_sick[df_sick['MemoType']=='IOD']))
    active = df_sick[df_sick['Status'].isin(['SICK', 'IOD_ACTIVE'])]
    m3.metric("Active Cases", len(active))
    
    st.divider()
    if not active.empty:
        sel_ret = st.selectbox("Mark FIT", active.apply(lambda r: f"{r['Name']} ({r['PF_Number']})", axis=1))
        if st.button("Confirm FIT"):
            target_id = active.iloc[0]['id'] # simplistic for demo
            db.collection(SICK_COLLECTION).document(target_id).update({"Status": "FIT", "ReturnDate": str(datetime.now().date())})
            st.rerun()

    st.subheader("📑 History")
    # FIX: Safety check for columns before displaying dataframe
    display_cols = ['Name', 'PF_Number', 'MemoType', 'StartDate', 'Status', 'ReturnDate']
    st.dataframe(df_sick[display_cols], use_container_width=True)
