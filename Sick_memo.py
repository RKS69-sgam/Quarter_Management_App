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
        # PF Number formatting fix (.0 hatane ke liye)
        raw_pf = emp.get('PF Number', '')
        emp['PF_Clean'] = str(raw_pf).split('.')[0].strip() if raw_pf else ""
        data.append(emp)
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

def generate_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Template not found: {template_path}")
        return None
    try:
        doc = Document(template_path)
        # Paragraphs aur Tables dono mein replacement
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
                    if placeholder in p.text: # Fallback
                        p.text = p.text.replace(placeholder, str(value))
        
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Docx Error: {e}")
        return None

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
        
        # Employees list for selection
        emp_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        
        st.subheader("Patient Details")
        selected_p = st.selectbox("Select Employee (Patient)", emp_options)
        p_pf = selected_p.split('(')[-1].strip(')')
        p_data = df_emp[df_emp['PF_Clean'] == p_pf].iloc[0]

        with st.form("memo_form"):
            c1, c2 = st.columns(2)
            memo_date = c1.date_input("Letter Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "SHAHDOL", "JABALPUR"])
            
            # IOD SPECIFIC FIELDS
            witness_info = {}
            injury_date, injury_time, injury_reason = None, "", ""
            
            if memo_type == "IOD":
                st.divider()
                st.subheader("Injury & Witness Details")
                ic1, ic2 = st.columns(2)
                injury_date = ic1.date_input("Injury Date")
                injury_time = ic2.text_input("Injury Time (HH:MM)")
                injury_reason = st.text_area("Reason of Injury")
                
                wc1, wc2 = st.columns(2)
                w1_select = wc1.selectbox("First Witness", ["Select Witness"] + emp_options)
                w2_select = wc2.selectbox("Second Witness", ["Select Witness"] + emp_options)
                
                # Extract Witness Data
                if w1_select != "Select Witness":
                    w1_pf = w1_select.split('(')[-1].strip(')')
                    w1_row = df_emp[df_emp['PF_Clean'] == w1_pf].iloc[0]
                    witness_info.update({"WITNESS1_NAME": w1_row['Employee Name'], "WITNESS1_DESIG": w1_row['Designation']})
                
                if w2_select != "Select Witness":
                    w2_pf = w2_select.split('(')[-1].strip(')')
                    w2_row = df_emp[df_emp['PF_Clean'] == w2_pf].iloc[0]
                    witness_info.update({"WITNESS2_NAME": w2_row['Employee Name'], "WITNESS2_DESIG": w2_row['Designation']})

            if st.form_submit_button("Save & Generate"):
                # Word Data Mapping
                word_data = {
                    "LETTER DATE": memo_date.strftime("%d/%m/%Y"),
                    "EMPLOYEE NAME": p_data.get('Employee Name in Hindi', p_data.get('Employee Name', '')),
                    "PF NUMBER": p_pf,
                    "DESIGNATION": p_data.get('Designation in Hindi', p_data.get('Designation', '')),
                    "UNIT": p_data.get('UNIT No.', ''),
                    **witness_info
                }
                
                if memo_type == "IOD":
                    word_data.update({
                        "INJURY DATE": injury_date.strftime("%d/%m/%Y") if injury_date else "",
                        "TIME": injury_time,
                        "INJURY REASON": injury_reason
                    })

                # Save to Firebase
                db.collection(SICK_COLLECTION).add({
                    "HRMS_ID": p_data.get('HRMS ID', ''),
                    "PF_Number": p_pf,
                    "Name": p_data.get('Employee Name', ''),
                    "MemoType": memo_type,
                    "StartDate": str(memo_date),
                    "Status": "SICK" if memo_type == "SICK" else "IOD_ACTIVE",
                    "Created": datetime.now()
                })
                
                # Document Generation
                t_path = IOD_TEMP if memo_type == "IOD" else SICK_TEMP
                st.session_state.memo_bytes = generate_docx(t_path, word_data)
                st.session_state.last_memo_name = f"{memo_type}_{p_pf}.docx"
                st.success("✅ Record Saved!")

        if 'memo_bytes' in st.session_state and st.session_state.memo_bytes:
            st.download_button("📥 Download Generated Memo", st.session_state.memo_bytes, 
                             st.session_state.last_memo_name, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with tab2:
    st.header("📊 Health Reports")
    df_sick = get_sick_records()
    
    if not df_sick.empty:
        # --- DASHBOARD ---
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Sick", len(df_sick[df_sick['MemoType']=='SICK']))
        m2.metric("Total IOD", len(df_sick[df_sick['MemoType']=='IOD']))
        active = df_sick[df_sick['Status'].isin(['SICK', 'IOD_ACTIVE'])]
        m3.metric("Current Active Cases", len(active))
        m4.metric("Total Fit", len(df_sick[df_sick['Status']=='FIT']))
        
        st.divider()
        
        # --- UPDATE RETURN ---
        st.subheader("Update Return (Mark FIT)")
        if not active.empty:
            active_list = active.apply(lambda r: f"{r['Name']} ({r['MemoType']} - {r['StartDate']})", axis=1).tolist()
            sel_ret = st.selectbox("Select Employee to Mark FIT", active_list)
            fit_date = st.date_input("FIT Date")
            if st.button("Confirm FIT Status"):
                idx = active_list.index(sel_ret)
                db.collection(SICK_COLLECTION).document(active.iloc[idx]['id']).update({
                    "Status": "FIT",
                    "ReturnDate": str(fit_date)
                })
                st.success("Status Updated!"); st.rerun()

        st.divider()
        st.subheader("📑 Detailed History")
        st.dataframe(df_sick[['Name', 'PF_Number', 'MemoType', 'StartDate', 'Status', 'ReturnDate']], use_container_width=True)
    else:
        st.info("No records found.")
