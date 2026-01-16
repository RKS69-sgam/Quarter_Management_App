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

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        # Search in Paragraphs and Tables
        all_elements = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_elements.extend(list(cell.paragraphs))
        
        for p in all_elements:
            for key, value in data.items():
                placeholder = f"[{key}]"
                if placeholder in p.text:
                    for run in p.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, str(value))
                    # Fallback for complex formatting
                    if placeholder in p.text:
                        p.text = p.text.replace(placeholder, str(value))
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
        
        selected_p = st.selectbox("Select Patient Employee", emp_options)
        p_pf = selected_p.split('(')[-1].strip(')')
        p_data = df_emp[df_emp['PF_Clean'] == p_pf].iloc[0]

        with st.form("memo_form"):
            st.subheader("General Details")
            c1, c2, c3 = st.columns(3)
            memo_date = c1.date_input("Letter Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "SHAHDOL", "JABALPUR"])
            age = c3.text_input("Age", value=str(p_data.get('Age', '')))

            # IOD SPECIFIC FIELDS
            iod_data = {}
            if memo_type == "IOD":
                st.divider()
                st.subheader("Injury & Witness Details")
                
                ic1, ic2, ic3 = st.columns(3)
                injury_date = ic1.date_input("Injury Date")
                injury_time = ic2.text_input("Injury Time (e.g. 10:30 AM)")
                acc_place = ic3.text_input("Accident Place (e.g. Yard, Station)")
                
                nature = st.radio("Nature of Injury:", ["साधारण", "गंभीर"], horizontal=True)
                reason = st.text_area("How did the injury happen? (Reason)")
                
                wc1, wc2 = st.columns(2)
                w1_sel = wc1.selectbox("First Witness (Gawah 1)", ["Select Witness"] + emp_options)
                w2_sel = wc2.selectbox("Second Witness (Gawah 2)", ["Select Witness"] + emp_options)
                
                # Witness Data Mapping
                if w1_sel != "Select Witness":
                    w1_row = df_emp[df_emp['PF_Clean'] == w1_sel.split('(')[-1].strip(')')].iloc[0]
                    iod_data["FIRST WITNESS"] = f"{w1_row['Employee Name']}, {w1_row['Designation']}"
                
                if w2_sel != "Select Witness":
                    w2_row = df_emp[df_emp['PF_Clean'] == w2_sel.split('(')[-1].strip(')')].iloc[0]
                    iod_data["SECOND WITNESS"] = f"{w2_row['Employee Name']}, {w2_row['Designation']}"

                iod_data.update({
                    "ACCIDENT PLACE": acc_place,
                    "NATURE OF INJURY": nature,
                    "INJURY DATE": injury_date.strftime("%d/%m/%Y"),
                    "TIME": injury_time,
                    "INJURY REASON": reason
                })

            if st.form_submit_button("Save & Generate Memo"):
                # Unified Data Mapping for Word Placeholders
                word_data = {
                    "LETTER DATE": memo_date.strftime("%d/%m/%Y"),
                    "EMPLOYEE NAME": p_data.get('Employee Name in Hindi', p_data.get('Employee Name', '')),
                    "PF NUMBER": p_pf,
                    "DESIGNATION": p_data.get('Designation in Hindi', p_data.get('Designation', '')),
                    "UNIT NUMBER": p_data.get('UNIT No.', p_data.get('Unit', '')),
                    "WORKING STATION": p_data.get('STATION', 'SGAM'),
                    "DOA": p_data.get('DOA', ''), # Date of Appointment
                    "AGE": age,
                    **iod_data
                }

                # Save Record to Firebase
                db.collection(SICK_COLLECTION).add({
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
                st.success("✅ Saved Successfully!")

        if 'memo_bytes' in st.session_state and st.session_state.memo_bytes:
            st.download_button("📥 Download Memo (DOCX)", st.session_state.memo_bytes, st.session_state.last_memo_name)

# =================================================================
# --- TAB 2: DASHBOARD & HISTORY ---
# =================================================================
with tab2:
    st.header("📊 Health Dashboard & Return Status")
    docs = db.collection(SICK_COLLECTION).stream()
    sick_data = []
    for d in docs:
        item = d.to_dict()
        item['id'] = d.id
        sick_data.append(item)
    
    if sick_data:
        df_sick = pd.DataFrame(sick_data)
        
        # Dashboard Counters
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Sick", len(df_sick[df_sick['MemoType']=='SICK']))
        m2.metric("Total IOD", len(df_sick[df_sick['MemoType']=='IOD']))
        active = df_sick[df_sick['Status'].isin(['SICK', 'IOD_ACTIVE'])]
        m3.metric("Currently Sick/IOD", len(active))
        
        st.divider()
        
        # Mark Fit
        if not active.empty:
            st.subheader("Mark FIT (Return to Duty)")
            sel_fit = st.selectbox("Select Employee", active.apply(lambda r: f"{r['Name']} ({r['PF_Number']})", axis=1))
            if st.button("Confirm Return (FIT)"):
                idx = active[active.apply(lambda r: f"{r['Name']} ({r['PF_Number']})", axis=1) == sel_fit].iloc[0]['id']
                db.collection(SICK_COLLECTION).document(idx).update({"Status": "FIT", "ReturnDate": str(datetime.now().date())})
                st.rerun()

        st.subheader("📑 Full History")
        st.dataframe(df_sick, use_container_width=True)
    else:
        st.info("No records found in database.")
