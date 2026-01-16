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
            st.error(f"Firebase Init Error: {e}"); st.stop()
    return firestore.client()

db = init_db()

# =================================================================
# --- 1. UTILITIES ---
# =================================================================
def calculate_age(dob_val):
    try:
        if not dob_val or pd.isna(dob_val): return ""
        dob = pd.to_datetime(dob_val)
        today = datetime.now()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return str(age)
    except: return ""

def format_date_dmy(date_val):
    try:
        if not date_val or pd.isna(date_val): return ""
        return pd.to_datetime(date_val).strftime("%d/%m/%Y")
    except: return str(date_val)

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
        data.append(item)
    cols = ['Name', 'PF_Number', 'MemoType', 'StartDate', 'Status', 'ReturnDate']
    if not data: return pd.DataFrame(columns=cols)
    df = pd.DataFrame(data)
    for c in cols:
        if c not in df.columns: df[c] = ""
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
                    if placeholder in p.text: p.text = p.text.replace(placeholder, str(value))
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
            else: st.error("Invalid Credentials")
    st.stop()

# =================================================================
# --- 3. MAIN UI ---
# =================================================================
tab1, tab2 = st.tabs(["📝 Generate Memo", "📊 Dashboard & History"])

with tab1:
    st.header("Memo Generation")
    df_emp = get_employees()
    df_sick_check = get_sick_records()
    
    if not df_emp.empty:
        memo_type = st.radio("Choose Memo Type:", ["SICK", "IOD"], horizontal=True)
        emp_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        
        selected_p = st.selectbox("Select Employee (Patient)", emp_options)
        p_pf = selected_p.split('(')[-1].strip(')')
        p_data = df_emp[df_emp['PF_Clean'] == p_pf].iloc[0]

        # --- VALIDATION: Check if already Sick/IOD ---
        is_already_active = False
        if not df_sick_check.empty:
            active_match = df_sick_check[(df_sick_check['PF_Number'] == p_pf) & 
                                         (df_sick_check['Status'].isin(['SICK', 'IOD_ACTIVE']))]
            if not active_match.empty:
                is_already_active = True
                current_status = active_match.iloc[0]['Status']
                st.warning(f"⚠️ **Dhyan Dein:** {p_data['Employee Name']} pehle se **{current_status}** mein hain. Jab tak unhe FIT mark nahi kiya jata, naya letter generate nahi ho sakta.")

        with st.form("memo_form"):
            c1, c2, c3 = st.columns(3)
            memo_date = c1.date_input("Letter Date", value=datetime.now())
            hospital = c2.selectbox("Hospital", ["BEOHARI", "NEW KATNI", "SHAHDOL", "JABALPUR"])
            age_val = calculate_age(p_data.get('DOB'))
            age = c3.text_input("Age (Auto)", value=age_val)

            iod_data = {}
            if memo_type == "IOD":
                st.divider()
                st.subheader("Injury & Witness Details")
                ic1, ic2, ic3 = st.columns(3)
                injury_date = ic1.date_input("Injury Date")
                injury_time = ic2.text_input("Injury Time (HH:MM)")
                acc_place = ic3.text_input("Accident Place")
                nature = st.radio("Nature of Injury:", ["साधारण", "गंभीर"], horizontal=True)
                reason = st.text_area("Reason of Injury")
                wc1, wc2 = st.columns(2)
                w1_sel = wc1.selectbox("First Witness (Hindi)", ["Select"] + emp_options)
                w2_sel = wc2.selectbox("Second Witness (Hindi)", ["Select"] + emp_options)
                
                if w1_sel != "Select":
                    w1_row = df_emp[df_emp['PF_Clean'] == w1_sel.split('(')[-1].strip(')')].iloc[0]
                    iod_data["FIRST WITNESS"] = f"{w1_row.get('Employee Name in Hindi', w1_row['Employee Name'])}, {w1_row.get('Designation in Hindi', w1_row['Designation'])}"
                if w2_sel != "Select":
                    w2_row = df_emp[df_emp['PF_Clean'] == w2_sel.split('(')[-1].strip(')')].iloc[0]
                    iod_data["SECOND WITNESS"] = f"{w2_row.get('Employee Name in Hindi', w2_row['Employee Name'])}, {w2_row.get('Designation in Hindi', w2_row['Designation'])}"
                
                iod_data.update({"ACCIDENT PLACE": acc_place, "NATURE OF INJURY": nature, "INJURY DATE": injury_date.strftime("%d/%m/%Y"), "TIME": injury_time, "INJURY REASON": reason})

            # Submit button state based on validation
            submit_btn = st.form_submit_button("Save & Generate Memo", disabled=is_already_active)

            if submit_btn:
                word_data = {
                    "LETTER DATE": memo_date.strftime("%d/%m/%Y"),
                    "EMPLOYEE NAME": p_data.get('Employee Name in Hindi', p_data.get('Employee Name', '')),
                    "PF NUMBER": p_pf,
                    "DESIGNATION": p_data.get('Designation in Hindi', p_data.get('Designation', '')),
                    "UNIT NUMBER": p_data.get('UNIT No.', ''),
                    "WORKING STATION": p_data.get('STATION', 'SGAM'),
                    "DOA": format_date_dmy(p_data.get('DOA')),
                    "AGE": age, **iod_data
                }
                db.collection(SICK_COLLECTION).add({
                    "PF_Number": p_pf, "Name": p_data.get('Employee Name', ''), "MemoType": memo_type,
                    "StartDate": str(memo_date), "Status": "SICK" if memo_type == "SICK" else "IOD_ACTIVE",
                    "Created": datetime.now()
                })
                t_path = IOD_TEMP if memo_type == "IOD" else SICK_TEMP
                st.session_state.memo_bytes = generate_docx(t_path, word_data)
                st.session_state.last_memo_name = f"{memo_type}_{p_pf}.docx"
                st.success("✅ Record Saved!"); st.rerun()

        if 'memo_bytes' in st.session_state and st.session_state.memo_bytes:
            st.download_button("📥 Download Memo", st.session_state.memo_bytes, st.session_state.last_memo_name)

with tab2:
    st.header("📊 Reports & Dashboard")
    df_sick = get_sick_records()
    if not df_sick.empty:
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Sick", len(df_sick[df_sick['MemoType']=='SICK']))
        m2.metric("Total IOD", len(df_sick[df_sick['MemoType']=='IOD']))
        active_cases = df_sick[df_sick['Status'].isin(['SICK', 'IOD_ACTIVE'])]
        m3.metric("Active Cases", len(active_cases))
        
        st.divider()
        if not active_cases.empty:
            st.subheader("🔄 Mark FIT")
            sel_list = active_cases.apply(lambda r: f"{r['Name']} ({r['PF_Number']})", axis=1).tolist()
            returning = st.selectbox("Wapas aane wala karmchari select karein:", sel_list)
            if st.button("Confirm FIT Status"):
                doc_id = active_cases.iloc[sel_list.index(returning)]['id']
                db.collection(SICK_COLLECTION).document(doc_id).update({"Status": "FIT", "ReturnDate": str(datetime.now().date())})
                st.success("Karmchari FIT mark ho gaya!"); st.rerun()

        st.subheader("📑 Full History")
        st.dataframe(df_sick[['Name', 'PF_Number', 'MemoType', 'StartDate', 'Status', 'ReturnDate']], use_container_width=True)
