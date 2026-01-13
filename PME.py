import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. PATH & DATABASE SETUP ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "pme memo temp.docx")

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

# --- 1. CORE UTILITIES ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_pme_history():
    """Tab 2 ke liye history fetch karne ka function"""
    docs = db.collection("pme_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def calculate_service_and_age(dob_val, doa_val):
    now = datetime.now()
    age, s_year, s_month = "0", "0", "0"
    try:
        if dob_val and str(dob_val) != 'nan':
            dob = pd.to_datetime(dob_val).to_pydatetime()
            age = str(relativedelta(now, dob).years)
        if doa_val and str(doa_val) != 'nan':
            doa = pd.to_datetime(doa_val).to_pydatetime()
            diff = relativedelta(now, doa)
            s_year, s_month = str(diff.years), str(diff.months)
    except: pass
    return age, s_year, s_month

def replace_text_logic(doc, data):
    """Placeholder format: {{ key }} replacement logic"""
    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            for key, value in data.items():
                placeholder = "{{ " + key + " }}" # Format as per your requirement
                if placeholder in p.text:
                    full_text = "".join(run.text for run in p.runs)
                    if placeholder in full_text:
                        new_text = full_text.replace(placeholder, str(value))
                        for i, run in enumerate(p.runs):
                            run.text = new_text if i == 0 else ""

    replace_in_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_paragraphs(cell.paragraphs)

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Template missing: {template_path}")
        return None
    try:
        doc = Document(template_path)
        replace_text_logic(doc, data)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Doc Error: {e}"); return None

# --- 2. MAIN UI ---
st.set_page_config(layout="wide", page_title="PME Management System")
df_emp = get_employees()

if 'pme_file' not in st.session_state: st.session_state.pme_file = None

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME History", "🛠 Update Database"])

with tab1:
    st.header("PME Memo Generation")
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            
            # DB fields matching: DOB, DOA, FATHER'S NAME, Medical category
            dob_val = emp_data.get('DOB', '')
            doa_val = emp_data.get('DOA', '')
            age, s_year, s_month = calculate_service_and_age(dob_val, doa_val)
            
            pme_vals = {
                "dob": dob_val,
                "doa": doa_val,
                "name": emp_data.get('Employee Name', ''),
                "age": age,
                "father_name": emp_data.get("FATHER'S NAME", ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical category', ''),
                "current_date": c_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_data.get('Last PME', 'N/A'),
                "last_place": emp_data.get('Last PME Place', 'NKJ'),
                "examiner": "ACMS/NKJ",
                "service_year": s_year,
                "service_month": s_month
            }

            if st.form_submit_button("Generate Memo"):
                out = generate_pme_docx(TEMPLATE_PATH, pme_vals)
                if out:
                    # Save history to DB
                    db.collection("pme_history").add({**pme_vals, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                    st.session_state.pme_file = out
                    st.success("✅ Memo Generated!")

        if st.session_state.pme_file:
            st.download_button("📥 Download PME Memo", st.session_state.pme_file, f"PME_{h_id}.docx")

with tab2:
    st.header("📊 PME Generation History")
    df_h = get_pme_history()
    if not df_h.empty:
        # Displaying key columns from history
        st.dataframe(df_h[['Timestamp', 'name', 'designation', 'current_date', 'last_examined_date']], use_container_width=True)
    else:
        st.info("Abhi tak koi memo generate nahi hua hai.")

with tab3:
    st.header("🛠 Update Employee Data")
    if not df_emp.empty:
        t_sel = st.selectbox("Select Employee to Update", emp_list, key="upd_pme_final")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("final_upd_form"):
            col1, col2 = st.columns(2)
            m1 = col1.text_input("Physical Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = col2.text_input("Physical Mark 2", t_row.get('Physical Mark 2', ''))
            lp_date = st.text_input("Last Examination Date (DD/MM/YYYY)", t_row.get('Last PME', ''))
            lp_place = st.text_input("Last Examination Place", t_row.get('Last PME Place', 'NKJ'))
            
            if st.form_submit_button("Save to Database"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_date, "Last PME Place": lp_place
                })
                st.success("✅ Database Updated!")
                st.rerun()
