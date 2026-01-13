import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Setup ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "pme memo temp.docx")
EMP_COLLECTION = "employees"
PME_HISTORY_COLLECTION = "pme_history"

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

# --- 1. Utilities ---
def get_employees():
    docs = db.collection(EMP_COLLECTION).stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_pme_history():
    docs = db.collection(PME_HISTORY_COLLECTION).order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def calculate_service(doa_val):
    try:
        if isinstance(doa_val, str):
            doa = datetime.strptime(doa_val, "%Y-%m-%d")
        else:
            doa = datetime.combine(doa_val, datetime.min.time())
        diff = relativedelta(datetime.now(), doa)
        return diff.years, diff.months
    except:
        return 0, 0

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path): return None
    try:
        doc = Document(template_path)
        # [span_0](start_span)Placeholder mapping from[span_0](end_span)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data.items():
                        p1, p2 = "{{ " + key + " }}", "{{" + key + "}}"
                        if p1 in cell.text: cell.text = cell.text.replace(p1, str(value))
                        if p2 in cell.text: cell.text = cell.text.replace(p2, str(value))
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Doc Error: {e}")
        return None

# --- 2. UI ---
st.set_page_config(layout="wide", page_title="PME System")

# Session state for file storage (Fix for Download Button)
if 'pme_file' not in st.session_state: st.session_state.pme_file = None
if 'pme_filename' not in st.session_state: st.session_state.pme_filename = ""

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME Records", "🛠 Update Marks"])
df_emp = get_employees()

with tab1:
    st.header("PME Memo Generation")
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            doa = emp_data.get('Date of Appointment', '')
            s_year, s_month = calculate_service(doa)
            
            pme_data = {
                "name": emp_data.get('Employee Name', ''),
                "age": emp_data.get('Age', ''),
                "father_name": emp_data.get('Father Name', ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical Category', ''),
                "dob": emp_data.get('Date of Birth', ''),
                "doa": doa,
                "current_date": c_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_data.get('Last PME', 'N/A'), # From DB
                "last_place": "NKJ", "examiner": "ACMS/NKJ",
                "service_year": s_year, "service_month": s_month
            }

            if st.form_submit_button("Generate Memo"):
                out = generate_pme_docx(TEMPLATE_PATH, pme_data)
                if out:
                    # History save
                    db.collection(PME_HISTORY_COLLECTION).add({**pme_data, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                    # Save to session state
                    st.session_state.pme_file = out
                    st.session_state.pme_filename = f"PME_{h_id}.docx"
                    st.success("✅ Memo Taiyar hai! Niche se download karein.")

        # Download Button form ke BAHAR
        if st.session_state.pme_file:
            st.download_button(
                label="📥 Download Generated Memo",
                data=st.session_state.pme_file,
                file_name=st.session_state.pme_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Database empty.")

# (Tab 2 & 3 logic remain same)
with tab2:
    st.header("📊 History")
    df_h = get_pme_history()
    if not df_h.empty: st.dataframe(df_h[['Timestamp', 'name', 'current_date', 'last_examined_date']], use_container_width=True)

with tab3:
    st.header("🛠 Update")
    if not df_emp.empty:
        t_sel = st.selectbox("Select to Update", emp_list)
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        with st.form("upd_f"):
            m1 = st.text_input("Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = st.text_input("Mark 2", t_row.get('Physical Mark 2', ''))
            lp = st.text_input("Last PME", t_row.get('Last PME', ''))
            if st.form_submit_button("Save"):
                db.collection(EMP_COLLECTION).document(t_row['id']).update({"Physical Mark 1": m1, "Physical Mark 2": m2, "Last PME": lp})
                st.success("Updated!")
