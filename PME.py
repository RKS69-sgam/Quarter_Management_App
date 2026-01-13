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

# --- 1. Utilities ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_pme_history():
    docs = db.collection("pme_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    data = [{**d.to_dict(), 'id': d.id} for d in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

def calculate_metrics(dob_val, doa_val):
    """Age aur Service ka sahi calculation aur date formatting"""
    now = datetime.now()
    res = {"age": "N/A", "s_yr": "0", "s_mn": "0", "dob_f": "", "doa_f": ""}
    
    try:
        if dob_val and str(dob_val).lower() != 'nan':
            dt_dob = pd.to_datetime(dob_val).to_pydatetime()
            res["age"] = str(relativedelta(now, dt_dob).years)
            res["dob_f"] = dt_dob.strftime("%d/%m/%Y")
            
        if doa_val and str(doa_val).lower() != 'nan':
            dt_doa = pd.to_datetime(doa_val).to_pydatetime()
            diff = relativedelta(now, dt_doa)
            res["s_yr"] = str(diff.years)
            res["s_mn"] = str(diff.months)
            res["doa_f"] = dt_doa.strftime("%d/%m/%Y")
    except:
        pass
    return res

def replace_text_logic(doc, data):
    """Word doc ke paragraphs aur tables mein placeholders replace karna"""
    all_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paras.extend(cell.paragraphs)
    
    for p in all_paras:
        for key, value in data.items():
            placeholder = "{{ " + str(key) + " }}"
            if placeholder in p.text:
                # Formatting preserve karne ke liye run-level merge
                full_text = "".join(run.text for run in p.runs)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(value))
                    for i, run in enumerate(p.runs):
                        run.text = new_text if i == 0 else ""

def generate_pme_docx(template_path, data):
    if not os.path.exists(template_path):
        st.error(f"Template not found at: {template_path}")
        return None
    try:
        doc = Document(template_path)
        replace_text_logic(doc, data)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception as e:
        st.error(f"Doc Generation Error: {e}")
        return None

# --- 2. UI ---
st.set_page_config(layout="wide", page_title="Railway PME System")
df_emp = get_employees()

# Session State for Downloads
if 'dl_data' not in st.session_state: st.session_state.dl_data = None
if 'dl_name' not in st.session_state: st.session_state.dl_name = ""

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME History", "🛠 Update Database"])

with tab1:
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Select Employee", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("pme_form"):
            c_date = st.date_input("Memo Date", value=datetime.now())
            
            # Metrics Calculation
            m = calculate_metrics(emp_data.get('DOB'), emp_data.get('DOA'))
            
            pme_vals = {
                "dob": m["dob_f"],
                "doa": m["doa_f"],
                "name": emp_data.get('Employee Name', ''),
                "age": m["age"],
                "father_name": emp_data.get("FATHER'S NAME", ''),
                "designation": emp_data.get('Designation', ''),
                "medical_category": emp_data.get('Medical category', ''),
                "current_date": c_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_data.get('Last PME', 'N/A'),
                "last_place": emp_data.get('Last PME Place', 'NKJ'),
                "examiner": "ACMS/NKJ",
                "service_year": m["s_yr"],
                "service_month": m["s_mn"]
            }

            if st.form_submit_button("Generate Memo"):
                out = generate_pme_docx(TEMPLATE_PATH, pme_vals)
                if out:
                    db.collection("pme_history").add({**pme_vals, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                    st.session_state.dl_data = out
                    st.session_state.dl_name = f"PME_{h_id}.docx"
                    st.success(f"✅ Memo Generated! Age: {m['age']}, Service: {m['s_yr']}y {m['s_mn']}m")

        # Download button form ke bahar
        if st.session_state.dl_data:
            st.download_button("📥 Download PME Memo", st.session_state.dl_data, st.session_state.dl_name)
    else:
        st.warning("Database is empty.")

with tab2:
    st.header("📊 Generation History")
    df_h = get_pme_history()
    if not df_h.empty:
        st.dataframe(df_h[['Timestamp', 'name', 'current_date', 'age', 'service_year']], use_container_width=True)

with tab3:
    st.header("🛠 Update Employee Records")
    if not df_emp.empty:
        t_sel = st.selectbox("Select Employee to Update", emp_list, key="upd_pme_tab")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("update_db_form"):
            col1, col2 = st.columns(2)
            m1 = col1.text_input("Physical Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = col2.text_input("Physical Mark 2", t_row.get('Physical Mark 2', ''))
            lp_date = st.text_input("Last PME Date", t_row.get('Last PME', ''))
            lp_place = st.text_input("Last PME Place", t_row.get('Last PME Place', 'NKJ'))
            
            if st.form_submit_button("Save to Database"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_date, "Last PME Place": lp_place
                })
                st.success("✅ Database Updated!")
                st.rerun()
