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
                # Local file fallback
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}")
            st.stop()
    return firestore.client()

# Database initialize karein
db = init_db()

# --- 1. Robust Utilities ---
def get_employees():
    try:
        docs = db.collection("employees").stream()
        return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])
    except:
        return pd.DataFrame()

def get_pme_history():
    try:
        docs = db.collection("pme_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
        data = [{**d.to_dict(), 'id': d.id} for d in docs]
        return pd.DataFrame(data) if data else pd.DataFrame()
    except:
        return pd.DataFrame()

def parse_and_format_date(date_val):
    """Har tarah ke date format (Timestamp/String) ko theek karne ke liye"""
    if not date_val or str(date_val).lower() == 'nan':
        return None, ""
    try:
        # Agar Firebase Timestamp hai
        if hasattr(date_val, 'to_datetime'):
            dt = date_val.to_datetime().replace(tzinfo=None)
        else:
            # Agar string hai (e.g., '1987-03-09...')
            dt = pd.to_datetime(date_val).to_pydatetime().replace(tzinfo=None)
        return dt, dt.strftime("%d/%m/%Y")
    except:
        return None, ""

def calculate_pme_metrics(dob_raw, doa_raw):
    [span_0](start_span)"""Age aur Service Length ka sahi calculation[span_0](end_span)"""
    now = datetime.now()
    metrics = {"age": "N/A", "s_yr": "0", "s_mn": "0", "dob_f": "", "doa_f": ""}
    
    dt_dob, dob_str = parse_and_format_date(dob_raw)
    dt_doa, doa_str = parse_and_format_date(doa_raw)
    
    metrics["dob_f"] = dob_str
    metrics["doa_f"] = doa_str
    
    if dt_dob:
        metrics["age"] = str(relativedelta(now, dt_dob).years)
        
    if dt_doa:
        diff = relativedelta(now, dt_doa)
        metrics["s_yr"] = str(diff.years)
        metrics["s_mn"] = str(diff.months)
        
    return metrics

def replace_placeholders(doc, mapping):
    [span_1](start_span)[span_2](start_span)"""Document ke paragraphs aur tables mein replacement[span_1](end_span)[span_2](end_span)"""
    target_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                target_paras.extend(cell.paragraphs)
                
    for p in target_paras:
        for key, val in mapping.items():
            placeholder = "{{ " + str(key) + " }}"
            if placeholder in p.text:
                full_text = "".join(run.text for run in p.runs)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(val))
                    for i, run in enumerate(p.runs):
                        run.text = new_text if i == 0 else ""

# --- 2. Streamlit UI ---
st.set_page_config(layout="wide", page_title="PME Management")

# Data fetch karein
df_emp = get_employees()

if 'memo_ready' not in st.session_state: st.session_state.memo_ready = None

tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 PME History", "🛠 Update Database"])

with tab1:
    if not df_emp.empty:
        emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected_emp = st.selectbox("Select Employee", emp_names)
        hrms_id = selected_emp.split('(')[-1].strip(')')
        emp_row = df_emp[df_emp['HRMS ID'] == hrms_id].iloc[0]

        with st.form("pme_gen_form"):
            memo_date = st.date_input("Memo Date", value=datetime.now())
            
            # [span_3](start_span)Metrics Calculation[span_3](end_span)
            m = calculate_pme_metrics(emp_row.get('DOB'), emp_row.get('DOA'))
            
            final_mapping = {
                "dob": m["dob_f"],
                "doa": m["doa_f"],
                "name": emp_row.get('Employee Name', ''),
                "age": m["age"],
                "father_name": emp_row.get("FATHER'S NAME", ''),
                "designation": emp_row.get('Designation', ''),
                "medical_category": emp_row.get('Medical category', ''),
                "current_date": memo_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_row.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_row.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_row.get('Last PME', 'N/A'),
                "last_place": emp_row.get('Last PME Place', 'NKJ'),
                "examiner": "ACMS/NKJ",
                "service_year": m["s_yr"],
                "service_month": m["s_mn"]
            }

            if st.form_submit_button("Generate Document"):
                if os.path.exists(TEMPLATE_PATH):
                    doc = Document(TEMPLATE_PATH)
                    replace_placeholders(doc, final_mapping)
                    buf = io.BytesIO()
                    doc.save(buf)
                    st.session_state.memo_ready = buf.getvalue()
                    
                    # History entry
                    db.collection("pme_history").add({**final_mapping, "Timestamp": datetime.now(), "HRMS_ID": hrms_id})
                    st.success(f"✅ Generated! Age: {m['age']}, Service: {m['s_yr']}y {m['s_mn']}m")
                else:
                    st.error("Template file missing!")

        if st.session_state.memo_ready:
            st.download_button("📥 Download PME Memo", st.session_state.memo_ready, f"PME_{hrms_id}.docx")
    else:
        st.warning("Database records not found.")

with tab2:
    st.header("📊 History")
    df_h = get_pme_history()
    if not df_h.empty:
        st.dataframe(df_h[['Timestamp', 'name', 'age', 'service_year', 'service_month']], use_container_width=True)

with tab3:
    st.header("🛠 Update Records")
    if not df_emp.empty:
        t_sel = st.selectbox("Select Employee", emp_names, key="upd_pme")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        with st.form("upd_form"):
            m1 = st.text_input("Mark 1", t_row.get('Physical Mark 1', ''))
            m2 = st.text_input("Mark 2", t_row.get('Physical Mark 2', ''))
            lp_d = st.text_input("Last PME Date", t_row.get('Last PME', ''))
            lp_p = st.text_input("Last Place", t_row.get('Last PME Place', 'NKJ'))
            if st.form_submit_button("Save"):
                db.collection("employees").document(t_row['id']).update({
                    "Physical Mark 1": m1, "Physical Mark 2": m2,
                    "Last PME": lp_d, "Last PME Place": lp_p
                })
                st.rerun()
