import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Firebase Initialization ---
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
            return None
    return firestore.client()

db = init_db()

# --- 1. Robust Date & Calculation Logic ---
def get_safe_date(date_val):
    """Timestamp ya String kisi bhi format se sahi datetime object nikalne ke liye"""
    if not date_val or str(date_val).lower() == 'nan':
        return None
    try:
        # Agar Firebase ka Timestamp object hai
        if hasattr(date_val, 'to_datetime'):
            return date_val.to_datetime().replace(tzinfo=None)
        # Agar string format mein hai
        return pd.to_datetime(date_val).to_pydatetime().replace(tzinfo=None)
    except:
        return None

def calculate_pme_metrics(dob_raw, doa_raw):
    now = datetime.now()
    # Default values
    res = {"age": "N/A", "s_yr": "0", "s_mn": "0", "dob_f": "", "doa_f": ""}
    
    dt_dob = get_safe_date(dob_raw)
    dt_doa = get_safe_date(doa_raw)
    
    if dt_dob:
        res["age"] = str(relativedelta(now, dt_dob).years)
        res["dob_f"] = dt_dob.strftime("%d/%m/%Y")
    
    if dt_doa:
        diff = relativedelta(now, dt_doa)
        res["s_yr"] = str(diff.years)
        res["s_mn"] = str(diff.months)
        res["doa_f"] = dt_doa.strftime("%d/%m/%Y")
        
    return res

# --- 2. Word Document Helper ---
def replace_placeholders(doc, mapping):
    # Paragraphs aur Tables dono ke paragraphs nikalna
    all_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paras.extend(cell.paragraphs)
                
    for p in all_paras:
        for key, val in mapping.items():
            placeholder = "{{ " + str(key) + " }}"
            if placeholder in p.text:
                # Full text merge to handle split runs
                full_text = "".join(run.text for run in p.runs)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(val))
                    for i, run in enumerate(p.runs):
                        run.text = new_text if i == 0 else ""

# --- 3. Main Streamlit Interface ---
st.set_page_config(layout="wide", page_title="PME Management")

if db is not None:
    # Employee list fetch karna
    docs = db.collection("employees").stream()
    df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

    if 'memo_file' not in st.session_state: st.session_state.memo_file = None

    tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 History", "🛠 Update"])

    with tab1:
        if not df_emp.empty:
            emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            selected = st.selectbox("Select Employee", emp_names)
            h_id = selected.split('(')[-1].strip(')')
            emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

            with st.form("pme_form"):
                memo_date = st.date_input("Memo Date", value=datetime.now())
                
                # Metrics calculate karna
                m = calculate_pme_metrics(emp_data.get('DOB'), emp_data.get('DOA'))
                
                final_vals = {
                    "dob": m["dob_f"],
                    "doa": m["doa_f"],
                    "name": emp_data.get('Employee Name', ''),
                    "age": m["age"],
                    "father_name": emp_data.get("FATHER'S NAME", ''),
                    "designation": emp_data.get('Designation', ''),
                    "medical_category": emp_data.get('Medical category', ''),
                    "current_date": memo_date.strftime("%d/%m/%Y"),
                    "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                    "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                    "last_examined_date": emp_data.get('Last PME', 'N/A'),
                    "last_place": emp_data.get('Last PME Place', 'NKJ'),
                    "examiner": "ACMS/NKJ",
                    "service_year": m["s_yr"],
                    "service_month": m["s_mn"]
                }

                if st.form_submit_button("Generate Memo"):
                    path = os.path.join(BASE_DIR, "assets", "pme memo temp.docx")
                    if os.path.exists(path):
                        doc = Document(path)
                        replace_placeholders(doc, final_vals)
                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state.memo_file = buf.getvalue()
                        db.collection("pme_history").add({**final_vals, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                        st.success(f"Generated! Age: {m['age']}, Service: {m['s_yr']}y {m['s_mn']}m")
                    else:
                        st.error("Template not found!")

            if st.session_state.memo_file:
                st.download_button("📥 Download PME Memo", st.session_state.memo_file, f"PME_{h_id}.docx")
        else:
            st.warning("No employees in database.")
            
    with tab2:
        h_docs = db.collection("pme_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).limit(10).stream()
        h_list = [{**d.to_dict()} for d in h_docs]
        if h_list:
            st.table(pd.DataFrame(h_list)[['Timestamp', 'name', 'age', 'service_year']])
else:
    st.error("Database connection failed.")
