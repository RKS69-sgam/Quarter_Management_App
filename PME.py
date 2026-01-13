import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Path aur Firebase Setup ---
# BASE_DIR ko yahan define kiya gaya hai taaki NameError na aaye
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
            st.error(f"Firebase Error: {e}")
            return None
    return firestore.client()

db = init_db()

# --- 1. Date aur Calculation Logic (Fixed for N/A Issue) ---
def get_safe_date(date_val):
    """Har tarah ke date format ko datetime object mein badalne ke liye"""
    if not date_val or str(date_val).lower() == 'nan' or date_val == "":
        return None
    try:
        # Agar Firebase Timestamp hai
        if hasattr(date_val, 'to_datetime'):
            return date_val.to_datetime().replace(tzinfo=None)
        # Agar string hai toh use parse karein
        return pd.to_datetime(date_val).to_pydatetime().replace(tzinfo=None)
    except:
        return None

def calculate_pme_metrics(dob_raw, doa_raw):
    now = datetime.now()
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

# --- 2. Word Document Replacement Logic ---
def replace_placeholders(doc, mapping):
    # Paragraphs aur Tables dono cover honge
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

# --- 3. Streamlit UI ---
st.set_page_config(layout="wide", page_title="PME Management")

if db is not None:
    # Employees fetch karein
    docs = db.collection("employees").stream()
    df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

    if 'memo_bytes' not in st.session_state: st.session_state.memo_bytes = None

    tab1, tab2, tab3 = st.tabs(["📝 Generate PME Memo", "📊 History", "🛠 Update"])

    with tab1:
        if not df_emp.empty:
            emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            selected = st.selectbox("Select Employee", emp_names)
            h_id = selected.split('(')[-1].strip(')')
            emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

            with st.form("pme_gen_form"):
                memo_date = st.date_input("Memo Date", value=datetime.now())
                
                # Sahi Metrics Calculate karein
                m = calculate_pme_metrics(emp_data.get('DOB'), emp_data.get('DOA'))
                
                final_mapping = {
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
                    if os.path.exists(TEMPLATE_PATH):
                        doc = Document(TEMPLATE_PATH)
                        replace_placeholders(doc, final_mapping)
                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state.memo_bytes = buf.getvalue()
                        # History mein save karein
                        db.collection("pme_history").add({**final_mapping, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                        st.success(f"✅ Generated! Age: {m['age']}, Service: {m['s_yr']}y {m['s_mn']}m")
                    else:
                        st.error(f"Template nahi mila: {TEMPLATE_PATH}")

            if st.session_state.memo_bytes:
                st.download_button("📥 Download PME Memo", st.session_state.memo_bytes, f"PME_{h_id}.docx")
        else:
            st.warning("Database khali hai.")

    with tab2:
        st.header("History Records")
        h_docs = db.collection("pme_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).limit(20).stream()
        h_list = [{**d.to_dict()} for d in h_docs]
        if h_list:
            st.dataframe(pd.DataFrame(h_list)[['Timestamp', 'name', 'age', 'service_year']], use_container_width=True)
else:
    st.error("Database connection fail ho gaya.")
