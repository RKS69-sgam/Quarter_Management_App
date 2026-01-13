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

# --- 1. Robust Date Utilities ---
def parse_date(date_val):
    """Firebase Timestamp ya String ko datetime object mein badalne ke liye"""
    if not date_val or str(date_val).lower() == 'nan':
        return None
    try:
        # Agar already Timestamp hai (Firebase specific)
        if hasattr(date_val, 'to_datetime'):
            return date_val.to_datetime().replace(tzinfo=None)
        # Agar string hai toh pandas se parse karein
        return pd.to_datetime(date_val).to_pydatetime().replace(tzinfo=None)
    except:
        return None

def calculate_pme_metrics(dob_raw, doa_raw):
    """Age aur Service Length ka nishchit calculation"""
    now = datetime.now()
    data = {"age": "N/A", "s_yr": "0", "s_mn": "0", "dob_f": "", "doa_f": ""}
    
    dt_dob = parse_date(dob_raw)
    dt_doa = parse_date(doa_raw)
    
    if dt_dob:
        data["age"] = str(relativedelta(now, dt_dob).years)
        data["dob_f"] = dt_dob.strftime("%d/%m/%Y")
        
    if dt_doa:
        diff = relativedelta(now, dt_doa)
        data["s_yr"] = str(diff.years)
        data["s_mn"] = str(diff.months)
        data["doa_f"] = dt_doa.strftime("%d/%m/%Y")
        
    return data

def replace_placeholders(doc, mapping):
    """Template ke paragraphs aur tables mein text replace karne ka sahi tareeka"""
    # Paragraphs + Table cells ke saare paragraphs ek saath
    target_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                target_paras.extend(cell.paragraphs)
                
    for p in target_paras:
        for key, val in mapping.items():
            # Aapke template ke hisaab se format: {{ key }}
            placeholder = "{{ " + str(key) + " }}"
            if placeholder in p.text:
                # Run-level replacement formatting bachane ke liye (Bold/Italic fix)
                full_text = "".join(run.text for run in p.runs)
                if placeholder in full_text:
                    new_text = full_text.replace(placeholder, str(val))
                    # Saare runs clear karke pehle run mein naya text daalna
                    for i, run in enumerate(p.runs):
                        run.text = new_text if i == 0 else ""

# --- 2. Main Streamlit App ---
st.set_page_config(layout="wide", page_title="PME Management")
df_emp = get_employees() if 'db' in locals() else pd.DataFrame()

# Download data handle karne ke liye state
if 'memo_bytes' not in st.session_state: st.session_state.memo_bytes = None

tab1, tab2, tab3 = st.tabs(["📝 Memo Generate", "📊 History", "🛠 Update Data"])

with tab1:
    if not df_emp.empty:
        # Searchable selectbox
        emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected_emp = st.selectbox("Employee Chunein", emp_names)
        hrms_id = selected_emp.split('(')[-1].strip(')')
        emp_row = df_emp[df_emp['HRMS ID'] == hrms_id].iloc[0]

        with st.form("gen_pme_form"):
            memo_date = st.date_input("Memo Ki Tarikh", value=datetime.now())
            
            # Calculation yahan ho rahi hai
            metrics = calculate_pme_metrics(emp_row.get('DOB'), emp_row.get('DOA'))
            
            # Final mapping for Document
            final_map = {
                "dob": metrics["dob_f"],
                "doa": metrics["doa_f"],
                "name": emp_row.get('Employee Name', ''),
                "age": metrics["age"],
                "father_name": emp_row.get("FATHER'S NAME", ''),
                "designation": emp_row.get('Designation', ''),
                "medical_category": emp_row.get('Medical category', ''),
                "current_date": memo_date.strftime("%d/%m/%Y"),
                "first_physical_mark": emp_row.get('Physical Mark 1', 'N/A'),
                "second_physical_mark": emp_row.get('Physical Mark 2', 'N/A'),
                "last_examined_date": emp_row.get('Last PME', 'N/A'),
                "last_place": emp_row.get('Last PME Place', 'NKJ'),
                "examiner": "ACMS/NKJ",
                "service_year": metrics["s_yr"],
                "service_month": metrics["s_mn"]
            }

            if st.form_submit_button("Generate Document"):
                if os.path.exists(TEMPLATE_PATH):
                    doc = Document(TEMPLATE_PATH)
                    replace_placeholders(doc, final_map)
                    
                    buf = io.BytesIO()
                    doc.save(buf)
                    st.session_state.memo_bytes = buf.getvalue()
                    
                    # History entry
                    db.collection("pme_history").add({**final_map, "Timestamp": datetime.now(), "HRMS_ID": hrms_id})
                    st.success(f"✅ Taiyar! Age: {metrics['age']}, Service: {metrics['s_yr']} saal")
                else:
                    st.error("Template file nahi mili!")

        # Form ke bahar download button
        if st.session_state.memo_bytes:
            st.download_button("📥 Download PME Memo", st.session_state.memo_bytes, f"PME_{hrms_id}.docx")
    else:
        st.warning("Database khali hai.")

# ... (Tab 2 aur Tab 3 ka code pehle jaisa simple rahega)


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
