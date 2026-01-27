import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. Path & Firebase Setup ---
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
            st.error(f"Firebase Error: {e}"); return None
    return firestore.client()

db = init_db()

# --- 1. Utilities ---
def get_safe_date(date_val):
    if not date_val or str(date_val).lower() == 'nan' or date_val == "":
        return None
    try:
        if hasattr(date_val, 'to_datetime'):
            return date_val.to_datetime().replace(tzinfo=None)
        return pd.to_datetime(date_val, dayfirst=True).to_pydatetime().replace(tzinfo=None)
    except: return None

def calculate_next_pme(last_pme_raw, dob_raw, medical_cat):
    """
    Railway Medical Rules (IRMM):
    A1, A2, A3: <45 (4yr), 45-55 (2yr), 55+ (1yr)
    B1, B2: First at 45 years, then every 5 years.
    """
    last_pme = get_safe_date(last_pme_raw)
    dob = get_safe_date(dob_raw)
    if not dob: return None
    
    cat = str(medical_cat).upper().strip()
    
    # Logic for A-Categories
    if any(x in cat for x in ["A1", "A2", "A3"]):
        if not last_pme: return None
        age_at_pme = relativedelta(last_pme, dob).years
        interval = 4 if age_at_pme < 45 else (2 if age_at_pme < 55 else 1)
        return last_pme + relativedelta(years=interval)
    
    # Logic for B-Categories (B1/B2)
    elif "B" in cat:
        # If never examined, due on 45th birthday
        due_45 = dob + relativedelta(years=45)
        if not last_pme: return due_45
        
        age_at_pme = relativedelta(last_pme, dob).years
        if age_at_pme < 45: 
            return due_45
        else: 
            return last_pme + relativedelta(years=5)
            
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

def replace_placeholders(doc, mapping):
    all_paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paras.extend(cell.paragraphs)
    for p in all_paras:
        for key, val in mapping.items():
            placeholder = "{{ " + str(key) + " }}"
            if placeholder in p.text:
                new_text = p.text.replace(placeholder, str(val))
                p.text = new_text

# --- 2. Main Interface ---
st.set_page_config(layout="wide", page_title="Railway PME System")

if db is not None:
    docs = db.collection("employees").stream()
    df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

    # --- 📢 PME ALERT DASHBOARD ---
    st.subheader("⚠️ PME Alerts (Next 15 Days)")
    if not df_emp.empty:
        today = datetime.now()
        alert_window = today + timedelta(days=15)
        alerts = []

        for _, row in df_emp.iterrows():
            next_pme = calculate_next_pme(row.get('Last PME'), row.get('DOB'), row.get('Medical category'))
            if next_pme and next_pme <= alert_window:
                alerts.append({
                    "Name": row.get('Employee Name'),
                    "Category": row.get('Medical category'),
                    "Next PME Due": next_pme.strftime("%d/%m/%Y"),
                    "Status": "OVERDUE" if next_pme < today else "DUE SOON"
                })
        
        if alerts:
            alert_df = pd.DataFrame(alerts)
            st.warning(f"Total {len(alerts)} employees need medical attention.")
            st.table(alert_df)
        else:
            st.success("✅ No PME due in the next 15 days.")

    tab1, tab2, tab3 = st.tabs(["📝 Generate Memo", "📊 History", "🛠 Update Database"])

    with tab1:
        if not df_emp.empty:
            emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            selected = st.selectbox("Select Employee", emp_names)
            h_id = selected.split('(')[-1].strip(')')
            emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

            with st.form("pme_gen_form"):
                memo_date = st.date_input("Memo Date", value=datetime.now())
                m = calculate_pme_metrics(emp_data.get('DOB'), emp_data.get('DOA'))
                
                final_mapping = {
                    "dob": m["dob_f"], "doa": m["doa_f"],
                    "name": emp_data.get('Employee Name', ''), "age": m["age"],
                    "father_name": emp_data.get("FATHER'S NAME", ''),
                    "designation": emp_data.get('Designation', ''),
                    "medical_category": emp_data.get('Medical category', ''),
                    "current_date": memo_date.strftime("%d/%m/%Y"),
                    "first_physical_mark": emp_data.get('Physical Mark 1', 'N/A'),
                    "second_physical_mark": emp_data.get('Physical Mark 2', 'N/A'),
                    "last_examined_date": emp_data.get('Last PME', 'N/A'),
                    "last_place": emp_data.get('Last PME Place', 'N/A'),
                    "examiner": emp_data.get('Last Examiner', 'ACMS/NKJ'),
                    "service_year": m["s_yr"], "service_month": m["s_mn"]
                }

                if st.form_submit_button("Generate & Log"):
                    if os.path.exists(TEMPLATE_PATH):
                        doc = Document(TEMPLATE_PATH)
                        replace_placeholders(doc, final_mapping)
                        buf = io.BytesIO()
                        doc.save(buf)
                        st.session_state.memo_bytes = buf.getvalue()
                        db.collection("pme_history").add({**final_mapping, "Timestamp": datetime.now(), "HRMS_ID": h_id})
                        st.success("✅ Memo Generated!")
                        st.rerun()
                    else: st.error("Template not found!")

            if st.session_state.get('memo_bytes'):
                st.download_button("📥 Download", st.session_state.memo_bytes, f"PME_{h_id}.docx")

    with tab3:
        st.header("🛠 Medical Data Update")
        t_sel = st.selectbox("Select Employee", emp_names, key="upd_pme")
        t_id = t_sel.split('(')[-1].strip(')')
        t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
        
        with st.form("update_medical"):
            lp_date = st.text_input("New PME Date (DD/MM/YYYY)", value=t_row.get('Last PME', ''))
            lp_cat = st.text_input("Medical Category", value=t_row.get('Medical category', ''))
            lp_place = st.text_input("Place", value=t_row.get('Last PME Place', ''))
            
            if st.form_submit_button("Update Firebase"):
                db.collection("employees").document(t_row['id']).update({
                    "Last PME": lp_date, "Medical category": lp_cat, "Last PME Place": lp_place
                })
                st.success("Database Updated! Dashboard refreshing...")
                st.rerun()
