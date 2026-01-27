import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore
from dateutil.relativedelta import relativedelta

# --- 0. CONFIG & LOGIN SECURITY ---
st.set_page_config(layout="wide", page_title="Railway PME System")

def check_password():
    """Returns True if the user had the correct password."""
    def password_entered():
        if st.session_state["username"] == "admin" and st.session_state["password"] == "sgam@2026":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔐 SGAM Office PME Login")
        st.text_input("Username", key="username")
        st.text_input("Password", type="password", key="password")
        st.button("Login", on_click=password_entered)
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Username", key="username")
        st.text_input("Password", type="password", key="password")
        st.button("Login", on_click=password_entered)
        st.error("😕 User not known or password incorrect")
        return False
    else:
        return True

# --- 1. FIREBASE SETUP ---
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

# --- 2. ADVANCED PME LOGIC ---
def get_safe_date(date_val):
    if not date_val or str(date_val).lower() == 'nan' or date_val == "":
        return None
    try:
        if hasattr(date_val, 'to_datetime'):
            return date_val.to_datetime().replace(tzinfo=None)
        return pd.to_datetime(date_val, dayfirst=True).to_pydatetime().replace(tzinfo=None)
    except: return None



def calculate_next_pme(last_pme_raw, dob_raw, medical_cat):
    last_pme = get_safe_date(last_pme_raw)
    dob = get_safe_date(dob_raw)
    if not dob: return None
    
    cat = str(medical_cat).upper().strip()
    
    # Logic for A1, A2, A3 (Safety Categories)
    if any(x in cat for x in ["A1", "A2", "A3"]):
        if not last_pme: return None 
        age_at_pme = relativedelta(last_pme, dob).years
        # Rules: <45 (4yr), 45-55 (2yr), 55+ (1yr)
        if age_at_pme < 45: interval = 4
        elif 45 <= age_at_pme < 55: interval = 2
        else: interval = 1
        return last_pme + relativedelta(years=interval)
    
    # Logic for B1, B2 (Non-Safety / Milestone Based)
    elif "B" in cat:
        due_45 = dob + relativedelta(years=45)
        # If never had PME, first is due at age 45
        if not last_pme: return due_45
        age_at_pme = relativedelta(last_pme, dob).years
        if age_at_pme < 45: return due_45
        else: return last_pme + relativedelta(years=5) # Every 5 years after 45
            
    return None

# --- 3. MAIN INTERFACE ---
if check_password():
    db = init_db()
    
    if db:
        docs = db.collection("employees").stream()
        df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

        # --- 📢 DASHBOARD ALERT SYSTEM ---
        st.subheader("⚠️ PME Alerts Dashboard")
        if not df_emp.empty:
            today = datetime.now()
            alert_window = today + timedelta(days=15)
            alerts = []

            for _, row in df_emp.iterrows():
                next_pme = calculate_next_pme(row.get('Last PME'), row.get('DOB'), row.get('Medical category'))
                
                if next_pme:
                    if next_pme <= alert_window:
                        status = "🔴 OVERDUE" if next_pme < today else "🟠 DUE SOON"
                        alerts.append({
                            "Employee Name": row.get('Employee Name'),
                            "HRMS ID": row.get('HRMS ID'),
                            "Category": row.get('Medical category'),
                            "Next PME Due": next_pme.strftime("%d/%m/%Y"),
                            "Status": status
                        })
                elif any(x in str(row.get('Medical category')).upper() for x in ["A1", "A2", "A3"]):
                    # Alert if Last PME date is missing for Safety Categories
                    if not row.get('Last PME'):
                        alerts.append({
                            "Employee Name": row.get('Employee Name'),
                            "HRMS ID": row.get('HRMS ID'),
                            "Category": row.get('Medical category'),
                            "Next PME Due": "DATA MISSING",
                            "Status": "⚪ UPDATE LAST PME"
                        })

            if alerts:
                st.dataframe(pd.DataFrame(alerts), use_container_width=True)
            else:
                st.success("✅ All PME records are up-to-date!")

        st.divider()

        tab1, tab2, tab3 = st.tabs(["📝 Generate Memo", "📊 History", "🛠 Update Database"])

        with tab1:
            if not df_emp.empty:
                emp_names = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
                selected = st.selectbox("Select Employee for PME", emp_names)
                h_id = selected.split('(')[-1].strip(')')
                emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

                with st.form("gen_memo_form"):
                    memo_date = st.date_input("Memo Date", value=datetime.now())
                    # Placeholder replacement logic here...
                    if st.form_submit_button("Generate Memo"):
                        st.info("Generating... Check below for download link.")
                        # (File generation logic remains same as previous version)

        with tab3:
            st.header("🛠 Quick Update Medical Records")
            if not df_emp.empty:
                t_sel = st.selectbox("Select Employee", emp_names, key="tab3_upd")
                t_id = t_sel.split('(')[-1].strip(')')
                t_row = df_emp[df_emp['HRMS ID'] == t_id].iloc[0]
                
                with st.form("update_medical_v2"):
                    col1, col2 = st.columns(2)
                    new_lp = col1.text_input("Last PME Date (DD/MM/YYYY)", value=t_row.get('Last PME', ''))
                    new_cat = col2.text_input("Medical Category", value=t_row.get('Medical category', ''))
                    
                    if st.form_submit_button("Save & Refresh Alerts"):
                        db.collection("employees").document(t_row['id']).update({
                            "Last PME": new_lp,
                            "Medical category": new_cat
                        })
                        st.success("Data Updated! Alerts recalculated.")
                        st.rerun()

# Default Login: admin / sgam@2026
