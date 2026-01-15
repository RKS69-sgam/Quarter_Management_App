import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 1. Security Login ---
ADMIN_USER = "admin"
ADMIN_PASS = "Sgam@4321"

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    
    if not st.session_state.logged_in:
        st.title("🔐 Railway Quarter Management Login")
        with st.form("login_form"):
            user = st.text_input("Username")
            pas = st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                if user == ADMIN_USER and pas == ADMIN_PASS:
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ Galat Username ya Password")
        return False
    return True

# --- 2. Firebase Setup ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        try:
            if "firebase_config" in st.secrets:
                cred = credentials.Certificate(dict(st.secrets["firebase_config"]))
            else:
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

# --- 3. Data Sync & Fetch Logic ---
def get_synced_data():
    db = init_db()
    # Fetch Employees for matching
    emp_docs = db.collection("employees").stream()
    emp_map = {str(d.to_dict().get('PF Number')): d.to_dict() for d in emp_docs if d.to_dict().get('PF Number')}
    
    # Fetch Quarter History
    hist_docs = db.collection("quarter_history").stream()
    data = []
    for d in hist_docs:
        item = d.to_dict()
        item['id'] = d.id
        pf = str(item.get('pf_number', ''))
        
        # PF se Designation aur Unit fetch karna agar available ho
        if pf in emp_map:
            item['designation'] = emp_map[pf].get('Designation', item.get('designation', ''))
            item['unit'] = emp_map[pf].get('Unit', item.get('unit', ''))
        
        # Display dates
        item['allot_disp'] = item['allotment_date'].strftime('%d-%m-%Y') if item.get('allotment_date') and hasattr(item['allotment_date'], 'strftime') else "N/A"
        item['vacat_disp'] = "🔴 Occupied" if item.get('is_current') else (item['vacation_date'].strftime('%d-%m-%Y') if item.get('vacation_date') and hasattr(item['vacation_date'], 'strftime') else "🟢 Vacant")
        data.append(item)
    return pd.DataFrame(data), emp_map

# --- 4. Main UI ---
def main():
    if not check_login(): return

    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    db = init_db()
    
    # Logout button in sidebar
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()

    df_hist, emp_map = get_synced_data()
    
    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Master Report"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Quarter Allotment")
        if not df_hist.empty:
            # Employee selection from synced employee map
            emp_list = [f"{v['Employee Name']} ({k})" for k, v in emp_map.items()]
            selected_emp = st.selectbox("Staff Chunein", emp_list)
            
            # Logic: Karmchari ke station par vacant quarter
            pf_key = selected_emp.split('(')[-1].strip(')')
            emp_station = emp_map[pf_key].get('Station', 'सरईग्राम') # Default agar station na ho
            
            st.info(f"Karmchari ka Station: {emp_station}")

            # Filter: Is station ke woh quarters jo abhi 'is_current' nahi hain
            station_qs = df_hist[df_hist['station'] == emp_station]
            occ_qs = station_qs[station_qs['is_current'] == True]['quarter_number'].tolist()
            available_qs = [q for q in station_qs['quarter_number'].unique() if q not in occ_qs]

            if not available_qs:
                st.warning(f"{emp_station} par koi vacant quarter nahi mila.")
            
            sel_q = st.selectbox("Vacant Quarters (at selected station)", available_qs)
            a_date = st.date_input("Allotment Date")

            if st.button("Allot & Generate Letter"):
                emp_info = emp_map[pf_key]
                new_allot = {
                    "employee_name": emp_info['Employee Name'],
                    "designation": emp_info['Designation'],
                    "unit": emp_info.get('Unit', 'SSE/PW/SGAM'),
                    "hrms_id": emp_info.get('HRMS ID', ''),
                    "pf_number": pf_key,
                    "quarter_number": sel_q,
                    "station": emp_station,
                    "allotment_date": datetime.combine(a_date, datetime.min.time()),
                    "is_current": True,
                    "vacation_date": None
                }
                db.collection("quarter_history").add(new_allot)
                st.success(f"Success! {sel_q} allotted to {emp_info['Employee Name']}")
                st.rerun()

    # --- TAB 3: REPORT (With Station and Synced Data) ---
    with tab3:
        st.header("📊 Master Quarter Report")
        if not df_hist.empty:
            # Table display
            disp_df = df_hist[['quarter_number', 'employee_name', 'designation', 'unit', 'station', 'allot_disp', 'vacat_disp']].copy()
            disp_df.columns = ['Quarter', 'Name', 'Designation', 'Unit', 'Station', 'Allotted', 'Status']
            st.dataframe(disp_df, use_container_width=True)
            
            csv = disp_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Export CSV", csv, "Quarter_Report.csv", "text/csv")

if __name__ == "__main__":
    main()
