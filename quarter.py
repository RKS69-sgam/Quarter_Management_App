import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 1. Security Login Configuration ---
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

# --- 3. Advance Data Fetching & PF Sync Logic ---
def get_master_data():
    db = init_db()
    # Employees fetch karein (Matching aur Station ke liye)
    emp_docs = db.collection("employees").stream()
    # PF Number key ke sath map karein
    emp_map = {str(d.to_dict().get('PF Number')): d.to_dict() for d in emp_docs if d.to_dict().get('PF Number')}
    
    # Quarter History fetch karein
    hist_docs = db.collection("quarter_history").stream()
    data = []
    for d in hist_docs:
        item = d.to_dict()
        item['id'] = d.id
        pf = str(item.get('pf_number', ''))
        
        # PF se Sync: Designation aur Unit ko employee list se fetch karein
        if pf in emp_map:
            item['designation'] = emp_map[pf].get('Designation', item.get('designation', ''))
            item['unit'] = emp_map[pf].get('Unit', item.get('unit', ''))
            item['station'] = item.get('station', emp_map[pf].get('Station', 'N/A'))
        
        # Date Formatting
        item['allot_disp'] = item['allotment_date'].strftime('%d-%m-%Y') if item.get('allotment_date') and hasattr(item['allotment_date'], 'strftime') else "N/A"
        item['status_disp'] = "🔴 Occupied" if item.get('is_current') else "🟢 Vacant"
        data.append(item)
        
    return pd.DataFrame(data), emp_map

# --- 4. Template Generator ---
def fill_template(temp_path, data, date_str):
    doc = Document(temp_path)
    mapping = {
        "EMPLOYEE_NAME": str(data.get('employee_name', '')),
        "DESIGNATION": str(data.get('designation', '')),
        "PF_Number": str(data.get('pf_number', '')),
        "HRMS_ID": str(data.get('hrms_id', '')),
        "QUARTER_NUMBER": str(data.get('quarter_number', '')),
        "STATION": str(data.get('station', '')),
        "UNIT": str(data.get('unit', '')),
        "DATE": date_str
    }
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in mapping.items():
            if f"{{{{{k}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{k}}}}}", v)
    return doc

# --- 5. Main UI Interface ---
def main():
    if not check_login(): return

    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    db = init_db()
    
    if st.sidebar.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.rerun()

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    df_hist, emp_map = get_master_data()
    
    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Master Report"])

    # --- TAB 1: ALLOTMENT (SMART FILTER LOGIC) ---
    with tab1:
        st.header("New Allotment Form")
        if not df_hist.empty and emp_map:
            # 1. Karmchari Chunein
            emp_list = [f"{v['Employee Name']} ({k})" for k, v in emp_map.items()]
            selected_emp = st.selectbox("Staff Select Karein", emp_list)
            
            # 2. Selected Karmchari ka Station fetch karein
            pf_key = selected_emp.split('(')[-1].strip(')')
            staff_station = str(emp_map[pf_key].get('STATION', '')).strip()
            
            st.info(f"📍 Karmchari ka Allotted Station: **{staff_station}**")

            # 3. Vacant Quarter Filter: Keval Staff ke Station par
            station_qs = df_hist[df_hist['station'].str.strip() == staff_station]
            
            # Jo abhi Occupied hain unhe hatayein
            occupied_qs = station_qs[station_qs['is_current'] == True]['quarter_number'].unique().tolist()
            available_qs = [q for q in station_qs['quarter_number'].unique() if q not in occupied_qs]

            if available_qs:
                sel_q = st.selectbox(f"Vacant Quarters at {staff_station}", available_qs)
                a_date = st.date_input("Allotment Date", value=datetime.now())

                if st.button("Allot & Generate Letter"):
                    emp_info = emp_map[pf_key]
                    allot_entry = {
                        "employee_name": emp_info['Employee Name'],
                        "designation": emp_info['Designation'],
                        "unit": emp_info.get('Unit', ''),
                        "hrms_id": emp_info.get('HRMS ID', ''),
                        "pf_number": pf_key,
                        "quarter_number": sel_q,
                        "station": staff_station,
                        "allotment_date": datetime.combine(a_date, datetime.min.time()),
                        "is_current": True,
                        "vacation_date": None
                    }
                    db.collection("quarter_history").add(allot_entry)
                    st.success(f"✅ Allotted {sel_q} to {emp_info['Employee Name']}")
                    
                    doc = fill_template(ALLOT_TEMP, allot_entry, a_date.strftime("%d/%m/%Y"))
                    buf = io.BytesIO(); doc.save(buf)
                    st.download_button("📥 Download Letter", buf.getvalue(), f"Allotment_{sel_q}.docx")
            else:
                st.warning(f"⚠️ {staff_station} par filhal koi Vacant Quarter available nahi hai.")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Vacation Process")
        occ_df = df_hist[df_hist['is_current'] == True]
        if not occ_df.empty:
            v_options = occ_df.apply(lambda r: f"{r['quarter_number']} - {r['employee_name']}", axis=1).tolist()
            sel_v = st.selectbox("Select Occupied Quarter", v_options)
            v_date = st.date_input("Vacation Date")
            
            if st.button("Process Vacation"):
                idx = v_options.index(sel_v)
                q_row = occ_df.iloc[idx]
                db.collection("quarter_history").document(q_row['id']).update({
                    "is_current": False,
                    "vacation_date": datetime.combine(v_date, datetime.min.time())
                })
                st.success("Quarter Marked Vacant!")
                st.rerun()
        else:
            st.info("Koi quarter occupied nahi hai.")

    # --- TAB 3: MASTER REPORT ---
    with tab3:
        st.header("📊 Master Database (Synced with Employees)")
        if not df_hist.empty:
            # Sync check and Display
            disp_df = df_hist[['quarter_number', 'employee_name', 'designation', 'unit', 'station', 'allot_disp', 'status_disp']].copy()
            disp_df.columns = ['Quarter', 'Name', 'Designation', 'Unit', 'Station', 'Allotted', 'Status']
            
            st.dataframe(disp_df, use_container_width=True)
            
            csv = disp_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Download Full CSV", csv, "Quarter_Report.csv", "text/csv")

if __name__ == "__main__":
    main()
