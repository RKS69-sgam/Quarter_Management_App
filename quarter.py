import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Login & Firebase Setup ---
ADMIN_USER = "admin"
ADMIN_PASS = "Sgam@4321"

def check_login():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if not st.session_state.logged_in:
        st.title("🔐 Railway Quarter System Login")
        with st.form("login"):
            u, p = st.text_input("Username"), st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                if u == ADMIN_USER and p == ADMIN_PASS:
                    st.session_state.logged_in = True
                    st.rerun()
                else: st.error("❌ Galat Credentials")
        return False
    return True

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

# --- 1. Data Fetching Logic (Image ke columns ke anusar) ---
@st.cache_data(ttl=60)
def get_employees_cached():
    db = init_db()
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_full_quarter_history():
    db = init_db()
    try:
        docs = db.collection("quarter_history").stream()
        data = []
        for d in docs:
            item = d.to_dict()
            item['id'] = d.id
            # Date Handling
            for d_field in ['allotment_date', 'vacation_date']:
                val = item.get(d_field)
                if val and hasattr(val, 'strftime'):
                    item[f'{d_field}_disp'] = val.strftime('%d-%m-%Y')
                else:
                    item[f'{d_field}_disp'] = "N/A"
            data.append(item)
        return pd.DataFrame(data)
    except Exception as e:
        return pd.DataFrame()

# --- 2. Template Filling Logic ---
def fill_template(template_path, db_item, manual_date):
    doc = Document(template_path)
    mapping = {
        "EMPLOYEE_NAME": str(db_item.get('employee_name', '')),
        "DESIGNATION": str(db_item.get('designation', '')),
        "PF_Number": str(db_item.get('pf_number', '')),
        "HRMS_ID": str(db_item.get('hrms_id', '')),
        "QUARTER_NUMBER": str(db_item.get('quarter_number', '')),
        "STATION": str(db_item.get('station', '')),
        "UNIT": str(db_item.get('unit', 'NA')),
        "DATE": manual_date
    }
    for obj in list(doc.paragraphs) + [p for table in doc.tables for row in table.rows for cell in row.cells for p in cell.paragraphs]:
        for key, val in mapping.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in obj.text:
                obj.text = obj.text.replace(placeholder, val)
    return doc

# --- 3. Main App ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    if not check_login(): return

    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Report & Dashboard"])

    # Data Loading
    df_emp = get_employees_cached()
    df_hist = get_full_quarter_history()

    # Master List of Quarters & Stations (Aap ise database se bhi link kar sakte hain)
    ALL_STATIONS = ["सरईग्राम", "ब्योहारी", "विजयसोता", "छतैनी"]
    # Maan lijiye ye aapke pas available quarters ki master list hai
    MASTER_QUARTERS = ["T-01/A", "T-01/B", "T-02/A", "T-02/B", "RB-I/12/B", "RB-II/05/A"]

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Quarter Allotment")
        if not df_emp.empty:
            emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name', 'Unknown')} ({r.get('HRMS ID', 'NA')})", axis=1).tolist()
            selected_emp = st.selectbox("Staff Select Karein", emp_options)
            
            # Logic: Jo quarters abhi Occupied hain unhe list se hatana
            occupied_list = []
            if not df_hist.empty and 'is_current' in df_hist.columns:
                occupied_list = df_hist[df_hist['is_current'] == True]['quarter_number'].tolist()
            
            available_quarters = [q for q in MASTER_QUARTERS if q not in occupied_list]
            
            col1, col2 = st.columns(2)
            with col1:
                selected_q = st.selectbox("Khali Quarter Select Karein", available_quarters)
            with col2:
                selected_stn = st.selectbox("Station Select Karein", ALL_STATIONS)
            
            allot_date = st.date_input("Allotment Date")

            if st.button("Generate Allotment Letter"):
                h_id = selected_emp.split('(')[-1].strip(')')
                emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                new_entry = {
                    "employee_name": str(emp_row['Employee Name']),
                    "designation": str(emp_row['Designation']),
                    "hrms_id": h_id,
                    "pf_number": str(emp_row.get('PF Number', '')),
                    "quarter_number": selected_q,
                    "station": selected_stn,
                    "unit": str(emp_row.get('Unit', 'NA')),
                    "allotment_date": datetime.combine(allot_date, datetime.min.time()),
                    "is_current": True,
                    "vacation_date": None
                }
                db.collection("quarter_history").add(new_entry)
                st.cache_data.clear() # Cache clear taaki list update ho jaye
                
                doc = fill_template(ALLOT_TEMP, new_entry, allot_date.strftime("%d/%m/%Y"))
                buf = io.BytesIO(); doc.save(buf)
                st.success(f"✅ Quarter {selected_q} successfully allotted!")
                st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{selected_q}.docx")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation")
        if not df_hist.empty and 'is_current' in df_hist.columns:
            occupied = df_hist[df_hist['is_current'] == True]
            if not occupied.empty:
                q_list = occupied.apply(lambda r: f"{r['quarter_number']} - {r['employee_name']}", axis=1).tolist()
                sel_vac = st.selectbox("Select Quarter to Vacate", q_list)
                v_date = st.date_input("Vacation Date")

                if st.button("Process Vacation"):
                    q_row = occupied.iloc[q_list.index(sel_vac)]
                    db.collection("quarter_history").document(q_row['id']).update({
                        "is_current": False,
                        "vacation_date": datetime.combine(v_date, datetime.min.time())
                    })
                    st.cache_data.clear()
                    doc = fill_template(VACATE_TEMP, q_row.to_dict(), v_date.strftime("%d/%m/%Y"))
                    buf = io.BytesIO(); doc.save(buf)
                    st.success("✅ Quarter Vacated!")
                    st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{q_row['quarter_number']}.docx")
            else: st.info("Abhi koi quarter occupied nahi hai.")

    # --- TAB 3: REPORT ---
    with tab3:
        st.header("📊 Quarter Master Report")
        if not df_hist.empty:
            # Dashboard Counters
            total_occ = len(df_hist[df_hist['is_current'] == True])
            st.metric("Total Occupied Quarters", total_occ)

            # Table Display with Station
            disp_df = df_hist[['quarter_number', 'employee_name', 'station', 'allotment_date_disp', 'vacation_date_disp', 'is_current']].copy()
            disp_df.rename(columns={
                'quarter_number': 'Quarter No', 'employee_name': 'Staff Name',
                'station': 'Station', 'allotment_date_disp': 'Allotted', 
                'vacation_date_disp': 'Vacated', 'is_current': 'Status'
            }, inplace=True)

            st.dataframe(disp_df, use_container_width=True)
            
            # CSV Export (Ensuring Station is included)
            csv = disp_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Download CSV Report (with Station)", csv, "Quarter_Report.csv", "text/csv")
        else: st.warning("Database khali hai.")

if __name__ == "__main__":
    main()
