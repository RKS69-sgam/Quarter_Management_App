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

# --- 3. Advance Data Fetching & Sync ---
def get_master_data():
    db = init_db()
    # Employees Fetch for Sync
    emp_docs = db.collection("employees").stream()
    emp_map = {str(d.to_dict().get('PF Number')): d.to_dict() for d in emp_docs if d.to_dict().get('PF Number')}
    
    # History Fetch
    hist_docs = db.collection("quarter_history").stream()
    data = []
    for d in hist_docs:
        item = d.to_dict()
        item['id'] = d.id
        pf = str(item.get('pf_number', ''))
        
        # Live Sync Designation & Unit from Employees Collection
        if pf in emp_map:
            item['designation'] = emp_map[pf].get('Designation', item.get('designation', ''))
            item['unit'] = emp_map[pf].get('Unit', item.get('unit', ''))
            if not item.get('station'):
                item['station'] = emp_map[pf].get('STATION', 'N/A')
        
        # Dates for Display
        item['allot_disp'] = item['allotment_date'].strftime('%d-%m-%Y') if item.get('allotment_date') and hasattr(item['allotment_date'], 'strftime') else "N/A"
        item['vacat_disp'] = item['vacation_date'].strftime('%d-%m-%Y') if item.get('vacation_date') and hasattr(item['vacation_date'], 'strftime') else ("Occupied" if item.get('is_current') else "Vacant")
        data.append(item)
        
    return pd.DataFrame(data), emp_map

# --- 4. Template Generator ---
def fill_template(temp_path, data, date_str):
    if not os.path.exists(temp_path):
        st.error(f"Template file nahi mili: {temp_path}")
        return None
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

# --- 5. Main UI ---
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
    
    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Dashboard & History"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Allotment Form")
        if not df_hist.empty and emp_map:
            emp_list = sorted([f"{v['Employee Name']} ({k})" for k, v in emp_map.items()])
            selected_emp = st.selectbox("Staff Select Karein", emp_list)
            
            pf_key = selected_emp.split('(')[-1].strip(')')
            
            # Check for Double Allotment
            already_has = df_hist[(df_hist['pf_number'] == pf_key) & (df_hist['is_current'] == True)]
            
            if not already_has.empty:
                st.error(f"❌ Is karmchari ko pehle se Quarter No. {already_has.iloc[0]['quarter_number']} allot hai.")
            else:
                staff_station = str(emp_map[pf_key].get('STATION', '')).strip()
                st.info(f"📍 Station: **{staff_station}**")

                # Filter Vacant Quarters ONLY at this Station
                station_qs = df_hist[df_hist['station'].str.strip() == staff_station]
                occupied_list = station_qs[station_qs['is_current'] == True]['quarter_number'].tolist()
                available_qs = [q for q in station_qs['quarter_number'].unique() if q not in occupied_list]

                if available_qs:
                    sel_q = st.selectbox(f"Vacant Quarters ({staff_station})", available_qs)
                    a_date = st.date_input("Allotment Date", value=datetime.now())

                    if st.button("Generate Allotment"):
                        emp_info = emp_map[pf_key]
                        entry = {
                            "employee_name": emp_info['Employee Name'],
                            "designation": emp_info['Designation'],
                            "unit": emp_info.get('Unit', ''),
                            "hrms_id": emp_info.get('HRMS ID', ''),
                            "pf_number": pf_key,
                            "quarter_number": sel_q,
                            "station": staff_station,
                            "allotment_date": datetime.combine(a_date, datetime.min.time()),
                            "is_current": True
                        }
                        db.collection("quarter_history").add(entry)
                        st.success(f"✅ {sel_q} allot kar diya gaya.")
                        
                        doc = fill_template(ALLOT_TEMP, entry, a_date.strftime("%d/%m/%Y"))
                        if doc:
                            buf = io.BytesIO(); doc.save(buf)
                            st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{sel_q}.docx")
                else:
                    st.warning(f"⚠️ {staff_station} par koi vacant quarter nahi hai.")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation")
        occupied_df = df_hist[df_hist['is_current'] == True]
        if not occupied_df.empty:
            v_list = occupied_df.apply(lambda r: f"{r['station']} | {r['quarter_number']} - {r['employee_name']}", axis=1).tolist()
            sel_v = st.selectbox("Select Quarter to Vacate", v_list)
            v_date = st.date_input("Vacation Date")
            
            if st.button("Process & Generate Vacation Letter"):
                idx = v_list.index(sel_v)
                q_row = occupied_df.iloc[idx]
                
                # Update DB
                db.collection("quarter_history").document(q_row['id']).update({
                    "is_current": False,
                    "vacation_date": datetime.combine(v_date, datetime.min.time())
                })
                
                # Generate Letter
                doc = fill_template(VACATE_TEMP, q_row.to_dict(), v_date.strftime("%d/%m/%Y"))
                if doc:
                    buf = io.BytesIO(); doc.save(buf)
                    st.success("✅ Quarter Khali ho gaya hai.")
                    st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{q_row['quarter_number']}.docx")
        else:
            st.info("Koi occupied quarter nahi hai.")

    # --- TAB 3: DASHBOARD & HISTORY ---
    with tab3:
        st.header("📊 Station-wise Analytics")
        if not df_hist.empty:
            # Correct Counting: Combined Station + Quarter
            df_hist['unique_key'] = df_hist['station'].str.strip() + "_" + df_hist['quarter_number'].str.strip()
            
            total_unique = df_hist['unique_key'].nunique()
            total_occ = df_hist[df_hist['is_current'] == True]['unique_key'].nunique()
            total_vac = total_unique - total_occ
            
            c1, c2, c3 = st.columns(3)
            c1.metric("🏠 Total Unique Quarters", total_unique)
            c2.metric("🔴 Occupied", total_occ)
            c3.metric("🟢 Vacant", total_vac)
            
            

            st.markdown("---")
            st.subheader("🔍 Check Quarter History")
            h1, h2 = st.columns(2)
            with h1:
                search_stn = st.selectbox("Station Chunein", sorted(df_hist['station'].unique()))
            with h2:
                search_q = st.selectbox("Quarter Number", sorted(df_hist[df_hist['station'] == search_stn]['quarter_number'].unique()))
            
            q_hist = df_hist[(df_hist['station'] == search_stn) & (df_hist['quarter_number'] == search_q)].sort_values(by='allotment_date', ascending=False)
            
            if not q_hist.empty:
                st.table(q_hist[['employee_name', 'pf_number', 'allot_disp', 'vacat_disp']])

            st.markdown("---")
            st.subheader("Full Master Report")
            st.dataframe(df_hist[['station', 'quarter_number', 'employee_name', 'allot_disp', 'status_disp']], use_container_width=True)

if __name__ == "__main__":
    main()
