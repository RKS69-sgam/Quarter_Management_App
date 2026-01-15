import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Firebase Setup ---
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

# --- 1. Admin Logic (Clean & Import) ---
def clean_and_import(df):
    db = init_db()
    # Purana data delete karein
    docs = db.collection("quarter_history").limit(500).stream()
    for doc in docs:
        doc.reference.delete()
    
    # Naya data upload karein
    for _, row in df.iterrows():
        is_occ = True if str(row.get('Remark', '')).lower() == 'occupation' else False
        data = {
            "station": str(row.get('Station', '')),
            "quarter_number": str(row.get('Quarter Number', '')),
            "pf_number": str(row.get('PF NUMBER', '')),
            "hrms_id": str(row.get('HRMS ID', '')),
            "employee_name": str(row.get('Employee Name', '')),
            "allotment_date": pd.to_datetime(row['Occupation Date']).to_pydatetime() if pd.notna(row.get('Occupation Date')) else None,
            "vacation_date": pd.to_datetime(row['Vacant Date']).to_pydatetime() if pd.notna(row.get('Vacant Date')) else None,
            "is_current": is_occ,
            "designation": "", # Allotment ke samay employee collection se aayega
            "unit": ""         # Allotment ke samay employee collection se aayega
        }
        db.collection("quarter_history").add(data)
    return len(df)

# --- 2. Data Fetching ---
@st.cache_data(ttl=60)
def get_employees():
    db = init_db()
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_history():
    db = init_db()
    docs = db.collection("quarter_history").stream()
    data = []
    for d in docs:
        item = d.to_dict()
        item['id'] = d.id
        item['allot_disp'] = item['allotment_date'].strftime('%d-%m-%Y') if item.get('allotment_date') else "N/A"
        item['vacat_disp'] = item['vacation_date'].strftime('%d-%m-%Y') if item.get('vacation_date') else ("Occupied" if item.get('is_current') else "Vacant")
        data.append(item)
    return pd.DataFrame(data)

# --- 3. Docx Template Logic ---
def fill_docx(temp_path, data, date_str):
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

# --- 4. Main App Interface ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    # --- Sidebar Admin Tools ---
    st.sidebar.header("⚙️ Admin Settings")
    if st.sidebar.checkbox("Data Import/Clean"):
        csv_file = st.sidebar.file_uploader("Upload Quarter CSV", type="csv")
        if st.sidebar.button("🔥 Delete & Import"):
            if csv_file:
                df_new = pd.read_csv(csv_file)
                count = clean_and_import(df_new)
                st.sidebar.success(f"Successfully Imported {count} records!")
                st.cache_data.clear()
                st.rerun()
            else: st.sidebar.error("Upload a file first")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Master Report"])

    df_emp = get_employees()
    df_hist = get_history()

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Allotment Form")
        if not df_emp.empty and not df_hist.empty:
            emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            sel_emp = st.selectbox("Select Employee", emp_list)
            
            # Khali Quarter Filter
            occ_q = df_hist[df_hist['is_current'] == True]['quarter_number'].tolist()
            vacant_q = [q for q in df_hist['quarter_number'].unique() if q not in occ_q]
            
            c1, c2 = st.columns(2)
            with c1: q_no = st.selectbox("Available Quarters", vacant_q)
            with c2: stn = st.selectbox("Station", df_hist['station'].unique())
            
            a_date = st.date_input("Allotment Date", value=datetime.now())

            if st.button("Allot & Generate Letter"):
                h_id = sel_emp.split('(')[-1].strip(')')
                emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                allot_data = {
                    "employee_name": emp_row.get('Employee Name'),
                    "designation": emp_row.get('Designation'),
                    "unit": emp_row.get('Unit', 'SSE/PW/SGAM'),
                    "hrms_id": h_id,
                    "pf_number": emp_row.get('PF Number', ''),
                    "quarter_number": q_no,
                    "station": stn,
                    "allotment_date": datetime.combine(a_date, datetime.min.time()),
                    "is_current": True
                }
                db.collection("quarter_history").add(allot_data)
                st.cache_data.clear()
                
                doc = fill_doc(ALLOT_TEMP, allot_data, a_date.strftime("%d/%m/%Y"))
                buf = io.BytesIO(); doc.save(buf)
                st.success(f"Allotted {q_no}")
                st.download_button("📥 Download Letter", buf.getvalue(), f"Allotment_{q_no}.docx")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Vacation Form")
        if not df_hist.empty:
            occ_only = df_hist[df_hist['is_current'] == True]
            if not occ_only.empty:
                v_sel = st.selectbox("Select Quarter", occ_only.apply(lambda r: f"{r['quarter_number']} - {r['employee_name']}", axis=1))
                v_date = st.date_input("Vacation Date")
                if st.button("Process Vacation"):
                    q_row = occ_only.iloc[0] # logic based on selection
                    db.collection("quarter_history").document(q_row['id']).update({
                        "is_current": False, "vacation_date": datetime.combine(v_date, datetime.min.time())
                    })
                    st.cache_data.clear()
                    doc = fill_doc(VACATE_TEMP, q_row.to_dict(), v_date.strftime("%d/%m/%Y"))
                    buf = io.BytesIO(); doc.save(buf)
                    st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{q_row['quarter_number']}.docx")

    # --- TAB 3: REPORT ---
    with tab3:
        st.header("📋 Master Quarter Database")
        if not df_hist.empty:
            st.metric("Total Occupied", len(df_hist[df_hist['is_current'] == True]))
            
            # Report with Station
            disp = df_hist[['quarter_number', 'employee_name', 'station', 'allot_disp', 'vacat_disp', 'is_current']].copy()
            disp.columns = ['Quarter', 'Name', 'Station', 'Allotted', 'Vacated', 'Current Status']
            
            st.dataframe(disp, use_container_width=True)
            
            # CSV with Station
            csv = disp.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Export CSV (Full History)", csv, "Quarter_Report.csv", "text/csv")

if __name__ == "__main__":
    main()
