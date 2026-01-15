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
            # First check for secrets, then local file
            if "firebase_config" in st.secrets:
                cred = credentials.Certificate(dict(st.secrets["firebase_config"]))
            else:
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

# --- 1. Admin Logic (Excel Import) ---
def clean_and_import_excel(excel_file):
    db = init_db()
    # 1. Purana data delete karein
    docs = db.collection("quarter_history").limit(500).stream()
    for doc in docs:
        doc.reference.delete()
    
    # 2. Excel read karein (using openpyxl)
    df = pd.read_excel(excel_file)
    
    count = 0
    for _, row in df.iterrows():
        # Remark 'Occupation' hai toh is_current True hoga
        remark = str(row.get('Remark', '')).strip().lower()
        is_occ = True if remark == 'occupation' else False
        
        data = {
            "station": str(row.get('Station', '')),
            "quarter_number": str(row.get('Quarter Number', '')),
            "pf_number": str(row.get('PF NUMBER', '')) if pd.notna(row.get('PF NUMBER')) else "",
            "hrms_id": str(row.get('HRMS ID', '')) if pd.notna(row.get('HRMS ID')) else "",
            "employee_name": str(row.get('Employee Name', '')) if pd.notna(row.get('Employee Name')) else "",
            "allotment_date": row['Occupation Date'].to_pydatetime() if pd.notna(row.get('Occupation Date')) and hasattr(row['Occupation Date'], 'to_pydatetime') else None,
            "vacation_date": row['Vacant Date'].to_pydatetime() if pd.notna(row.get('Vacant Date')) and hasattr(row['Vacant Date'], 'to_pydatetime') else None,
            "is_current": is_occ,
            "designation": "", # Employees collection se fetch hoga naye allotment par
            "unit": ""
        }
        db.collection("quarter_history").add(data)
        count += 1
    return count

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
        # Formatting for UI display
        item['allot_disp'] = item['allotment_date'].strftime('%d-%m-%Y') if item.get('allotment_date') and hasattr(item['allotment_date'], 'strftime') else "N/A"
        
        if item.get('is_current'):
            item['vacat_disp'] = "🔴 Occupied"
        else:
            item['vacat_disp'] = item['vacation_date'].strftime('%d-%m-%Y') if item.get('vacation_date') and hasattr(item['vacation_date'], 'strftime') else "🟢 Vacant"
        
        data.append(item)
    return pd.DataFrame(data)

# --- 3. Docx Logic ---
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
    # Replace in Paragraphs and Tables
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in mapping.items():
            if f"{{{{{k}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{k}}}}}", v)
    return doc

# --- 4. Main UI ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter Management")
    db = init_db()
    
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    # --- Sidebar Admin Tools (EXCEL IMPORT) ---
    st.sidebar.header("🛠️ Admin Data Tools")
    if st.sidebar.checkbox("Enable Excel Import/Clean"):
        excel_file = st.sidebar.file_uploader("Upload 'Quarter Register.xlsx'", type=["xlsx"])
        if st.sidebar.button("🔥 Clean & Import from Excel"):
            if excel_file:
                with st.spinner("Processing Excel Data..."):
                    count = clean_and_import_excel(excel_file)
                    st.sidebar.success(f"Cleaned & Imported {count} records!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.sidebar.error("Kripya .xlsx file chunein")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Master Report"])

    df_emp = get_employees()
    df_hist = get_history()

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("Quarter Allotment Form")
        if not df_emp.empty and not df_hist.empty:
            # Employee selection
            emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            sel_emp_str = st.selectbox("Staff Chunein", emp_options)
            
            # Khali Quarter Filter Logic
            occ_q_list = df_hist[df_hist['is_current'] == True]['quarter_number'].tolist()
            all_q_list = sorted(df_hist['quarter_number'].unique().tolist())
            vacant_q_list = [q for q in all_q_list if q not in occ_q_list]
            
            col1, col2 = st.columns(2)
            with col1:
                sel_q = st.selectbox("Khali Quarter (Vacant Only)", vacant_q_list)
            with col2:
                # Station selection (History se unique stations)
                stn_list = sorted(df_hist['station'].unique().tolist())
                sel_stn = st.selectbox("Station Name", stn_list)
            
            a_date = st.date_input("Allotment Effective Date")

            if st.button("Allot & Generate Allotment Letter"):
                h_id = sel_emp_str.split('(')[-1].strip(')')
                emp_data = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                new_allot = {
                    "employee_name": str(emp_data.get('Employee Name', '')),
                    "designation": str(emp_data.get('Designation', '')), # Employee collection se
                    "unit": str(emp_data.get('Unit', 'SSE/PW/SGAM')),   # Employee collection se
                    "hrms_id": h_id,
                    "pf_number": str(emp_data.get('PF Number', '')),
                    "quarter_number": sel_q,
                    "station": sel_stn,
                    "allotment_date": datetime.combine(a_date, datetime.min.time()),
                    "is_current": True,
                    "vacation_date": None
                }
                db.collection("quarter_history").add(new_allot)
                st.cache_data.clear()
                
                doc = fill_template(ALLOT_TEMP, new_allot, a_date.strftime("%d/%m/%Y"))
                buf = io.BytesIO(); doc.save(buf)
                st.success(f"Allotted {sel_q} to {new_allot['employee_name']}")
                st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{sel_q}.docx")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation Form")
        if not df_hist.empty:
            occupied = df_hist[df_hist['is_current'] == True]
            if not occupied.empty:
                v_options = occupied.apply(lambda r: f"{r['quarter_number']} - {r['employee_name']}", axis=1).tolist()
                sel_v_str = st.selectbox("Select Quarter to Vacate", v_options)
                v_date = st.date_input("Vacation Effective Date")
                
                if st.button("Vacate & Generate Vacation Letter"):
                    v_row = occupied.iloc[v_options.index(sel_v_str)]
                    db.collection("quarter_history").document(v_row['id']).update({
                        "is_current": False,
                        "vacation_date": datetime.combine(v_date, datetime.min.time())
                    })
                    st.cache_data.clear()
                    
                    doc = fill_template(VACATE_TEMP, v_row.to_dict(), v_date.strftime("%d/%m/%Y"))
                    buf = io.BytesIO(); doc.save(buf)
                    st.success(f"Quarter {v_row['quarter_number']} marked as Vacant")
                    st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{v_row['quarter_number']}.docx")
            else:
                st.info("Abhi koi quarter occupied nahi hai.")

    # --- TAB 3: MASTER REPORT ---
    with tab3:
        st.header("📊 Full Quarter History & Report")
        if not df_hist.empty:
            # Stats Dashboard
            c1, c2 = st.columns(2)
            c1.metric("Total Quarters in Registry", len(df_hist['quarter_number'].unique()))
            c2.metric("Currently Occupied", len(df_hist[df_hist['is_current'] == True]))

            # Table for report
            report_df = df_hist[['quarter_number', 'employee_name', 'station', 'allot_disp', 'vacat_disp', 'is_current']].copy()
            report_df.columns = ['Quarter No', 'Employee Name', 'Station', 'Allotted On', 'Vacated On', 'Current Status']
            
            st.dataframe(report_df, use_container_width=True)
            
            # CSV Download with Station
            csv = report_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Download Master Report (CSV)", csv, "Quarter_Master_Report.csv", "text/csv")
        else:
            st.warning("Database khali hai. Sidebar se Excel import karein.")

if __name__ == "__main__":
    main()
