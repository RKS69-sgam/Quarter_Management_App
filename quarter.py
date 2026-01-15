import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Configuration & Login ---
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

# --- 1. Firebase Initialization ---
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
                json_path = 'sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json'
                cred = credentials.Certificate(json_path)
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

# --- 2. Data Fetching with Caching (To save Quota) ---
@st.cache_data(ttl=300)
def get_employees_cached():
    db = init_db()
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_full_quarter_history():
    db = init_db()
    docs = db.collection("quarter_history").order_by("Allotment_Date", direction=firestore.Query.DESCENDING).stream()
    data = []
    for d in docs:
        item = d.to_dict()
        item['id'] = d.id
        # Formatting dates for DataFrame display
        if item.get('Allotment_Date'):
            item['Allotment_Date_Disp'] = item['Allotment_Date'].strftime('%d-%m-%Y')
        if item.get('Vacation_Date'):
            item['Vacation_Date_Disp'] = item['Vacation_Date'].strftime('%d-%m-%Y')
        else:
            item['Vacation_Date_Disp'] = "Occupied"
        data.append(item)
    return pd.DataFrame(data)

# --- 3. Template Filling Logic ---
def fill_quarter_template(template_path, data_map):
    doc = Document(template_path)
    # Paragraphs replacement
    for p in doc.paragraphs:
        for key, val in data_map.items():
            placeholder = "{{" + key + "}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(val))
    # Tables replacement
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(val))
    return doc

# --- 4. Main App Interface ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    if not check_login(): return

    if st.sidebar.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.rerun()

    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Report & Full History"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("Quarter Allotment Form")
        df_emp = get_employees_cached()
        if not df_emp.empty:
            emp_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['HRMS ID']})", axis=1).tolist()
            selected_emp_str = st.selectbox("Select Employee", emp_options)
            q_number = st.text_input("Quarter Number", placeholder="e.g. T-15/B")
            station = st.text_input("Station", value="सरईग्राम")
            allot_date = st.date_input("Allotment Effective Date")

            if st.button("Allot & Generate Letter"):
                h_id = selected_emp_str.split('(')[-1].strip(')')
                emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                data_map = {
                    "EMPLOYEE_NAME": str(emp_row['Employee Name']),
                    "DESIGNATION": str(emp_row['Designation']),
                    "UNIT": str(emp_row.get('Unit', 'SSE/P.Way/SGAM')),
                    "PF_Number": str(emp_row.get('PF Number', '')),
                    "HRMS_ID": h_id,
                    "QUARTER_NUMBER": q_number,
                    "STATION": station,
                    "DATE": allot_date.strftime("%d/%m/%Y")
                }

                # Save to quarter_history
                db.collection("quarter_history").add({
                    **data_map,
                    "Allotment_Date": datetime.combine(allot_date, datetime.min.time()),
                    "Status": "Occupied",
                    "Vacation_Date": None
                })
                
                doc = fill_quarter_template(ALLOT_TEMP, data_map)
                buf = io.BytesIO()
                doc.save(buf)
                st.success(f"Quarter {q_number} allotted to {data_map['EMPLOYEE_NAME']}")
                st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{h_id}.docx")
        else:
            st.error("Employee database is empty.")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation Form")
        df_hist = get_full_quarter_history()
        occupied_q = df_hist[df_hist['Status'] == "Occupied"] if not df_hist.empty else pd.DataFrame()

        if not occupied_q.empty:
            q_options = occupied_q.apply(lambda r: f"{r['QUARTER_NUMBER']} - {r['EMPLOYEE_NAME']}", axis=1).tolist()
            selected_vac = st.selectbox("Select Quarter to Vacate", q_options)
            vac_date = st.date_input("Vacation Effective Date")

            if st.button("Vacate & Generate Letter"):
                idx = q_options.index(selected_vac)
                orig_data = occupied_q.iloc[idx]
                doc_id = orig_data['id']
                
                vac_data_map = orig_data.to_dict()
                vac_data_map['DATE'] = vac_date.strftime("%d/%m/%Y") # Update placeholder date

                # Update Firestore
                db.collection("quarter_history").document(doc_id).update({
                    "Status": "Vacated",
                    "Vacation_Date": datetime.combine(vac_date, datetime.min.time())
                })
                
                doc = fill_quarter_template(VACATE_TEMP, vac_data_map)
                buf = io.BytesIO()
                doc.save(buf)
                st.success(f"Quarter {vac_data_map['QUARTER_NUMBER']} marked as Vacated.")
                st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{vac_data_map['HRMS_ID']}.docx")
        else:
            st.info("No occupied quarters found to vacate.")

    # --- TAB 3: REPORT ---
    with tab3:
        st.header("📋 Full Quarter History Database")
        df_report = get_full_quarter_history()
        
        if not df_report.empty:
            # Stats
            total_q = len(df_report['QUARTER_NUMBER'].unique())
            occupied_count = len(df_report[df_report['Status'] == "Occupied"])
            
            c1, c2 = st.columns(2)
            c1.metric("Unique Quarters Handled", total_q)
            c2.metric("Currently Occupied", occupied_count)

            # Search / Filter
            search = st.text_input("🔍 Search by Quarter No, Name or HRMS ID")
            if search:
                df_report = df_report[
                    df_report['QUARTER_NUMBER'].str.contains(search, case=False) | 
                    df_report['EMPLOYEE_NAME'].str.contains(search, case=False) |
                    df_report['HRMS_ID'].str.contains(search, case=False)
                ]

            # Display Table
            display_df = df_report[[
                'QUARTER_NUMBER', 'EMPLOYEE_NAME', 'DESIGNATION', 
                'HRMS_ID', 'Allotment_Date_Disp', 'Vacation_Date_Disp', 'Status'
            ]].rename(columns={
                'Allotment_Date_Disp': 'Allotted On',
                'Vacation_Date_Disp': 'Vacated On'
            })

            st.dataframe(display_df, use_container_width=True)

            # CSV Download
            csv = df_report.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Export Full History to CSV", csv, "quarter_master_report.csv", "text/csv")
        else:
            st.warning("No records found in quarter_history collection.")

if __name__ == "__main__":
    main()
