
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
                else: st.error("Galat Credentials")
        return False
    return True

@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        if "firebase_config" in st.secrets:
            cred = credentials.Certificate(dict(st.secrets["firebase_config"]))
        else:
            cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

# --- 1. Helper Functions ---
@st.cache_data(ttl=300)
def get_employees():
    db = init_db()
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_quarter_history():
    db = init_db()
    docs = db.collection("quarter_history").order_by("Allotment_Date", direction=firestore.Query.DESCENDING).stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def fill_template(template_path, data):
    doc = Document(template_path)
    # Paragraph replacement
    for p in doc.paragraphs:
        for key, value in data.items():
            placeholder = "{{" + key + "}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
    # Table replacement
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in data.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))
    return doc

# --- 2. Main UI ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter Management")
    if not check_login(): return

    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    tab1, tab2, tab3 = st.tabs(["🏠 Quarter Allotment", "🗝️ Quarter Vacation", "📊 Quarter History & Status"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Quarter Allotment")
        df_emp = get_employees()
        if not df_emp.empty:
            emp_list = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['HRMS ID']})", axis=1).tolist()
            selected_emp = st.selectbox("Employee Select Karein", emp_list)
            q_no = st.text_input("Quarter Number (e.g. T-10/A)")
            stn = st.text_input("Station", value="सरईग्राम")
            allot_date = st.date_input("Allotment Date", value=datetime.now())

            if st.button("Allot Quarter & Generate Letter"):
                h_id = selected_emp.split('(')[-1].strip(')')
                emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                allot_data = {
                    "EMPLOYEE_NAME": str(emp_row['Employee Name']),
                    "DESIGNATION": str(emp_row['Designation']),
                    "PF_Number": str(emp_row.get('PF Number', h_id)),
                    "HRMS_ID": h_id,
                    "QUARTER_NUMBER": q_no,
                    "STATION": stn,
                    "DATE": allot_date.strftime("%d/%m/%Y"),
                    "UNIT": "SSE/P.Way/SGAM"
                }

                # Save to History
                db.collection("quarter_history").add({
                    **allot_data,
                    "Allotment_Date": datetime.combine(allot_date, datetime.min.time()),
                    "Status": "Occupied",
                    "Vacation_Date": None
                })
                
                doc = fill_template(ALLOT_TEMP, allot_data)
                buf = io.BytesIO()
                doc.save(buf)
                st.success(f"Quarter {q_no} successfully allotted to {allot_data['EMPLOYEE_NAME']}")
                st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{q_no}.docx")
        
    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation (Khali Karna)")
        df_hist = get_quarter_history()
        # Sirf wahi quarters jo occupied hain
        occupied = df_hist[df_hist['Status'] == "Occupied"] if not df_hist.empty else pd.DataFrame()
        
        if not occupied.empty:
            q_list = occupied.apply(lambda r: f"{r['QUARTER_NUMBER']} - {r['EMPLOYEE_NAME']}", axis=1).tolist()
            selected_q = st.selectbox("Vacate karne ke liye Quarter chunein", q_list)
            vac_date = st.date_input("Vacation Date", value=datetime.now())

            if st.button("Vacate Quarter & Generate Letter"):
                q_id = occupied.iloc[q_list.index(selected_q)]['id']
                q_data = occupied.iloc[q_list.index(selected_q)].to_dict()
                
                vac_info = {
                    **q_data,
                    "DATE": vac_date.strftime("%d/%m/%Y")
                }

                # Update Status in DB
                db.collection("quarter_history").document(q_id).update({
                    "Status": "Vacated",
                    "Vacation_Date": datetime.combine(vac_date, datetime.min.time())
                })
                
                doc = fill_template(VACATE_TEMP, vac_info)
                buf = io.BytesIO()
                doc.save(buf)
                st.success(f"Quarter {q_data['QUARTER_NUMBER']} has been vacated.")
                st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{q_data['QUARTER_NUMBER']}.docx")
        else:
            st.info("Abhi koi quarter occupied nahi hai.")

    # --- TAB 3: HISTORY & STATUS ---
    with tab3:
        st.header("Quarter Master Report")
        df_full = get_quarter_history()
        if not df_full.empty:
            # Current Status Summary
            st.subheader("Current Occupancy Status")
            st.dataframe(df_full[['QUARTER_NUMBER', 'EMPLOYEE_NAME', 'Status', 'Allotment_Date', 'Vacation_Date']], use_container_width=True)
            
            # Individual History search
            search_q = st.text_input("Quarter Number se search karein (History dekhne ke liye)")
            if search_q:
                hist = df_full[df_full['QUARTER_NUMBER'].str.contains(search_q, case=False)]
                st.write(f"History for {search_q}:")
                st.table(hist[['EMPLOYEE_NAME', 'Allotment_Date', 'Vacation_Date', 'Status']])
        else:
            st.warning("Koi history record nahi mila.")

if __name__ == "__main__":
    main()
