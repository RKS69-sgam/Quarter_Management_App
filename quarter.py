import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Login Configuration ---
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

# --- 1. Firebase Setup ---
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

# --- 2. Data Fetching Logic (Optimized for Quota & Errors) ---
@st.cache_data(ttl=300)
def get_employees_cached():
    db = init_db()
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_full_quarter_history():
    """quarter_history collection se poora data fetch karna (Headers & Data)"""
    db = init_db()
    try:
        # Bina order_by ke stream kar rahe hain taaki missing fields error na dein
        docs = db.collection("quarter_history").stream()
        data = []
        for d in docs:
            item = d.to_dict()
            item['id'] = d.id
            
            # Allotment Date Handling
            allot_date = item.get('Allotment_Date')
            if allot_date:
                item['Allotment_Date_Disp'] = allot_date.strftime('%d-%m-%Y') if hasattr(allot_date, 'strftime') else str(allot_date)
            else:
                item['Allotment_Date_Disp'] = "N/A"
            
            # Vacation Date Handling
            vac_date = item.get('Vacation_Date')
            if vac_date:
                item['Vacation_Date_Disp'] = vac_date.strftime('%d-%m-%Y') if hasattr(vac_date, 'strftime') else str(vac_date)
            else:
                item['Vacation_Date_Disp'] = "Occupied" if item.get('Status') == "Occupied" else "N/A"
                
            data.append(item)
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Data Fetch Error: {e}")
        return pd.DataFrame()

# --- 3. Template Filling Logic ---
def fill_template(template_path, data_map):
    doc = Document(template_path)
    # Paragraphs aur Tables mein placeholders {{KEY}} replace karna
    for p in doc.paragraphs:
        for key, val in data_map.items():
            if f"{{{{{key}}}}}" in p.text:
                p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in data_map.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(val))
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

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Full History Report"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("Quarter Allotment Form")
        df_emp = get_employees_cached()
        if not df_emp.empty:
            emp_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['HRMS ID']})", axis=1).tolist()
            selected_emp = st.selectbox("Select Staff", emp_options)
            q_no = st.text_input("Quarter No.")
            stn = st.text_input("Station", value="सरईग्राम")
            allot_date = st.date_input("Allotment Date")

            if st.button("Allot & Generate Letter"):
                h_id = selected_emp.split('(')[-1].strip(')')
                row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                data = {
                    "EMPLOYEE_NAME": str(row['Employee Name']),
                    "DESIGNATION": str(row['Designation']),
                    "PF_Number": str(row.get('PF Number', '')),
                    "HRMS_ID": h_id,
                    "UNIT": str(row.get('Unit', 'SSE/P.Way/SGAM')),
                    "QUARTER_NUMBER": q_no,
                    "STATION": stn,
                    "DATE": allot_date.strftime("%d/%m/%Y")
                }
                
                db.collection("quarter_history").add({
                    **data,
                    "Allotment_Date": datetime.combine(allot_date, datetime.min.time()),
                    "Status": "Occupied",
                    "Vacation_Date": None
                })
                
                doc = fill_template(ALLOT_TEMP, data)
                buf = io.BytesIO()
                doc.save(buf)
                st.success("Allotment Done!")
                st.download_button("📥 Download Letter", buf.getvalue(), f"Allotment_{q_no}.docx")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation Form")
        df_h = get_full_quarter_history()
        occupied = df_h[df_h['Status'] == "Occupied"] if not df_h.empty else pd.DataFrame()
        
        if not occupied.empty:
            q_list = occupied.apply(lambda r: f"{r['QUARTER_NUMBER']} - {r['EMPLOYEE_NAME']}", axis=1).tolist()
            sel_q = st.selectbox("Select Quarter to Vacate", q_list)
            v_date = st.date_input("Vacation Date")

            if st.button("Vacate & Generate Letter"):
                q_row = occupied.iloc[q_list.index(sel_q)]
                v_data = q_row.to_dict()
                v_data['DATE'] = v_date.strftime("%d/%m/%Y")
                
                db.collection("quarter_history").document(q_row['id']).update({
                    "Status": "Vacated",
                    "Vacation_Date": datetime.combine(v_date, datetime.min.time())
                })
                
                doc = fill_template(VACATE_TEMP, v_data)
                buf = io.BytesIO()
                doc.save(buf)
                st.success("Vacation Processed!")
                st.download_button("📥 Download Letter", buf.getvalue(), f"Vacation_{v_data['QUARTER_NUMBER']}.docx")

    # --- TAB 3: FULL REPORT ---
    with tab3:
        st.header("📊 Quarter History Master Database")
        df_full = get_full_quarter_history()
        
        if not df_full.empty:
            # Header Formatting
            disp_df = df_full[[
                'QUARTER_NUMBER', 'EMPLOYEE_NAME', 'DESIGNATION', 
                'HRMS_ID', 'Allotment_Date_Disp', 'Vacation_Date_Disp', 'Status'
            ]].rename(columns={
                'QUARTER_NUMBER': 'Quarter No.', 'EMPLOYEE_NAME': 'Staff Name',
                'Allotment_Date_Disp': 'Allotted On', 'Vacation_Date_Disp': 'Vacated On'
            })

            # Universal Search
            search = st.text_input("🔍 Search Quarter, Name or ID")
            if search:
                disp_df = disp_df[disp_df.apply(lambda r: r.astype(str).str.contains(search, case=False).any(), axis=1)]

            # Table Display
            st.dataframe(disp_df, use_container_width=True, column_config={
                "Status": st.column_config.BadgeColumn(map={"Occupied": "🔴 Occupied", "Vacated": "🟢 Vacated"})
            })
            
            # Export
            st.download_button("📥 Export CSV", disp_df.to_csv(index=False).encode('utf-8-sig'), "Quarter_Report.csv", "text/csv")
        else:
            st.info("No records found in quarter_history collection.")

if __name__ == "__main__":
    main()
