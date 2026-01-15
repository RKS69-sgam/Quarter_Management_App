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

# --- 1. Data Fetching Logic (Fixed) ---
@st.cache_data(ttl=300)
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
            # Date format handle karna
            for d_field in ['Allotment_Date', 'Vacation_Date']:
                val = item.get(d_field)
                if val:
                    item[f'{d_field}_Disp'] = val.strftime('%d-%m-%Y') if hasattr(val, 'strftime') else str(val)
                else:
                    item[f'{d_field}_Disp'] = "N/A"
            data.append(item)
        
        df = pd.DataFrame(data)
        
        # KEYERROR PREVENTER: Agar Status column nahi hai toh khali column bana do
        if not df.empty and 'Status' not in df.columns:
            df['Status'] = "Unknown"
        return df
    except Exception as e:
        st.error(f"Data Fetch Error: {e}")
        return pd.DataFrame()

# --- 2. Template Logic ---
def fill_template(template_path, data_map):
    doc = Document(template_path)
    for p in doc.paragraphs:
        for k, v in data_map.items():
            if f"{{{{{k}}}}}" in p.text: p.text = p.text.replace(f"{{{{{k}}}}}", str(v))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in data_map.items():
                        if f"{{{{{k}}}}}" in p.text: p.text = p.text.replace(f"{{{{{k}}}}}", str(v))
    return doc

# --- 3. Main UI ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter Management")
    if not check_login(): return

    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Report Tab"])

    # --- TAB 1: ALLOTMENT ---
    with tab1:
        st.header("New Quarter Allotment")
        df_emp = get_employees_cached()
        if not df_emp.empty:
            emp_list = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['HRMS ID']})", axis=1).tolist()
            selected_emp = st.selectbox("Select Employee", emp_list)
            q_no = st.text_input("Quarter No.")
            stn = st.text_input("Station", value="सरईग्राम")
            dt = st.date_input("Allotment Date")

            if st.button("Allot Quarter"):
                h_id = selected_emp.split('(')[-1].strip(')')
                row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                data = {
                    "EMPLOYEE_NAME": str(row['Employee Name']),
                    "DESIGNATION": str(row['Designation']),
                    "PF_Number": str(row.get('PF Number', '')),
                    "HRMS_ID": h_id,
                    "UNIT": str(row.get('Unit', 'SSE/P.Way/SGAM')),
                    "QUARTER_NUMBER": q_no, "STATION": stn, "DATE": dt.strftime("%d/%m/%Y")
                }
                db.collection("quarter_history").add({
                    **data, "Allotment_Date": datetime.combine(dt, datetime.min.time()),
                    "Status": "Occupied", "Vacation_Date": None
                })
                st.success(f"Allotted {q_no} to {data['EMPLOYEE_NAME']}")
                doc = fill_template(ALLOT_TEMP, data)
                buf = io.BytesIO(); doc.save(buf)
                st.download_button("📥 Download", buf.getvalue(), f"Allotment_{q_no}.docx")

    # --- TAB 2: VACATION ---
    with tab2:
        st.header("Quarter Vacation")
        df_h = get_full_quarter_history()
        # Safe filtering to avoid KeyError
        if not df_h.empty and 'Status' in df_h.columns:
            occupied = df_h[df_h['Status'] == "Occupied"]
            if not occupied.empty:
                q_list = occupied.apply(lambda r: f"{r['QUARTER_NUMBER']} - {r['EMPLOYEE_NAME']}", axis=1).tolist()
                sel = st.selectbox("Select Quarter to Vacate", q_list)
                v_dt = st.date_input("Vacation Date")
                if st.button("Process Vacation"):
                    q_row = occupied.iloc[q_list.index(sel)]
                    v_data = q_row.to_dict()
                    v_data['DATE'] = v_dt.strftime("%d/%m/%Y")
                    db.collection("quarter_history").document(q_row['id']).update({
                        "Status": "Vacated", "Vacation_Date": datetime.combine(v_dt, datetime.min.time())
                    })
                    st.success(f"Vacated {v_data['QUARTER_NUMBER']}")
                    doc = fill_template(VACATE_TEMP, v_data)
                    buf = io.BytesIO(); doc.save(buf)
                    st.download_button("📥 Download", buf.getvalue(), f"Vacation_{v_data['QUARTER_NUMBER']}.docx")
            else: st.info("Abhi koi quarter occupied nahi hai.")
        else: st.warning("Database mein 'Status' field nahi mila.")

    # --- TAB 3: REPORT ---
    with tab3:
        st.header("📊 Full History & Master Report")
        df_full = get_full_quarter_history()
        if not df_full.empty:
            # Re-order and select columns for the report
            cols = ['QUARTER_NUMBER', 'EMPLOYEE_NAME', 'DESIGNATION', 'HRMS_ID', 'Allotment_Date_Disp', 'Vacation_Date_Disp', 'Status']
            # Filter only existing columns to be safe
            actual_cols = [c for c in cols if c in df_full.columns]
            
            st.dataframe(df_full[actual_cols], use_container_width=True)
            
            csv = df_full.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Download Full Report CSV", csv, "quarter_history.csv", "text/csv")
        else:
            st.info("Database khali hai.")

if __name__ == "__main__":
    main()
