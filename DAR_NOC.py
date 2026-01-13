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
        st.title("🔐 Railway DAR System Login")
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

# --- 2. Table Logic (Error Handled) ---
def fill_bulk_template(doc, selected_data, report_date):
    # Header Date Replace
    for p in doc.paragraphs:
        if "[Date]" in p.text:
            p.text = p.text.replace("[Date]", str(report_date))

    if not doc.tables:
        st.error("Template mein koi table nahi mili!")
        return doc
    
    # SADAR PRESHIT ke baad wali table (Last Table)
    table = doc.tables[-1] 
    num_cols = len(table.columns)

    for i, emp in enumerate(selected_data):
        # Index 1 placeholders wali row hai (S.No 1 wali)
        if i == 0 and len(table.rows) >= 2:
            row_cells = table.rows[1].cells
        else:
            row_cells = table.add_row().cells
            
        # str() wrap kiya gaya hai taaki TypeError na aaye
        if num_cols > 0: row_cells[0].text = str(i + 1)
        if num_cols > 1: row_cells[1].text = str(emp.get('name', ''))
        if num_cols > 2: row_cells[2].text = str(emp.get('desig', ''))
        if num_cols > 3: row_cells[3].text = str(emp.get('pf', 'N/A')) # 'N/A' if None
        if num_cols > 4: row_cells[4].text = "कर्मचारी के विरूद्ध डी.ए.आर. एवं विजिलेंस केश लम्बित नहीं है"
            
    return doc

# --- 3. Main App ---
def main():
    st.set_page_config(layout="wide", page_title="DAR NOC Management")
    if not check_login(): return

    if st.sidebar.button("🚪 Logout"):
        st.session_state.logged_in = False
        st.rerun()

    db = init_db()
    
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "DAR NOC temp.docx")

    # Data fetch
    docs = db.collection("employees").stream()
    df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

    tab1, tab2 = st.tabs(["📝 Generate Joint NOC", "📊 Records History"])

    with tab1:
        st.header("Bulk DAR NOC Generation")
        if not df_emp.empty:
            emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            selected_names = st.multiselect("Select Staff", emp_options)
            rep_date = st.date_input("NOC Date", value=datetime.now())

            if st.button("Generate Joint NOC"):
                if not selected_names:
                    st.warning("Staff select karein.")
                elif not os.path.exists(TEMPLATE_PATH):
                    st.error(f"Template nahi mila: {TEMPLATE_PATH}")
                else:
                    selected_data = []
                    for name_str in selected_names:
                        h_id = name_str.split('(')[-1].strip(')')
                        row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                        # Safe extraction to avoid NoneType
                        selected_data.append({
                            "name": str(row.get('Employee Name', '')),
                            "desig": str(row.get('Designation', '')),
                            "pf": str(row.get('PF Number') if row.get('PF Number') else h_id)
                        })
                    
                    doc = Document(TEMPLATE_PATH)
                    filled_doc = fill_bulk_template(doc, selected_data, rep_date.strftime("%d/%m/%Y"))
                    
                    buf = io.BytesIO()
                    filled_doc.save(buf)
                    
                    # Save History
                    for emp in selected_data:
                        db.collection("dar_history").add({
                            "Employee Name": emp['name'],
                            "PF Number": emp['pf'],
                            "Generated_At": datetime.now(),
                            "Type": "Joint NOC"
                        })
                    
                    st.success("✅ NOC Generated Successfully!")
                    st.download_button("📥 Download Document", buf.getvalue(), f"NOC_{datetime.now().strftime('%d%m%Y')}.docx")
        else:
            st.warning("Database khali hai.")

    with tab2:
        st.header("📋 Recent DAR Records")
        h_docs = db.collection("dar_history").order_by("Generated_At", direction=firestore.Query.DESCENDING).limit(50).stream()
        h_list = [{**d.to_dict()} for d in h_docs]
        if h_list:
            st.dataframe(pd.DataFrame(h_list)[['Generated_At', 'Employee Name', 'PF Number']], use_container_width=True)

if __name__ == "__main__":
    main()
