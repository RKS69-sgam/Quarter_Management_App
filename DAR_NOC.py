import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- Login & Setup ---
ADMIN_USER = "admin"
ADMIN_PASS = "Sgam@4321"

def check_login():
    if "logged_in" not in st.session_state: st.session_state.logged_in = False
    if not st.session_state.logged_in:
        st.title("🔐 Railway DAR System Login")
        with st.form("login"):
            u, p = st.text_input("Username"), st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                if u == ADMIN_USER and p == ADMIN_PASS:
                    st.session_state.logged_in = True
                    st.rerun()
                else: st.error("Galat Credentials")
        return False
    return True

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "DAR NOC temp.docx")

@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

# --- Table Target Logic (Fixed) ---
def fill_bulk_template(doc, selected_data, report_date):
    # 1. Date Replace
    for p in doc.paragraphs:
        if "[Date]" in p.text:
            p.text = p.text.replace("[Date]", report_date)

    # 2. Sahi Table Pakadna (Last Table)
    if not doc.tables:
        st.error("Template mein table nahi mili!")
        return doc
    
    # Aapke temp mein 'सादर प्रेषित है' ke baad wali table aamtaur par last table hoti hai
    table = doc.tables[-1] 
    num_cols = len(table.columns)

    for i, emp in enumerate(selected_data):
        # Index 1 placeholders wali row hai (S.No 1 wali)
        if i == 0 and len(table.rows) >= 2:
            row_cells = table.rows[1].cells
        else:
            row_cells = table.add_row().cells
            
        if num_cols > 0: row_cells[0].text = str(i + 1)
        if num_cols > 1: row_cells[1].text = emp['name']
        if num_cols > 2: row_cells[2].text = emp['desig']
        if num_cols > 3: row_cells[3].text = emp['pf']
        if num_cols > 4: row_cells[4].text = "कर्मचारी के विरूद्ध डी.ए.आर. एवं विजिलेंस केश लम्बित नहीं है"
            
    return doc

# --- Main App ---
def main():
    st.set_page_config(layout="wide")
    if not check_login(): return

    db = init_db()
    docs = db.collection("employees").stream()
    df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

    tab1, tab2 = st.tabs(["📝 Generate NOC", "📊 History"])

    with tab1:
        if not df_emp.empty:
            emp_opts = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
            selected = st.multiselect("Select Staff", emp_opts)
            rep_date = st.date_input("NOC Date", value=datetime.now())

            if st.button("Generate Joint NOC"):
                if not selected: st.warning("Staff chunein")
                else:
                    selected_data = []
                    for s in selected:
                        h_id = s.split('(')[-1].strip(')')
                        row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                        selected_data.append({"name": row.get('Employee Name'), "desig": row.get('Designation'), "pf": row.get('PF Number', h_id)})
                    
                    doc = Document(TEMPLATE_PATH)
                    filled_doc = fill_bulk_template(doc, selected_data, rep_date.strftime("%d/%m/%Y"))
                    
                    buf = io.BytesIO()
                    filled_doc.save(buf)
                    st.success("NOC Generated!")
                    st.download_button("📥 Download", buf.getvalue(), f"NOC_{datetime.now().strftime('%d%m%Y')}.docx")
        else: st.warning("DB Khali hai")

if __name__ == "__main__":
    main()
