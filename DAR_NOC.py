import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Path & Firebase Setup ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "assets", "DAR NOC temp.docx")

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
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}")
            st.stop()
    return firestore.client()

db = init_db()

# --- 1. Helper Functions ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def fill_bulk_template(doc, selected_data, report_date):
    """Template ki table mein multiple employees ki rows add karna"""
    # 1. [span_0](start_span)Header mein [Date] placeholder replace karein[span_0](end_span)
    for p in doc.paragraphs:
        if "[Date]" in p.text:
            p.text = p.text.replace("[Date]", report_date)

    # 2. [span_1](start_span)Table process karein[span_1](end_span)
    if len(doc.tables) > 0:
        table = doc.tables[0]
        
        # [span_2](start_span)Pehli row (index 1) placeholders wali hai[span_2](end_span)
        for i, emp in enumerate(selected_data):
            if i == 0:
                # Pehli row ke placeholders replace karein
                row_cells = table.rows[1].cells
            else:
                # Nayi row add karein
                new_row = table.add_row()
                row_cells = new_row.cells
            
            row_cells[0].text = str(i + 1)
            row_cells[1].text = emp['name']
            row_cells[2].text = emp['desig']
            row_cells[3].text = emp['pf']
            row_cells[4].text = "कर्मचारी के विरूद्ध डी.ए.आर. एवं विजिलेंस केश लम्बित नहीं है"
            
    return doc

# --- 2. Main Interface ---
st.set_page_config(layout="wide", page_title="Bulk DAR NOC System")
df_emp = get_employees()

if 'multi_dar_bytes' not in st.session_state: st.session_state.multi_dar_bytes = None

tab1, tab2 = st.tabs(["📝 Generate Multi-Employee NOC", "📊 Report Management"])

with tab1:
    st.header("Generate Single NOC for Multiple Staff")
    if not df_emp.empty:
        # Multiple employees select karne ka option
        emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected_names = st.multiselect("Employees Select Karein", emp_options)
        rep_date = st.date_input("NOC Report Date", value=datetime.now())

        if st.button("Generate Joint NOC"):
            if not selected_names:
                st.warning("Kripya kam se kam ek employee select karein.")
            elif not os.path.exists(TEMPLATE_PATH):
                st.error("Template 'DAR NOC temp.docx' nahi mila.")
            else:
                # Selected staff ka data prepare karein
                selected_data = []
                for name_str in selected_names:
                    h_id = name_str.split('(')[-1].strip(')')
                    row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                    selected_data.append({
                        "name": row.get('Employee Name', ''),
                        "desig": row.get('Designation', ''),
                        "pf": row.get('PF Number', h_id)
                    })
                
                # Document bharna
                doc = Document(TEMPLATE_PATH)
                filled_doc = fill_bulk_template(doc, selected_data, rep_date.strftime("%d/%m/%Y"))
                
                buf = io.BytesIO()
                filled_doc.save(buf)
                st.session_state.multi_dar_bytes = buf.getvalue()
                
                # History mein entry (har employee ke liye alag)
                for emp in selected_data:
                    db.collection("dar_history").add({
                        "Employee Name": emp['name'],
                        "Designation": emp['desig'],
                        "PF Number": emp['pf'],
                        "Generated_At": datetime.now(),
                        "Type": "Joint NOC"
                    })
                
                st.success(f"✅ {len(selected_data)} Employees ke liye Joint NOC taiyar hai!")

        if st.session_state.multi_dar_bytes:
            st.download_button(
                label="📥 Download Joint NOC Document",
                data=st.session_state.multi_dar_bytes,
                file_name=f"Joint_NOC_{datetime.now().strftime('%d%m%Y')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Database mein koi employee nahi mila.")

with tab2:
    st.header("📋 DAR NOC Generation Records")
    history_docs = db.collection("dar_history").order_by("Generated_At", direction=firestore.Query.DESCENDING).stream()
    h_list = [{**d.to_dict()} for d in history_docs]
    
    if h_list:
        df_h = pd.DataFrame(h_list)
        st.dataframe(df_h[['Generated_At', 'Employee Name', 'PF Number', 'Type']], use_container_width=True)
    else:
        st.info("Abhi tak koi record nahi hai.")
