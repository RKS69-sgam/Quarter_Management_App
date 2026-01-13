
import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Setup & Paths ---
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
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

db = init_db()

# --- 1. Helper Functions ---
def get_employees():
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def replace_dar_placeholders(doc, data):
    """Template ke placeholders [Employee Name], [Designation], [PF Number] ko replace karna"""
    for p in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"[{key}]"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(value))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in data.items():
                        placeholder = f"[{key}]"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(value))

# --- 2. Streamlit UI ---
st.set_page_config(layout="wide", page_title="DAR NOC Management")
df_emp = get_employees()

if 'dar_bytes' not in st.session_state: st.session_state.dar_bytes = None
if 'dar_filename' not in st.session_state: st.session_state.dar_filename = ""

tab1, tab2 = st.tabs(["📄 Generate DAR NOC", "📊 Report & Records"])

with tab1:
    st.header("Generate DAR/Vigilance NOC")
    if not df_emp.empty:
        emp_list = df_emp.apply(lambda r: f"{r.get('Employee Name')} ({r.get('HRMS ID')})", axis=1).tolist()
        selected = st.selectbox("Karmchari Chunein", emp_list)
        h_id = selected.split('(')[-1].strip(')')
        emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]

        with st.form("dar_form"):
            current_dt = st.date_input("Report Date", value=datetime.now())
            
            dar_data = {
                "Employee Name": emp_row.get('Employee Name', ''),
                "Designation": emp_row.get('Designation', ''),
                "PF Number": emp_row.get('PF Number', emp_row.get('HRMS ID', '')), # PF No na hone par HRMS ID use karega
                "Date": current_dt.strftime("%d/%m/%Y")
            }

            if st.form_submit_button("Generate NOC Document"):
                if os.path.exists(TEMPLATE_PATH):
                    doc = Document(TEMPLATE_PATH)
                    replace_dar_placeholders(doc, dar_data)
                    
                    buf = io.BytesIO()
                    doc.save(buf)
                    st.session_state.dar_bytes = buf.getvalue()
                    st.session_state.dar_filename = f"DAR_NOC_{dar_data['Employee Name'].replace(' ', '_')}.docx"
                    
                    # Firebase History mein save karein
                    db.collection("dar_history").add({
                        **dar_data,
                        "Generated_At": datetime.now(),
                        "Status": "Clear"
                    })
                    st.success(f"✅ NOC Taiyar hai: {dar_data['Employee Name']}")
                else:
                    st.error("Template file 'DAR NOC temp.docx' assets folder mein nahi mili.")

        if st.session_state.dar_bytes:
            st.download_button(
                label="📥 Download NOC",
                data=st.session_state.dar_bytes,
                file_name=st.session_state.dar_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Database mein koi karmchari nahi mila.")

with tab2:
    st.header("📋 DAR NOC Generation Records")
    try:
        history_docs = db.collection("dar_history").order_by("Generated_At", direction=firestore.Query.DESCENDING).stream()
        history_data = [{**d.to_dict()} for d in history_docs]
        
        if history_data:
            df_hist = pd.DataFrame(history_data)
            # Display relevant columns
            cols = ['Generated_At', 'Employee Name', 'Designation', 'PF Number', 'Status']
            st.dataframe(df_hist[cols], use_container_width=True)
            
            # CSV Download option for reports
            csv = df_hist.to_csv(index=False).encode('utf-8')
            st.download_button("📊 Export Report (CSV)", csv, "DAR_NOC_Report.csv", "text/csv")
        else:
            st.info("Abhi tak koi record generate nahi kiya gaya hai.")
    except Exception as e:
        st.error(f"Error fetching records: {e}")
