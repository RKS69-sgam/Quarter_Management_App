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

# --- 2. Data Fetching (Using Lowercase Fields from Image) ---
@st.cache_data(ttl=300)
def get_employees_cached():
    db = init_db()
    # Maan ke chal rahe hain ki employee collection mein fields HRMS ID ke sath hain
    docs = db.collection("employees").stream()
    return pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in docs])

def get_full_quarter_history():
    """quarter_history se data fetch karna (Image ke columns ke anusar)"""
    db = init_db()
    try:
        docs = db.collection("quarter_history").stream()
        data = []
        for d in docs:
            item = d.to_dict()
            item['id'] = d.id
            
            # Date Handling (allotment_date aur vacation_date as per image)
            for d_field in ['allotment_date', 'vacation_date']:
                val = item.get(d_field)
                if val and hasattr(val, 'strftime'):
                    item[f'{d_field}_disp'] = val.strftime('%d-%m-%Y')
                elif val:
                    item[f'{d_field}_disp'] = str(val)
                else:
                    item[f'{d_field}_disp'] = "Occupied" if d_field == 'vacation_date' and item.get('is_current') == True else "N/A"
            data.append(item)
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Fetch Error: {e}")
        return pd.DataFrame()

# --- 3. Template Filling (Mapping DB Fields to Template Placeholders) ---
def fill_template(template_path, db_item, manual_date):
    doc = Document(template_path)
    # [span_0](start_span)[span_1](start_span)Mapping: DB Field Name -> Template Placeholder[span_0](end_span)[span_1](end_span)
    mapping = {
        "EMPLOYEE_NAME": str(db_item.get('employee_name', '')),
        "DESIGNATION": str(db_item.get('designation', '')),
        "PF_Number": str(db_item.get('pf_number', '')),
        "HRMS_ID": str(db_item.get('hrms_id', '')),
        "QUARTER_NUMBER": str(db_item.get('quarter_number', '')),
        "STATION": str(db_item.get('station', '')),
        "UNIT": str(db_item.get('unit', 'NA')),
        "DATE": manual_date # User input date
    }

    for p in doc.paragraphs:
        for key, val in mapping.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, val)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in mapping.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, val)
    return doc

# --- 4. Main App ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter MS")
    if not check_login(): return

    db = init_db()
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ALLOT_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Allotment_Template.docx")
    VACATE_TEMP = os.path.join(BASE_DIR, "assets", "Quarter_Vacation_Template.docx")

    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Full Report"])

    with tab1:
        st.header("New Quarter Allotment")
        df_emp = get_employees_cached()
        if not df_emp.empty:
            # Employee list loading (Case-sensitive check required based on your 'employees' collection)
            emp_options = df_emp.apply(lambda r: f"{r.get('Employee Name', 'Unknown')} ({r.get('HRMS ID', 'NA')})", axis=1).tolist()
            selected_emp = st.selectbox("Staff Chunein", emp_options)
            q_no = st.text_input("Quarter No.")
            stn = st.text_input("Station Name")
            a_date = st.date_input("Allotment Date")

            if st.button("Generate Allotment"):
                h_id = selected_emp.split('(')[-1].strip(')')
                emp_row = df_emp[df_emp['HRMS ID'] == h_id].iloc[0]
                
                # Firestore entry based on image columns
                new_entry = {
                    "employee_name": str(emp_row['Employee Name']),
                    "designation": str(emp_row['Designation']),
                    "hrms_id": h_id,
                    "pf_number": str(emp_row.get('PF Number', '')),
                    "quarter_number": q_no,
                    "station": stn,
                    "unit": str(emp_row.get('Unit', 'NA')),
                    "allotment_date": datetime.combine(a_date, datetime.min.time()),
                    "is_current": True,
                    "vacation_date": None
                }
                db.collection("quarter_history").add(new_entry)
                
                doc = fill_template(ALLOT_TEMP, new_entry, a_date.strftime("%d/%m/%Y"))
                buf = io.BytesIO(); doc.save(buf)
                st.success("Record Saved!")
                st.download_button("📥 Download Allotment Letter", buf.getvalue(), f"Allotment_{q_no}.docx")

    with tab2:
        st.header("Quarter Vacation")
        df_h = get_full_quarter_history()
        # is_current field ka upyog occupied check karne ke liye (As per image)
        if not df_h.empty and 'is_current' in df_h.columns:
            occupied = df_h[df_h['is_current'] == True]
            if not occupied.empty:
                q_list = occupied.apply(lambda r: f"{r['quarter_number']} - {r['employee_name']}", axis=1).tolist()
                sel_q = st.selectbox("Khali karne ke liye Quarter chunein", q_list)
                v_date = st.date_input("Vacation Date")

                if st.button("Process Vacation"):
                    q_row = occupied.iloc[q_list.index(sel_q)]
                    db.collection("quarter_history").document(q_row['id']).update({
                        "is_current": False,
                        "vacation_date": datetime.combine(v_date, datetime.min.time())
                    })
                    doc = fill_template(VACATE_TEMP, q_row.to_dict(), v_date.strftime("%d/%m/%Y"))
                    buf = io.BytesIO(); doc.save(buf)
                    st.success("Quarter Vacated!")
                    st.download_button("📥 Download Vacation Letter", buf.getvalue(), f"Vacation_{q_row['quarter_number']}.docx")
            else: st.info("Koi occupied quarter nahi mila.")

    with tab3:
        st.header("📊 Quarter Master Database")
        df_full = get_full_quarter_history()
        if not df_full.empty:
            # Columns rename for better display as per your image
            disp_df = df_full[[
                'quarter_number', 'employee_name', 'designation', 
                'hrms_id', 'allotment_date_disp', 'vacation_date_disp', 'is_current'
            ]].rename(columns={
                'quarter_number': 'Quarter No', 'employee_name': 'Name',
                'allotment_date_disp': 'Allotted', 'vacation_date_disp': 'Vacated',
                'is_current': 'Occupied Now'
            })
            st.dataframe(disp_df, use_container_width=True)
            st.download_button("📥 Export CSV", disp_df.to_csv(index=False).encode('utf-8-sig'), "Report.csv")
        else:
            st.warning("Database empty hai.")

if __name__ == "__main__":
    main()
