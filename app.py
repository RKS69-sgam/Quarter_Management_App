import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import io
from docx import Document
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta

# --- 1. FIREBASE INITIALIZATION ---
if not firebase_admin._apps:
    try:
        cred_dict = dict(st.secrets["firebase_config"])
        if isinstance(cred_dict.get('private_key'), str):
            cred_dict['private_key'] = cred_dict['private_key'].replace('\\n', '\n')
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"Firebase Setup Error: {e}")

db = firestore.client()

# --- 2. DATA FETCH HELPERS ---
def get_cloud_data(collection_name):
    docs = db.collection(collection_name).stream()
    data = [doc.to_dict() for doc in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. DOCUMENT GENERATION ENGINE ---
def generate_docx(template_name, context):
    template_path = f"assets/{template_name}.docx"
    if not os.path.exists(template_path):
        st.error(f"Template not found: {template_path}")
        return None
    
    doc = Document(template_path)
    
    def replace_all(text, ctx):
        for k, v in ctx.items():
            # [Key] और {{ Key }} दोनों को सपोर्ट करता है
            text = text.replace(f"[{k}]", str(v)).replace(f"{{{{ {k} }}}}", str(v)).replace(f"{{{{{k}}}}}", str(v))
        return text

    for p in doc.paragraphs:
        p.text = replace_all(p.text, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.text = replace_all(p.text, context)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. MAIN UI ---
st.set_page_config(page_title="Railway Admin Cloud", layout="wide")
st.title("🚂 SSE/PW/SGAM - Cloud Integrated Management")

if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    # Firestore से डेटा लोड करना
    emp_df = get_cloud_data('employees')
    q_df = get_cloud_data('master_quarters')

    if not emp_df.empty:
        # 'PF No.' और 'Employee Name in Hindi' का उपयोग करके लिस्ट बनाना
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        menu = st.sidebar.radio("Menu", ["Letter Generation", "Quarter Management", "SF-11 Records"])

        # --- TAB: LETTER GENERATION ---
        if menu == "Letter Generation":
            letter = st.selectbox("Select Letter Type", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "DAR NOC temp", "pme_memo_temp"
            ])
            
            sel_display = st.selectbox("Select Employee", emp_df['Display'])
            row = emp_df[emp_df['Display'] == sel_display].iloc[0]

            # Firestore के सटीक स्पेस वाले हेडर्स का उपयोग
            ctx = {
                "EmployeeName": row.get('Employee Name in Hindi', ''),
                "Designation": row.get('Designation in Hindi', ''),
                "PFNumber": row.get('PF No.', ''),
                "Unit": row.get('UNIT / MUSTER NUMBER', '')[:2] if row.get('UNIT / MUSTER NUMBER') else '',
                "UnitNumber": row.get('UNIT / MUSTER NUMBER', ''),
                "LetterDate": date.today().strftime("%d-%m-%Y"),
                "Date": date.today().strftime("%d-%m-%Y")
            }

            # PME के लिए विशेष गणना (DOB और DOA)
            if "pme" in letter.lower():
                dob = pd.to_datetime(row.get('DOB')).date()
                doa = pd.to_datetime(row.get('DOA')).date()
                age = relativedelta(date.today(), dob).years
                ctx.update({
                    "dob": dob.strftime("%d-%m-%Y"),
                    "doa": doa.strftime("%d-%m-%Y"),
                    "age": age,
                    "father_name": row.get("FATHER'S NAME", 'N/A')
                })

            if st.button("Generate & Sync"):
                doc_io = generate_docx(letter, ctx)
                if doc_io:
                    st.download_button("⬇️ Download Document", doc_io, file_name=f"{letter}.docx")
                    
                    # SF-11 का डेटा Firestore में अपडेट करना
                    if "SF-11" in letter:
                        db.collection("sf11_register").add({
                            "PF No.": ctx["PFNumber"],
                            "Employee Name": ctx["EmployeeName"],
                            "Letter Type": letter,
                            "Timestamp": datetime.now(),
                            "Status": "Generated"
                        })
                        st.success("Firestore SF-11 Register Updated!")

        # --- TAB: QUARTER MANAGEMENT ---
        elif menu == "Quarter Management":
            st.subheader("🏠 Quarter Allotment (Live Status)")
            # यहाँ भी स्पेस वाले हेडर्स 'STATION' और 'QUARTER NO.'
            vacant_qs = q_df[q_df['STATUS'] == 'Vacant']
            
            q_choice = st.selectbox("Select Quarter", vacant_qs['QUARTER NO.'] if not vacant_qs.empty else ["No Vacancy"])
            emp_choice = st.selectbox("Assign To", emp_df['Display'])
            target_row = emp_df[emp_df['Display'] == emp_choice].iloc[0]

            if st.button("Allot & Update Cloud"):
                # Firestore में क्वार्टर का स्टेटस बदलना
                q_doc_id = f"{target_row['UNIT / MUSTER NUMBER']}_{q_choice}"
                db.collection("master_quarters").document(q_doc_id).update({
                    "STATUS": "Occupied",
                    "PF No.": target_row['PF No.'],
                    "OCCUPIED DATE": date.today().strftime("%d-%m-%Y")
                })
                st.success(f"Quarter {q_choice} is now Occupied by {target_row['Employee Name in Hindi']}")

        # --- TAB: SF-11 RECORDS ---
        elif menu == "SF-11 Records":
            st.subheader("📊 Cloud SF-11 Register")
            sf_logs = get_cloud_data("sf11_register")
            st.dataframe(sf_logs, use_container_width=True)

else:
    st.warning("Please enter correct password to access cloud database.")
