import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import io
import base64
from docx import Document
from docx.shared import Inches
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
        st.error(f"Firebase Error: {e}")

db = firestore.client()

# --- 2. DATA FETCH HELPERS ---
def get_cloud_data(collection):
    docs = db.collection(collection).stream()
    data = [doc.to_dict() for doc in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. HELPER FUNCTIONS ---
def format_date_safe(date_val):
    if pd.isna(date_val): return "N/A"
    if isinstance(date_val, datetime): return date_val.strftime("%d-%m-%Y")
    return str(date_val)

def generate_word_file(template_path, context, letter_type):
    if not os.path.exists(template_path):
        st.error(f"Template missing: {template_path}")
        return None
    
    doc = Document(template_path)

    # --- SPECIAL CASE: Table Insertion for NOCs (As per your old logic) ---
    if letter_type in ["Exam NOC Letter temp", "DAR NOC temp"] and "EmployeeData" in context:
        for paragraph in doc.paragraphs:
            # Table Logic for Exam NOC
            if letter_type == "Exam NOC Letter temp" and "[PFNumber]" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                table = doc.add_table(rows=1, cols=6)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                headers = ["Sr.", "PF Number", "Name", "Desig", "Exam", "Term"]
                for i, h in enumerate(headers): hdr[i].text = h
                
                for idx, emp in enumerate(context["EmployeeData"]):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(idx + 1)
                    row_cells[1].text = str(emp["PF Number"])
                    row_cells[2].text = emp["Employee Name"]
                    row_cells[3].text = emp["Designation"]
                    row_cells[4].text = emp["Exam Name"]
                    row_cells[5].text = emp["Term"]
                break

    # General Placeholder Replacement
    def replace_tags(text, ctx):
        for k, v in ctx.items():
            text = text.replace(f"[{k}]", str(v)).replace(f"{{{{ {k} }}}}", str(v)).replace(f"{{{{{k}}}}}", str(v))
        return text

    for p in doc.paragraphs: p.text = replace_tags(p.text, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: p.text = replace_tags(p.text, context)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. MAIN UI ---
st.set_page_config(page_title="SSE/PW/SGAM Cloud", layout="wide")
st.title("🚂 SSE/PW/SGAM - Complete Management System")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    q_df = get_cloud_data('master_quarters')

    if not emp_df.empty:
        # Firestore headers with spaces
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        tab_choice = st.sidebar.radio("Navigation", ["Letters", "Quarter System", "Registers"])

        # --- TAB: LETTERS ---
        if tab_choice == "Letters":
            letter_opt = st.selectbox("Letter Type", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "DAR NOC temp", "pme_memo_temp"
            ])

            # Multi-select for NOCs, Single for others
            if "NOC" in letter_opt:
                selected_names = st.multiselect("Select Employees", emp_df['Display'])
                selected_rows = emp_df[emp_df['Display'].isin(selected_names)]
            else:
                selected_name = st.selectbox("Select Employee", emp_df['Display'])
                selected_rows = emp_df[emp_df['Display'] == selected_name]

            if not selected_rows.empty:
                first_row = selected_rows.iloc[0]
                ctx = {
                    "EmployeeName": first_row.get('Employee Name in Hindi', ''),
                    "Designation": first_row.get('Designation in Hindi', ''),
                    "PFNumber": first_row.get('PF No.', ''),
                    "Unit": str(first_row.get('UNIT / MUSTER NUMBER', ''))[:2],
                    "UnitNumber": first_row.get('UNIT / MUSTER NUMBER', ''),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "Date": date.today().strftime("%d-%m-%Y")
                }

                # Conditional Inputs
                if "Exam NOC" in letter_opt:
                    emp_data_list = []
                    for _, r in selected_rows.iterrows():
                        exam = st.text_input(f"Exam for {r['PF No.']}", key=f"ex_{r['PF No.']}")
                        term = st.text_input(f"Term for {r['PF No.']}", value="2024-25", key=f"tr_{r['PF No.']}")
                        emp_data_list.append({"PF Number": r['PF No.'], "Employee Name": r['Employee Name in Hindi'], "Designation": r['Designation in Hindi'], "Exam Name": exam, "Term": term})
                    ctx["EmployeeData"] = emp_data_list

                if "pme_memo" in letter_opt:
                    dob = pd.to_datetime(first_row.get('DOB')).date()
                    doa = pd.to_datetime(first_row.get('DOA')).date()
                    age = relativedelta(date.today(), dob).years
                    ctx.update({"dob": dob.strftime("%d-%m-%Y"), "age": age, "medical_category": first_row.get("Medical category", "A3")})

                if st.button("Generate Letter & Sync"):
                    res = generate_word_file(f"assets/{letter_opt}.docx", ctx, letter_opt)
                    if res:
                        st.download_button("Download", res, file_name=f"{letter_opt}.docx")
                        
                        # SF-11 Register Update in Firestore
                        if "SF-11" in letter_opt:
                            db.collection("sf11_register").add({
                                "PF No.": ctx["PFNumber"],
                                "Name": ctx["EmployeeName"],
                                "Letter Type": letter_opt,
                                "Generated At": datetime.now()
                            })
                            st.success("SF-11 Register Updated on Cloud!")

        # --- TAB: QUARTER SYSTEM ---
        elif tab_choice == "Quarter System":
            st.subheader("🏠 Quarter Management")
            vacant_qs = q_df[q_df['STATUS'] == 'Vacant']
            sel_q = st.selectbox("Quarter No.", vacant_qs['QUARTER NO.'] if not vacant_qs.empty else ["No Vacant"])
            sel_emp = st.selectbox("Assign To", emp_df['Display'])
            target = emp_df[emp_df['Display'] == sel_emp].iloc[0]

            if st.button("Allot & Update Database"):
                q_id = f"{target['UNIT / MUSTER NUMBER']}_{sel_q}"
                db.collection("master_quarters").document(q_id).update({
                    "STATUS": "Occupied",
                    "PF No.": target['PF No.'],
                    "OCCUPIED DATE": date.today().strftime("%d-%m-%Y")
                })
                st.success("Quarter Database Updated!")

        # --- TAB: REGISTERS ---
        elif tab_choice == "Registers":
            st.subheader("📊 Live Cloud Registers")
            reg_type = st.selectbox("View Register", ["SF-11 Logs", "Quarter Master"])
            if reg_type == "SF-11 Logs":
                st.dataframe(get_cloud_data("sf11_register"), use_container_width=True)
            else:
                st.dataframe(q_df, use_container_width=True)

else:
    st.info("Sidebar me password enter karein.")
