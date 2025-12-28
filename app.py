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
        st.error(f"Firebase Error: {e}")

db = firestore.client()

# --- 2. DATA HELPERS ---
def get_cloud_data(collection):
    docs = db.collection(collection).stream()
    data = [doc.to_dict() for doc in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. DOCUMENT ENGINE ---
def generate_office_letter(template_name, context, letter_type):
    template_path = f"assets/{template_name}.docx"
    if not os.path.exists(template_path):
        st.error(f"Template not found: {template_path}")
        return None
    
    doc = Document(template_path)

    # NOC Table Logic
    if "NOC" in letter_type and "EmployeeData" in context:
        for paragraph in doc.paragraphs:
            if "[PFNumber]" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                table = doc.add_table(rows=1, cols=6)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                headers = ["Sr.", "PF Number", "Name", "Desig", "Exam/Purpose", "Term/Details"]
                for i, h in enumerate(headers): hdr[i].text = h
                
                for idx, emp in enumerate(context["EmployeeData"]):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(idx + 1)
                    row_cells[1].text = str(emp.get("PF", ""))
                    row_cells[2].text = emp.get("Name", "")
                    row_cells[3].text = emp.get("Desig", "")
                    row_cells[4].text = emp.get("Subject", "")
                    row_cells[5].text = emp.get("Details", "")
                break

    def replace_placeholder(text, ctx):
        for k, v in ctx.items():
            text = text.replace(f"[{k}]", str(v)).replace(f"{{{{ {k} }}}}", str(v)).replace(f"{{{{{k}}}}}", str(v))
        return text

    for p in doc.paragraphs: p.text = replace_placeholder(p.text, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: p.text = replace_placeholder(p.text, context)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. MAIN INTERFACE ---
st.set_page_config(page_title="Railway Admin Portal", layout="wide")
st.title("🚂 SSE/PW/SGAM - एकीकृत क्लाउड पोर्टल")

if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    q_df = get_cloud_data('master_quarters')

    if not emp_df.empty:
        # Space-based Headers from Firestore
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        menu = st.sidebar.radio("Menu", ["Letter Generation", "Quarter Master", "Digital Registers"])

        if menu == "Letter Generation":
            letter_opt = st.selectbox("Select Letter", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "DAR NOC temp", "pme_memo_temp"
            ])

            # Selection Logic
            if "NOC" in letter_opt:
                selected_names = st.multiselect("Select Employees", emp_df['Display'])
                target_rows = emp_df[emp_df['Display'].isin(selected_names)]
            else:
                selected_name = st.selectbox("Select Employee", emp_df['Display'])
                target_rows = emp_df[emp_df['Display'] == selected_name]

            if not target_rows.empty:
                first_row = target_rows.iloc[0]
                # Default Context
                ctx = {
                    "EmployeeName": first_row.get('Employee Name in Hindi', ''),
                    "Designation": first_row.get('Designation in Hindi', ''),
                    "PFNumber": first_row.get('PF No.', ''),
                    "Unit": str(first_row.get('UNIT / MUSTER NUMBER', ''))[:2],
                    "UnitNumber": first_row.get('UNIT / MUSTER NUMBER', ''),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "Date": date.today().strftime("%d-%m-%Y"),
                    "ShortName": "STF"
                }

                # --- DYNAMIC INPUTS ---
                if "Absent" in letter_opt:
                    col1, col2 = st.columns(2)
                    f_date = col1.date_input("Absent From Date")
                    t_date = col2.date_input("Absent To Date")
                    d_date = st.date_input("Reporting Duty Date", value=t_date + timedelta(days=1))
                    ctx.update({
                        "FromDate": f_date.strftime("%d-%m-%Y"),
                        "ToDate": t_date.strftime("%d-%m-%Y"),
                        "DutyDate": d_date.strftime("%d-%m-%Y")
                    })

                elif "SF-11" in letter_opt:
                    ctx["Memo"] = st.text_area("विवरण / Charge Details (Memo)", height=150)
                    if "Punishment" in letter_opt:
                        ctx["Dandadesh"] = st.text_input("दंड आदेश नंबर", value="SGAM/SF-11/Order/01")
                        ctx["SF-11Date"] = st.date_input("पुराने SF-11 की तारीख").strftime("%d-%m-%Y")
                        ctx["LetterNo."] = st.text_input("पुरानी चार्जशीट फाइल नंबर")

                elif "NOC" in letter_opt:
                    emp_data_list = []
                    for _, r in target_rows.iterrows():
                        subj = st.text_input(f"Exam/Subject for {r['PF No.']}")
                        dtl = st.text_input(f"Term/Details for {r['PF No.']}", value="2024-25")
                        emp_data_list.append({"PF": r['PF No.'], "Name": r['Employee Name in Hindi'], "Desig": r['Designation in Hindi'], "Subject": subj, "Details": dtl})
                    ctx["EmployeeData"] = emp_data_list

                elif "pme" in letter_opt:
                    dob = pd.to_datetime(first_row.get('DOB')).date()
                    ctx.update({
                        "name": first_row.get('Employee Name in English', ''),
                        "age": relativedelta(date.today(), dob).years,
                        "father_name": first_row.get("FATHER'S NAME", ''),
                        "medical_category": first_row.get("Medical category", "A3"),
                        "dob": dob.strftime("%d-%m-%Y")
                    })

                if st.button("Generate & Sync to Cloud"):
                    doc_result = generate_office_letter(letter_opt, ctx, letter_opt)
                    if doc_result:
                        st.download_button("Download Letter", doc_result, file_name=f"{letter_opt}.docx")
                        
                        # Register Update
                        if "SF-11" in letter_opt:
                            db.collection("sf11_register").add({
                                "PF No.": ctx["PFNumber"],
                                "Name": ctx["EmployeeName"],
                                "Memo": ctx.get("Memo", ""),
                                "Date": datetime.now(),
                                "Type": letter_opt
                            })
                            st.success("Firestore Updated!")

        elif menu == "Quarter Master":
            st.subheader("🏠 Quarter Allotment")
            v_qs = q_df[q_df['STATUS'] == 'Vacant']
            sel_q = st.selectbox("Quarter", v_qs['QUARTER NO.'] if not v_qs.empty else ["No Vacancy"])
            sel_e = st.selectbox("Assign To", emp_df['Display'])
            e_row = emp_df[emp_df['Display'] == sel_e].iloc[0]

            if st.button("Allot Quarter"):
                q_id = f"{e_row['UNIT / MUSTER NUMBER']}_{sel_q}"
                db.collection("master_quarters").document(q_id).update({
                    "STATUS": "Occupied", "PF No.": e_row['PF No.'], "DATE": date.today().strftime("%d-%m-%Y")
                })
                st.success("Quarter Database Sync Complete.")

        elif menu == "Digital Registers":
            st.subheader("📊 Live Database Views")
            reg = st.radio("Choose Register", ["SF-11 History", "Quarter Status"], horizontal=True)
            if reg == "SF-11 History":
                st.dataframe(get_cloud_data("sf11_register"), use_container_width=True)
            else:
                st.dataframe(q_df, use_container_width=True)
else:
    st.warning("Side menu mein password enter karein.")
