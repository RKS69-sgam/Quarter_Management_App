import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import io
from docx import Document
from datetime import date, datetime, timedelta

# --- 1. FIREBASE INITIALIZATION ---
if not firebase_admin._apps:
    try:
        cred_dict = dict(st.secrets["firebase_config"])
        if isinstance(cred_dict.get('private_key'), str):
            cred_dict['private_key'] = cred_dict['private_key'].replace('\\n', '\n')
        cred = credentials.Certificate(cred_dict)
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"Firebase Connection Error: {e}")

db = firestore.client()

# --- 2. DATA UTILITIES ---
def get_cloud_data(collection):
    docs = db.collection(collection).stream()
    data = []
    for doc in docs:
        d = doc.to_dict()
        d['doc_id'] = doc.id 
        data.append(d)
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. DOCUMENT ENGINE ---
def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template missing: {path}")
        return None
    doc = Document(path)
    def replace_tags(text, ctx):
        for k, v in ctx.items():
            val = str(v) if v is not None else ""
            text = text.replace(f"[{k}]", val).replace(f"{{{{ {k} }}}}", val).replace(f"{{{{{k}}}}}", val)
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
st.set_page_config(page_title="Railway Admin Portal", layout="wide")
st.title("🚂 SSE/PW/SGAM - एकीकृत एडमिन पोर्टल")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    sf_reg_df = get_cloud_data('sf11_register')

    # Column Validation for SF-11 Register
    required_cols = ['दण्‍डादेश क्रमांक', 'अपील का दिनांक', 'अपील निर्णय', 'रिमार्क']
    for col in required_cols:
        if not sf_reg_df.empty and col not in sf_reg_df.columns:
            sf_reg_df[col] = ""

    tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers"])

    if tab == "Letter Generation":
        letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
            "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
            "SICK MEMO temp.", "Exam NOC Letter temp", "pme_memo_temp"
        ])

        # कर्मचारी डेटा की सफाई
        if not emp_df.empty:
            emp_df['PF_Clean'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False)
            emp_df['Full_Disp'] = emp_df['PF_Clean'] + " - " + emp_df['Employee Name in Hindi'].astype(str)
            
            # --- CASE 1: SF-11 PUNISHMENT ORDER ---
            if "Punishment order" in letter_opt:
                if not sf_reg_df.empty:
                    pending_sf = sf_reg_df[sf_reg_df['दण्‍डादेश क्रमांक'].isna() | (sf_reg_df['दण्‍डादेश क्रमांक'].astype(str) == '')].copy()
                    if not pending_sf.empty:
                        pending_sf['PF_Disp'] = pending_sf['पी.एफ. क्रमांक'].astype(str).str.replace('.0', '', regex=False)
                        selected_val = st.selectbox("पेंडिंग चार्जशीट चुनें", pending_sf['PF_Disp'] + " | " + pending_sf['कर्मचारी का नाम'])
                        sf_row = pending_sf[(pending_sf['PF_Disp'] + " | " + pending_sf['कर्मचारी का नाम']) == selected_val].iloc[0]
                        
                        col1, col2 = st.columns(2)
                        o_date = col1.date_input("दण्डादेश दिनांक").strftime("%d-%m-%Y")
                        o_num = col1.text_input("दण्डादेश क्रमांक", value="SGAM/SF-11/Order/")
                        punishment_memo = col2.selectbox("दंड का प्रकार", [
                            "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                            "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                            "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता।"
                        ])

                        ctx = {
                            "EmployeeName": sf_row['कर्मचारी का नाम'], "PFNumber": sf_row['PF_Disp'],
                            "LetterNo.": sf_row['पत्र क्र.'], "SF-11Date": sf_row['दिनांक'],
                            "LetterDate": o_date, "Dandadesh": o_num, "Memo": punishment_memo, "Unit": "SGAM"
                        }

                        if st.button("Generate & Update"):
                            doc = generate_doc(letter_opt, ctx)
                            if doc:
                                st.download_button("⬇️ Download", doc, file_name=f"Order_{ctx['PFNumber']}.docx")
                                db.collection("sf11_register").document(sf_row['doc_id']).update({
                                    "दण्‍डादेश क्रमांक": o_num, "दण्‍डादेश जारी करने का दिनांक": o_date,
                                    "दण्‍ड का विवरण": punishment_memo, "रिमार्क": "Punishment Issued"
                                })
                                st.success("Updated!")
                    else: st.info("No pending SF-11 found.")

            # --- CASE 2: ABSENT CASE (DUTY + SF-11) ---
            elif "Absent" in letter_opt:
                sel_emp = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
                e_row = emp_df[emp_df['Full_Disp'] == sel_emp].iloc[0]
                
                f_dt = st.date_input("अनुपस्थिति से", value=date.today() - timedelta(days=6))
                t_dt = st.date_input("अनुपस्थिति तक", value=date.today())
                
                ctx = {
                    "EmployeeName": e_row['Employee Name in Hindi'],
                    "Designation": e_row['Designation in Hindi'],
                    "PFNumber": e_row['PF_Clean'],
                    "Unit": str(e_row.get('UNIT / MUSTER NUMBER', '')).replace('.0', ''),
                    "FromDate": f_dt.strftime("%d-%m-%Y"),
                    "ToDate": t_dt.strftime("%d-%m-%Y"),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime("%d-%m-%Y"),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "LetterNo.": f"SGAM/SF11/ABS/{e_row['PF_Clean']}"
                }
                ctx["Memo"] = st.text_area("SF-11 आरोप", value=f"आप दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक अनुपस्थित थे...")

                if st.button("Generate Documents"):
                    d_doc = generate_doc("Absent Duty letter temp", ctx)
                    s_doc = generate_doc("SF-11 temp", ctx)
                    if d_doc: st.download_button("⬇️ Duty Letter", d_doc, file_name="Duty.docx")
                    if s_doc: 
                        st.download_button("⬇️ SF-11", s_doc, file_name="SF11.docx")
                        db.collection("sf11_register").add({
                            "पी.एफ. क्रमांक": ctx["PFNumber"], "कर्मचारी का नाम": ctx["EmployeeName"],
                            "दिनांक": ctx["LetterDate"], "पत्र क्र.": ctx["LetterNo."], "आरोप का विवरण": ctx["Memo"]
                        })

            # --- CASE 3: GENERAL LETTERS (SF-11, SICK, NOC, PME) ---
            else:
                sel_emp = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
                e_row = emp_df[emp_df['Full_Disp'] == sel_emp].iloc[0]
                
                ctx = {
                    "EmployeeName": e_row['Employee Name in Hindi'],
                    "Designation": e_row['Designation in Hindi'],
                    "PFNumber": e_row['PF_Clean'],
                    "Unit": str(e_row.get('UNIT / MUSTER NUMBER', '')).replace('.0', ''),
                    "Date": date.today().strftime("%d-%m-%Y"),
                    "LetterDate": date.today().strftime("%d-%m-%Y")
                }
                
                # Extra fields for specific letters
                if "SICK" in letter_opt:
                    ctx["Time"] = st.text_input("समय (Time)", value=datetime.now().strftime("%H:%M"))
                
                if st.button("Generate Letter"):
                    doc = generate_doc(letter_opt, ctx)
                    if doc:
                        st.download_button("⬇️ Download Letter", doc, file_name=f"{letter_opt}.docx")

    elif tab == "Digital Registers":
        st.subheader("📊 SF-11 मास्टर रजिस्टर")
        if not sf_reg_df.empty:
            st.dataframe(sf_reg_df.drop(columns=['doc_id']).astype(str), use_container_width=True)
else:
    st.warning("Password please.")
