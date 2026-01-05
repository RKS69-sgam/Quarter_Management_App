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
        st.error(f"Firebase Connection Error: {e}")

db = firestore.client()

# --- 2. DATA UTILITIES ---
def get_cloud_data(collection):
    docs = db.collection(collection).stream()
    data = [doc.to_dict() for doc in docs]
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. DOCUMENT GENERATION ENGINE ---
def generate_doc(template_name, context, letter_type):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template not found: {path}")
        return None
    
    doc = Document(path)

    # NOC Table Logic (Table Generation for Multi-Employee)
    if "NOC" in letter_type and "EmployeeData" in context:
        for paragraph in doc.paragraphs:
            if "[PFNumber]" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                table = doc.add_table(rows=1, cols=6)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                headers = ["Sr.", "PF Number", "Name", "Desig", "Subject", "Details"]
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
st.set_page_config(page_title="Railway Admin Cloud", layout="wide")
st.title("🚂 SSE/PW/SGAM - एकीकृत क्लाउड प्रबंधन पोर्टल")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    q_df = get_cloud_data('master_quarters')

    if not emp_df.empty:
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers", "Quarter Management"])

        if tab == "Letter Generation":
            letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "DAR NOC temp", "pme_memo_temp"
            ])

            if "NOC" in letter_opt:
                selected_names = st.multiselect("कर्मचारी चुनें", emp_df['Display'])
                rows = emp_df[emp_df['Display'].isin(selected_names)]
            else:
                selected_name = st.selectbox("कर्मचारी चुनें", emp_df['Display'])
                rows = emp_df[emp_df['Display'] == selected_name]

            if not rows.empty:
                first = rows.iloc[0]
                ctx = {
                    "EmployeeName": first.get('Employee Name in Hindi', ''),
                    "Designation": first.get('Designation in Hindi', ''),
                    "PFNumber": first.get('PF No.', ''),
                    "Unit": str(first.get('UNIT / MUSTER NUMBER', ''))[:2],
                    "UnitNumber": first.get('UNIT / MUSTER NUMBER', ''),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "Date": date.today().strftime("%d-%m-%Y"),
                    "ShortName": "STF"
                }

                gen_sf11_also = False
                # --- ABSENT LOGIC WITH SF-11 OPTION ---
                if "Absent" in letter_opt:
                    c1, c2 = st.columns(2)
                    # DEFAULT DATE LOGIC: From = Current - 6, To = Current
                    f_dt = c1.date_input("अनुपस्थिति से (From)", value=date.today() - timedelta(days=6))
                    t_dt = c2.date_input("अनुपस्थिति तक (To)", value=date.today())
                    total_absent_days = (t_dt - f_dt).days + 1
                    ctx.update({
                        "FromDate": f_dt.strftime("%d-%m-%Y"),
                        "ToDate": t_dt.strftime("%d-%m-%Y"),
                        "DutyDate": (t_dt + timedelta(days=1)).strftime("%d-%m-%Y")
                    })
                    st.info(f"कुल अनुपस्थिति दिवस: {total_absent_days}")
                    
                    st.divider()
                    gen_sf11_also = st.checkbox("क्या इस अनुपस्थिति के लिए SF-11 चार्जशीट भी बनानी है?")
                    if gen_sf11_also:
                        absent_memo = f"आप बिना किसी पूर्व सूचना के दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक कुल {total_absent_days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
                        ctx["Memo"] = st.text_area("SF-11 आरोप (Memo)", value=absent_memo, height=150)

                elif "Punishment" in letter_opt:
                    ctx["Memo"] = st.text_area("दण्ड का विवरण (Punishment Details)")
                    ctx["Dandadesh"] = st.text_input("दण्ड आदेश क्रमांक")
                    ctx["LetterNo."] = st.text_input("पत्र क्र. (Charge-sheet Ref)")
                    ctx["SF-11Date"] = st.date_input("पुराने चार्जशीट की दिनांक").strftime("%d-%m-%Y")

                elif "SF-11 temp" in letter_opt:
                    ctx["Memo"] = st.text_area("आरोप का विवरण (Memo)")
                    ctx["LetterNo."] = st.text_input("नया पत्र क्रमांक")

                # PME Auto-Calculation
                elif "pme" in letter_opt:
                    dob = pd.to_datetime(first.get('DOB')).date()
                    ctx.update({
                        "name": first.get('Employee Name in English', ''),
                        "age": relativedelta(date.today(), dob).years,
                        "father_name": first.get("FATHER'S NAME", ''),
                        "medical_category": first.get("Medical category", "A3"),
                        "dob": dob.strftime("%d-%m-%Y")
                    })

                # --- GENERATION & SF-11 REGISTER SYNC ---
                if st.button("Generate & Update Register"):
                    main_file = generate_doc(letter_opt, ctx, letter_opt)
                    if main_file:
                        st.download_button(f"⬇️ Download {letter_opt}", main_file, file_name=f"{letter_opt}.docx")
                        
                        # Register Sync Logic (Standard 16-Column Format)
                        if "SF-11" in letter_opt or gen_sf11_also:
                            reg_data = {
                                "स.क्र.": datetime.now().strftime("%Y%m%d%H%M"),
                                "पी.एफ. क्रमांक": ctx["PFNumber"],
                                "कर्मचारी का नाम": ctx["EmployeeName"],
                                "पिता का नाम": first.get("FATHER'S NAME", ""),
                                "पदनाम": ctx["Designation"],
                                "पत्र क्र.": ctx.get("LetterNo.", "SF11/GEN"),
                                "दिनांक": ctx["Date"],
                                "आरोप का विवरण": ctx.get("Memo", ""),
                                "पावती का दिनांक": "", 
                                "प्रत्‍युत्तर दिनांक": "",
                                "दण्‍डादेश क्रमांक": ctx.get("Dandadesh", ""),
                                "दण्‍डादेश जारी करने का दिनांक": ctx["Date"] if "Punishment" in letter_opt else "",
                                "दण्‍ड का विवरण": ctx.get("Memo", "") if "Punishment" in letter_opt else "",
                                "अपील का दिनांक": "",
                                "अपील निर्णय": "",
                                "रिमार्क": "Generated via Portal"
                            }
                            db.collection("sf11_register").add(reg_data)
                            st.success("SF-11 रजिस्टर में जानकारी सुरक्षित कर दी गई है।")

        elif tab == "Digital Registers":
            st.subheader("📊 SF-11 मास्टर रजिस्टर")
            sf_logs = get_cloud_data("sf11_register")
            if not sf_logs.empty:
                cols = ["स.क्र.", "पी.एफ. क्रमांक", "कर्मचारी का नाम", "पिता का नाम", "पदनाम", "पत्र क्र.", "दिनांक", "आरोप का विवरण", "पावती का दिनांक", "प्रत्‍युत्तर दिनांक", "दण्‍डादेश क्रमांक", "दण्‍डादेश जारी करने का दिनांक", "दण्‍ड का विवरण", "अपील का दिनांक", "अपील निर्णय", "रिमार्क"]
                st.dataframe(sf_logs.reindex(columns=cols), use_container_width=True)

        elif tab == "Quarter Management":
            st.subheader("🏠 डिजिटल क्वार्टर आवंटन")
            vacant = q_df[q_df['STATUS'] == 'Vacant']
            sel_q = st.selectbox("खाली क्वार्टर", vacant['QUARTER NO.'] if not vacant.empty else ["No Vacancy"])
            sel_emp = st.selectbox("आवंटी कर्मचारी", emp_df['Display'])
            e_row = emp_df[emp_df['Display'] == sel_emp].iloc[0]

            if st.button("Allot & Update Cloud"):
                q_id = f"{e_row['UNIT / MUSTER NUMBER']}_{sel_q}"
                db.collection("master_quarters").document(q_id).update({
                    "STATUS": "Occupied", 
                    "PF No.": e_row['PF No.'], 
                    "DATE": date.today().strftime("%d-%m-%Y")
                })
                st.success(f"Quarter {sel_q} आवंटित हो गया है।")
else:
    st.warning("साइडबार में पासवर्ड दर्ज करें।")
