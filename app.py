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

    # NOC Table Logic (For Exam/DAR NOC)
    if "NOC" in letter_type and "EmployeeData" in context:
        for paragraph in doc.paragraphs:
            if "[PFNumber]" in paragraph.text:
                p = paragraph._element
                p.getparent().remove(p)
                table = doc.add_table(rows=1, cols=6)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                for i, h in enumerate(["Sr.", "PF Number", "Name", "Desig", "Subject", "Details"]):
                    hdr[i].text = h
                
                for idx, emp in enumerate(context["EmployeeData"]):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(idx + 1)
                    row_cells[1].text = str(emp["PF"])
                    row_cells[2].text = emp["Name"]
                    row_cells[3].text = emp["Desig"]
                    row_cells[4].text = emp["Subject"]
                    row_cells[5].text = emp["Details"]
                break

    # Placeholder Replacement
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
st.title("🚂 SSE/PW/SGAM - एकीकृत कार्यालय प्रबंधन")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    q_df = get_cloud_data('master_quarters')

    if not emp_df.empty:
        # Firestore Headers matching (Spaces included)
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        tab = st.sidebar.radio("Navigation", ["Letter Generation", "Quarter Management", "Digital Registers"])

        if tab == "Letter Generation":
            letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "DAR NOC temp", "pme_memo_temp"
            ])

            # Selection Logic (Multi-select for NOC, Single for others)
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
                    "Unit": str(first.get('Unit', ''))[:2],
                    "UnitNumber": first.get('Unit', ''),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "Date": date.today().strftime("%d-%m-%Y"),
                    "ShortName": "STF"
                }

                # --- SPECIAL LOGIC FOR ABSENT & SF-11 ---
                gen_sf11_also = False
                if "Absent" in letter_opt:
                    c1, c2 = st.columns(2)
                    f_dt = c1.date_input("अनुपस्थिति से (From)", value=date(2025, 10, 4))
                    t_dt = c2.date_input("अनुपस्थिति तक (To)", value=date(2025, 11, 4))
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
                        # आपके द्वारा दिया गया सटीक आरोप टेक्स्ट
                        absent_memo = f"आप बिना किसी पूर्व सूचना के दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक कुल {total_absent_days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
                        ctx["Memo"] = st.text_area("SF-11 आरोप (Memo)", value=absent_memo, height=150)

                elif "Punishment" in letter_opt:
                    ctx["Memo"] = st.text_area("दंड आदेश विवरण (Memo)")
                    ctx["Dandadesh"] = st.text_input("दंड आदेश संख्या")
                    ctx["LetterNo."] = st.text_input("पुरानी चार्जशीट फाइल नंबर")
                    ctx["SF-11Date"] = st.date_input("चार्जशीट की तारीख").strftime("%d-%m-%Y")

                elif "SF-11 temp" in letter_opt:
                    ctx["Memo"] = st.text_area("आरोप का विवरण (Memo)")

                elif "pme" in letter_opt:
                    dob = pd.to_datetime(first.get('DOB')).date()
                    ctx.update({
                        "name": first.get('Employee Name in English', ''),
                        "age": relativedelta(date.today(), dob).years,
                        "father_name": first.get("FATHER'S NAME", ''),
                        "medical_category": first.get("Medical category", "A3"),
                        "dob": dob.strftime("%d-%m-%Y")
                    })

                elif "NOC" in letter_opt:
                    emp_list = []
                    for _, r in rows.iterrows():
                        subj = st.text_input(f"Exam/Subject for {r['PF No.']}")
                        dtl = st.text_input(f"Term/Details for {r['PF No.']}", value="2024-25")
                        emp_list.append({"PF": r['PF No.'], "Name": r['Employee Name in Hindi'], "Desig": r['Designation in Hindi'], "Subject": subj, "Details": dtl})
                    ctx["EmployeeData"] = emp_list

                # --- EXECUTION ---
                if st.button("Generate & Sync"):
                    # 1. Main Letter
                    main_file = generate_doc(letter_opt, ctx, letter_opt)
                    if main_file:
                        st.download_button(f"⬇️ Download {letter_opt}", main_file, file_name=f"{letter_opt}.docx")
                    
                    # 2. Parallel SF-11
                    if gen_sf11_also:
                        sf11_file = generate_doc("SF-11 temp", ctx, "SF-11 temp")
                        if sf11_file:
                            st.download_button("⬇️ Download SF-11 (Absence)", sf11_file, file_name="SF-11_Absence.docx")
                            # Sync to Cloud Register
                            db.collection("sf11_register").add({
                                "PF No.": ctx["PFNumber"],
                                "Name": ctx["EmployeeName"],
                                "Type": "Absence Action",
                                "Date": datetime.now(),
                                "Memo": ctx["Memo"]
                            })
                            st.success("SF-11 Record synced to Digital Register!")

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
                    "ALLOTMENT DATE": date.today().strftime("%d-%m-%Y")
                })
                st.success(f"Quarter {sel_q} आवंटित कर दिया गया है।")

        elif tab == "Digital Registers":
            st.subheader("📊 लाइव क्लाउड रजिस्टर")
            choice = st.radio("रजिस्टर चुनें", ["SF-11 History", "Quarter Master"], horizontal=True)
            if choice == "SF-11 History":
                st.dataframe(get_cloud_data("sf11_register"), use_container_width=True)
            else:
                st.dataframe(q_df, use_container_width=True)

else:
    st.warning("कृपया ऐप एक्सेस करने के लिए साइडबार में सही पासवर्ड डालें।")

