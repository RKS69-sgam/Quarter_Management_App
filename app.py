import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import io
from docx import Document
from datetime import date, datetime, timedelta
from docx.shared import Inches
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
    data = []
    for doc in docs:
        d = doc.to_dict()
        d['doc_id'] = doc.id 
        data.append(d)
    return pd.DataFrame(data) if data else pd.DataFrame()

def parse_date_safe(date_val):
    if pd.isna(date_val) or date_val == "": return None
    if isinstance(date_val, (date, datetime)): return date_val
    try: return datetime.strptime(str(date_val).split()[0], "%Y-%m-%d").date()
    except: return None

# --- 3. DOCUMENT ENGINE ---
def replace_placeholders(doc, ctx):
    """दस्तावेज़ के पैराग्राफ और टेबल में टैग्स को पूरी तरह बदलता है।"""
    for p in doc.paragraphs:
        for run in p.runs:
            for k, v in ctx.items():
                val = str(v) if v is not None else ""
                run.text = run.text.replace(f"[{k}]", val).replace(f"{{{{{k}}}}}", val)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        for k, v in ctx.items():
                            val = str(v) if v is not None else ""
                            run.text = run.text.replace(f"[{k}]", val).replace(f"{{{{{k}}}}}", val)

def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template not found: {path}")
        return None
    doc = Document(path)
    
    # NOC/DAR के लिए ऑटोमैटिक टेबल जनरेशन
    if context.get("LetterType") in ["Exam NOC", "DAR NOC"] and context.get("EmployeeData"):
        for p in doc.paragraphs:
            if "[TableGoesHere]" in p.text or "[PFNumber]" in p.text:
                p.text = "" 
                cols = 6 if context["LetterType"] == "Exam NOC" else 5
                table = doc.add_table(rows=1, cols=cols)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                if context["LetterType"] == "Exam NOC":
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text, hdr[5].text = "स.क्र.", "PF No.", "नाम", "पदनाम", "परीक्षा", "अवधि"
                else:
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text = "स.क्र.", "नाम", "पदनाम", "PF No.", "डी.ए.आर. स्थिति"
                
                for idx, emp in enumerate(context["EmployeeData"]):
                    row = table.add_row().cells
                    row[0].text = str(idx+1)
                    if context["LetterType"] == "Exam NOC":
                        row[1].text, row[2].text, row[3].text, row[4].text, row[5].text = emp["PF"], emp["Name"], emp["Desg"], emp["Exam"], emp["Term"]
                    else:
                        row[1].text, row[2].text, row[3].text, row[4].text = emp["Name"], emp["Desg"], emp["PF"], emp["Status"]
                break

    replace_placeholders(doc, context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. UI INTERFACE ---
st.set_page_config(page_title="Railway Admin Cloud Portal", layout="wide")
st.title("🚂 SSE/PW/SGAM - पूर्ण क्लाउड एडमिनिस्ट्रेशन")

if st.sidebar.text_input("Access Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    # डेटा लोड करना
    emp_df = get_cloud_data('employees')
    sf_reg_df = get_cloud_data('sf11_register')
    q_df = get_cloud_data('quarters')

    tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers", "Quarter Management"])

    if not emp_df.empty:
        emp_df['PF_Clean'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False)
        emp_df['Full_Disp'] = emp_df['PF_Clean'] + " - " + emp_df['Employee Name in Hindi'].astype(str)

        if tab == "Letter Generation":
            option = st.selectbox("पत्र चुनें", [
                "Absent Case (Duty + SF11)", 
                "SF-11 Punishment Order", 
                "Exam NOC (Multi-Employee)", 
                "DAR NOC (Multi-Employee)", 
                "PME Memo", 
                "Sick Memo", 
                "Quarter Allotment Letter"
            ])

            # --- CASE 1: ABSENT ACTION ---
            if option == "Absent Case (Duty + SF11)":
                sel = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
                row = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
                c1, c2 = st.columns(2)
                f_dt = c1.date_input("अनुपस्थिति से")
                t_dt = c2.date_input("अनुपस्थिति तक")
                memo_text = f"आप बिना किसी पूर्व सूचना के दिनांक {f_dt.strftime('%d-%m-%Y')} से {t_dt.strftime('%d-%m-%Y')} तक कुल {(t_dt-f_dt).days+1} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है।"
                
                if st.button("Generate Absent Duty & SF-11 Letters"):
                    ctx = {
                        "EmployeeName": row['Employee Name in Hindi'], "Designation": row['Designation in Hindi'],
                        "PFNumber": row['PF_Clean'], "Unit": str(row.get('UNIT / MUSTER NUMBER','')).replace('.0',''),
                        "FromDate": f_dt.strftime('%d-%m-%Y'), "ToDate": t_dt.strftime('%d-%m-%Y'),
                        "LetterDate": date.today().strftime('%d-%m-%Y'), "Memo": memo_text,
                        "LetterNo": f"SGAM/SF11/ABS/{row['PF_Clean']}", "DutyDate": (t_dt + timedelta(days=1)).strftime('%d-%m-%Y')
                    }
                    d_doc = generate_doc("Absent Duty letter temp", ctx)
                    s_doc = generate_doc("SF-11 temp", ctx)
                    st.download_button("⬇️ Duty Letter", d_doc, f"Duty_{ctx['PFNumber']}.docx")
                    st.download_button("⬇️ SF-11 Charge-sheet", s_doc, f"SF11_{ctx['PFNumber']}.docx")
                    db.collection("sf11_register").add({
                        "कर्मचारी का नाम": ctx["EmployeeName"], "पी.एफ. क्रमांक": ctx["PFNumber"],
                        "दिनांक": ctx["LetterDate"], "पत्र क्र.": ctx["LetterNo"], "आरोप का विवरण": ctx["Memo"],
                        "दण्‍डादेश क्रमांक": ""
                    })
                    st.success("रजिस्टर में एंट्री कर दी गई है।")

            # --- CASE 2: PUNISHMENT (Full Sentences) ---
            elif option == "SF-11 Punishment Order":
                pending = sf_reg_df[sf_reg_df['दण्‍डादेश क्रमांक'].isna() | (sf_reg_df['दण्‍डादेश क्रमांक'] == '')]
                if not pending.empty:
                    sel_p = st.selectbox("पेंडिंग चार्जशीट चुनें", pending['पी.एफ. क्रमांक'].astype(str) + " - " + pending['कर्मचारी का नाम'])
                    sf_row = pending[(pending['पी.एफ. क्रमांक'].astype(str) + " - " + pending['कर्मचारी का नाम']) == sel_p].iloc[0]
                    
                    punishment_choice = st.selectbox("दंड का पूर्ण विवरण चुनें", [
                        "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                        "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                        "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                        "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                        "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                        "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता।"
                    ])
                    
                    if st.button("Generate Punishment Order"):
                        ctx = {
                            "EmployeeName": sf_row['कर्मचारी का नाम'], "PFNumber": sf_row['पी.एफ. क्रमांक'],
                            "Memo": punishment_choice, "Dandadesh": f"{sf_row['पत्र क्र.']}/D-1",
                            "LetterDate": date.today().strftime('%d-%m-%Y'), "SF-11Date": sf_row['दिनांक']
                        }
                        doc = generate_doc("SF-11 Punishment order temp", ctx)
                        st.download_button("⬇️ Download Order", doc, f"Order_{ctx['PFNumber']}.docx")
                        db.collection("sf11_register").document(sf_row['doc_id']).update({
                            "दण्‍डादेश क्रमांक": ctx["Dandadesh"], "दण्‍ड का विवरण": punishment_choice, "रिमार्क": "Final Order Issued"
                        })
                        st.rerun()
                else: st.info("कोई पेंडिंग केस नहीं मिला।")

            # --- CASE 3: EXAM NOC (Multi) ---
            elif option == "Exam NOC (Multi-Employee)":
                selected = st.multiselect("कर्मचारियों को चुनें", emp_df['Full_Disp'])
                if selected:
                    noc_data = []
                    for name in selected:
                        r = emp_df[emp_df['Full_Disp'] == name].iloc[0]
                        ex = st.text_input(f"Exam Name for {r['PF_Clean']}", key=f"ex_{r['PF_Clean']}")
                        term = st.text_input(f"NOC Term", value="2025-26", key=f"tm_{r['PF_Clean']}")
                        noc_data.append({"PF": r['PF_Clean'], "Name": r['Employee Name in Hindi'], "Desg": r['Designation in Hindi'], "Exam": ex, "Term": term})
                    
                    if st.button("Generate Multi-Employee NOC"):
                        ctx = {"LetterType": "Exam NOC", "EmployeeData": noc_data, "LetterDate": date.today().strftime('%d-%m-%Y')}
                        doc = generate_doc("Exam NOC Letter temp", ctx)
                        st.download_button("⬇️ Download Multi-NOC", doc, "Exam_NOC_List.docx")

            # --- CASE 4: PME MEMO (Full Calculation) ---
            elif option == "PME Memo":
                sel = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
                r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
                dob = parse_date_safe(r.get('DOB'))
                doa = parse_date_safe(r.get('DOA'))
                age = relativedelta(date.today(), dob).years if dob else "N/A"
                service = relativedelta(date.today(), doa).years if doa else "N/A"
                
                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'], "PFNumber": r['PF_Clean'],
                    "Age": age, "ServiceYears": service, "Designation": r['Designation in Hindi'],
                    "MedicalCategory": r.get('Medical category', 'A3'), "LetterDate": date.today().strftime('%d-%m-%Y')
                }
                if st.button("Generate PME Memo"):
                    doc = generate_doc("pme_memo_temp", ctx)
                    st.download_button("⬇️ Download PME Memo", doc, f"PME_{ctx['PFNumber']}.docx")

        elif tab == "Digital Registers":
            st.subheader("📊 लाइव डिजिटल रजिस्टर्स (Cloud)")
            reg_view = st.selectbox("रजिस्टर चुनें", ["SF-11 Register", "Quarter Register"])
            data = sf_reg_df if reg_view == "SF-11 Register" else q_df
            st.dataframe(data.astype(str), use_container_width=True)

        elif tab == "Quarter Management":
            st.subheader("🏠 क्वार्टर अलॉटमेंट")
            vacant = q_df[q_df['STATUS'] == 'VACANT']
            if not vacant.empty:
                sel_q = st.selectbox("खाली क्वार्टर चुनें", vacant['STATION'] + " - " + vacant['QUARTER NO.'])
                sel_e = st.selectbox("कर्मचारी को असाइन करें", emp_df['Full_Disp'])
                if st.button("Allot & Generate Letter"):
                    q_row = q_df[(q_df['STATION'] + " - " + q_df['QUARTER NO.']) == sel_q].iloc[0]
                    e_row = emp_df[emp_df['Full_Disp'] == sel_e].iloc[0]
                    ctx = {
                        "EmployeeName": e_row['Employee Name in Hindi'], "PFNumber": e_row['PF_Clean'],
                        "QuarterNo": q_row['QUARTER NO.'], "Station": q_row['STATION'],
                        "LetterDate": date.today().strftime('%d-%m-%Y')
                    }
                    doc = generate_doc("Quarter Allotment temp", ctx)
                    st.download_button("⬇️ Download Allotment", doc, "Allotment.docx")
                    db.collection("quarters").document(q_row['doc_id']).update({
                        "STATUS": "OCCUPIED", "PF No.": e_row['PF_Clean'], "EMPLOYEE NAME": e_row['Employee Name in Hindi']
                    })
                    st.success("क्लाउड डेटा अपडेट हो गया!")
                    st.rerun()
            else: st.info("कोई क्वार्टर खाली नहीं है।")
else:
    st.warning("कृपया एक्सेस के लिए पासवर्ड डालें।")
