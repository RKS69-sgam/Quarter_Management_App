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
def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template missing: {path}")
        return None
    
    doc = Document(path)
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
    sf_reg_df = get_cloud_data('sf11_register')

    if not emp_df.empty:
        emp_df['Display'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers", "Quarter Management"])

        if tab == "Letter Generation":
            letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "pme_memo_temp"
            ])

            # --- CASE 1: SF-11 PUNISHMENT ORDER (Data from Register) ---
            if "Punishment order" in letter_opt:
                if not sf_reg_df.empty:
                    # Filter for display
                    sf_reg_df['SF_Display'] = sf_reg_df['पी.एफ. क्रमांक'] + " | " + sf_reg_df['पत्र क्र.'] + " (" + sf_reg_df['दिनांक'] + ")"
                    selected_sf = st.selectbox("रजिस्टर से चार्जशीट चुनें (Auto-fill)", sf_reg_df['SF_Display'])
                    sf_row = sf_reg_df[sf_reg_df['SF_Display'] == selected_sf].iloc[0]
                    
                    ctx = {
                        "EmployeeName": sf_row['कर्मचारी का नाम'],
                        "Designation": sf_row['पदनाम'],
                        "PFNumber": sf_row['पी.एफ. क्रमांक'],
                        "LetterNo.": sf_row['पत्र क्र.'],
                        "SF-11Date": sf_row['दिनांक'],
                        "Unit": "SGAM",
                        "LetterDate": st.date_input("दण्डादेश जारी करने की दिनांक").strftime("%d-%m-%Y"),
                        "Dandadesh": st.text_input("दण्डादेश क्रमांक (Order No.)", value="SGAM/SF-11/Order/"),
                        "Memo": st.text_area("दण्ड का विवरण (Punishment Description)")
                    }

                    if st.button("Generate Punishment Order & Update Register"):
                        p_doc = generate_doc(letter_opt, ctx)
                        if p_doc:
                            st.download_button("⬇️ Download Punishment Order", p_doc, file_name=f"Punishment_{ctx['PFNumber']}.docx")
                            # Sync to existing register format
                            db.collection("sf11_register").add({
                                **sf_row.to_dict(),
                                "दण्‍डादेश क्रमांक": ctx["Dandadesh"],
                                "दण्‍डादेश जारी करने का दिनांक": ctx["LetterDate"],
                                "दण्‍ड का विवरण": ctx["Memo"],
                                "रिमार्क": "Punishment Issued"
                            })
                else:
                    st.warning("SF-11 रजिस्टर में कोई पिछला रिकॉर्ड नहीं मिला।")

            # --- CASE 2: ABSENT DUTY + SF-11 ---
            elif "Absent" in letter_opt:
                sel_name = st.selectbox("कर्मचारी चुनें", emp_df['Display'])
                row = emp_df[emp_df['Display'] == sel_name].iloc[0]
                
                c1, c2 = st.columns(2)
                f_dt = c1.date_input("अनुपस्थिति से", value=date.today() - timedelta(days=6))
                t_dt = c2.date_input("अनुपस्थिति तक", value=date.today())
                total_days = (t_dt - f_dt).days + 1
                
                ctx = {
                    "EmployeeName": row.get('Employee Name in Hindi', ''),
                    "Designation": row.get('Designation in Hindi', ''),
                    "PFNumber": row.get('PF No.', ''),
                    "UnitNumber": row.get('UNIT / MUSTER NUMBER', ''),
                    "FromDate": f_dt.strftime("%d-%m-%Y"),
                    "ToDate": t_dt.strftime("%d-%m-%Y"),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime("%d-%m-%Y"),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "Date": date.today().strftime("%d-%m-%Y"),
                    "ShortName": "STF", "Unit": "SGAM"
                }

                st.divider()
                gen_sf11 = st.checkbox("साथ में SF-11 चार्जशीट भी जनरेट करें?", value=True)
                if gen_sf11:
                    absent_memo = f"आप बिना किसी पूर्व सूचना के दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक कुल {total_days} दिवस कार्य से अनुपस्थित थे, जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"
                    ctx["Memo"] = st.text_area("SF-11 आरोप (Memo)", value=absent_memo, height=150)
                    ctx["LetterNo."] = f"SGAM/SF11/ABS/{ctx['PFNumber']}"

                if st.button("Generate Documents"):
                    # Generate Duty Letter
                    d_doc = generate_doc("Absent Duty letter temp", ctx)
                    if d_doc:
                        st.download_button("⬇️ Download Duty Letter", d_doc, file_name="Duty_Letter.docx")
                    
                    # Generate SF-11
                    if gen_sf11:
                        s_doc = generate_doc("SF-11 temp", ctx)
                        if s_doc:
                            st.download_button("⬇️ Download SF-11 Charge-sheet", s_doc, file_name="SF11_ChargeSheet.docx")
                            # Sync Register
                            db.collection("sf11_register").add({
                                "स.क्र.": datetime.now().strftime("%Y%m%d%H%M"),
                                "पी.एफ. क्रमांक": ctx["PFNumber"],
                                "कर्मचारी का नाम": ctx["EmployeeName"],
                                "पदनाम": ctx["Designation"],
                                "दिनांक": ctx["Date"],
                                "पत्र क्र.": ctx["LetterNo."],
                                "आरोप का विवरण": ctx["Memo"],
                                "रिमार्क": "Absent Case Action"
                            })
                            st.success("SF-11 Register Updated!")

            # --- CASE 3: OTHER LETTERS (PME, NOC, SICK) ---
            else:
                sel_name = st.selectbox("कर्मचारी चुनें", emp_df['Display'])
                row = emp_df[emp_df['Display'] == sel_name].iloc[0]
                # Standard generation logic... (जैसा ऊपर दिया गया है)

        elif tab == "Digital Registers":
            st.subheader("📊 SF-11 मास्टर रजिस्टर")
            sf_logs = get_cloud_data("sf11_register")
            if not sf_logs.empty:
                cols = ["स.क्र.", "पी.एफ. क्रमांक", "कर्मचारी का नाम", "पदनाम", "पत्र क्र.", "दिनांक", "आरोप का विवरण", "दण्‍डादेश क्रमांक", "दण्‍डादेश जारी करने का दिनांक", "दण्‍ड का विवरण", "रिमार्क"]
                st.dataframe(sf_logs.reindex(columns=cols), use_container_width=True)

        elif tab == "Quarter Management":
            # (क्वार्टर मैनेजमेंट का पुराना कोड यहाँ रहेगा)
            pass
else:
    st.warning("Side menu में पासवर्ड दर्ज करें।")
