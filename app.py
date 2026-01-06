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

# --- 3. DOCUMENT GENERATION ENGINE ---
def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template missing: {path}")
        return None
    
    doc = Document(path)
    def replace_tags(text, ctx):
        for k, v in ctx.items():
            val = str(v) if v is not None else ""
            # Replace all tag formats
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
st.title("🚂 SSE/PW/SGAM - एकीकृत क्लाउड प्रबंधन")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    sf_reg_df = get_cloud_data('sf11_register')

    # Ensure required columns exist to prevent KeyErrors
    required_cols = ['दण्‍डादेश क्रमांक', 'अपील का दिनांक', 'अपील निर्णय', 'रिमार्क']
    for col in required_cols:
        if not sf_reg_df.empty and col not in sf_reg_df.columns:
            sf_reg_df[col] = ""

    tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers", "Quarter Management"])

    if tab == "Letter Generation":
        letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
            "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
            "SICK MEMO temp.", "Exam NOC Letter temp", "pme_memo_temp"
        ])

        # --- CASE 1: SF-11 PUNISHMENT ORDER (Smart Filtering & Dropdown) ---
        if "Punishment order" in letter_opt:
            if not sf_reg_df.empty:
                # Filter: Only show records where Punishment is not yet issued
                pending_sf = sf_reg_df[sf_reg_df['दण्‍डादेश क्रमांक'].isna() | (sf_reg_df['दण्‍डादेश क्रमांक'].astype(str) == '')].copy()
                
                if not pending_sf.empty:
                    pending_sf['PF_Disp'] = pending_sf['पी.एफ. क्रमांक'].astype(str).str.replace('.0', '', regex=False)
                    pending_sf['List_Display'] = pending_sf['PF_Disp'] + " | " + pending_sf['कर्मचारी का नाम'].astype(str) + " (" + pending_sf['दिनांक'].astype(str) + ")"
                    
                    selected_val = st.selectbox("पेंडिंग चार्जशीट चुनें (Auto-fill)", pending_sf['List_Display'])
                    sf_row = pending_sf[pending_sf['List_Display'] == selected_val].iloc[0]
                    
                    st.divider()
                    col1, col2 = st.columns(2)
                    with col1:
                        o_date = st.date_input("दण्डादेश दिनांक (Order Date)").strftime("%d-%m-%Y")
                        o_num = st.text_input("दण्डादेश क्रमांक", value="SGAM/SF-11/Order/")
                    
                    with col2:
                        punishment_memo = st.selectbox("दंड का प्रकार (Punishment Type)", [
                            "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                            "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                            "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                            "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता।"
                        ])

                    ctx = {
                        "EmployeeName": sf_row['कर्मचारी का नाम'],
                        "Designation": sf_row.get('पदनाम', ''),
                        "PFNumber": sf_row['PF_Disp'],
                        "LetterNo.": sf_row['पत्र क्र.'],
                        "SF-11Date": sf_row['दिनांक'],
                        "LetterDate": o_date,
                        "Dandadesh": o_num,
                        "Memo": punishment_memo,
                        "Unit": "SGAM"
                    }

                    if st.button("Generate Punishment Order & Update Cloud"):
                        p_doc = generate_doc(letter_opt, ctx)
                        if p_doc:
                            st.download_button("⬇️ Download Punishment Order", p_doc, file_name=f"Punishment_{ctx['PFNumber']}.docx")
                            # Update existing record in Firestore using its doc_id
                            db.collection("sf11_register").document(sf_row['doc_id']).update({
                                "दण्‍डादेश क्रमांक": ctx["Dandadesh"],
                                "दण्‍डादेश जारी करने का दिनांक": ctx["LetterDate"],
                                "दण्‍ड का विवरण": ctx["Memo"],
                                "रिमार्क": "Punishment Issued"
                            })
                            st.success("रजिस्टर अपडेट कर दिया गया है। कर्मचारी पेंडिंग लिस्ट से हट गया है।")
                            st.rerun()
                else:
                    st.info("सभी रिकॉर्ड्स के दण्डादेश जारी हो चुके हैं।")
            else:
                st.warning("SF-11 रजिस्टर में कोई पिछला रिकॉर्ड नहीं है।")

        # --- CASE 2: ABSENT DUTY LETTER + SF-11 ---
        elif "Absent" in letter_opt:
            emp_df['PF_Clean'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False)
            emp_df['Full_Disp'] = emp_df['PF_Clean'] + " - " + emp_df['Employee Name in Hindi']
            sel_emp = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
            e_row = emp_df[emp_df['Full_Disp'] == sel_emp].iloc[0]
            
            c1, c2 = st.columns(2)
            f_dt = c1.date_input("अनुपस्थिति से", value=date.today() - timedelta(days=6))
            t_dt = c2.date_input("अनुपस्थिति तक", value=date.today())
            total_days = (t_dt - f_dt).days + 1
            
            ctx = {
                "EmployeeName": e_row['Employee Name in Hindi'],
                "Designation": e_row['Designation in Hindi'],
                "PFNumber": e_row['PF_Clean'],
                "FromDate": f_dt.strftime("%d-%m-%Y"),
                "ToDate": t_dt.strftime("%d-%m-%Y"),
                "DutyDate": (t_dt + timedelta(days=1)).strftime("%d-%m-%Y"),
                "LetterDate": date.today().strftime("%d-%m-%Y"),
                "Date": date.today().strftime("%d-%m-%Y"),
                "UnitNumber": e_row.get('UNIT / MUSTER NUMBER', ''),
                "ShortName": "STF", "Unit": "SGAM"
            }

            st.divider()
            ctx["Memo"] = st.text_area("SF-11 आरोप (Memo)", value=f"आप बिना किसी पूर्व सूचना के दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक कुल {total_days} दिवस कार्य से अनुपस्थित थे...")
            ctx["LetterNo."] = f"SGAM/SF11/ABS/{ctx['PFNumber']}"

            if st.button("Generate Documents"):
                d_doc = generate_doc("Absent Duty letter temp", ctx)
                s_doc = generate_doc("SF-11 temp", ctx)
                
                if d_doc: st.download_button("⬇️ Download Duty Letter", d_doc, file_name="Duty_Letter.docx")
                if s_doc: 
                    st.download_button("⬇️ Download SF-11 Charge-sheet", s_doc, file_name="SF11.docx")
                    # Initial Register Entry
                    db.collection("sf11_register").add({
                        "स.क्र.": datetime.now().strftime("%Y%m%d%H%M"),
                        "पी.एफ. क्रमांक": ctx["PFNumber"],
                        "कर्मचारी का नाम": ctx["EmployeeName"],
                        "पदनाम": ctx["Designation"],
                        "दिनांक": ctx["Date"],
                        "पत्र क्र.": ctx["LetterNo."],
                        "आरोप का विवरण": ctx["Memo"],
                        "दण्‍डादेश क्रमांक": "", # Blank for pending
                        "रिमार्क": "Absent Case Action"
                    })
                    st.success("SF-11 रजिस्टर में एंट्री कर दी गई है।")

    elif tab == "Digital Registers":
        st.subheader("📊 SF-11 मास्टर रजिस्टर")
        if not sf_reg_df.empty:
            sf_reg_df = sf_reg_df.fillna('')
            st.dataframe(sf_reg_df.drop(columns=['doc_id']).astype(str), use_container_width=True)
            
            st.divider()
            st.subheader("📝 अपील एवं रिमार्क संपादित करें")
            edit_pf = st.selectbox("संशोधन के लिए कर्मचारी चुनें", sf_reg_df['पी.एफ. क्रमांक'].astype(str) + " - " + sf_reg_df['कर्मचारी का नाम'])
            edit_row = sf_reg_df[sf_reg_df['पी.एफ. क्रमांक'].astype(str) + " - " + sf_reg_df['कर्मचारी का नाम'] == edit_pf].iloc[0]
            
            c1, c2 = st.columns(2)
            ap_date = c1.text_input("अपील का दिनांक", value=edit_row.get('अपील का दिनांक', ''))
            ap_dec = c2.text_area("अपील निर्णय / रिमार्क", value=edit_row.get('अपील निर्णय', ''))
            
            if st.button("Save Appeal Details"):
                db.collection("sf11_register").document(edit_row['doc_id']).update({
                    "अपील का दिनांक": ap_date,
                    "अपील निर्णय": ap_dec,
                    "रिमार्क": "Updated after Appeal"
                })
                st.success("विवरण अपडेट कर दिया गया!")
                st.rerun()

    elif tab == "Quarter Management":
        st.info("Quarter management modules are active.")
        # (Your existing quarter logic goes here)

else:
    st.warning("कृपया पासवर्ड दर्ज करें।")
