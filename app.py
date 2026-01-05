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
            text = text.replace(f"[{k}]", str(v if v is not None else ""))
            text = text.replace(f"{{{{ {k} }}}}", str(v if v is not None else ""))
            text = text.replace(f"{{{{{k}}}}}", str(v if v is not None else ""))
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
        emp_df['Display'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False) + " - " + emp_df['Employee Name in Hindi'].astype(str)
        
        tab = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers", "Quarter Management"])

        if tab == "Letter Generation":
            letter_opt = st.selectbox("पत्र का प्रकार चुनें", [
                "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp", 
                "SICK MEMO temp.", "Exam NOC Letter temp", "pme_memo_temp"
            ])

            # --- CASE 1: SF-11 PUNISHMENT ORDER ---
            if "Punishment order" in letter_opt:
                if not sf_reg_df.empty:
                    # Clean and format for display
                    sf_reg_df = sf_reg_df.fillna('')
                    sf_reg_df['पी.एफ. क्रमांक'] = sf_reg_df['पी.एफ. क्रमांक'].astype(str).str.replace('.0', '', regex=False)
                    sf_reg_df['SF_Display'] = sf_reg_df['पी.एफ. क्रमांक'] + " | " + sf_reg_df['पत्र क्र.'].astype(str) + " (" + sf_reg_df['दिनांक'].astype(str) + ")"
                    
                    selected_sf = st.selectbox("रजिस्टर से चार्जशीट चुनें (Auto-fill)", sf_reg_df['SF_Display'])
                    sf_row = sf_reg_df[sf_reg_df['SF_Display'] == selected_sf].iloc[0]
                    
                    ctx = {
                        "EmployeeName": sf_row['कर्मचारी का नाम'],
                        "Designation": sf_row['पदनाम'],
                        "PFNumber": sf_row['पी.एफ. क्रमांक'],
                        "LetterNo.": sf_row['पत्र क्र.'],
                        "SF-11Date": sf_row['दिनांक'],
                        "LetterDate": st.date_input("दण्डादेश जारी करने की दिनांक").strftime("%d-%m-%Y"),
                        "Dandadesh": st.text_input("दण्डादेश क्रमांक", value="SGAM/SF-11/Order/"),
                        "Memo": st.text_area("दण्ड का विवरण"),
                        "Unit": "SGAM"
                    }

                    if st.button("Generate Punishment Order & Update"):
                        p_doc = generate_doc(letter_opt, ctx)
                        if p_doc:
                            st.download_button("⬇️ Download Document", p_doc, file_name=f"Punishment_{ctx['PFNumber']}.docx")
                            
                            # CRITICAL FIX: Convert Pandas Row to pure Python Dictionary and handle NaT/NaN
                            entry_dict = sf_row.to_dict()
                            # Clean dictionary for Firestore
                            clean_entry = {k: (str(v) if pd.notna(v) else "") for k, v in entry_dict.items()}
                            clean_entry.update({
                                "दण्‍डादेश क्रमांक": ctx["Dandadesh"],
                                "दण्‍डादेश जारी करने का दिनांक": ctx["LetterDate"],
                                "दण्‍ड का विवरण": ctx["Memo"],
                                "रिमार्क": "Punishment Issued"
                            })
                            # Delete temporary helper key before saving
                            if "SF_Display" in clean_entry: del clean_entry["SF_Display"]
                            
                            db.collection("sf11_register").add(clean_entry)
                            st.success("रजिस्टर में दण्डादेश की जानकारी अपडेट कर दी गई है।")
                else:
                    st.warning("रजिस्टर खाली है।")

            # --- CASE 2: ABSENT CASE (DUTY + SF-11) ---
            elif "Absent" in letter_opt:
                sel_name = st.selectbox("कर्मचारी चुनें", emp_df['Display'])
                row = emp_df[emp_df['Display'] == sel_name].iloc[0]
                
                c1, c2 = st.columns(2)
                f_dt = c1.date_input("अनुपस्थिति से", value=date.today() - timedelta(days=6))
                t_dt = c2.date_input("अनुपस्थिति तक", value=date.today())
                total_days = (t_dt - f_dt).days + 1
                
                ctx = {
                    "EmployeeName": row['Employee Name in Hindi'],
                    "Designation": row['Designation in Hindi'],
                    "PFNumber": str(row['PF No.']).replace('.0', ''),
                    "FromDate": f_dt.strftime("%d-%m-%Y"),
                    "ToDate": t_dt.strftime("%d-%m-%Y"),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime("%d-%m-%Y"),
                    "LetterDate": date.today().strftime("%d-%m-%Y"),
                    "ShortName": "STF", "Unit": "SGAM"
                }

                gen_sf11 = st.checkbox("SF-11 भी जनरेट करें?", value=True)
                if gen_sf11:
                    ctx["Memo"] = st.text_area("SF-11 आरोप", value=f"आप दिनांक {ctx['FromDate']} से {ctx['ToDate']} तक कुल {total_days} दिवस अनुपस्थित थे...")
                    ctx["LetterNo."] = f"SGAM/SF11/ABS/{ctx['PFNumber']}"

                if st.button("Generate Documents"):
                    d_doc = generate_doc("Absent Duty letter temp", ctx)
                    if d_doc: st.download_button("⬇️ Download Duty Letter", d_doc, file_name="Duty.docx")
                    
                    if gen_sf11:
                        s_doc = generate_doc("SF-11 temp", ctx)
                        if s_doc:
                            st.download_button("⬇️ Download SF-11", s_doc, file_name="SF11.docx")
                            db.collection("sf11_register").add({
                                "स.क्र.": datetime.now().strftime("%Y%m%d%H%M"),
                                "पी.एफ. क्रमांक": ctx["PFNumber"],
                                "कर्मचारी का नाम": ctx["EmployeeName"],
                                "दिनांक": ctx["LetterDate"],
                                "पत्र क्र.": ctx["LetterNo."],
                                "आरोप का विवरण": ctx["Memo"],
                                "रिमार्क": "Absent Action"
                            })
                            st.success("रजिस्टर अपडेट किया गया।")

        elif tab == "Digital Registers":
            st.subheader("📊 SF-11 मास्टर रजिस्टर")
            sf_logs = get_cloud_data("sf11_register")
            if not sf_logs.empty:
                sf_logs = sf_logs.fillna('').astype(str)
                st.dataframe(sf_logs, use_container_width=True)
else:
    st.warning("Side menu में पासवर्ड दर्ज करें।")
