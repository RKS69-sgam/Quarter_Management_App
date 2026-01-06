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
        st.error(f"Firebase Error: {e}")

db = firestore.client()

# --- 2. LOGGING & UTILS ---
def log_activity(action, details):
    db.collection("activity_reports").add({
        "timestamp": datetime.now(),
        "action": action,
        "details": details
    })

def safe_replace(paragraph, context):
    inline = paragraph.runs
    if not inline: return
    full_text = "".join([run.text for run in inline])
    original_text = full_text
    for key, val in context.items():
        placeholder = f"[{key}]"
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, str(val) if val is not None else "")
    if full_text != original_text:
        for i in range(len(inline)): inline[i].text = ""
        inline[0].text = full_text

# --- 3. CORE LOGIC ---
def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path): return None
    doc = Document(path)
    for p in doc.paragraphs: safe_replace(p, context)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: safe_replace(p, context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. INTERFACE ---
st.set_page_config(page_title="Railway Admin Pro", layout="wide")

if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    # Load Employee Data
    emp_stream = db.collection('employees').stream()
    emp_list = [d.to_dict() for d in emp_stream]
    emp_df = pd.DataFrame(emp_list) if emp_list else pd.DataFrame()

    if not emp_df.empty:
        emp_df['PF_Clean'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False)
        emp_df['Full_Disp'] = emp_df['PF_Clean'] + " - " + emp_df['Employee Name in Hindi'].astype(str)

        tab = st.sidebar.radio("Navigation", ["Absent Case (Duty+SF11)", "Other SF-11/Order", "Appeal Management", "SF-11 Register", "Activity Reports"])

        # SF-11 के लिए प्रोफेशनल एंडिंग स्ट्रिंग
        LEGAL_ENDING = " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

        if tab == "Absent Case (Duty+SF11)":
            st.subheader("📝 अनुपस्थिति प्रकरण")
            sel = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
            r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
            
            c1, c2 = st.columns(2)
            f_dt = c1.date_input("अनुपस्थिति से")
            t_dt = c2.date_input("अनुपस्थिति तक")
            
            if st.button("Generate Documents & Update Register"):
                # Building the context
                unit_val = str(r.get('UNIT / MUSTER NUMBER', '')).replace('.0', '')
                short_name_val = f"{r.get('Short Name', '')} / {unit_val}"
                
                memo_main = f"आप दिनांक {f_dt.strftime('%d-%m-%Y')} से {t_dt.strftime('%d-%m-%Y')} तक बिना किसी पूर्व सूचना के अपने कार्य से अनुपस्थित रहे,"
                full_memo = memo_main + LEGAL_ENDING

                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'],
                    "Designation": r['Designation in Hindi'],
                    "PFNumber": r['PF_Clean'],
                    "Unit": unit_val,
                    "ShortName": short_name_val,
                    "FromDate": f_dt.strftime('%d-%m-%Y'),
                    "ToDate": t_dt.strftime('%d-%m-%Y'),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime('%d-%m-%Y'),
                    "LetterDate": date.today().strftime('%d-%m-%Y'),
                    "LetterNo": f"SGAM/SF-11/{r['PF_Clean']}",
                    "Memo": full_memo
                }

                d_doc = generate_doc("Absent Duty letter temp", ctx)
                s_doc = generate_doc("SF-11 temp", ctx)
                
                st.download_button("⬇️ Download Duty Letter", d_doc, f"Duty_{ctx['PFNumber']}.docx")
                st.download_button("⬇️ Download SF-11", s_doc, f"SF11_{ctx['PFNumber']}.docx")
                
                db.collection("sf11_register").add({
                    "कर्मचारी का नाम": ctx["EmployeeName"],
                    "पी.एफ. क्रमांक": ctx["PFNumber"],
                    "दिनांक": ctx["LetterDate"],
                    "पत्र क्र.": ctx["LetterNo"],
                    "आरोप": ctx["Memo"],
                    "स्थिति": "जारी"
                })
                log_activity("ABSENT CASE GENERATED", f"PF: {ctx['PFNumber']} - {ctx['EmployeeName']}")

        elif tab == "Other SF-11/Order":
            mode = st.radio("प्रकार चुनें", ["अन्य आरोप SF-11", "दण्‍डादेश (Order)"])
            
            if mode == "अन्य आरोप SF-11":
                sel = st.selectbox("कर्मचारी", emp_df['Full_Disp'])
                r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
                user_memo = st.text_area("आरोप लिखें (जैसे: ड्यूटी के दौरान मोबाइल का उपयोग करना, आदि)")
                
                if st.button("Generate SF-11"):
                    full_memo = user_memo + "," + LEGAL_ENDING
                    ctx = {
                        "EmployeeName": r['Employee Name in Hindi'],
                        "Designation": r['Designation in Hindi'],
                        "PFNumber": r['PF_Clean'],
                        "ShortName": f"{r.get('Short Name', '')} / {str(r.get('UNIT / MUSTER NUMBER','')).replace('.0','')}",
                        "LetterDate": date.today().strftime('%d-%m-%Y'),
                        "Memo": full_memo,
                        "LetterNo": f"SGAM/SF-11/OTH/{r['PF_Clean']}"
                    }
                    doc = generate_doc("SF-11 temp", ctx)
                    st.download_button("Download SF-11", doc, "SF11_Other.docx")
                    db.collection("sf11_register").add({**ctx, "स्थिति": "जारी"})
                    log_activity("OTHER SF11 ISSUED", f"PF: {ctx['PFNumber']}")

        # ... (Activity Reports and Register tabs remain same as previous version)

else:
    st.info("Side menu में पासवर्ड डालें।")
