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

# --- 2. DOCUMENT ENGINE ---
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

def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path):
        st.error(f"Template not found: {path}")
        return None
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

# --- 3. MAIN INTERFACE ---
st.set_page_config(page_title="Railway Admin Pro", layout="wide")

if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    # मास्टर कर्मचारी डेटा लोड करें
    emp_stream = db.collection('employees').stream()
    emp_list = [d.to_dict() for d in emp_stream]
    emp_df = pd.DataFrame(emp_list) if emp_list else pd.DataFrame()

    tab = st.sidebar.radio("Navigation", ["Absent Case (Duty+SF11)", "Other SF-11/Order", "SF-11 Register & Import"])

    LEGAL_ENDING = " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

    # --- TAB 1: ABSENT CASE ---
    if tab == "Absent Case (Duty+SF11)":
        st.subheader("📝 अनुपस्थिति प्रकरण (Absent Case)")
        if not emp_df.empty:
            emp_df['Full_Disp'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
            sel_emp = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'].unique())
            r = emp_df[emp_df['Full_Disp'] == sel_emp].iloc[0]
            
            c1, c2 = st.columns(2)
            f_dt = c1.date_input("अनुपस्थिति से")
            t_dt = c2.date_input("अनुपस्थिति तक")
            
            if st.button("Generate Documents & Save"):
                memo_main = f"आप दिनांक {f_dt.strftime('%d-%m-%Y')} से {t_dt.strftime('%d-%m-%Y')} तक बिना किसी पूर्व सूचना के अपने कार्य से अनुपस्थित रहे,"
                full_memo = memo_main + LEGAL_ENDING
                
                # Unit Column से डेटा निकालें
                unit_val = str(r.get('Unit', ''))[:2]
                
                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'],
                    "Designation": r['Designation in Hindi'],
                    "Unit": unit_val,
                    "PFNumber": str(r['PF No.']).strip(),
                    "FromDate": f_dt.strftime('%d-%m-%Y'),
                    "ToDate": t_dt.strftime('%d-%m-%Y'),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime('%d-%m-%Y'),
                    "LetterDate": date.today().strftime('%d-%m-%Y'),
                    "LetterNo": f"SGAM/SF-11/{r['PF No.']}",
                    "Memo": full_memo
                }
                d_doc = generate_doc("Absent Duty letter temp", ctx)
                s_doc = generate_doc("SF-11 temp", ctx)
                if d_doc: st.download_button("⬇️ Duty Letter", d_doc, f"Duty_{ctx['PFNumber']}.docx")
                if s_doc: st.download_button("⬇️ SF-11", s_doc, f"SF11_{ctx['PFNumber']}.docx")
                db.collection("sf11_register").add({**ctx, "status": "Issued", "timestamp": datetime.now()})

    # --- TAB 2: OTHER SF-11 & ORDER ---
    elif tab == "Other SF-11/Order":
        mode = st.radio("प्रकार", ["नया SF-11 जारी करें", "दण्‍डादेश (Punishment Order)"])
        
        if mode == "दण्‍डादेश (Punishment Order)":
            st.subheader("🔨 दण्‍डादेश (NIP) जनरेट")
            docs = db.collection("sf11_register").stream()
            reg_list = []
            for d in docs:
                item = d.to_dict(); item['doc_id'] = d.id
                if not item.get('OrderNo') or str(item.get('OrderNo')) == "":
                    reg_list.append(item)
            
            if reg_list:
                reg_df = pd.DataFrame(reg_list)
                reg_df['Select_Disp'] = reg_df['PFNumber'].astype(str) + " - " + reg_df['EmployeeName'] + " - Dt: " + reg_df['LetterDate']
                sel_text = st.selectbox("केस चुनें", reg_df['Select_Disp'].unique())
                case = reg_df[reg_df['Select_Disp'] == sel_text].iloc[0]
                
                # Unit Extraction from employees collection
                unit_extracted = ""
                if not emp_df.empty:
                    match = emp_df[emp_df['PF No.'].astype(str).str.strip() == str(case['PFNumber']).strip()]
                    if not match.empty:
                        unit_extracted = str(match.iloc[0].get('Unit', ''))[:2]
                
                c1, c2 = st.columns(2)
                dandadesh_no = c1.text_input("दण्‍डादेश क्रमांक", value=f"SGAM/NIP/{case['PFNumber']}")
                order_date = c2.date_input("दण्‍डादेश दिनांक", value=date.today())
                
                punishment_text = st.selectbox("दण्ड का विवरण", [
                    "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                    "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                    "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
                ])

                if st.button("Generate & Update Database"):
                    # Mapping logic as requested
                    ctx = {
                        "EmployeeName": case.get('EmployeeName', ''),
                        "Designation": case.get('Designation', ''),
                        "Unit": unit_extracted,
                        "Dandadesh": dandadesh_no,
                        "LetterNo.": case.get('LetterNo', ''),
                        "SF-11Date": case.get('LetterDate', ''),
                        "LetterDate": order_date.strftime('%d-%m-%Y'), # [LetterDate] me Order Date
                        "OrderDate": order_date.strftime('%d-%m-%Y'),
                        "Memo": punishment_text
                    }
                    doc_bio = generate_doc("SF-11 Punishment order temp", ctx)
                    if doc_bio:
                        db.collection("sf11_register").document(case['doc_id']).update({
                            "OrderNo": dandadesh_no, "OrderDate": order_date.strftime('%d-%m-%Y'),
                            "PunishmentDetails": punishment_text, "status": "Closed"
                        })
                        st.success(f"NIP तैयार! Unit: {unit_extracted}")
                        st.download_button("⬇️ Download NIP", doc_bio, f"NIP_{case['PFNumber']}.docx")
            else: st.warning("कोई पेंडिंग केस नहीं मिला।")

    # --- TAB 3: REGISTER ---
    elif tab == "SF-11 Register & Import":
        st.subheader("📊 रजिस्टर")
        all_reg = [d.to_dict() for d in db.collection("sf11_register").stream()]
        if all_reg: st.dataframe(pd.DataFrame(all_reg))

else:
    st.info("Side menu में पासवर्ड डालें।")
