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

# --- 2. UTILS & DOCUMENT ENGINE ---
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
    
    emp_stream = db.collection('employees').stream()
    emp_list = [d.to_dict() for d in emp_stream]
    emp_df = pd.DataFrame(emp_list) if emp_list else pd.DataFrame()

    tab = st.sidebar.radio("Navigation", [
        "Absent Case (Duty+SF11)", 
        "Other SF-11/Order", 
        "SF-11 Register & Import", 
        "Activity Reports"
    ])

    LEGAL_ENDING = " जो कि रेल सेवक होने के नाते आपकी रेल सेवा निष्ठा के प्रति घोर लापरवाही को प्रदर्शित करता है। अतः आप कामों व भूलो के फेहरिस्त धारा 1, 2 एवं 3 के उल्लंघन के दोषी पाए जाते है।"

    # --- TAB 1: ABSENT CASE ---
    if tab == "Absent Case (Duty+SF11)":
        st.subheader("📝 अनुपस्थिति प्रकरण (Absent Case)")
        if not emp_df.empty:
            emp_df['Full_Disp'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
            sel = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
            r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
            
            c1, c2 = st.columns(2)
            f_dt = c1.date_input("अनुपस्थिति से")
            t_dt = c2.date_input("अनुपस्थिति तक")
            
            if st.button("Generate Documents"):
                memo_main = f"आप दिनांक {f_dt.strftime('%d-%m-%Y')} से {t_dt.strftime('%d-%m-%Y')} तक बिना किसी पूर्व सूचना के अपने कार्य से अनुपस्थित रहे,"
                full_memo = memo_main + LEGAL_ENDING
                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'],
                    "Designation": r['Designation in Hindi'],
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
        mode = st.radio("प्रकार चुनें", ["नया SF-11 जारी करें", "दण्‍डादेश (Punishment Order)"])
        
        if mode == "नया SF-11 जारी करें":
            if not emp_df.empty:
                emp_df['Full_Disp'] = emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'].astype(str)
                sel = st.selectbox("कर्मचारी", emp_df['Full_Disp'])
                r = emp_df[emp_df['PF No.'].astype(str) == sel.split(" - ")[0]].iloc[0]
                user_memo = st.text_area("आरोप का विवरण लिखें")
                if st.button("Generate SF-11"):
                    ctx = {
                        "EmployeeName": r['Employee Name in Hindi'],
                        "Designation": r['Designation in Hindi'],
                        "PFNumber": str(r['PF No.']).strip(),
                        "LetterDate": date.today().strftime('%d-%m-%Y'),
                        "Memo": user_memo + LEGAL_ENDING,
                        "LetterNo": f"SGAM/SF-11/OTH/{r['PF No.']}"
                    }
                    doc = generate_doc("SF-11 temp", ctx)
                    if doc:
                        st.download_button("Download SF-11", doc, f"SF11_{ctx['PFNumber']}.docx")
                        db.collection("sf11_register").add({**ctx, "status": "Issued", "timestamp": datetime.now()})

        elif mode == "दण्‍डादेश (Punishment Order)":
            st.subheader("🔨 दण्‍डादेश (NIP) जनरेट और अपडेट")
            docs = db.collection("sf11_register").stream()
            reg_list = []
            for d in docs:
                item = d.to_dict(); item['doc_id'] = d.id
                if not item.get('OrderNo') or str(item.get('OrderNo')).lower() == 'nan' or item.get('OrderNo') == "":
                    reg_list.append(item)
            
            if reg_list:
                reg_df = pd.DataFrame(reg_list)
                # डिस्प्ले में Name + Date
                reg_df['Select_Disp'] = (reg_df['PFNumber'].astype(str) + " - " + 
                                       reg_df['EmployeeName'].astype(str) + " - SF11 Date: " + 
                                       reg_df['LetterDate'].astype(str))
                
                sel_text = st.selectbox("पेंडिंग केस चुनें", reg_df['Select_Disp'].unique())
                selected_pf = sel_text.split(" - ")[0]
                selected_date = sel_text.split(" - SF11 Date: ")[1]
                case = reg_df[(reg_df['PFNumber'] == selected_pf) & (reg_df['LetterDate'] == selected_date)].iloc[0]
                
                c1, c2 = st.columns(2)
                order_no = c1.text_input("दण्‍डादेश क्रमांक", value=f"SGAM/NIP/{case['PFNumber']}")
                order_date = c2.date_input("दण्‍डादेश दिनांक", value=date.today())
                
                punishment_text = st.selectbox("दण्ड का प्रकार चुनें", [
                    "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                    "आगामी देय एक वर्ष की वेतन वृद्धि संचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                    "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                    "आगामी देय एक सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                    "आगामी देय दो सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।",
                    "आगामी देय दो सेट PTO तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
                ])

                if st.button("Generate Order & Update Firebase"):
                    ctx = {**case, "OrderNo": order_no, "OrderDate": order_date.strftime('%d-%m-%Y'), "PunishmentDetails": punishment_text}
                    
                    # सुधार: यहाँ आपकी इमेज के अनुसार नया फाइल नाम इस्तेमाल किया गया है
                    doc_bio = generate_doc("SF-11 Punishment order temp", ctx)
                    
                    if doc_bio:
                        db.collection("sf11_register").document(case['doc_id']).update({
                            "OrderNo": order_no, "OrderDate": order_date.strftime('%d-%m-%Y'),
                            "PunishmentDetails": punishment_text, "status": "Closed/Punished"
                        })
                        st.success("डेटाबेस अपडेट हुआ!")
                        st.download_button("⬇️ Download Order", doc_bio, f"Order_{case['PFNumber']}.docx")
            else: st.warning("कोई पेंडिंग केस नहीं मिला।")

    # --- TAB 3: REGISTER & IMPORT ---
    elif tab == "SF-11 Register & Import":
        st.subheader("📊 रजिस्टर और डेटा मैनेजमेंट")
        with st.expander("📥 Excel से पुराना डेटा इंपोर्ट करें"):
            file = st.file_uploader("Upload Excel", type=["xlsx"])
            if file:
                try:
                    df_imp = pd.read_excel(file, dtype=str, engine='openpyxl')
                    if st.button("Confirm Bulk Upload"):
                        for _, row in df_imp.iterrows():
                            db.collection("sf11_register").add({
                                "PFNumber": str(row.get('पी.एफ. क्रमांक', '')).strip(),
                                "EmployeeName": str(row.get('कर्मचारी का नाम', '')).strip(),
                                "Designation": str(row.get('पदनाम', '')).strip(),
                                "LetterNo": str(row.get('पत्र क्र.', '')).strip(),
                                "LetterDate": str(row.get('दिनांक', '')).strip(),
                                "Memo": str(row.get('आरोप का विवरण', '')).strip(),
                                "OrderNo": str(row.get('दण्‍डादेश क्रमांक', '')).replace('nan',''),
                                "PunishmentDetails": str(row.get('दण्‍ड का विवरण', '')).replace('nan',''),
                                "status": "Imported", "timestamp": datetime.now()
                            })
                        st.success("इंपोर्ट पूरा हुआ!")
                except Exception as e: st.error(f"Error: {e}")
        
        all_reg = [d.to_dict() for d in db.collection("sf11_register").stream()]
        if all_reg: st.dataframe(pd.DataFrame(all_reg))

else:
    st.info("Side menu में पासवर्ड डालें।")
