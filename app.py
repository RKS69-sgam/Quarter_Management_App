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

def generate_doc(template_name, context):
    path = f"assets/{template_name}.docx"
    if not os.path.exists(path): 
        st.error(f"Template missing: {path}")
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

# --- 3. INTERFACE ---
st.set_page_config(page_title="Railway Admin Pro", layout="wide")

# Sidebar Password
if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    # Load Employee Master Data
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
        st.subheader("📝 अनुपस्थिति प्रकरण (SF-11)")
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
                
                if d_doc: st.download_button("⬇️ Download Duty Letter", d_doc, f"Duty_{ctx['PFNumber']}.docx")
                if s_doc: st.download_button("⬇️ Download SF-11", s_doc, f"SF11_{ctx['PFNumber']}.docx")
                
                db.collection("sf11_register").add({**ctx, "status": "Issued", "timestamp": datetime.now()})
                log_activity("ABSENT_CASE", ctx['PFNumber'])

    # --- TAB 2: OTHER SF-11 & PUNISHMENT ORDER ---
    elif tab == "Other SF-11/Order":
        mode = st.radio("प्रकार चुनें", ["नया SF-11 जारी करें", "दण्‍डादेश (Punishment Order)"])
        
        if mode == "नया SF-11 जारी करें":
            sel = st.selectbox("कर्मचारी", emp_df['PF No.'] + " - " + emp_df['Employee Name in Hindi'])
            r = emp_df.iloc[0] # Simplification for logic
            user_memo = st.text_area("आरोप का विवरण लिखें")
            if st.button("Generate SF-11"):
                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'],
                    "PFNumber": str(r['PF No.']).strip(),
                    "LetterDate": date.today().strftime('%d-%m-%Y'),
                    "Memo": user_memo + LEGAL_ENDING,
                    "LetterNo": f"SGAM/SF-11/OTH/{r['PF No.']}"
                }
                doc = generate_doc("SF-11 temp", ctx)
                st.download_button("Download SF-11", doc, "SF11_Custom.docx")
                db.collection("sf11_register").add({**ctx, "status": "Issued"})

        elif mode == "दण्‍डादेश (Punishment Order)":
            st.info("यह सेक्शन रजिस्टर से डेटा उठाकर दंडादेश बनाता है।")
            reg_stream = db.collection("sf11_register").stream()
            reg_list = [d.to_dict() for d in reg_stream]
            
            if reg_list:
                reg_df = pd.DataFrame(reg_list)
                sel_pf = st.selectbox("पेंडिंग SF-11 चुनें (PF No)", reg_df['PFNumber'].unique())
                case = reg_df[reg_df['PFNumber'] == sel_pf].iloc[-1] # नवीनतम केस
                
                order_no = st.text_input("दण्‍डादेश क्रमांक (Order No)")
                punishment = st.text_area("दण्‍ड का विवरण (Punishment Details)")
                
                if st.button("Generate Order"):
                    ctx = {
                        **case,
                        "OrderNo": order_no,
                        "PunishmentDetails": punishment,
                        "OrderDate": date.today().strftime('%d-%m-%Y')
                    }
                    doc = generate_doc("Order temp", ctx)
                    st.download_button("Download Order", doc, f"Order_{sel_pf}.docx")

    # --- TAB 3: REGISTER & EXCEL IMPORT ---
    elif tab == "SF-11 Register & Import":
        st.subheader("📊 SF-11 रजिस्टर")
        
        # Excel Import Section
        with st.expander("📥 Excel से पुराना डेटा इंपोर्ट करें"):
            file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
            if file:
                df_imp = pd.read_excel(file, dtype=str) # Alphanumeric support
                st.dataframe(df_imp.head())
                if st.button("Confirm Upload to Firebase"):
                    for _, row in df_imp.iterrows():
                        data = {
                            "PFNumber": str(row.get('पी.एफ. क्रमांक', '')).strip(),
                            "EmployeeName": str(row.get('कर्मचारी का नाम', '')).strip(),
                            "Designation": str(row.get('पदनाम', '')).strip(),
                            "LetterNo": str(row.get('पत्र क्र.', '')).strip(),
                            "LetterDate": str(row.get('दिनांक', '')).strip(),
                            "Memo": str(row.get('आरोप का विवरण', '')).strip(),
                            "OrderNo": str(row.get('दण्‍डादेश क्रमांक', '')).strip(),
                            "PunishmentDetails": str(row.get('दण्‍ड का विवरण', '')).strip(),
                            "status": "Imported"
                        }
                        db.collection("sf11_register").add(data)
                    st.success("Data Imported!")

        # Display Register
        reg_data = [d.to_dict() for d in db.collection("sf11_register").stream()]
        if reg_data:
            st.dataframe(pd.DataFrame(reg_data))
        else:
            st.warning("कोई रिकॉर्ड नहीं मिला।")

else:
    st.title("🚂 Railway Admin Pro")
    st.info("कृपया आगे बढ़ने के लिए साइडबार में पासवर्ड दर्ज करें।")
