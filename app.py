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

# --- 2. LOGGING UTILITY (Activity Tracking) ---
def log_activity(user_action, details):
    """हर गतिविधि को टाइमस्टैम्प के साथ रिपोर्ट टैब के लिए सेव करता है।"""
    db.collection("activity_reports").add({
        "timestamp": datetime.now(),
        "action": user_action,
        "details": details
    })

# --- 3. WORD ENGINE (Placeholder Fix) ---
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

# --- 4. MAIN INTERFACE ---
st.set_page_config(page_title="Railway Admin Portal", layout="wide")
st.title("🚂 SSE/PW/SGAM - डिजिटल एडमिनिस्ट्रेशन")

if st.sidebar.text_input("Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = db.collection('employees').stream()
    employees = [doc.to_dict() for doc in emp_df]
    emp_df = pd.DataFrame(employees) if employees else pd.DataFrame()

    if not emp_df.empty:
        emp_df['PF_Clean'] = emp_df['PF No.'].astype(str).str.replace('.0', '', regex=False)
        emp_df['Full_Disp'] = emp_df['PF_Clean'] + " - " + emp_df['Employee Name in Hindi'].astype(str)

        tab = st.sidebar.radio("Navigation", [
            "Absent Case (Duty+SF11)", 
            "Other SF-11/Order", 
            "Appeal Management",
            "SF-11 Register", 
            "Activity Reports"
        ])

        # --- TAB 1: ABSENT CASE ---
        if tab == "Absent Case (Duty+SF11)":
            st.subheader("📝 अनुपस्थिति प्रकरण (Duty Letter + SF-11)")
            sel = st.selectbox("कर्मचारी चुनें", emp_df['Full_Disp'])
            r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
            c1, c2 = st.columns(2)
            f_dt, t_dt = c1.date_input("Absent From"), c2.date_input("To")
            
            if st.button("Generate & Update Register"):
                ctx = {
                    "EmployeeName": r['Employee Name in Hindi'], "Designation": r['Designation in Hindi'],
                    "PFNumber": r['PF_Clean'], "Unit": str(r.get('UNIT / MUSTER NUMBER','')).replace('.0',''),
                    "FromDate": f_dt.strftime('%d-%m-%Y'), "ToDate": t_dt.strftime('%d-%m-%Y'),
                    "DutyDate": (t_dt + timedelta(days=1)).strftime('%d-%m-%Y'),
                    "LetterDate": date.today().strftime('%d-%m-%Y'), "LetterNo": f"SGAM/SF11/{r['PF_Clean']}",
                    "Memo": f"आप दिनांक {f_dt} से {t_dt} तक बिना सूचना अनुपस्थित थे।"
                }
                d_doc, s_doc = generate_doc("Absent Duty letter temp", ctx), generate_doc("SF-11 temp", ctx)
                st.download_button("⬇️ Duty Letter", d_doc, "Duty.docx")
                st.download_button("⬇️ SF-11", s_doc, "SF11.docx")
                db.collection("sf11_register").add({
                    "कर्मचारी का नाम": ctx["EmployeeName"], "पी.एफ. क्रमांक": ctx["PFNumber"],
                    "दिनांक": ctx["LetterDate"], "पत्र क्र.": ctx["LetterNo"], "आरोप": ctx["Memo"], "स्थिति": "जारी"
                })
                log_activity("ABSENT CASE GENERATED", f"PF: {ctx['PFNumber']}, Name: {ctx['EmployeeName']}")

        # --- TAB 2: OTHER SF-11 & ORDERS ---
        elif tab == "Other SF-11/Order":
            mode = st.radio("चुनें", ["अन्य आरोप हेतु SF-11", "दण्‍डादेश (Punishment Order)"])
            
            if mode == "अन्य आरोप हेतु SF-11":
                sel = st.selectbox("कर्मचारी", emp_df['Full_Disp'])
                r = emp_df[emp_df['Full_Disp'] == sel].iloc[0]
                memo = st.text_area("आरोप का विवरण (पूरा वाक्य लिखें)")
                if st.button("Issue SF-11"):
                    ctx = {"EmployeeName": r['Employee Name in Hindi'], "Designation": r['Designation in Hindi'], "PFNumber": r['PF_Clean'], "Unit": str(r.get('UNIT / MUSTER NUMBER','')).replace('.0',''), "LetterDate": date.today().strftime('%d-%m-%Y'), "Memo": memo, "LetterNo": f"SGAM/SF11/OTH/{r['PF_Clean']}"}
                    doc = generate_doc("SF-11 temp", ctx)
                    st.download_button("Download", doc, "SF11_Other.docx")
                    db.collection("sf11_register").add({**ctx, "स्थिति": "जारी"})
                    log_activity("OTHER SF11 ISSUED", f"PF: {ctx['PFNumber']}, Reason: {memo[:30]}...")

            elif mode == "दण्‍डादेश (Punishment Order)":
                sf_data = [d.to_dict() | {"id": d.id} for d in db.collection("sf11_register").where("स्थिति", "==", "जारी").stream()]
                if sf_data:
                    sel_sf = st.selectbox("पेंडिंग चार्जशीट चुनें", [f"{d['पी.एफ. क्रमांक']} - {d['कर्मचारी का नाम']}" for d in sf_data])
                    sf_row = next(d for d in sf_data if f"{d['पी.एफ. क्रमांक']} - {d['कर्मचारी का नाम']}" == sel_sf)
                    punishment = st.selectbox("दंड चुनें", [
                        "आगामी देय एक वर्ष की वेतन वृद्धि असंचयी प्रभाव से रोके जाने के अर्थदंड से दंडित किया जाता है।",
                        "आगामी देय एक सेट सुविधा पास तत्काल प्रभाव से रोके जाने के दंड से दंडित किया जाता है।"
                    ])
                    if st.button("Generate Punishment Order"):
                        ctx = {"EmployeeName": sf_row['कर्मचारी का नाम'], "PFNumber": sf_row['पी.एफ. क्रमांक'], "Memo": punishment, "Dandadesh": f"{sf_row['LetterNo']}/D-1", "LetterDate": date.today().strftime('%d-%m-%Y'), "SF-11Date": sf_row['दिनांक']}
                        doc = generate_doc("SF-11 Punishment order temp", ctx)
                        st.download_button("Download Order", doc, "Order.docx")
                        db.collection("sf11_register").document(sf_row['id']).update({"स्थिति": "दंडित", "दण्‍ड का विवरण": punishment})
                        log_activity("PUNISHMENT ORDER ISSUED", f"PF: {ctx['PFNumber']}, Dand: {punishment[:20]}...")
                else: st.info("कोई पेंडिंग केस नहीं है।")

        # --- TAB 3: APPEAL MANAGEMENT ---
        elif tab == "Appeal Management":
            st.subheader("⚖️ अपील प्रबंधन")
            punished = [d.to_dict() | {"id": d.id} for d in db.collection("sf11_register").where("स्थिति", "==", "दंडित").stream()]
            if punished:
                sel = st.selectbox("अपील हेतु कर्मचारी चुनें", [f"{d['पी.एफ. क्रमांक']} - {d['कर्मचारी का नाम']}" for d in punished])
                sf_row = next(d for d in punished if f"{d['पी.एफ. क्रमांक']} - {d['कर्मचारी का नाम']}" == sel)
                remark = st.text_area("अपील का विवरण/रिमार्क")
                if st.button("Update Appeal Status"):
                    db.collection("sf11_register").document(sf_row['id']).update({"स्थिति": "अपील में", "अपील रिमार्क": remark})
                    log_activity("APPEAL FILED", f"PF: {sf_row['पी.एफ. क्रमांक']}, Remark: {remark}")
                    st.success("स्थिति अपडेट की गई।")

        # --- TAB 4: SF-11 REGISTER ---
        elif tab == "SF-11 Register":
            st.subheader("📊 डिजिटल SF-11 रजिस्टर")
            regs = [d.to_dict() for d in db.collection("sf11_register").stream()]
            if regs: st.dataframe(pd.DataFrame(regs), use_container_width=True)

        # --- TAB 5: ACTIVITY REPORTS (The Heart of the System) ---
        elif tab == "Activity Reports":
            st.subheader("📑 गतिविधि रिपोर्ट (Timestamp के साथ)")
            reports = [d.to_dict() for d in db.collection("activity_reports").order_by("timestamp", direction=firestore.Query.DESCENDING).stream()]
            if reports:
                df_rep = pd.DataFrame(reports)
                df_rep['timestamp'] = pd.to_datetime(df_rep['timestamp']).dt.strftime('%Y-%m-%d %H:%M:%S')
                st.table(df_rep)
