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
        d['doc_id'] = doc.id # Store document ID for editing
        data.append(d)
    return pd.DataFrame(data) if data else pd.DataFrame()

# --- 3. DOCUMENT ENGINE ---
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
st.title("🚂 SSE/PW/SGAM - Smart Register & Admin")

if st.sidebar.text_input("Admin Password", type="password") == st.secrets.get("PASSWORD", "sgam@4321"):
    
    emp_df = get_cloud_data('employees')
    sf_reg_df = get_cloud_data('sf11_register')

    menu = st.sidebar.radio("Navigation", ["Letter Generation", "Digital Registers"])

    if menu == "Letter Generation":
        letter_opt = st.selectbox("Patra Chunen", [
            "Absent Duty letter temp", "SF-11 temp", "SF-11 Punishment order temp"
        ])

        # --- PUNISHMENT ORDER LOGIC (FILTERED) ---
        if "Punishment order" in letter_opt:
            if not sf_reg_df.empty:
                # Filter: Only show where 'दण्‍डादेश क्रमांक' is empty
                pending_sf = sf_reg_df[sf_reg_df.get('दण्‍डादेश क्रमांक', '').isna() | (sf_reg_df.get('दण्‍डादेश क्रमांक', '') == '')]
                
                if not pending_sf.empty:
                    pending_sf['PF_Clean'] = pending_sf['पी.एफ. क्रमांक'].astype(str).str.replace('.0', '', regex=False)
                    pending_sf['Display'] = pending_sf['PF_Clean'] + " | " + pending_sf['कर्मचारी का नाम'] + " (" + pending_sf['दिनांक'] + ")"
                    
                    selected = st.selectbox("Select Pending SF-11", pending_sf['Display'])
                    row = pending_sf[pending_sf['Display'] == selected].iloc[0]
                    
                    ctx = {
                        "EmployeeName": row['कर्मचारी का नाम'],
                        "PFNumber": row['PF_Clean'],
                        "LetterNo.": row['पत्र क्र.'],
                        "SF-11Date": row['दिनांक'],
                        "Designation": row.get('पदनाम', ''),
                        "LetterDate": st.date_input("Dandadesh Date").strftime("%d-%m-%Y"),
                        "Dandadesh": st.text_input("Dandadesh No.", value="SGAM/SF-11/Order/"),
                        "Memo": st.text_area("Punishment Details")
                    }

                    if st.button("Generate & Issue Punishment"):
                        p_doc = generate_doc(letter_opt, ctx)
                        if p_doc:
                            st.download_button("Download Order", p_doc, file_name=f"Order_{ctx['PFNumber']}.docx")
                            # Update existing record in Firestore
                            db.collection("sf11_register").document(row['doc_id']).update({
                                "दण्‍डादेश क्रमांक": ctx["Dandadesh"],
                                "दण्‍डादेश जारी करने का दिनांक": ctx["LetterDate"],
                                "दण्‍ड का विवरण": ctx["Memo"],
                                "रिमार्क": "Punishment Issued"
                            })
                            st.success("Punishment issued and removed from pending list!")
                            st.rerun()
                else:
                    st.info("Sabhi SF-11 ke Dandadesh issue ho chuke hain.")
            else:
                st.warning("Register khali hai.")

        # --- ABSENT CASE ---
        elif "Absent" in letter_opt:
            sel_emp = st.selectbox("Employee Chunen", emp_df['PF No.'].astype(str) + " - " + emp_df['Employee Name in Hindi'])
            e_row = emp_df.iloc[st.selectbox("Index", range(len(emp_df)), format_func=lambda x: emp_df.iloc[x]['Employee Name in Hindi'])] # Simplified for space
            
            f_dt = st.date_input("From", value=date.today() - timedelta(days=6))
            t_dt = st.date_input("To", value=date.today())
            
            ctx = {"EmployeeName": e_row['Employee Name in Hindi'], "PFNumber": str(e_row['PF No.']).replace('.0',''), "FromDate": f_dt.strftime("%d-%m-%Y"), "ToDate": t_dt.strftime("%d-%m-%Y")}
            
            if st.button("Generate Documents"):
                # Same logic as before to generate and add to register...
                pass

    elif menu == "Digital Registers":
        st.subheader("📊 SF-11 Master Register (Edit Appeal)")
        if not sf_reg_df.empty:
            sf_reg_df = sf_reg_df.fillna('')
            st.dataframe(sf_reg_df.drop(columns=['doc_id']), use_container_width=True)
            
            st.divider()
            st.write("### 📝 Edit Appeal / Remarks")
            edit_id = st.selectbox("Edit karne ke liye PF No Chunen", sf_reg_df['पी.एफ. क्रमांक'] + " - " + sf_reg_df['कर्मचारी का नाम'])
            edit_row = sf_reg_df[sf_reg_df['पी.एफ. क्रमांक'] + " - " + sf_reg_df['कर्मचारी का नाम'] == edit_id].iloc[0]
            
            col1, col2 = st.columns(2)
            a_date = col1.text_input("Appeal ka Dinank", value=edit_row.get('अपील का दिनांक', ''))
            a_res = col2.text_area("Appeal Nirnay", value=edit_row.get('अपील निर्णय', ''))
            
            if st.button("Save Changes"):
                db.collection("sf11_register").document(edit_row['doc_id']).update({
                    "अपील का दिनांक": a_date,
                    "अपील निर्णय": a_res,
                    "रिमार्क": "Updated after Appeal"
                })
                st.success("Record Updated!")
                st.rerun()
else:
    st.warning("Enter Password")
