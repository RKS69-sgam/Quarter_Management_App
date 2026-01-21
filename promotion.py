import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# =================================================================
# --- 0. CONFIG & AUTHENTICATION ---
# =================================================================
if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login"):
        u = st.text_input("User")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Invalid credentials")
    st.stop()

# =================================================================
# --- 1. FIREBASE & PAY MATRIX ---
# =================================================================
# [span_0](start_span)[span_1](start_span)Pay Matrix as per 7th CPC[span_0](end_span)[span_1](end_span)
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900]
}

if not firebase_admin._apps:
    cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
    firebase_admin.initialize_app(cred)
db = firestore.client()

# --- Utility Functions ---
def get_next_increment_date(promo_date):
    if promo_date.month <= 6:
        return f"01/01/{promo_date.year + 1}"
    else:
        return f"01/07/{promo_date.year + 1}"

def find_cell_in_level(level, target_value):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_value:
            idx = cells.index(val) + 1
            nxt_val = cells[cells.index(val) + 1] if (cells.index(val) + 1) < len(cells) else val
            return val, idx, nxt_val
    return target_value, 1, target_value

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    for p in list(doc.paragraphs):
        for k, v in data.items():
            if f"[{k}]" in p.text:
                p.text = p.text.replace(f"[{k}]", str(v))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =================================================================
# --- 2. MAIN APP ---
# =================================================================
tab1, tab2 = st.tabs(["🚀 Promotion Entry", "📜 History Report"])

# Fetch Employees
emp_docs = db.collection("employees").stream()
df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in emp_docs])

with tab1:
    if not df_emp.empty:
        selected_pf = st.selectbox("Select PF Number", df_emp['PF Number'].unique())
        emp_data = df_emp[df_emp['PF Number'] == selected_pf].iloc[0]
        
        initial_basic = float(emp_data.get('Basic Pay', 0))
        initial_level = str(emp_data.get('Level', '1'))

        with st.form("promo_form"):
            st.subheader(f"Promotion for: {emp_data['Employee Name']}")
            c1, c2, c3 = st.columns(3)
            
            # Editable Old Basic Pay
            old_basic = c1.number_input("Old Basic Pay (Editable)", value=initial_basic)
            promo_date = c2.date_input("Promotion Date")
            [span_2](start_span)order_no = c3.text_input("Order Number")[span_2](end_span)
            
            c4, c5 = st.columns(2)
            [span_3](start_span)new_desig = c4.text_input("New Designation", value=emp_data.get('Designation', ''))[span_3](end_span)
            [span_4](start_span)[span_5](start_span)target_level = c5.selectbox("Select New Level", list(PAY_MATRIX.keys()))[span_4](end_span)[span_5](end_span)

            # [span_6](start_span)Logic[span_6](end_span)
            notional_pay = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, new_idx, next_val = find_cell_in_level(target_level, notional_pay)
            [span_7](start_span)next_incr_date = get_next_increment_date(promo_date)[span_7](end_span)

            st.info(f"New Basic: ₹{final_basic} | Level: {target_level} | Next Incr: {next_incr_date}")

            if st.form_submit_button("Generate & Update"):
                # 1. Update Employee
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic,
                    "Level": target_level,
                    "Designation": new_desig
                })
                
                # 2. History
                db.collection("promotion_history").add({
                    "PF Number": selected_pf,
                    "Name": emp_data['Employee Name'],
                    "New Basic": final_basic,
                    "Promotion Date": str(promo_date),
                    "Timestamp": datetime.now()
                })

                # 3. [span_8](start_span)[span_9](start_span)Word Data Mapping[span_8](end_span)[span_9](end_span)
                word_mapping = {
                    "PFNUMBER": selected_pf,
                    "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDDESIGNATION": emp_data.get('Designation', ''),
                    "NEWDESIGNATION": new_desig,
                    "OLDBASICPAY": old_basic,
                    "NEWBASICPAY": final_basic,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                    "PROMOTIONORDERNUMBER": order_no,
                    "OLDLEVEL": initial_level,
                    "NEWLEVEL": target_level,
                    "NEXTINCRDATE": next_incr_date,
                    "STATION": emp_data.get('STATION', 'SGAM'),
                    "MROUND100OLDBASICPAY*103%": notional_pay
                }
                
                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.promo_file = generate_docx(t_path, word_mapping)
                st.success("Record Updated Successfully!")
                st.rerun()

    if 'promo_file' in st.session_state:
        st.download_button("📥 Download Promotion Memo", st.session_state.promo_file, "Promotion_Memo.docx")

with tab2:
    hist_docs = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    st.table([d.to_dict() for d in hist_docs])
