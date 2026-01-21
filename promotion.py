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
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500, 24200],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300, 34300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100, 39200],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200, 47600],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400]
}

if not firebase_admin._apps:
    cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
    firebase_admin.initialize_app(cred)
db = firestore.client()

# --- Helper Functions ---
def get_next_increment_date(promo_date):
    # Logic: Before 1st July -> Jan Next Year | After 30 June -> July Next Year
    if promo_date.month <= 6:
        return f"01/01/{promo_date.year + 1}"
    else:
        return f"01/07/{promo_date.year + 1}"

def find_cell_in_level(level, target_value):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_value:
            idx = cells.index(val) + 1
            # [span_3](start_span)Next Index Basic for Increment calculation[span_3](end_span)
            nxt_val = cells[cells.index(val) + 1] if (cells.index(val) + 1) < len(cells) else val
            return val, idx, nxt_val
    return target_value, 1, target_value

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    for p in list(doc.paragraphs):
        for k, v in data.items():
            placeholder = f"[{k}]"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(v))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in data.items():
                        placeholder = f"[{k}]"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(v))
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =================================================================
# --- 2. MAIN UI ---
# =================================================================
tab1, tab2 = st.tabs(["🚀 Promotion Entry", "📜 History Report"])

emp_docs = db.collection("employees").stream()
df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in emp_docs])

with tab1:
    if not df_emp.empty:
        selected_pf = st.selectbox("Select PF Number", df_emp['PF Number'].unique())
        emp_data = df_emp[df_emp['PF Number'] == selected_pf].iloc[0]
        
        initial_basic = float(emp_data.get('Basic Pay', 0))
        initial_level = str(emp_data.get('Level', '1'))

        with st.form("promo_form"):
            st.subheader(f"Processing Promotion: {emp_data['Employee Name']}")
            c1, c2, c3 = st.columns(3)
            
            # Old Basic editable mode
            old_basic = c1.number_input("Old Basic Pay", value=initial_basic)
            promo_date = c2.date_input("Promotion Date")
            order_no = c3.text_input("Order Number")
            
            c4, c5 = st.columns(2)
            new_desig = c4.text_input("New Designation", value=emp_data.get('Designation', ''))
            target_level = c5.selectbox("Target Level", list(PAY_MATRIX.keys()))

            # --- Calculation Logic ---
            # 1. [span_4](start_span)1.03 Notional increment rounded to nearest 100[span_4](end_span)
            notional_pay = math.ceil((old_basic * 1.03) / 100) * 100
            
            # 2. [span_5](start_span)Find next higher basic in New Level[span_5](end_span)
            final_basic, new_idx, next_val = find_cell_in_level(target_level, notional_pay)
            
            # 3. [span_6](start_span)Next Increment Date[span_6](end_span)
            next_incr_date = get_next_increment_date(promo_date)

            st.info(f"Summary: New Basic ₹{final_basic} in Level {target_level}")

            if st.form_submit_button("Generate & Update"):
                # Update Employee Database
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic,
                    "Level": target_level,
                    "Designation": new_desig,
                    "LastPromotionDate": str(promo_date)
                })
                
                # Log History in Separate Collection
                db.collection("promotion_history").add({
                    "PFNUMBER": selected_pf,
                    "Name": emp_data['Employee Name'],
                    "OldBasic": old_basic,
                    "NewBasic": final_basic,
                    "PromoDate": str(promo_date),
                    "OrderNo": order_no,
                    "Timestamp": datetime.now()
                })

                # [span_7](start_span)[span_8](start_span)[span_9](start_span)Map data to Word Template placeholders[span_7](end_span)[span_8](end_span)[span_9](end_span)
                word_mapping = {
                    "PFNUMBER": selected_pf,
                    "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDDESIGNATION": emp_data.get('Designation', ''),
                    "NEWDESIGNATION": new_desig,
                    "OLDBASICPAY": f"{old_basic}/-",
                    "NEWBASICPAY": f"{final_basic}/-",
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                    "PROMOTIONORDERNUMBER": order_no,
                    "OLDLEVEL": initial_level,
                    "NEWLEVEL": target_level,
                    "NEXTINCRDATE": next_incr_date,
                    "STATION": emp_data.get('STATION', 'SGAM'),
                    "MROUND100OLDBASICPAY*103%": notional_pay,
                    "MROUND100ONEWBASICPAY*103%": math.ceil((final_basic * 1.03) / 100) * 100
                }
                
                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.promo_file = generate_docx(t_path, word_mapping)
                st.success("Successfully Processed!")
                st.rerun()

    if 'promo_file' in st.session_state:
        st.download_button("📥 Download Promotion Document", st.session_state.promo_file, f"Promotion_{selected_pf}.docx")

with tab2:
    st.subheader("Promotion & MACP Logs")
    hist = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    logs = [d.to_dict() for d in hist]
    if logs:
        st.dataframe(pd.DataFrame(logs), use_container_width=True)
