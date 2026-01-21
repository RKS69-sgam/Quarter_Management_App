import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 0. CONFIG & AUTHENTICATION ---
st.set_page_config(page_title="Railway Promotion System", layout="wide")

if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login_form"):
        u = st.text_input("User")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Invalid Credentials")
    st.stop()

# --- 1. FIREBASE CONNECTION ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        try:
            # Pehle Streamlit Secrets check karein
            if "firebase_config" in st.secrets:
                cred_dict = dict(st.secrets["firebase_config"])
                cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_dict)
            else:
                # Local testing ke liye JSON file
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Init Error: {e}")
            st.stop()
    return firestore.client()

db = init_db()

# --- 2. PAY MATRIX & LOGIC ---
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600]
}

def get_next_increment_date(promo_date):
    # Rule: <= 30 June -> Jan | > 30 June -> July
    return f"01/01/{promo_date.year + 1}" if promo_date.month <= 6 else f"01/07/{promo_date.year + 1}"

def find_cell_in_level(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val:
            idx = cells.index(val) + 1
            # Next Index Basic
            next_val = cells[cells.index(val)+1] if (cells.index(val)+1) < len(cells) else val
            return val, idx, next_val
    return target_val, 1, target_val

# --- 3. UI & PROCESSING ---
tab1, tab2 = st.tabs(["🚀 Promotion Process", "📜 History Report"])

emp_docs = db.collection("employees").stream()
df_emp = pd.DataFrame([{**d.to_dict(), 'id': d.id} for d in emp_docs])

with tab1:
    if not df_emp.empty:
        selected_pf = st.selectbox("Select PF Number", df_emp['PF Number'].unique())
        emp_data = df_emp[df_emp['PF Number'] == selected_pf].iloc[0]

        with st.form("promo_form"):
            st.subheader(f"Employee: {emp_data['Employee Name']}")
            c1, c2, c3 = st.columns(3)
            
            # Form Fields
            old_basic = c1.number_input("Old Basic Pay (Editable)", value=float(emp_data.get('Basic Pay', 0)))
            promo_date = c2.date_input("Promotion Date", value=datetime.now())
            order_no = c3.text_input("Order Number")
            
            c4, c5 = st.columns(2)
            new_desig = c4.text_input("New Designation", value=emp_data.get('Designation', ''))
            target_lvl = c5.selectbox("Select New Level", list(PAY_MATRIX.keys()), index=1)

            # Calculation Logic
            # Notional Increment 3% rounded to nearest 100
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, new_idx, next_incr_pay = find_cell_in_level(target_lvl, notional)
            next_date = get_next_increment_date(promo_date)

            st.info(f"Summary: New Basic ₹{final_basic} | Next Increment Date: {next_date}")

            if st.form_submit_button("Generate & Update Database"):
                # 1. Update Employee Record
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic,
                    "Level": target_lvl,
                    "Designation": new_desig,
                    "LastPromotionDate": str(promo_date)
                })
                
                # 2. History Entry
                db.collection("promotion_history").add({
                    "PF": selected_pf, "Name": emp_data['Employee Name'],
                    "OldBasic": old_basic, "NewBasic": final_basic,
                    "Date": str(promo_date), "Timestamp": datetime.now()
                })
                
                st.success("Record updated successfully!")
                st.rerun()

with tab2:
    st.subheader("Promotion Logs")
    history = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    st.dataframe(pd.DataFrame([d.to_dict() for d in history]), use_container_width=True)
