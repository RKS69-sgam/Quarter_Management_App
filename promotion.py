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
st.set_page_config(page_title="Railway Promotion System", layout="wide")

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
                st.error("Ghalat user ya password")
    st.stop()

# =================================================================
# --- 1. PAY MATRIX & FIREBASE ---
# =================================================================
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
    try:
        cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    except:
        pass
db = firestore.client()

# --- Helpers ---
def get_next_increment_date(promo_date):
    if promo_date.month <= 6:
        return f"01/01/{promo_date.year + 1}"
    else:
        return f"01/07/{promo_date.year + 1}"

def get_nearest_100(value):
    return math.ceil(value / 100) * 100

def find_cell_in_level(level, target_value):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_value:
            index = cells.index(val) + 1
            # [span_0](start_span)Next index basic for next increment[span_0](end_span)
            next_val = cells[cells.index(val) + 1] if (cells.index(val) + 1) < len(cells) else val
            return val, index, next_val
    return target_value, 1, target_value

# =================================================================
# --- 2. MAIN APP ---
# =================================================================
tab1, tab2 = st.tabs(["🚀 Promotion Entry", "📜 History Report"])

emp_docs = db.collection("employees").stream()
emp_list = [ {**d.to_dict(), 'id': d.id} for d in emp_docs ]
df_emp = pd.DataFrame(emp_list)

with tab1:
    st.header("Promotion Processing (Secure Mode)")
    
    if not df_emp.empty:
        selected_pf = st.selectbox("Select Employee PF Number", df_emp['PF Number'].unique())
        emp_data = df_emp[df_emp['PF Number'] == selected_pf].iloc[0]
        
        # Initial values for auto-fill
        initial_basic = float(emp_data.get('Basic Pay', 0))
        initial_level = str(emp_data.get('Level', '1'))
        
        with st.form("promo_form"):
            st.subheader(f"Processing: {emp_data['Employee Name']}")
            c1, c2, c3 = st.columns(3)
            
            # Editable Old Basic Pay as per request
            [span_1](start_span)old_basic = c1.number_input("Old Basic Pay (Editable)", value=initial_basic)[span_1](end_span)
            [span_2](start_span)[span_3](start_span)promo_date = c2.date_input("Promotion Date", value=datetime.now())[span_2](end_span)[span_3](end_span)
            [span_4](start_span)order_no = c3.text_input("Promotion Order Number")[span_4](end_span)
            
            c4, c5 = st.columns(2)
            [span_5](start_span)new_desig = c4.text_input("New Designation", value=emp_data.get('Designation', ''))[span_5](end_span)
            [span_6](start_span)[span_7](start_span)new_level = c5.selectbox("Target Promotion Level", list(PAY_MATRIX.keys()), index=int(initial_level)-1 if initial_level.isdigit() else 0)[span_6](end_span)[span_7](end_span)

            # [span_8](start_span)Calculations Logic[span_8](end_span)
            # 1. 3% Notional Increment rounded to nearest 100
            notional_pay = get_nearest_100(old_basic * 1.03)
            # 2. Find matching basic in new level
            final_basic, new_index, next_incr_basic = find_cell_in_level(new_level, notional_pay)
            [span_9](start_span)next_incr_date = get_next_increment_date(promo_date)[span_9](end_span)
            
            st.divider()
            sc1, sc2, sc3 = st.columns(3)
            sc1.write(f"**Calculated New Basic:** ₹{final_basic}")
            sc2.write(f"**Index in Level {new_level}:** {new_index}")
            sc3.write(f"**Next Increment Date:** {next_incr_date}")

            if st.form_submit_button("Update Records & Generate"):
                # 1. Update Employee Collection
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic,
                    "Level": new_level,
                    "Designation": new_desig,
                    "Last Promotion Date": str(promo_date)
                })
                
                # 2. Save to History
                db.collection("promotion_history").add({
                    "PF Number": selected_pf,
                    "Name": emp_data['Employee Name'],
                    "Old Basic": old_basic,
                    "New Basic": final_basic,
                    "Old Level": initial_level,
                    "New Level": new_level,
                    "Order Number": order_no,
                    "Promotion Date": str(promo_date),
                    "Timestamp": datetime.now()
                })
                
                st.success("Record update kar diya gaya hai!")
                st.rerun()

with tab2:
    st.header("Promotion History Report")
    hist_docs = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    hist_data = [d.to_dict() for d in hist_docs]
    if hist_data:
        st.dataframe(pd.DataFrame(hist_data), use_container_width=True)
    else:
        st.info("Koi history nahi mili.")

