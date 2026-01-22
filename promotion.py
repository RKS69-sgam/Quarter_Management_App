import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 1. DATA TABLES (Pay Band & GP Mapping) ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
}

# Pay Matrix Cells
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600]
}

# --- 2. AUTH & DATABASE INIT ---
if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login"):
        u, p = st.text_input("User"), st.text_input("Password", type="password")
        if st.form_submit_button("Login") and u == "admin" and p == "Sgam@4321":
            st.session_state.auth = True
            st.rerun()
    st.stop()

@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_db()

# --- 3. DATA FETCHING ---
emp_docs = db.collection("employees").stream()
data_list = []
designations_en = set()
designations_hi = set()

for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    raw_pf = str(item.get('PF Number', '')).split('.')[0].strip()
    item['PF_Clean'] = raw_pf
    data_list.append(item)
    if item.get('Designation'): designations_en.add(item['Designation'])
    if item.get('Designation in Hindi'): designations_hi.add(item['Designation in Hindi'])

df_emp = pd.DataFrame(data_list)
all_desig_en = sorted(list(designations_en))
all_desig_hi = sorted(list(designations_hi))

# --- 4. MAIN UI ---
st.header("🚀 Promotion & MACP Process")

if not df_emp.empty:
    search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
    selected_option = st.selectbox("Search Employee", search_options)
    selected_pf = selected_option.split('(')[-1].strip(')')
    emp_data = df_emp[df_emp['PF_Clean'] == selected_pf].iloc[0]

    with st.form("promotion_form"):
        # Section 1: Basic Selection
        st.subheader(f"Processing: {emp_data['Employee Name']}")
        c1, c2, c3 = st.columns(3)
        old_basic = c1.number_input("Old Basic Pay", value=float(emp_data.get('Basic Pay', 0)))
        promo_date = c2.date_input("Promotion Date")
        order_no = c3.text_input("Order Number")

        st.divider()

        # Section 2: OLD Status (Auto from Collection)
        st.subheader("Current (Old) Details")
        oc1, oc2, oc3, oc4, oc5 = st.columns(5)
        old_lvl = str(emp_data.get('Level', '1'))
        oc1.text_input("Old Level", value=old_lvl, disabled=True)
        oc2.text_input("Old Pay Band", value=PAY_LEVEL_MAP.get(old_lvl, {}).get("PB", ""), disabled=True)
        oc3.text_input("Old Grade Pay", value=PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", ""), disabled=True)
        oc4.text_input("Old Designation (EN)", value=emp_data.get('Designation', ''), disabled=True)
        oc5.text_input("Old Designation (HI)", value=emp_data.get('Designation in Hindi', ''), disabled=True)

        st.divider()

        # Section 3: NEW Status (Manual Select)
        st.subheader("New Promotion Details")
        nc1, nc2, nc3, nc4, nc5 = st.columns(5)
        new_lvl = nc1.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=int(old_lvl))
        nc2.text_input("New Pay Band", value=PAY_LEVEL_MAP.get(new_lvl, {}).get("PB", ""))
        nc3.text_input("New Grade Pay", value=PAY_LEVEL_MAP.get(new_lvl, {}).get("GP", ""))
        new_desig_en = nc4.selectbox("New Designation (EN)", all_desig_en)
        new_desig_hi = nc5.selectbox("New Designation (HI)", all_desig_hi)

        # Math Logic
        notional = math.ceil((old_basic * 1.03) / 100) * 100
        lvl_cells = PAY_MATRIX.get(new_lvl, [])
        final_basic = next((val for val in lvl_cells if val >= notional), notional)
        
        st.info(f"Target Basic Calculated: ₹{final_basic}")

        if st.form_submit_button("Update Records & Generate Memo"):
            # 1. Update Employee Collection
            db.collection("employees").document(emp_data['id']).update({
                "Basic Pay": final_basic,
                "Level": new_lvl,
                "Designation": new_desig_en,
                "Designation in Hindi": new_desig_hi
            })
            
            # 2. History Save (for Report Tab)
            db.collection("promotion_history").add({
                "PF": selected_pf, "Name": emp_data['Employee Name'],
                "OldBasic": old_basic, "NewBasic": final_basic,
                "NewLevel": new_lvl, "Timestamp": datetime.now()
            })
            st.success("Database Updated Successfully!")
