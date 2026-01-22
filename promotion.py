import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 0. DATA TABLES (7th CPC Pay Matrix) ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
}

PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600]
}

# --- 1. HELPERS ---
def clean_int(val):
    try:
        if val is None or val == "": return 0
        return int(float(str(val).strip()))
    except: return 0

def find_matrix_details(level, target_val):
    """Returns (MatchedPay, IndexRowNumber)"""
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val:
            return int(val), cells.index(val) + 1
    return int(target_val), 1

def safe_replace(doc, data):
    """Para and Table replacement logic"""
    for k, v in data.items():
        tag = f"[{k}]"
        for p in doc.paragraphs:
            if tag in p.text: p.text = p.text.replace(tag, str(v))
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        if tag in p.text: p.text = p.text.replace(tag, str(v))

# --- 2. DB INIT ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        if "firebase_config" in st.secrets:
            cred_dict = dict(st.secrets["firebase_config"])
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        else:
            cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_db()

# --- 3. UI ---
st.title("🚀 Central Railway Promotion & MACP System")

emp_docs = db.collection("employees").stream()
data_list, desigs_en = [], []
for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    pf = str(item.get('PF Number', '')).strip().split('.')[0]
    item['PF_Clean'] = pf
    data_list.append(item)
    if item.get('Designation'): desigs_en.append(item['Designation'])

df_emp = pd.DataFrame(data_list)
sorted_en = sorted(list(set(desigs_en)))

if not df_emp.empty:
    search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
    sel_emp = st.selectbox("Search Employee", search_options)
    sel_pf = sel_emp.split('(')[-1].strip(')')
    emp_data = df_emp[df_emp['PF_Clean'] == sel_pf].iloc[0]

    with st.form("promotion_form_v3"):
        # Load from Uppercase Fields
        old_basic_val = clean_int(emp_data.get('BASIC PAY', 0))
        old_lvl_val = str(clean_int(emp_data.get('PAY LEVEL', 1)))
        
        c1, c2, c3 = st.columns(3)
        old_basic = c1.number_input("Old Basic Pay", value=old_basic_val, step=1)
        promo_date = c2.date_input("Promotion/MACP Date", value=datetime.now())
        order_no = c3.text_input("Order Number (PROMOTIONORDERNUMBER)")

        st.markdown("---")
        col_old, col_new = st.columns(2)

        with col_old:
            st.info("Current Status")
            old_desig = emp_data.get('Designation', '')
            st.text_input("Old Level", old_lvl_val, disabled=True)
            st.text_input("Old Pay Band", PAY_LEVEL_MAP.get(old_lvl_val, {}).get("PB", ""), disabled=True)
            st.text_input("Old Designation", old_desig, disabled=True)
            # Fetch OLDINDEX from Matrix
            _, old_index = find_matrix_details(old_lvl_val, old_basic)
            st.text_input("Old Index", old_index, disabled=True)

        with col_new:
            st.success("New Status")
            # New Level Logic (Old+1)
            def_lvl_idx = int(old_lvl_val) if int(old_lvl_val) < 7 else int(old_lvl_val) - 1
            new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=def_lvl_idx)
            
            # Smart Designation Select (Index - 1)
            def_des_idx = (sorted_en.index(old_desig) - 1) if old_desig in sorted_en and sorted_en.index(old_desig) > 0 else 0
            new_desig = st.selectbox("New Designation", sorted_en, index=def_des_idx)
            new_pb = PAY_LEVEL_MAP.get(new_lvl, {}).get("PB", "")

        # Fixation Math
        notional_pay = math.ceil((old_basic * 1.03) / 100) * 100
        final_pay, new_index = find_matrix_details(new_lvl, notional_pay)

        st.divider()
        st.write(f"### Proposed New Basic: ₹{final_pay} (Index: {new_index})")

        if st.form_submit_button("Update Records & Export Word"):
            # Update DB
            db.collection("employees").document(emp_data['id']).update({
                "BASIC PAY": int(final_pay),
                "PAY LEVEL": new_lvl,
                "Designation": new_desig
            })
            
            # Mapping for Word Template
            mapping = {
                "PFNUMBER": sel_pf,
                "EMPLOYEENAME": emp_data['Employee Name'],
                "OLDBASICPAY": int(old_basic),
                "NEWBASICPAY": int(final_pay),
                "OLDPAYBAND": PAY_LEVEL_MAP.get(old_lvl_val, {}).get("PB", ""),
                "NEWPAYBAND": new_pb,
                "OLD DESIGNATION": old_desig,
                "NEWDESIGNATION": new_desig,
                "PROMOTIONORDERNUMBER": order_no,
                "OLDLEVEL": old_lvl_val,
                "NEWLEVEL": new_lvl,
                "OLDINDEX": old_index,
                "NEWINDEX": new_index,
                "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                "MROUND100OLDBASICPAY*103%": int(notional_pay)
            }

            t_path = os.path.join("assets", "General Promotion MACP temp.docx")
            if os.path.exists(t_path):
                doc = Document(t_path)
                safe_replace(doc, mapping)
                bio = io.BytesIO()
                doc.save(bio)
                st.session_state.memo = bio.getvalue()
                st.success("Database Updated and File Ready!")
                st.rerun()

if 'memo' in st.session_state:
    st.download_button("📥 Download Fixation Order", st.session_state.memo, f"Fixation_{sel_pf}.docx")
