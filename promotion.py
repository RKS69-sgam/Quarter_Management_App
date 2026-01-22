import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 1. CONFIG & SECURITY ---
st.set_page_config(page_title="Railway Secure Promotion", layout="wide")

def check_password():
    """Returns True if the user had the correct password."""
    def password_entered():
        if st.session_state["username"] == "admin" and st.session_state["password"] == "sgam123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔐 SGAM Office Security")
        st.text_input("Username", on_change=None, key="username")
        st.text_input("Password", type="password", on_change=None, key="password")
        st.button("Log In", on_click=password_entered)
        return False
    elif not st.session_state["password_correct"]:
        st.error("😕 User not known or password incorrect")
        return False
    else:
        return True

# --- 2. DATA TABLES ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
}

PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500, 24200, 24900],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300, 34300, 35300],
}

# --- 3. DATABASE & HELPERS ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        # Load from secrets or local JSON
        if "firebase_config" in st.secrets:
            cred_dict = dict(st.secrets["firebase_config"])
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        else:
            cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

def clean_int(val):
    try: return int(float(str(val).strip())) if val else 0
    except: return 0

def find_matrix_details(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val: return int(val), cells.index(val) + 1
    return int(target_val), 1

def safe_replace(doc, data):
    for k, v in data.items():
        tag = f"[{k}]"
        for p in doc.paragraphs:
            if tag in p.text: p.text = p.text.replace(tag, str(v))
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if tag in p.text: p.text = p.text.replace(tag, str(v))

# --- 4. MAIN APPLICATION ---
if check_password():
    db = init_db()
    st.sidebar.success("Logged in as Admin")
    if st.sidebar.button("Log Out"):
        del st.session_state["password_correct"]
        st.rerun()

    st.title("🚀 Smart Promotion & Secure Fixation")
    
    # Fetch Data
    emp_docs = db.collection("employees").stream()
    data_list, desigs_en = [], []
    for d in emp_docs:
        item = d.to_dict(); item['id'] = d.id
        data_list.append(item)
        if item.get('Designation'): desigs_en.append(item['Designation'])

    df_emp = pd.DataFrame(data_list)
    sorted_en = sorted(list(set(desigs_en)))

    if not df_emp.empty:
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({str(r.get('PF Number','')).split('.')[0]})", axis=1).tolist()
        sel_emp = st.selectbox("Search Employee", search_options)
        emp_data = df_emp[df_emp['Employee Name'].str.contains(sel_emp.split(' (')[0])].iloc[0]
        sel_pf = str(emp_data.get('PF Number','')).split('.')[0]

        with st.form("secure_promotion_form"):
            old_basic_db = clean_int(emp_data.get('BASIC PAY', 0))
            old_lvl_db = str(clean_int(emp_data.get('PAY LEVEL', 1)))
            
            c1, c2, c3 = st.columns(3)
            old_basic = c1.number_input("Old Basic Pay", value=old_basic_val if 'old_basic_val' in locals() else old_basic_db, step=1)
            promo_date = c2.date_input("Fixation Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            col_old, col_new = st.columns(2)
            with col_old:
                st.info("Current Status")
                old_desig = emp_data.get('Designation', '')
                st.text_input("Old Designation", old_desig, disabled=True)
                old_gp = PAY_LEVEL_MAP.get(old_lvl_db, {}).get("GP", "")
                _, old_index = find_matrix_details(old_lvl_db, old_basic)
                st.write(f"Old Index: **{old_index}**")

            with col_new:
                st.success("Promotion Status")
                next_lvl = str(int(old_lvl_db) + 1) if int(old_lvl_db) < 4 else old_lvl_db
                new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=int(next_lvl)-1)
                new_gp = PAY_LEVEL_MAP.get(new_lvl, {}).get("GP", "")
                def_idx = (sorted_en.index(old_desig) - 1) if old_desig in sorted_en and sorted_en.index(old_desig) > 0 else 0
                new_desig = st.selectbox("New Designation", sorted_en, index=def_idx)

            # Fixation Calculation
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_pay, new_index = find_matrix_details(new_lvl, notional)
            next_date = f"01.01.{promo_date.year + 1}" if promo_date.month <= 6 else f"01.07.{promo_date.year + 1}"
            next_inc_pay, _ = find_matrix_details(new_lvl, final_pay * 1.03)

            if st.form_submit_button("Securely Update & Export"):
                # Database Update
                db.collection("employees").document(emp_data['id']).update({
                    "BASIC PAY": int(final_pay),
                    "PAY LEVEL": new_lvl,
                    "Designation": new_desig
                })
                
                # Hindi Template Mapping
                mapping = {
                    "PFNUMBER": sel_pf,
                    "EMPLOYEENAME": emp_data['Employee Name'],
                    "OLDDESIGNATION": old_desig,
                    "STATION": emp_data.get('STATION', 'SGAM'),
                    "OLDGP": old_gp, "NEWGP": new_gp,
                    "OLDBASICPAY": int(old_basic),
                    "NEWBASICPAY": int(final_pay),
                    "NEWBASICPAY= (MROUND100OLDBASICPAY*103%=<INNEWGP": int(final_pay),
                    "OLDINDEX": old_index, "NEWINDEX": new_index,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                    "NEXTINCRDATE": next_date,
                    "MROUND100ONEWBASICPAY*103%": int(next_inc_pay)
                }

                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                if os.path.exists(t_path):
                    doc = Document(t_path)
                    safe_replace(doc, mapping)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.session_state.memo = bio.getvalue()
                    st.success("Database and Memo Updated Successfully!")
                    st.rerun()

    if 'memo' in st.session_state:
        st.download_button("📥 Download Secure Memo", st.session_state.memo, f"Fixation_{sel_pf}.docx")

