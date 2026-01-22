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
    if "password_correct" not in st.session_state:
        st.title("🔐 Secure Login - SGAM Office")
        user = st.text_input("Username", key="login_user")
        pwd = st.text_input("Password", type="password", key="login_pwd")
        if st.button("Login"):
            if user == "admin" and pwd == "sgam123":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Invalid Credentials")
        return False
    return True

# --- 2. EXTENDED DATA TABLES (LEVEL 1 TO 8) ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
    "8": {"PB": "9300-34800", "GP": "4800"},
}

PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500, 24200, 24900, 25600, 26400, 27200],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200, 31100, 32000, 33000],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300, 34300, 35300, 36400, 37500, 38600],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100, 39200, 40400, 41600, 42800, 44100],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200, 47600, 49000, 50500, 52000, 53600],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000],
    "8": [47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000, 70000, 72100],
}

# --- 3. CORE FUNCTIONS ---
def find_details(level, pay):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= pay: return int(val), cells.index(val) + 1
    return int(pay), 1

def powerful_replace(doc, data):
    """Deep search and replace in paragraphs and tables"""
    for k, v in data.items():
        tag = f"[{k}]"
        # Search Paragraphs
        for p in doc.paragraphs:
            if tag in p.text: p.text = p.text.replace(tag, str(v))
        # Search Tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if tag in p.text: p.text = p.text.replace(tag, str(v))

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

# --- 4. MAIN APP ---
if check_password():
    db = init_db()
    docs = db.collection("employees").stream()
    data_list, desigs = [], []
    for d in docs:
        item = d.to_dict(); item['id'] = d.id
        data_list.append(item)
        if item.get('Designation'): desigs.append(item['Designation'])
    
    df = pd.DataFrame(data_list)
    sorted_desigs = sorted(list(set(desigs)))

    st.title("📋 Railway Promotion Fixation System (L1-L8)")

    if not df.empty:
        search_list = df.apply(lambda r: f"{r['Employee Name']} ({str(r.get('PF Number','')).split('.')[0]})", axis=1).tolist()
        sel_emp_str = st.selectbox("Search Employee", search_list)
        emp_data = df[df['Employee Name'] == sel_emp_str.split(' (')[0]].iloc[0]
        
        with st.form("fixation_form_v_final"):
            old_pay = int(float(emp_data.get('BASIC PAY', 18000)))
            old_lvl = str(int(float(emp_data.get('PAY LEVEL', 1))))
            
            c1, c2, c3 = st.columns(3)
            curr_pay = c1.number_input("Current Basic Pay", value=old_pay)
            fix_date = c2.date_input("Fixation Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Old Status")
                old_desig = emp_data.get('Designation', '')
                st.text_input("Designation", old_desig, disabled=True)
                old_gp = PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", "")
                _, old_idx = find_details(old_lvl, curr_pay)
                st.write(f"Old Index: **{old_idx}** | Old GP: **{old_gp}**")
            
            with col2:
                st.subheader("New Status")
                next_l_idx = int(old_lvl) if int(old_lvl) < 8 else 7
                new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=next_l_idx)
                new_gp = PAY_LEVEL_MAP.get(new_lvl, {}).get("GP", "")
                
                def_d_idx = (sorted_desigs.index(old_desig)-1) if old_desig in sorted_desigs and sorted_desigs.index(old_desig)>0 else 0
                new_desig = st.selectbox("New Designation", sorted_desigs, index=def_idx if 'def_idx' in locals() else def_d_idx)

            # Fixation Maths
            notional = math.ceil((curr_pay * 1.03) / 100) * 100
            final_pay, new_idx = find_details(new_lvl, notional)
            
            # Increment logic
            inc_month = "01.07" if fix_date.month <= 6 else "01.01"
            inc_year = fix_date.year + (1 if fix_date.month > 6 else 0)
            inc_date_str = f"{inc_month}.{inc_year}"
            inc_pay, _ = find_details(new_lvl, final_pay * 1.03)

            if st.form_submit_button("Update & Export Document"):
                hindi_name = emp_data.get('Employee Name in Hindi', emp_data['Employee Name'])
                
                # Database Update
                db.collection("employees").document(emp_data['id']).update({
                    "BASIC PAY": int(final_pay),
                    "PAY LEVEL": new_lvl,
                    "Designation": new_desig
                })
                
                # FULL MAPPING (Including [OLDGP] and [NEWGP])
                mapping = {
                    "PFNUMBER": str(emp_data.get('PF Number','')).split('.')[0],
                    "EMPLOYEENAME": hindi_name,
                    "OLDDESIGNATION": old_desig,
                    "NEWDESIGNATION": new_desig,
                    "STATION": emp_data.get('STATION', 'GNDI'),
                    "OLDGP": old_gp,
                    "NEWGP": new_gp,
                    "OLDBASICPAY": int(curr_pay),
                    "NEWBASICPAY": int(final_pay),
                    "OLDLEVEL": old_lvl,
                    "NEWLEVEL": new_lvl,
                    "OLDPAYBAND": PAY_LEVEL_MAP[old_lvl]["PB"],
                    "NEWPAYBAND": PAY_LEVEL_MAP[new_lvl]["PB"],
                    "OLDINDEX": old_idx,
                    "NEWINDEX": new_idx,
                    "PROMOTIONORDERNUMBER": order_no,
                    "PROMOTIONDATE": fix_date.strftime("%d.%m.%Y"),
                    "NEXTINCRDATE": inc_date_str,
                    "MROUND100OLDBASICPAY*103%": int(notional),
                    "MROUND100ONEWBASICPAY*103%": int(inc_pay)
                }

                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                if os.path.exists(t_path):
                    doc = Document(t_path)
                    powerful_replace(doc, mapping)
                    bio = io.BytesIO()
                    doc.save(bio)
                    st.session_state.word_file = bio.getvalue()
                    st.success(f"Order for {hindi_name} is ready!")
                    st.rerun()

    if "word_file" in st.session_state:
        st.download_button("📥 Download Fixation Order", st.session_state.word_file, "Fixation_Done.docx")
