import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- DATA TABLES (Pay Band, GP & Matrix) ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
}

# Pay Matrix (Sample values, add full matrix as needed)
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900]
}

# --- AUTHENTICATION ---
if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login"):
        u, p = st.text_input("User"), st.text_input("Password", type="password")
        if st.form_submit_button("Login") and u == "admin" and p == "Sgam@4321":
            st.session_state.auth = True
            st.rerun()
    st.stop()

# --- DB INIT ---
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

# --- HELPERS ---
def get_next_incr_date(promo_date):
    return f"01/01/{promo_date.year + 1}" if promo_date.month <= 6 else f"01/07/{promo_date.year + 1}"

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in data.items():
            if f"[{k}]" in p.text:
                p.text = p.text.replace(f"[{k}]", str(v))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI & LOGIC ---
st.header("🚀 Promotion & MACP System")
tab1, tab2 = st.tabs(["Promotion Entry", "Reports"])

emp_docs = db.collection("employees").stream()
data_list = []
desigs_en, desigs_hi = set(), set()

for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    raw_pf = str(item.get('PF Number', '')).strip().split('.')[0]
    item['PF_Clean'] = raw_pf
    data_list.append(item)
    if item.get('Designation'): desigs_en.add(item['Designation'])
    if item.get('Designation in Hindi'): desigs_hi.add(item['Designation in Hindi'])

df_emp = pd.DataFrame(data_list)

with tab1:
    if not df_emp.empty:
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        sel_emp = st.selectbox("Search Employee", search_options)
        sel_pf = sel_emp.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['PF_Clean'] == sel_pf].iloc[0]

        with st.form("promo_form"):
            c1, c2, c3 = st.columns(3)
            old_basic = c1.number_input("Old Basic Pay", value=float(emp_data.get('Basic Pay', 0)))
            promo_date = c2.date_input("Promotion Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            st.write("### Comparison (Old vs New)")
            col_old, col_new = st.columns(2)

            with col_old:
                st.info("Current Status")
                old_lvl = str(emp_data.get('Level', '1'))
                st.text_input("Old Level", value=old_lvl, disabled=True)
                st.text_input("Old Grade Pay", value=PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", ""), disabled=True)
                st.text_input("Old Designation", value=emp_data.get('Designation', ''), disabled=True)

            with col_new:
                st.success("Target Status")
                new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=int(old_lvl))
                new_gp = st.text_input("New Grade Pay", value=PAY_LEVEL_MAP.get(new_lvl, {}).get("GP", ""))
                new_desig_en = st.selectbox("New Designation (EN)", sorted(list(desigs_en)))
                new_desig_hi = st.selectbox("New Designation (HI)", sorted(list(desigs_hi)))

            # Math Logic
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            lvl_cells = PAY_MATRIX.get(new_lvl, [])
            final_basic = next((v for v in lvl_cells if v >= notional), notional)
            next_date = get_next_incr_date(promo_date)

            st.divider()
            st.metric("New Basic Pay", f"₹{final_basic}")

            if st.form_submit_button("Update & Generate"):
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic, "Level": new_lvl, "Designation": new_desig_en
                })
                db.collection("promotion_history").add({
                    "PF": sel_pf, "Name": emp_data['Employee Name'], "NewBasic": final_basic, "Timestamp": datetime.now()
                })
                
                mapping = {
                    "PFNUMBER": sel_pf, "EMPLOYEENAME": emp_data['Employee Name'],
                    "OLDBASICPAY": old_basic, "NEWBASICPAY": final_basic,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"), "PROMOTIONORDERNUMBER": order_no,
                    "OLDLEVEL": old_lvl, "NEWLEVEL": new_lvl, "NEXTINCRDATE": next_date,
                    "MROUND100OLDBASICPAY*103%": notional, "NEWGP": new_gp, "OLDGP": PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", "")
                }
                
                path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.promo_file = generate_docx(path, mapping)
                st.success("Record Updated!")
                st.rerun()

    if 'promo_file' in st.session_state:
        st.download_button("📥 Download Memo", st.session_state.promo_file, "Promotion_Memo.docx")

with tab2:
    hist = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    st.dataframe(pd.DataFrame([d.to_dict() for d in hist]), use_container_width=True)
