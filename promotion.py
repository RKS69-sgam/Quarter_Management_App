import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 0. DATA TABLES (7th CPC) ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
}

# Sample Pay Matrix (Add your full rows here)
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900]
}

# --- 1. CONFIG & AUTH ---
st.set_page_config(page_title="Railway Promotion System", layout="wide")

if 'auth' not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔒 Admin Login")
    with st.form("login"):
        u = st.text_input("User")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else: st.error("Invalid credentials")
    st.stop()

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

# --- 3. HELPERS ---
def clean_val(val):
    """Decimal hatane ke liye function"""
    try:
        if isinstance(val, float) or (isinstance(val, str) and '.' in val):
            return str(int(float(val)))
        return str(val)
    except: return str(val)

def find_matrix_pay(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val:
            return val, cells.index(val) + 1
    return target_val, 1

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    # Paragraphs and Tables replacement
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in data.items():
            if f"[{k}]" in p.text:
                p.text = p.text.replace(f"[{k}]", str(v))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. DATA FETCH ---
emp_docs = db.collection("employees").stream()
data_list, desigs_en, desigs_hi = [], set(), set()

for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    pf = clean_val(item.get('PF Number', ''))
    item['PF_Clean'] = pf
    data_list.append(item)
    if item.get('Designation'): desigs_en.add(item['Designation'])
    if item.get('Designation in Hindi'): desigs_hi.add(item['Designation in Hindi'])

df_emp = pd.DataFrame(data_list)
sorted_desig_en = sorted(list(desigs_en))
sorted_desig_hi = sorted(list(desigs_hi))

# --- 5. MAIN UI ---
st.title("🚀 Promotion & MACP Fixation")
tab1, tab2 = st.tabs(["Promotion Entry", "History Logs"])

with tab1:
    if not df_emp.empty:
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        sel_emp = st.selectbox("Search Employee (Name or PF)", search_options)
        sel_pf = sel_emp.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['PF_Clean'] == sel_pf].iloc[0]

        with st.form("promo_form"):
            # Row 1: Input
            c1, c2, c3 = st.columns(3)
            # Auto-load Basic Pay as Integer
            old_basic = c1.number_input("Old Basic Pay (Auto-loaded)", value=int(float(emp_data.get('Basic Pay', 0))), step=1)
            promo_date = c2.date_input("Promotion Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            st.markdown("### Old vs New Status")
            col_old, col_new = st.columns(2)

            with col_old:
                st.info("Current (Old) Status")
                o_lvl = clean_val(emp_data.get('Level', '1'))
                st.text_input("Old Level", o_lvl, disabled=True)
                st.text_input("Old Pay Band", PAY_LEVEL_MAP.get(o_lvl, {}).get("PB", ""), disabled=True)
                st.text_input("Old Grade Pay", PAY_LEVEL_MAP.get(o_lvl, {}).get("GP", ""), disabled=True)
                st.text_input("Old Designation (EN)", emp_data.get('Designation', ''), disabled=True)
                st.text_input("Old Designation (HI)", emp_data.get('Designation in Hindi', ''), disabled=True)
                # Find current index
                _, o_idx = find_matrix_pay(o_lvl, old_basic)

            with col_new:
                st.success("Promotion (New) Status")
                n_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=int(o_lvl)-1 if o_lvl.isdigit() else 0)
                n_pb = st.text_input("New Pay Band", PAY_LEVEL_MAP.get(n_lvl, {}).get("PB", ""))
                n_gp = st.text_input("New Grade Pay", PAY_LEVEL_MAP.get(n_lvl, {}).get("GP", ""))
                n_desig_en = st.selectbox("New Designation (EN)", sorted_desig_en)
                n_desig_hi = st.selectbox("New Designation (HI)", sorted_desig_hi)

            # Calculation
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, n_idx = find_matrix_pay(n_lvl, notional)
            
            st.divider()
            st.write(f"**Calculated New Basic:** ₹{final_basic} (Level {n_lvl}, Index {n_idx})")

            if st.form_submit_button("Update Records & Generate"):
                # 1. Update DB (As Integers)
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": int(final_basic),
                    "Level": n_lvl,
                    "Designation": n_desig_en,
                    "Designation in Hindi": n_desig_hi
                })
                # 2. Add History
                db.collection("promotion_history").add({
                    "PF": sel_pf, "Name": emp_data['Employee Name'], "NewBasic": final_basic, "Timestamp": datetime.now()
                })
                # 3. Word Mapping
                mapping = {
                    "PFNUMBER": sel_pf, "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDBASICPAY": clean_val(old_basic), "NEWBASICPAY": clean_val(final_basic),
                    "OLDLEVEL": o_lvl, "NEWLEVEL": n_lvl, "OLDINDEX": clean_val(o_idx), "NEWINDEX": clean_val(n_idx),
                    "OLDGP": PAY_LEVEL_MAP.get(o_lvl, {}).get("GP", ""), "NEWGP": n_gp,
                    "OLDPAYBAND": PAY_LEVEL_MAP.get(o_lvl, {}).get("PB", ""), "NEWPAYBAND": n_pb,
                    "OLDDESIGNATION": emp_data.get('Designation', ''), "NEWDESIGNATION": n_desig_en,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"), "PROMOTIONORDERNUMBER": order_no,
                    "MROUND100OLDBASICPAY*103%": clean_val(notional), "STATION": emp_data.get('STATION', 'SGAM')
                }
                
                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.memo = generate_docx(t_path, mapping)
                st.success("Database Updated Successfully!")
                st.rerun()

    if 'memo' in st.session_state:
        st.download_button("📥 Download Promotion Memo", st.session_state.memo, f"Promotion_{sel_pf}.docx")

with tab2:
    logs = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    st.dataframe(pd.DataFrame([d.to_dict() for d in logs]), use_container_width=True)
