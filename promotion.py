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

# Add your full Matrix here (Sample below)
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
    with st.form("login_ui"):
        u = st.text_input("User")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if u == "admin" and p == "Sgam@4321":
                st.session_state.auth = True
                st.rerun()
            else: st.error("Invalid credentials")
    st.stop()

# --- 2. DB INIT (Secrets Based) ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        if "firebase_config" in st.secrets:
            cred_dict = dict(st.secrets["firebase_config"])
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        else:
            # Local Testing
            cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

db = init_db()

# --- 3. HELPERS ---
def clean_int(val):
    """Safely convert any value to integer without .0"""
    try:
        if val is None or val == "": return 0
        return int(float(str(val).strip()))
    except: return 0

def find_matrix_pay(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val:
            return int(val), cells.index(val) + 1
    return int(target_val), 1

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    # Paragraphs and Tables replacement
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in data.items():
            placeholder = f"[{k}]"
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str(v))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 4. DATA FETCH & PROCESSING ---
emp_docs = db.collection("employees").stream()
data_list, desigs_en, desigs_hi = [], set(), set()

for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    # Clean PF for alphanumeric support
    pf_raw = str(item.get('PF Number', '')).strip()
    if pf_raw.endswith('.0'): pf_raw = pf_raw[:-2]
    item['PF_Clean'] = pf_raw
    data_list.append(item)
    if item.get('Designation'): desigs_en.add(item['Designation'])
    if item.get('Designation in Hindi'): desigs_hi.add(item['Designation in Hindi'])

df_emp = pd.DataFrame(data_list)
all_desig_en = sorted(list(desigs_en))
all_desig_hi = sorted(list(desigs_hi))

# --- 5. MAIN UI ---
st.title("🚀 Promotion & MACP Fixation")
tab1, tab2 = st.tabs(["Promotion Entry", "History Logs"])

with tab1:
    if not df_emp.empty:
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        sel_emp = st.selectbox("Search Employee (Name or PF Number)", search_options)
        sel_pf = sel_emp.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['PF_Clean'] == sel_pf].iloc[0]

        # Start Main Form
        with st.form("promotion_fixation_form"):
            st.subheader(f"Fixation for: {emp_data['Employee Name']}")
            
            c1, c2, c3 = st.columns(3)
            # Auto-loaded Old Basic Pay as Integer
            curr_pay = clean_int(emp_data.get('Basic Pay', 0))
            old_basic = c1.number_input("Old Basic Pay", value=curr_pay, step=1)
            promo_date = c2.date_input("Promotion/MACP Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            st.markdown("---")
            col_old, col_new = st.columns(2)

            with col_old:
                st.info("Current (Old) Status")
                o_lvl = str(clean_int(emp_data.get('Level', 1)))
                st.text_input("Old Level", o_lvl, disabled=True)
                st.text_input("Old Pay Band", PAY_LEVEL_MAP.get(o_lvl, {}).get("PB", ""), disabled=True)
                st.text_input("Old Grade Pay", PAY_LEVEL_MAP.get(o_lvl, {}).get("GP", ""), disabled=True)
                st.text_input("Old Designation (EN)", emp_data.get('Designation', ''), disabled=True)
                st.text_input("Old Designation (HI)", emp_data.get('Designation in Hindi', ''), disabled=True)

            with col_new:
                st.success("Promotion (New) Status")
                # Level selection triggers PB/GP update
                n_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), 
                                     index=int(o_lvl)-1 if o_lvl.isdigit() and 0 < int(o_lvl) <= 7 else 0)
                n_pb = st.text_input("New Pay Band", PAY_LEVEL_MAP.get(n_lvl, {}).get("PB", ""))
                n_gp = st.text_input("New Grade Pay", PAY_LEVEL_MAP.get(n_lvl, {}).get("GP", ""))
                n_desig_en = st.selectbox("New Designation (EN)", all_desig_en)
                n_desig_hi = st.selectbox("New Designation (HI)", all_desig_hi)

            # Calculation Logic
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, n_idx = find_matrix_pay(n_lvl, notional)
            next_incr_date = f"01/01/{promo_date.year + 1}" if promo_date.month <= 6 else f"01/07/{promo_date.year + 1}"

            st.divider()
            st.write(f"### Proposed New Basic: ₹{final_basic}")

            # Submit Button (Essential for Form)
            submit_btn = st.form_submit_button("Update Records & Generate Memo")

            if submit_btn:
                # 1. Update Database (All as Clean Integers/Strings)
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": int(final_basic),
                    "Level": n_lvl,
                    "Designation": n_desig_en,
                    "Designation in Hindi": n_desig_hi
                })
                
                # 2. Add to History
                db.collection("promotion_history").add({
                    "PF": sel_pf, "Name": emp_data['Employee Name'], 
                    "OldBasic": int(old_basic), "NewBasic": int(final_basic), 
                    "Timestamp": datetime.now()
                })
                
                # 3. Word Mapping
                mapping = {
                    "PFNUMBER": sel_pf,
                    "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDBASICPAY": int(old_basic),
                    "NEWBASICPAY": int(final_basic),
                    "OLDLEVEL": o_lvl,
                    "NEWLEVEL": n_lvl,
                    "OLDGP": PAY_LEVEL_MAP.get(o_lvl, {}).get("GP", ""),
                    "NEWGP": n_gp,
                    "OLDPAYBAND": PAY_LEVEL_MAP.get(o_lvl, {}).get("PB", ""),
                    "NEWPAYBAND": n_pb,
                    "OLDDESIGNATION": emp_data.get('Designation', ''),
                    "NEWDESIGNATION": n_desig_en,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                    "PROMOTIONORDERNUMBER": order_no,
                    "MROUND100OLDBASICPAY*103%": int(notional),
                    "NEXTINCRDATE": next_incr_date,
                    "STATION": emp_data.get('STATION', 'SGAM')
                }
                
                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.memo_file = generate_docx(t_path, mapping)
                st.success("Database and History Updated!")
                st.rerun()

    if 'memo_file' in st.session_state:
        st.download_button("📥 Download Promotion Memo", st.session_state.memo_file, f"Promotion_{sel_pf}.docx")

with tab2:
    st.subheader("Promotion Logs")
    logs = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    df_logs = pd.DataFrame([d.to_dict() for d in logs])
    if not df_logs.empty:
        st.dataframe(df_logs, use_container_width=True)
    else:
        st.write("No history records found.")
