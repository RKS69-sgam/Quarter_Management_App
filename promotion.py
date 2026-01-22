import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore
from docx import Document
import io

# --- 0. DATA TABLES ---
PAY_LEVEL_MAP = {
    "1": {"PB": "5200-20200", "GP": "1800"},
    "2": {"PB": "5200-20200", "GP": "1900"},
    "3": {"PB": "5200-20200", "GP": "2000"},
    "4": {"PB": "5200-20200", "GP": "2400"},
    "5": {"PB": "5200-20200", "GP": "2800"},
    "6": {"PB": "9300-34800", "GP": "4200"},
    "7": {"PB": "9300-34800", "GP": "4600"},
}

# Pay Matrix 
PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900]
}

# --- 1. DB INIT ---
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

# --- 2. HELPERS ---
def clean_int(val):
    try:
        if val is None or val == "": return 0
        return int(float(str(val).strip()))
    except: return 0

def find_matrix_pay(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val: return int(val), cells.index(val) + 1
    return int(target_val), 1

def generate_docx(template_path, data):
    if not os.path.exists(template_path): return None
    doc = Document(template_path)
    for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        for k, v in data.items():
            tag = f"[{k}]"
            if tag in p.text:
                p.text = p.text.replace(tag, str(v))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. DATA PREPARATION ---
emp_docs = db.collection("employees").stream()
data_list, desigs_en, desigs_hi = [], [], []

for d in emp_docs:
    item = d.to_dict(); item['id'] = d.id
    pf = str(item.get('PF Number', '')).strip().split('.')[0]
    item['PF_Clean'] = pf
    data_list.append(item)
    if item.get('Designation'): desigs_en.append(item['Designation'])
    if item.get('Designation in Hindi'): desigs_hi.append(item['Designation in Hindi'])

df_emp = pd.DataFrame(data_list)
sorted_en = sorted(list(set(desigs_en)))
sorted_hi = sorted(list(set(desigs_hi)))

# --- 4. MAIN UI ---
st.title("🚀 Smart Promotion & Pay Fixation")
tab1, tab2 = st.tabs(["Promotion Entry", "History Logs"])

with tab1:
    if not df_emp.empty:
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        sel_emp = st.selectbox("Search Employee", search_options)
        sel_pf = sel_emp.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['PF_Clean'] == sel_pf].iloc[0]

        with st.form("promotion_fixation_form"):
            # FETCH OLD DATA
            old_basic_val = clean_int(emp_data.get('BASIC PAY', 0))
            old_lvl_val = str(clean_int(emp_data.get('PAY LEVEL', 1)))
            
            c1, c2, c3 = st.columns(3)
            # Displaying Old Basic as default but editable
            old_basic = c1.number_input("Old Basic Pay (Auto-loaded)", value=old_basic_val, step=1)
            promo_date = c2.date_input("Promotion Date", value=datetime.now())
            order_no = c3.text_input("Order Number")

            st.markdown("---")
            col_old, col_new = st.columns(2)

            with col_old:
                st.subheader("Current (Old) Status")
                st.text_input("Old Level", old_lvl_val, disabled=True)
                st.text_input("Old GP", PAY_LEVEL_MAP.get(old_lvl_val, {}).get("GP", ""), disabled=True)
                st.text_input("Old Designation (EN)", emp_data.get('Designation', ''), disabled=True)
                st.text_input("Old Designation (HI)", emp_data.get('Designation in Hindi', ''), disabled=True)

            with col_new:
                st.subheader("New (Promotion) Status")
                # Logical Next Level Selection
                next_lvl_idx = int(old_lvl_val) if int(old_lvl_val) < 7 else int(old_lvl_val) - 1
                new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=next_lvl_idx)
                
                # New GP auto-updates based on new_lvl
                new_gp = st.text_input("New Grade Pay", PAY_LEVEL_MAP.get(new_lvl, {}).get("GP", ""))
                
                # Auto-select next designations in list
                curr_en = emp_data.get('Designation', '')
                curr_hi = emp_data.get('Designation in Hindi', '')
                def_idx_en = (sorted_en.index(curr_en) + 1) if curr_en in sorted_en and (sorted_en.index(curr_en)+1) < len(sorted_en) else 0
                def_idx_hi = (sorted_hi.index(curr_hi) + 1) if curr_hi in sorted_hi and (sorted_hi.index(curr_hi)+1) < len(sorted_hi) else 0

                new_desig_en = st.selectbox("New Designation (EN)", sorted_en, index=def_idx_en)
                new_desig_hi = st.selectbox("New Designation (HI)", sorted_hi, index=def_idx_hi)

            # CALCULATION Logic
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, n_idx = find_matrix_pay(new_lvl, notional)
            next_incr_date = f"01/01/{promo_date.year + 1}" if promo_date.month <= 6 else f"01/07/{promo_date.year + 1}"

            st.divider()
            st.write(f"### Proposed New Basic Pay: ₹{final_basic}")

            submit = st.form_submit_button("Update Records & Generate Memo")

            if submit:
                # 1. Update Database
                db.collection("employees").document(emp_data['id']).update({
                    "BASIC PAY": int(final_basic),
                    "PAY LEVEL": new_lvl,
                    "Designation": new_desig_en,
                    "Designation in Hindi": new_desig_hi
                })
                
                # 2. Add to History
                db.collection("promotion_history").add({
                    "PF": sel_pf, "Name": emp_data['Employee Name'], 
                    "OldPay": old_basic, "NewPay": final_basic, "Timestamp": datetime.now()
                })

                # 3. Word Template Mapping
                mapping = {
                    "PFNUMBER": sel_pf,
                    "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDBASICPAY": old_basic, "NEWBASICPAY": final_basic,
                    "OLDLEVEL": old_lvl_val, "NEWLEVEL": new_lvl,
                    "OLDGP": PAY_LEVEL_MAP.get(old_lvl_val, {}).get("GP", ""), "NEWGP": new_gp,
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"),
                    "PROMOTIONORDERNUMBER": order_no,
                    "MROUND100OLDBASICPAY*103%": notional,
                    "NEXTINCRDATE": next_incr_date,
                    "NEWDESIGNATION": new_desig_en
                }
                
                t_path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.memo = generate_docx(t_path, mapping)
                st.success("Database and History Updated!")
                st.rerun()

    if 'memo' in st.session_state:
        st.download_button("📥 Download Document", st.session_state.memo, f"Fixation_{sel_pf}.docx")

with tab2:
    logs = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    df_logs = pd.DataFrame([d.to_dict() for d in logs])
    if not df_logs.empty:
        st.dataframe(df_logs, use_container_width=True)
