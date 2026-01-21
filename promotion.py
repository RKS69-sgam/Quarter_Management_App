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
            if "firebase_config" in st.secrets:
                cred_dict = dict(st.secrets["firebase_config"])
                cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
                cred = credentials.Certificate(cred_dict)
            else:
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Init Error: {e}"); st.stop()
    return firestore.client()

db = init_db()

# --- 2. PAY MATRIX DATA ---

PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600]
}

# --- 3. HELPER FUNCTIONS ---
def get_next_increment_date(promo_date):
    return f"01/01/{promo_date.year + 1}" if promo_date.month <= 6 else f"01/07/{promo_date.year + 1}"

def find_cell_in_level(level, target_val):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= target_val:
            idx = cells.index(val) + 1
            return val, idx
    return target_val, 1

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

# --- 4. MAIN UI ---
tab1, tab2 = st.tabs(["🚀 Promotion Process", "📜 History Report"])

# Load Data and treat PF as Text
emp_docs = db.collection("employees").stream()
data_list = []
for d in emp_docs:
    item = d.to_dict()
    item['id'] = d.id
    # Sabhi PF numbers ko string mein convert kar rahe hain (text support ke liye)
    raw_pf = str(item.get('PF Number', '')).strip()
    if raw_pf.endswith('.0'): raw_pf = raw_pf[:-2] # Float conversion cleanup
    item['PF_Clean'] = raw_pf
    data_list.append(item)

df_emp = pd.DataFrame(data_list)

with tab1:
    if not df_emp.empty:
        # Search by Name or Alphanumeric PF
        search_options = df_emp.apply(lambda r: f"{r['Employee Name']} ({r['PF_Clean']})", axis=1).tolist()
        selected_option = st.selectbox("Search Employee (Name or PF Number)", search_options)
        
        # Safe extraction of PF from string
        selected_pf = selected_option.split('(')[-1].strip(')')
        emp_data = df_emp[df_emp['PF_Clean'] == selected_pf].iloc[0]

        with st.form("promo_form"):
            st.subheader(f"Promotion: {emp_data['Employee Name']} | PF: {selected_pf}")
            c1, c2, c3 = st.columns(3)
            
            old_basic = c1.number_input("Old Basic Pay", value=float(emp_data.get('Basic Pay', 0)))
            promo_date = c2.date_input("Promotion Date")
            order_no = c3.text_input("Order Number")
            
            c4, c5 = st.columns(2)
            new_desig = c4.text_input("New Designation", value=emp_data.get('Designation', ''))
            target_lvl = c5.selectbox("Select New Level", list(PAY_MATRIX.keys()))

            # Increment Calculation
            notional = math.ceil((old_basic * 1.03) / 100) * 100
            final_basic, new_idx = find_cell_in_level(target_lvl, notional)
            next_date = get_next_increment_date(promo_date)

            if st.form_submit_button("Update & Generate"):
                # Database Updates
                db.collection("employees").document(emp_data['id']).update({
                    "Basic Pay": final_basic, "Level": target_lvl, "Designation": new_desig
                })
                db.collection("promotion_history").add({
                    "PF": selected_pf, "Name": emp_data['Employee Name'], 
                    "NewBasic": final_basic, "Date": str(promo_date), "Timestamp": datetime.now()
                })
                
                # Word Data
                mapping = {
                    "PFNUMBER": selected_pf, 
                    "EMPLOYEENAME": emp_data.get('Employee Name in Hindi', emp_data['Employee Name']),
                    "OLDBASICPAY": f"{old_basic}/-", "NEWBASICPAY": f"{final_basic}/-",
                    "PROMOTIONDATE": promo_date.strftime("%d.%m.%Y"), 
                    "PROMOTIONORDERNUMBER": order_no,
                    "OLDLEVEL": emp_data.get('Level', '1'), "NEWLEVEL": target_lvl, 
                    "NEXTINCRDATE": next_date, "STATION": emp_data.get('STATION', 'SGAM'),
                    "MROUND100OLDBASICPAY*103%": notional
                }
                
                path = os.path.join("assets", "General Promotion MACP temp.docx")
                st.session_state.promo_file = generate_docx(path, mapping)
                st.session_state.file_name = f"Promotion_{selected_pf}.docx"
                st.success("Record Updated Successfully!")
                st.rerun()

    if 'promo_file' in st.session_state:
        st.download_button("📥 Download Document", st.session_state.promo_file, st.session_state.file_name)

with tab2:
    st.subheader("Promotion Logs")
    hist = db.collection("promotion_history").order_by("Timestamp", direction=firestore.Query.DESCENDING).stream()
    st.dataframe(pd.DataFrame([d.to_dict() for d in hist]), use_container_width=True)
