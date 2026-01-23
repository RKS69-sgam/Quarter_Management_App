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
st.set_page_config(page_title="Railway Promotion System", layout="wide")

def check_password():
    if "password_correct" not in st.session_state:
        st.title("🔐 SGAM Office Login")
        user = st.text_input("Username")
        pwd = st.text_input("Password", type="password")
        if st.button("Login"):
            if user == "admin" and pwd == "sgam123":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Invalid Credentials")
        return False
    return True

# --- 2. DATA TABLES (L1 TO L8) ---
PAY_LEVEL_MAP = {str(i): {"PB": "5200-20200" if i<6 else "9300-34800", "GP": gp} 
                 for i, gp in zip(range(1,9), [1800, 1900, 2000, 2400, 2800, 4200, 4600, 4800])}

PAY_MATRIX = {
    "1": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500, 24200, 24900, 25600],
    "2": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400],
    "3": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200, 31100],
    "4": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300, 34300, 35300, 36400],
    "5": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100, 39200, 40400, 41600],
    "6": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200, 47600, 49000, 50500],
    "7": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100],
    "8": [47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000],
}

# --- 3. HELPERS ---
def find_details(level, pay):
    cells = PAY_MATRIX.get(str(level), [])
    for val in cells:
        if val >= pay: return int(val), cells.index(val) + 1
    return int(pay), 1

def powerful_replace(doc, data):
    for k, v in data.items():
        tag = f"[{k}]"
        for p in doc.paragraphs:
            if tag in p.text: p.text = p.text.replace(tag, str(v))
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if tag in p.text: p.text = p.text.replace(tag, str(v))

@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

# --- 4. MAIN APP ---
if check_password():
    db = init_db()
    tab1, tab2 = st.tabs(["🚀 Promotion Tab", "📜 Promotion History"])

    # Load Employees
    emp_docs = db.collection("employees").stream()
    emp_list = [d.to_dict() | {"id": d.id} for d in emp_docs]
    df_emp = pd.DataFrame(emp_list)
    
    # Load History
    hist_docs = db.collection("promotion_history").order_by("timestamp", direction=firestore.Query.DESCENDING).stream()
    hist_list = [h.to_dict() for h in hist_docs]
    df_hist = pd.DataFrame(hist_list)

    with tab1:
        if not df_emp.empty:
            search_list = df_emp.apply(lambda r: f"{r['Employee Name']} ({str(r.get('PF Number','')).split('.')[0]})", axis=1).tolist()
            sel_emp_str = st.selectbox("Search Employee", search_list)
            emp_data = df_emp[df_emp['Employee Name'] == sel_emp_str.split(' (')[0]].iloc[0]
            
            promo_type = st.radio("Select Promotion Type", ["On Date Promotion", "Promotion On Next Increment"], horizontal=True)
            
            with st.form("promo_form"):
                old_pay = int(float(emp_data.get('BASIC PAY', 18000)))
                old_lvl = str(int(float(emp_data.get('PAY LEVEL', 1))))
                
                c1, c2, c3 = st.columns(3)
                curr_pay = c1.number_input("Current Basic Pay", value=old_pay)
                fix_date = c2.date_input("Promotion/Fixation Date", value=datetime.now())
                order_no = c3.text_input("Order Number")

                col1, col2 = st.columns(2)
                with col1:
                    st.info("Old Details")
                    old_desig = emp_data.get('Designation', '')
                    old_gp = PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", "")
                    _, old_idx = find_details(old_lvl, curr_pay)
                    st.write(f"GP: {old_gp} | Index: {old_idx}")
                
                with col2:
                    st.success("New Details")
                    all_desigs = sorted(list(set(df_emp['Designation'].dropna())))
                    new_lvl = st.selectbox("New Level", list(PAY_LEVEL_MAP.keys()), index=int(old_lvl) if int(old_lvl)<8 else 7)
                    new_gp = PAY_LEVEL_MAP[new_lvl]["GP"]
                    def_idx = (all_desigs.index(old_desig)-1) if old_desig in all_desigs and all_desigs.index(old_desig)>0 else 0
                    new_desig = st.selectbox("New Designation", all_desigs, index=def_idx)

                # Fixation Maths
                notional = math.ceil((curr_pay * 1.03) / 100) * 100
                final_pay, new_idx = find_details(new_lvl, notional)
                inc_date = f"01.07.{fix_date.year + (1 if fix_date.month > 6 else 0)}" if fix_date.month <= 6 else f"01.01.{fix_date.year + 1}"

                if st.form_submit_button("Process Promotion"):
                    hindi_name = emp_data.get('Employee Name in Hindi', emp_data['Employee Name'])
                    
                    # 1. Update Employee Collection
                    db.collection("employees").document(emp_data['id']).update({
                        "BASIC PAY": int(final_pay),
                        "PAY LEVEL": new_lvl,
                        "Designation": new_desig,
                        "Posting Status": new_desig  # Update Posting Status
                    })
                    
                    # 2. Add to History
                    history_entry = {
                        "PF Number": emp_data.get('PF Number'),
                        "Name": hindi_name,
                        "Old GP": old_gp,
                        "New GP": new_gp,
                        "Promotion Type": promo_type,
                        "Order No": order_no,
                        "timestamp": datetime.now()
                    }
                    db.collection("promotion_history").add(history_entry)

                    # 3. Generate Word
                    mapping = {
                        "PFNUMBER": str(emp_data.get('PF Number','')).split('.')[0],
                        "EMPLOYEENAME": hindi_name,
                        "OLDDESIGNATION": old_desig, "NEWDESIGNATION": new_desig,
                        "OLDGP": old_gp, "NEWGP": new_gp,
                        "OLDBASICPAY": int(curr_pay), "NEWBASICPAY": int(final_pay),
                        "OLDLEVEL": old_lvl, "NEWLEVEL": new_lvl,
                        "OLDPAYBAND": PAY_LEVEL_MAP[old_lvl]["PB"], "NEWPAYBAND": PAY_LEVEL_MAP[new_lvl]["PB"],
                        "OLDINDEX": old_idx, "NEWINDEX": new_idx,
                        "PROMOTIONORDERNUMBER": order_no,
                        "PROMOTIONDATE": fix_date.strftime("%d.%m.%Y"),
                        "NEXTINCRDATE": inc_date,
                        "MROUND100OLDBASICPAY*103%": int(notional)
                    }

                    template_file = "General Promotion MACP temp.docx" if promo_type == "On Date Promotion" else "On Increment Promotion temp.docx"
                    t_path = os.path.join("assets", template_file)
                    
                    if os.path.exists(t_path):
                        doc = Document(t_path)
                        powerful_replace(doc, mapping)
                        bio = io.BytesIO()
                        doc.save(bio)
                        # Smart Naming: Sunil_Kumar_1900.docx
                        clean_name = "".join([c for c in emp_data['Employee Name'] if c.isalnum() or c==' ']).replace(' ', '_')
                        st.session_state.file_name = f"{clean_name}_{new_gp}.docx"
                        st.session_state.word_file = bio.getvalue()
                        st.success(f"History updated for {hindi_name}")
                        st.rerun()

        if "word_file" in st.session_state:
            st.download_button(f"📥 Download {st.session_state.file_name}", 
                             st.session_state.word_file, 
                             st.session_state.file_name)

    with tab2:
        st.header("📋 Promotion History Report")
        if not df_hist.empty:
            # Filter options
            p_type_filter = st.multiselect("Filter by Type", df_hist['Promotion Type'].unique(), default=df_hist['Promotion Type'].unique())
            filtered_hist = df_hist[df_hist['Promotion Type'].isin(p_type_filter)]
            
            st.dataframe(filtered_hist[['Name', 'PF Number', 'Old GP', 'New GP', 'Promotion Type', 'Order No', 'timestamp']], use_container_width=True)
            
            # Excel Export
            towrite = io.BytesIO()
            filtered_hist.to_excel(towrite, index=False)
            st.download_button("📂 Export History to Excel", towrite.getvalue(), "Promotion_History.xlsx")
        else:
            st.info("No promotion history records found.")
