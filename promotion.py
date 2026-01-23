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
        st.title("🔐 SGAM Office Secure Login")
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

# --- 2. EXTENDED PAY MATRIX (LEVEL 1 TO 8) ---
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

# --- 3. HELPER FUNCTIONS ---
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
        if "firebase_config" in st.secrets:
            cred_dict = dict(st.secrets["firebase_config"])
            cred_dict["private_key"] = cred_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cred_dict)
        else:
            # Fallback for local testing
            cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
        firebase_admin.initialize_app(cred)
    return firestore.client()

# --- 4. MAIN APP ---
if check_password():
    db = init_db()
    tab1, tab2 = st.tabs(["🚀 Promotion Entry", "📊 Promotion History Report"])

    # Fetch all employees
    docs = db.collection("employees").stream()
    emp_list = [d.to_dict() | {"id": d.id} for d in docs]
    df_emp = pd.DataFrame(emp_list)
    
    # --- TAB 1: PROMOTION LOGIC ---
    with tab1:
        if not df_emp.empty:
            search_list = df_emp.apply(lambda r: f"{r['Employee Name']} ({str(r.get('PF Number','')).split('.')[0]})", axis=1).tolist()
            sel_emp_str = st.selectbox("Search Employee", search_list)
            emp_data = df_emp[df_emp['Employee Name'] == sel_emp_str.split(' (')[0]].iloc[0]
            
            promo_type = st.radio("Choose Option", ["On Date Promotion", "Promotion On Next Increment"], horizontal=True)
            
            with st.form("promo_fixation_form"):
                old_pay = int(float(emp_data.get('BASIC PAY', 18000)))
                old_lvl = str(int(float(emp_data.get('PAY LEVEL', 1))))
                
                c1, c2, c3 = st.columns(3)
                curr_pay = c1.number_input("Current Basic Pay", value=old_pay)
                fix_date = c2.date_input("Fixation/Order Date", value=datetime.now())
                order_no = st.text_input("Promotion Order Number")

                col1, col2 = st.columns(2)
                with col1:
                    st.info("Current (Old) Status")
                    old_desig = emp_data.get('Designation', '')
                    old_gp = PAY_LEVEL_MAP.get(old_lvl, {}).get("GP", "1800")
                    _, old_idx = find_details(old_lvl, curr_pay)
                    st.write(f"GP: **{old_gp}** | Matrix Index: **{old_idx}**")
                
                with col2:
                    st.success("Promotion (New) Status")
                    all_desigs = sorted(list(set(df_emp['Designation'].dropna())))
                    new_lvl = st.selectbox("Select New Level", list(PAY_LEVEL_MAP.keys()), index=int(old_lvl) if int(old_lvl)<8 else 7)
                    new_gp = PAY_LEVEL_MAP[new_lvl]["GP"]
                    def_idx = (all_desigs.index(old_desig)-1) if old_desig in all_desigs and all_desigs.index(old_desig)>0 else 0
                    new_desig = st.selectbox("New Designation (NEWDESIGNATION)", all_desigs, index=def_idx)

                # Fixation Calculation
                notional = math.ceil((curr_pay * 1.03) / 100) * 100
                final_pay, new_idx = find_details(new_lvl, notional)
                inc_date = f"01.07.{fix_date.year + (1 if fix_date.month > 6 else 0)}" if fix_date.month <= 6 else f"01.01.{fix_date.year + 1}"

                if st.form_submit_button("Update & Generate Document"):
                    hindi_name = emp_data.get('Employee Name in Hindi', emp_data['Employee Name'])
                    
                    # 1. Update Employee Database
                    db.collection("employees").document(emp_data['id']).update({
                        "BASIC PAY": int(final_pay),
                        "PAY LEVEL": new_lvl,
                        "Designation": new_desig,
                        "Posting Status": new_desig  # Updating Posting Status as requested
                    })
                    
                    # 2. Add to History Collection
                    db.collection("promotion_history").add({
                        "PF_Number": emp_data.get('PF Number'),
                        "Name": hindi_name,
                        "Type": promo_type,
                        "Old_GP": old_gp,
                        "New_GP": new_gp,
                        "Order_No": order_no,
                        "Date": fix_date.strftime("%Y-%m-%d"),
                        "timestamp": datetime.now()
                    })

                    # 3. Generate Document
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

                    template_name = "General Promotion MACP temp.docx" if promo_type == "On Date Promotion" else "On Increment Promotion temp.docx"
                    t_path = os.path.join("assets", template_name)
                    
                    if os.path.exists(t_path):
                        doc = Document(t_path)
                        powerful_replace(doc, mapping)
                        bio = io.BytesIO()
                        doc.save(bio)
                        # File Name: Sunil_Kumar_Prajapati_1900.docx
                        safe_name = "".join([c if c.isalnum() else "_" for c in emp_data['Employee Name']])
                        st.session_state.file_output = bio.getvalue()
                        st.session_state.file_name = f"{safe_name}_{new_gp}.docx"
                        st.success("Database Updated and Document Prepared!")
                        st.rerun()
                    else:
                        st.error(f"Template not found at {t_path}")

        if "file_output" in st.session_state:
            st.download_button(f"📥 Download {st.session_state.file_name}", 
                             st.session_state.file_output, 
                             st.session_state.file_name)

    # --- TAB 2: HISTORY REPORT ---
    with tab2:
        st.header("📋 Promotion History Log")
        hist_docs = db.collection("promotion_history").order_by("timestamp", direction=firestore.Query.DESCENDING).stream()
        hist_data = [h.to_dict() for h in hist_docs]
        
        if hist_data:
            df_hist = pd.DataFrame(hist_data)
            # Filter by Promotion Type
            filter_type = st.multiselect("Filter by Promotion Type", df_hist['Type'].unique(), default=df_hist['Type'].unique())
            filtered_df = df_hist[df_hist['Type'].isin(filter_type)]
            
            st.dataframe(filtered_df[['Name', 'PF_Number', 'Old_GP', 'New_GP', 'Type', 'Order_No', 'Date']], use_container_width=True)
            
            # Export to Excel
            towrite = io.BytesIO()
            filtered_df.to_excel(towrite, index=False)
            st.download_button("📂 Export History to Excel", towrite.getvalue(), "Promotion_History.xlsx")
        else:
            st.info("No promotion history found yet.")
