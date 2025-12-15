import streamlit as st
import datetime
import os 
import pandas as pd
from docx import Document 
import io
import firebase_admin
from firebase_admin import credentials, firestore
import json
import uuid # Firestore Document ID जेनरेट करने के लिए
from datetime import date

# ----------------------------------------------------------------------
# 0. कॉन्फ़िगरेशन (Config)
# ----------------------------------------------------------------------

# NOTE: अब यह ऐप पूरी तरह से Firebase Firestore पर निर्भर करता है।
# PostgreSQL कनेक्शन और SQLalchemy की अब आवश्यकता नहीं है।
# Master Quarters और History Data इन CSVs से एक बार Firestore में लोड होगा।
INVENTORY_CSV_PATH = "data/Quarter_Register.csv" 
HISTORY_CSV_PATH = "data/Quarter_History_Initial.csv" # यदि आपके पास प्रारंभिक इतिहास CSV है
EMPLOYEE_COLLECTION = "employees" # कर्मचारी मास्टर डेटा
QUARTER_MASTER_COLLECTION = "master_quarters" # क्वार्टर इन्वेंट्री
QUARTER_HISTORY_COLLECTION = "quarter_history" # अलॉटमेंट हिस्ट्री

# --- SECURITY CONFIGURATION ---
CORRECT_PASSWORD = "Sgam@1234" 
# ------------------------------

st.set_page_config(layout="wide", page_title="रेलवे क्वार्टर प्रबंधन (Firebase Firestore)")

# ----------------------------------------------------------------------
# 1. FIREBASE CONNECTION & INITIAL DATA LOAD (नया)
# ----------------------------------------------------------------------

@st.cache_resource
def initialize_firebase():
    """Firebase SDK को इनिशियलाइज़ करता है और Firestore क्लाइंट लौटाता है।"""
    try:
        if not firebase_admin._apps:
            # secrets.toml में 'firebase_config' से क्रेडेंशियल्स लोड करना
            if st.secrets.get("firebase_config"):
                service_account_info_attrdict = st.secrets["firebase_config"]
                final_credentials = dict(service_account_info_attrdict)
                if isinstance(final_credentials.get('private_key'), str):
                     final_credentials['private_key'] = final_credentials['private_key'].replace('\\n', '\n')
                
                cred = credentials.Certificate(final_credentials)
            
            else:
                # लोकल टेस्टिंग: सर्विस अकाउंट फ़ाइल यहाँ होनी चाहिए
                SA_FILE = 'sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json'
                if os.path.exists(SA_FILE):
                     with open(SA_FILE) as f:
                        service_account_info = json.load(f)
                     cred = credentials.Certificate(service_account_info)
                else:
                    st.error("Firebase Service Account File not found locally.")
                    return None
            
            firebase_admin.initialize_app(cred)
            
        return firestore.client()
        
    except Exception as e:
        st.error(f"❌ Firebase कनेक्शन विफल। त्रुटि: {e}")
        return None

db = initialize_firebase()

# --- डेटा माइग्रेशन फ़ंक्शन (केवल एक बार चलाने के लिए) ---

def load_initial_data_to_firestore():
    """
    CSV से क्वार्टर इन्वेंट्री और इतिहास को Firestore में लोड करता है।
    यह जाँचता है कि कलेक्शन खाली है या नहीं।
    """
    if db is None: return False
    
    # 1. Master Quarters Loading
    try:
        master_count = len(list(db.collection(QUARTER_MASTER_COLLECTION).limit(1).get()))
        
        if master_count == 0:
            st.info(f"CSV से क्वार्टर इन्वेंट्री ({INVENTORY_CSV_PATH}) को Firestore में लोड किया जा रहा है...")
            df_inventory = pd.read_csv(INVENTORY_CSV_PATH)
            df_inventory.columns = df_inventory.columns.str.strip().str.upper()
            
            required_cols = ['QUARTER_NUMBER', 'STATION']
            if not all(col in df_inventory.columns for col in required_cols):
                 st.error("CSV Error: Quarter Inventory फ़ाइल में 'QUARTER_NUMBER' और 'STATION' कॉलम नहीं मिले।")
                 return False

            batch = db.batch()
            for index, row in df_inventory.iterrows():
                doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(f"{row['STATION']}_{row['QUARTER_NUMBER']}")
                batch.set(doc_ref, {
                    'quarter_number': str(row['QUARTER_NUMBER']).strip(),
                    'station': str(row['STATION']).strip().upper(),
                    'current_status': 'Vacant', # डिफ़ॉल्ट रूप से Vacant
                    'last_occupant_id': None,
                    'created_at': firestore.SERVER_TIMESTAMP
                })
            batch.commit()
            st.success(f"सफलता: कुल {len(df_inventory)} क्वार्टर मास्टर रजिस्टर में जोड़े गए।")
            st.cache_data.clear()
            
    except FileNotFoundError:
        st.warning(f"Warning: Quarter Register CSV file not found at {INVENTORY_CSV_PATH}. Master Quarters not loaded.")
    except Exception as e:
        st.error(f"Error loading Master Quarters to Firestore: {e}")
        return False
    
    # 2. History Loading (यदि आवश्यक हो)
    # यह खंड जटिल है, इसे केवल तभी उपयोग करें जब आपके पास सही HISTORY_CSV_PATH हो।
    # अन्यथा, नया इतिहास ऐप के माध्यम से बनाया जाएगा।
    # try:
    #     history_count = len(list(db.collection(QUARTER_HISTORY_COLLECTION).limit(1).get()))
    #     if history_count == 0 and os.path.exists(HISTORY_CSV_PATH):
    #         st.info(f"CSV से प्रारंभिक इतिहास ({HISTORY_CSV_PATH}) को Firestore में लोड किया जा रहा है...")
    #         # ... (History CSV प्रोसेसिंग और अपलोड लॉजिक यहाँ आएगा)
    # except Exception:
    #     pass # त्रुटियों को अनदेखा करें यदि इतिहास फ़ाइल मौजूद नहीं है।
    
    return True


# ----------------------------------------------------------------------
# 2. FIREBASE QUARTER DATA ACCESS (नया)
# ----------------------------------------------------------------------

@st.cache_data(ttl=5) 
def get_all_quarters():
    """Firestore से सभी क्वार्टर और उनके स्टेटस फ़ेच करता है।"""
    if db is None:
        # DB विफल होने पर भी कॉलम के साथ खाली DF लौटाएँ
        return pd.DataFrame(columns=['quarter_number', 'station', 'current_status'])
    
    try:
        docs = db.collection(QUARTER_MASTER_COLLECTION).stream()
        data = [doc.to_dict() for doc in docs]
        
        # 🚨 सुरक्षा जांच: यदि कोई डेटा नहीं है, तो आवश्यक कॉलम के साथ एक खाली DF लौटाएँ
        if not data:
            return pd.DataFrame(columns=['quarter_number', 'station', 'current_status'])
            
        df = pd.DataFrame(data) 
        
        # सुनिश्चित करें कि 'current_status' मौजूद है (डिफ़ॉल्ट मान दें यदि किसी दस्तावेज़ में वह फ़ील्ड छूट गया हो)
        if 'current_status' not in df.columns:
            df['current_status'] = 'Vacant'
        
        # सभी आवश्यक कॉलम सुनिश्चित करें (यदि कोई दस्तावेज़ अधूरा है)
        required_cols = ['quarter_number', 'station', 'current_status']
        for col in required_cols:
            if col not in df.columns:
                df[col] = 'N/A' # मिसिंग कॉलम को N/A से भरें
                
        # DataFrame को स्टेशन और क्वार्टर_नंबर से छाँटें
        df = df.sort_values(by=['station', 'quarter_number']).reset_index(drop=True)
            
        return df[required_cols] # केवल आवश्यक कॉलम लौटाएँ
        
    except Exception as e:
        st.error(f"Error fetching quarters from Firestore: {e}")
        # त्रुटि पर भी कॉलम के साथ खाली DF लौटाएँ
        return pd.DataFrame(columns=['quarter_number', 'station', 'current_status'])

@st.cache_data(ttl=5)
def get_quarter_history_df():
    """Firestore से सभी क्वार्टर इतिहास फ़ेच करता है।"""
    if db is None: return pd.DataFrame()
    
    try:
        docs = db.collection(QUARTER_HISTORY_COLLECTION).stream()
        data = []
        for doc in docs:
             record = doc.to_dict()
             record['id'] = doc.id
             data.append(record)
             
        df = pd.DataFrame(data)
        
        # DATE फ़ील्ड को स्ट्रिंग में बदलें (रिपोर्टिंग के लिए)
        if 'allotment_date' in df.columns:
             df['allotment_date'] = df['allotment_date'].apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, date) else x)
        if 'vacation_date' in df.columns:
             df['vacation_date'] = df['vacation_date'].apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, date) else x)
             
        return df.sort_values(by=['station', 'quarter_number', 'allotment_date'], ascending=[True, True, False])
        
    except Exception as e:
        st.error(f"Error fetching history from Firestore: {e}")
        return pd.DataFrame()

# ----------------------------------------------------------------------
# 3. FIREBASE EMPLOYEE DATA LOOKUP (अपरिवर्तित)
# ----------------------------------------------------------------------

@st.cache_data(ttl=3600)
def get_employee_details_from_firebase(hrms_id):
    """Firebase से HRMS ID द्वारा कर्मचारी का पूरा विवरण प्राप्त करता है।"""
    if db is None or not hrms_id: return None
    
    hrms_id_str = str(hrms_id).strip()
    
    try:
        docs = db.collection(EMPLOYEE_COLLECTION).where('HRMS ID', '==', hrms_id_str).limit(1).get()
        
        if not docs:
            return None
            
        record = docs[0].to_dict()
        
        return {
            'hrms_id': hrms_id_str,
            'employee_name_english': str(record.get('Employee Name', 'NA')),
            'designation_english': str(record.get('Designation', 'NA')),
            'employee_name_hindi': str(record.get('Employee Name in Hindi', str(record.get('Employee Name', 'NA')))),
            'designation_hindi': str(record.get('Designation in Hindi', str(record.get('Designation', 'NA')))),
            'pf_number': str(record.get('PF Number', 'NA')),
            'unit': str(record.get('Unit', 'NA'))
        }
    except Exception as e:
        st.error(f"Error fetching employee {hrms_id} from Firebase: {e}")
        return None

# ----------------------------------------------------------------------
# 4. WORD FILE GENERATION (अपरिवर्तित)
# ----------------------------------------------------------------------

def generate_word_file(template_name, data):
    # (पिछला कोड यहाँ)
    current_date = datetime.date.today()
    current_date_str_letter = current_date.strftime('%d/%m/%Y') 
    
    template_path = f'{template_name}.docx'
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}. Please upload it to your GitHub root.")
        return None

    try:
        document = Document(template_path)
        
        # Ensure all keys have a default value
        replacements = {
            '{{DATE}}': current_date_str_letter,
            '{{QUARTER_NUMBER}}': data.get('quarter_number', 'NA'),
            '{{STATION}}': data.get('station', 'NA'),
            '{{EMPLOYEE_NAME}}': data.get('employee_name_hindi', data.get('employee_name_english', 'NA')), 
            '{{DESIGNATION}}': data.get('designation_hindi', data.get('designation_english', 'NA')), 
            '{{HRMS_ID}}': data.get('hrms_id', 'NA'),
            '{{PF_Number}}': data.get('pf_number', 'NA'),
            '{{UNIT}}': data.get('unit', 'NA'),
        }

        # Replace text in paragraphs
        for paragraph in document.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
        
        # Replace text in tables
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for key, value in replacements.items():
                            if key in paragraph.text:
                                paragraph.text = paragraph.text.replace(key, str(value))
        
        file_stream = io.BytesIO()
        document.save(file_stream)
        file_stream.seek(0)
        
        return file_stream

    except Exception as e:
        st.error(f"Failed to generate Word file: {e}")
        return None

# ----------------------------------------------------------------------
# 5. CORE LOGIC (Firebase Update)
# ----------------------------------------------------------------------

def allot_quarter(quarter_num, station, hrms_id, allot_date):
    """क्वार्टर को अलॉट करता है और Firestore में मास्टर और हिस्ट्री अपडेट करता है।"""
    if db is None: return False, "Database connection failed."

    employee_details = get_employee_details_from_firebase(hrms_id)
    if employee_details is None:
        return False, f"Error: Employee details not found for HRMS ID: {hrms_id}."
    
    # Firestore में क्वार्टर का दस्तावेज़ संदर्भ (Document Reference)
    q_doc_id = f"{station}_{quarter_num}"
    q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)
    
    try:
        batch = db.batch()
        
        # 1. Check for Duplicate Allotment (History Collection)
        docs_dup = db.collection(QUARTER_HISTORY_COLLECTION)\
                     .where('hrms_id', '==', hrms_id)\
                     .where('is_current', '==', True).limit(1).get()
        if docs_dup:
            return False, f"Error: Employee ({hrms_id}) already occupies quarter {docs_dup[0].to_dict().get('quarter_number', 'N/A')}."

        # 2. Check quarter status (Master Collection)
        q_doc = q_doc_ref.get()
        if not q_doc.exists:
            return False, f"Error: Quarter {quarter_num} at {station} not found in Master Register."
            
        q_data = q_doc.to_dict()
        if q_data.get('current_status') == 'Occupied':
            return False, f"Warning: Quarter {quarter_num} at {station} is already occupied."

        # A. Master Register Update (Set status to Occupied)
        batch.update(q_doc_ref, {
            'current_status': 'Occupied', 
            'last_occupant_id': hrms_id,
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        # B. History Log Insert (New document with is_current = True)
        history_data = {
            'quarter_number': quarter_num, 
            'station': station, 
            'hrms_id': hrms_id, 
            'pf_number': employee_details['pf_number'], 
            'designation': employee_details['designation_english'], 
            'unit': employee_details['unit'], 
            'employee_name': employee_details['employee_name_english'], 
            'allotment_date': allot_date, # Date object
            'vacation_date': None,
            'is_current': True,
            'created_at': firestore.SERVER_TIMESTAMP
        }
        # नया डॉक्यूमेंट जेनरेट करें (कोई कस्टम ID नहीं)
        batch.set(db.collection(QUARTER_HISTORY_COLLECTION).document(), history_data) 
        
        batch.commit()
        
        # C. Generate Word Allotment Letter
        file_stream = generate_word_file("Allotment_Template", employee_details | {"quarter_number": quarter_num, "station": station})
        
        return True, file_stream

    except Exception as e:
        st.error(f"Allotment failed: {e}")
        return False, f"Allotment failed: {e}"


def vacate_quarter(quarter_num, station, vacate_date):
    """क्वार्टर को खाली करता है और Firestore में मास्टर और हिस्ट्री अपडेट करता है।"""
    if db is None: return False, "Database connection failed."

    q_doc_id = f"{station}_{quarter_num}"
    q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)

    try:
        batch = db.batch()
        
        # A. Get current occupant history record
        docs_history = db.collection(QUARTER_HISTORY_COLLECTION)\
                         .where('quarter_number', '==', quarter_num)\
                         .where('station', '==', station)\
                         .where('is_current', '==', True).limit(1).get()

        if not docs_history:
            return False, f"Warning: Quarter {quarter_num} at {station} is not currently occupied in the database."

        history_doc = docs_history[0]
        history_data = history_doc.to_dict()
        hrms_id = history_data['hrms_id']
        
        # B. Look up Hindi name/designation from FIREBASE (मेमो जेनरेट करने के लिए)
        employee_details_full = get_employee_details_from_firebase(hrms_id)
        
        if employee_details_full is None:
            employee_details_full = history_data
            employee_details_full['employee_name_english'] = employee_details_full.pop('employee_name')
            employee_details_full['designation_english'] = employee_details_full.pop('designation')
            st.warning("Firebase से कर्मचारी विवरण नहीं मिला। मेमो में इतिहास डेटा का उपयोग किया जाएगा।")

        # C. History Log Update: is_current को FALSE और vacation_date सेट करें
        batch.update(history_doc.reference, {
            'vacation_date': vacate_date, # Date object
            'is_current': False,
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        # D. Master Register Update: current_status को 'Vacant' सेट करें
        batch.update(q_doc_ref, {
            'current_status': 'Vacant', 
            # last_occupant_id को वर्तमान ऑक्यूपेंट की ID पर ही रखें या चाहें तो हटा दें
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        batch.commit()
        st.cache_data.clear() # कैश साफ़ करें
        
        # E. Generate Word Vacation Memo
        template_data = employee_details_full | {"quarter_number": quarter_num, "station": station}
        file_stream = generate_word_file("Vacation_Template", template_data)

        return True, file_stream

    except Exception as e:
        st.error(f"FATAL DB ERROR: Vacation transaction failed. Details: {e}") 
        return False, f"Vacation failed: {e}"

# ----------------------------------------------------------------------
# 6. REPORTING (Firebase Update)
# ----------------------------------------------------------------------

def generate_current_status_report():
    # यह फ़ंक्शन अब get_all_quarters() पर निर्भर करता है, जो Firebase से डेटा लाता है।
    df_master = get_all_quarters()
    df_history = get_quarter_history_df()
    
    if df_master.empty: return pd.DataFrame()
    
    # मास्टर और इतिहास को hrms_id के माध्यम से जोड़ें
    df_current_occupants = df_history[df_history['is_current'] == True].rename(columns={
        'employee_name': 'current_occupant',
        'hrms_id': 'occupant_hrms_id',
        'pf_number': 'pf_number',
        'designation': 'designation',
        'unit': 'unit',
        'allotment_date': 'allotment_date'
    })
    
    # केवल आवश्यक कॉलम रखें ताकि मर्ज साफ हो
    df_current_occupants = df_current_occupants[['quarter_number', 'station', 'current_occupant', 
                                                 'occupant_hrms_id', 'pf_number', 'designation', 
                                                 'unit', 'allotment_date']]

    # मास्टर और इतिहास को जोड़ें
    df_report = df_master.merge(
        df_current_occupants, 
        on=['quarter_number', 'station'], 
        how='left'
    ).rename(columns={'occupant_hrms_id': 'hrms_id'})
    
    # None/NaN को N/A से बदलें
    df_report = df_report.fillna('N/A')
    
    # अंतिम कॉलम क्रम
    display_cols = ['quarter_number', 'station', 'current_status', 'current_occupant', 
                    'hrms_id', 'pf_number', 'designation', 'unit', 'allotment_date']
    
    return df_report[[col for col in display_cols if col in df_report.columns]]

def generate_full_history_report():
    df_history = get_quarter_history_df()
    if df_history.empty: return pd.DataFrame()
    
    # रिपोर्ट के लिए कॉलम का क्रम और मान सेट करें
    df_history['vacation_date'] = df_history.apply(
        lambda row: 'CURRENTLY OCCUPIED' if row['is_current'] else row['vacation_date'], axis=1
    )
    
    df_history['record_type'] = df_history['is_current'].apply(
        lambda x: 'Current Occupant' if x else 'History Record'
    )
    
    display_cols = [
        'quarter_number', 'station', 'employee_name', 'hrms_id', 'pf_number', 
        'designation', 'unit', 'allotment_date', 'vacation_date', 'record_type'
    ]
    
    return df_history[[col for col in display_cols if col in df_history.columns]].fillna('N/A')

# ----------------------------------------------------------------------
# 7. AUTHENTICATION & UI
# ----------------------------------------------------------------------

def check_password(password):
    return password == CORRECT_PASSWORD

def authenticate_user():
    """लॉगिन UI प्रदर्शित करता है और प्रमाणीकरण संभालता है।"""
    
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        
        st.sidebar.markdown("## Login Required")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("## 🔑 रेलवे क्वार्टर प्रबंधन लॉगिन")
            
            with st.form("login_form"):
                st.markdown("---")
                password = st.text_input("पासवर्ड दर्ज करें", type="password")
                submitted = st.form_submit_button("लॉगिन")

                if submitted:
                    if check_password(password):
                        st.session_state.authenticated = True
                        st.success("Login Successful!")
                        st.rerun() 
                    else:
                        st.error("Invalid Password.")
    
    # --- UI Start after Authentication ---
    if st.session_state.authenticated:
        # 1. डेटा माइग्रेशन (एक बार चलाएँ)
        if not db:
            st.error("Firebase DB connection failed.")
            st.stop()
        
        load_initial_data_to_firestore() # CSV से Firestore में डेटा लोड करें
        
        # 2. मुख्य UI
        st.sidebar.markdown("---")
        if st.sidebar.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.rerun()

        # --- UI TABS ---
        tab_allot, tab_vacate, tab_report = st.tabs(["🔑 क्वार्टर अलॉट करें", "🔓 क्वार्टर खाली करें", "📈 रिपोर्ट"])

        # -----------------------------------------------------------
        # I. Allot Quarter Tab
        # -----------------------------------------------------------
        with tab_allot:
            st.header("🔑 नया क्वार्टर अलॉट करें")
            
            quarters_df = get_all_quarters()
            vacant_quarters = quarters_df[quarters_df['current_status'] == 'Vacant']
            
            if vacant_quarters.empty:
                st.warning("अलॉट करने के लिए कोई खाली क्वार्टर उपलब्ध नहीं है।")
                st.stop()

            vacant_quarters['Display'] = vacant_quarters['quarter_number'] + ' (' + vacant_quarters['station'] + ')'
            selected_quarter_display = st.selectbox("खाली क्वार्टर चुनें", vacant_quarters['Display'].tolist(), key='allot_q_select')

            if selected_quarter_display:
                selected_q_num = selected_quarter_display.split(' (')[0].strip()
                selected_station = selected_quarter_display.split('(')[-1].strip(')')
                
                st.info(f"चयनित: **{selected_q_num}** (स्टेशन: **{selected_station}**)")

                with st.form("allotment_form"):
                    hrms_id_input = st.text_input("कर्मचारी का HRMS ID दर्ज करें (Firebase Lookup)", key='allot_hrms_id').strip()
                    # अलॉटमेंट तिथि को date.today() के रूप में लें (यह Firestore में Date Object के रूप में सेव होगा)
                    allot_date = st.date_input("अलॉटमेंट तिथि", datetime.date.today(), key='allot_date') 
                    
                    st.markdown("---")
                    
                    employee_details_display = None
                    if hrms_id_input:
                        employee_details_display = get_employee_details_from_firebase(hrms_id_input)
                    
                    if hrms_id_input and employee_details_display:
                        st.success("✅ कर्मचारी विवरण Firebase से सफलतापूर्वक प्राप्त हुआ।")
                        st.json({
                            "Name": employee_details_display['employee_name_english'],
                            "Designation": employee_details_display['designation_english'],
                            "PF No.": employee_details_display['pf_number'],
                            "Unit": employee_details_display['unit']
                        })
                    elif hrms_id_input:
                         st.error(f"❌ HRMS ID: {hrms_id_input} का विवरण Firebase में नहीं मिला।")

                    
                    submitted = st.form_submit_button("🔑 क्वार्टर अलॉट करें और लेटर जेनरेट करें")

                    if submitted:
                        if hrms_id_input and employee_details_display:
                            with st.spinner("अलॉटमेंट संसाधित किया जा रहा है..."):
                                success, result = allot_quarter(selected_q_num, selected_station, hrms_id_input, allot_date)
                            
                            if success:
                                st.success("🎉 अलॉटमेंट सफलतापूर्वक पूरा हुआ!")
                                file_stream = result
                                st.download_button(
                                    label="डाउनलोड अलॉटमेंट लेटर (.docx)",
                                    data=file_stream,
                                    file_name=f"Allotment_Letter_{selected_q_num}_{hrms_id_input}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key='download_allotment'
                                )
                                st.cache_data.clear()
                                st.rerun()
                            else:
                                st.error(result)
                        else:
                            st.error("HRMS ID दर्ज करें और सुनिश्चित करें कि विवरण Firebase से सफलतापूर्वक प्राप्त हुआ है।")

        # -----------------------------------------------------------
        # II. Vacate Quarter Tab
        # -----------------------------------------------------------
        with tab_vacate:
            st.header("🔓 क्वार्टर खाली करें")
            
            quarters_df = get_all_quarters()
            occupied_quarters = quarters_df[quarters_df['current_status'] == 'Occupied']
            
            if occupied_quarters.empty:
                st.warning("खाली करने के लिए कोई ऑक्यूपाइड क्वार्टर नहीं है।")
                st.stop()

            occupied_quarters['Display'] = occupied_quarters['quarter_number'] + ' (' + occupied_quarters['station'] + ')'
            selected_occupied_display = st.selectbox("ऑक्यूपाइड क्वार्टर चुनें", occupied_quarters['Display'].tolist(), key='vacate_q_select')

            if selected_occupied_display:
                selected_q_num = selected_occupied_display.split(' (')[0].strip()
                selected_station = selected_occupied_display.split('(')[-1].strip(')')
                
                st.info(f"चयनित: **{selected_q_num}** (स्टेशन: **{selected_station}**)")

                with st.form("vacation_form"):
                    vacate_date = st.date_input("वेकेशन तिथि", datetime.date.today(), key='vacate_date')
                    st.warning("पुष्टि करें: यह क्वार्टर खाली हो जाएगा और इतिहास अपडेट हो जाएगा।")
                    
                    submitted = st.form_submit_button("🔓 क्वार्टर खाली करें और मेमो जेनरेट करें")

                    if submitted:
                        with st.spinner("वेकेशन संसाधित किया जा रहा है..."):
                            success, result = vacate_quarter(selected_q_num, selected_station, vacate_date)
                        
                        if success:
                            st.success("🎉 वेकेशन सफलतापूर्वक पूरा हुआ!")
                            file_stream = result
                            st.download_button(
                                label="डाउनलोड वेकेशन मेमो (.docx)",
                                data=file_stream,
                                file_name=f"Vacation_Memo_{selected_q_num}_{selected_station}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key='download_vacation'
                            )
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error(result)

        # -----------------------------------------------------------
        # III. Report Tab
        # -----------------------------------------------------------
        with tab_report:
            st.header("📈 क्वार्टर स्टेटस और इतिहास रिपोर्ट")
            
            st.subheader("1. वर्तमान क्वार्टर स्टेटस")
            df_status = generate_current_status_report()
            st.dataframe(df_status, width='stretch', hide_index=True)
            
            st.subheader("2. संपूर्ण क्वार्टर इतिहास (Allotment & Vacation)")
            df_history = generate_full_history_report()
            st.dataframe(df_history, width='stretch', hide_index=True)

            col_rep_dl1, col_rep_dl2 = st.columns(2)
            
            with col_rep_dl1:
                csv_status = df_status.to_csv(index=False, encoding='utf-8').encode('utf-8')
                st.download_button(
                    label="वर्तमान स्टेटस CSV डाउनलोड करें",
                    data=csv_status,
                    file_name='quarter_status_report.csv',
                    mime='text/csv',
                    key='dl_status'
                )
            
            with col_rep_dl2:
                csv_history = df_history.to_csv(index=False, encoding='utf-8').encode('utf-8')
                st.download_button(
                    label="संपूर्ण इतिहास CSV डाउनलोड करें",
                    data=csv_history,
                    file_name='quarter_history_report.csv',
                    mime='text/csv',
                    key='dl_history'
                )
# --- UI CODE END ---

authenticate_user()

