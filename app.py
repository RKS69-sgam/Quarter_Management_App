import streamlit as st
import datetime
import os 
import pandas as pd
from docx import Document 
import io
import firebase_admin
from firebase_admin import credentials, firestore
import json
from datetime import date
import uuid
import re # रेगुलर एक्सप्रेशन के लिए

# ----------------------------------------------------------------------
# 0. कॉन्फ़िगरेशन (Config)
# ----------------------------------------------------------------------

INVENTORY_CSV_PATH = "data/Quarter_Register.csv" 
EMPLOYEE_COLLECTION = "employees"          
QUARTER_MASTER_COLLECTION = "master_quarters"
QUARTER_HISTORY_COLLECTION = "quarter_history" 

# --- SECURITY CONFIGURATION ---
CORRECT_PASSWORD = "Sgam@1234" 
# ------------------------------

st.set_page_config(layout="wide", page_title="रेलवे क्वार्टर प्रबंधन (Firebase Firestore)")

# ----------------------------------------------------------------------
# 1. FIREBASE CONNECTION & INITIALIZATION
# ----------------------------------------------------------------------

@st.cache_resource
def initialize_firebase():
    """Firebase SDK को इनिशियलाइज़ करता है और Firestore क्लाइंट लौटाता है।"""
    try:
        if not firebase_admin._apps:
            # Secrets.toml से क्रेडेंशियल्स लोड करने का लॉजिक (जैसा कि पहले था)
            if st.secrets.get("firebase_config"):
                service_account_info_attrdict = st.secrets["firebase_config"]
                final_credentials = dict(service_account_info_attrdict)
                if isinstance(final_credentials.get('private_key'), str):
                     final_credentials['private_key'] = final_credentials['private_key'].replace('\\n', '\n')
                cred = credentials.Certificate(final_credentials)
                
                # Check for project_id to initialize app
                if not final_credentials.get('project_id'):
                    st.error("Project ID missing in firebase_config. Check secrets.toml.")
                    return None
            
                firebase_admin.initialize_app(cred, {'projectId': final_credentials['project_id']})
            else:
                st.error("Firebase Secrets not found. Please configure secrets.toml.")
                return None
            
        return firestore.client()
        
    except Exception as e:
        st.error(f"❌ Firebase कनेक्शन विफल। त्रुटि: {e}")
        return None

db = initialize_firebase()

# ----------------------------------------------------------------------
# 1.1. CLEANING FUNCTION
# ----------------------------------------------------------------------
def clean_data_string(s):
    """Firestore पाथ्स के लिए स्ट्रिंग को साफ़ करता है (स्लैश, स्पेस, हाइफ़न हटाता है)।"""
    if s is None:
        return 'NA'
    s = str(s).strip().upper()
    s = re.sub(r'[/\\-]', '_', s) # स्लैश, बैकस्लैश, हाइफ़न को अंडरस्कोर से बदलें
    s = re.sub(r'\s+', '', s)   # सभी स्पेस हटा दें
    return s

# ----------------------------------------------------------------------
# 1.2. SINGLE-PASS DATA MIGRATION TO FIRESTORE
# ----------------------------------------------------------------------

def load_initial_data_to_firestore():
    """
    CSV से क्वार्टर इन्वेंट्री को लोड करता है और ऑक्यूपेंसी के आधार पर history बनाता है।
    """
    if db is None: return False
    
    try:
        # केवल तभी चलाएँ जब Master Collection खाली हो
        master_count = len(list(db.collection(QUARTER_MASTER_COLLECTION).limit(1).get()))
        
        if master_count == 0:
            st.info(f"CSV ({INVENTORY_CSV_PATH}) से Master Quarters और History को लोड किया जा रहा है...")
            df_inventory = pd.read_csv(INVENTORY_CSV_PATH)
            df_inventory.columns = df_inventory.columns.str.strip().str.upper()
            
            required_cols = ['QUARTER_NUMBER', 'STATION', 'IS_OCCUPIED', 'HRMS_ID', 'ALLOTMENT_DATE', 'EMPLOYEE_NAME']
            if not all(col in df_inventory.columns for col in required_cols):
                 st.error(f"CSV Error: आवश्यक कॉलम {required_cols} नहीं मिले।")
                 return False

            batch = db.batch()
            master_count = 0
            history_count = 0
            
            for index, row in df_inventory.iterrows():
                # Firestore ID और Master Data के लिए डेटा साफ़ करें
                clean_station = clean_data_string(row['STATION'])
                clean_quarter = clean_data_string(row['QUARTER_NUMBER'])
                
                q_doc_id = f"{clean_station}_{clean_quarter}"
                q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)
                
                is_occupied_str = str(row['IS_OCCUPIED']).strip().upper()
                is_occupied = is_occupied_str in ('YES', 'TRUE', '1')

                # --- 1. Master Register Creation ---
                master_status = 'Occupied' if is_occupied else 'Vacant'
                last_occupant_id = clean_data_string(row['HRMS_ID']) if is_occupied else None
                
                master_data = {
                    'quarter_number': clean_quarter,
                    'station': clean_station,
                    'current_status': master_status, 
                    'last_occupant_id': last_occupant_id,
                    'created_at': firestore.SERVER_TIMESTAMP
                }
                batch.set(q_doc_ref, master_data)
                master_count += 1

                # --- 2. History Record Creation (if Occupied) ---
                if is_occupied:
                    try:
                        allot_date_str = str(row['ALLOTMENT_DATE']).strip()
                        allot_date_obj = datetime.datetime.strptime(allot_date_str, '%Y-%m-%d').date()
                    except ValueError:
                         st.error(f"Invalid ALLOTMENT_DATE format ({allot_date_str}) for quarter {q_doc_id}. Skipping history.")
                         continue
                         
                    # CSV से अतिरिक्त फ़ील्ड (यदि मौजूद हों)
                    pf_num = str(row.get('PF_NUMBER', 'NA')).strip()
                    designation = str(row.get('DESIGNATION', 'NA')).strip()

                    history_data = {
                        'quarter_number': clean_quarter, 
                        'station': clean_station, 
                        'hrms_id': clean_data_string(row['HRMS_ID']), 
                        'employee_name': str(row['EMPLOYEE_NAME']).strip(), 
                        'pf_number': pf_num, 
                        'designation': designation, 
                        'unit': row.get('UNIT', 'NA'), # यदि CSV में UNIT कॉलम है
                        'allotment_date': allot_date_obj,
                        'vacation_date': None,
                        'is_current': True,
                        'created_at': firestore.SERVER_TIMESTAMP
                    }
                    # history कलेक्शन में नया डॉक्यूमेंट सेट करें
                    batch.set(db.collection(QUARTER_HISTORY_COLLECTION).document(), history_data)
                    history_count += 1

            batch.commit()
            st.success(f"🎉 सफलता: {master_count} Master Quarters और {history_count} History Records लोड किए गए।")
            st.cache_data.clear()
            
    except FileNotFoundError:
        st.warning(f"Warning: Quarter Register CSV file not found at {INVENTORY_CSV_PATH}. Data not loaded.")
    except Exception as e:
        st.error(f"Error loading initial data to Firestore: {e}")
        return False
    
    return True

# ----------------------------------------------------------------------
# 2. FIREBASE DATA ACCESS (मॉडिफ़ाईड/रोबस्ट)
# ----------------------------------------------------------------------

# (यहां `get_all_quarters` और `get_quarter_history_df` फ़ंक्शंस पिछले अपडेटेड और रोबस्ट वर्शन से रहेंगे।)
# ... (पिछली बार के get_all_quarters और get_quarter_history_df फ़ंक्शन का कोड यहाँ कॉपी करें)
# ...

# ----------------------------------------------------------------------
# 3. FIREBASE EMPLOYEE DATA LOOKUP (संशोधित)
# ----------------------------------------------------------------------

# इस फ़ंक्शन को अब नाम या HRMS ID से खोज करने की अनुमति देने के लिए संशोधित किया गया है।
@st.cache_data(ttl=3600)
def search_employee_details_from_firebase(search_term):
    """Firebase में नाम या HRMS ID द्वारा कर्मचारी खोजता है।"""
    if db is None or not search_term: return pd.DataFrame()
    
    search_term = str(search_term).strip()
    results = []

    # HRMS ID से exact मैच की जाँच करें
    try:
        docs_id = db.collection(EMPLOYEE_COLLECTION)\
                    .where('HRMS ID', '==', search_term)\
                    .limit(1).get()
        if docs_id:
            results.append(docs_id[0].to_dict())
            
    except Exception:
        # यदि HRMS ID सर्च विफल होता है तो आगे बढ़ें
        pass
        
    # नाम से खोज (Contains logic Firestore में मुश्किल है, इसलिए हम 'starts with' या क्लाइंट साइड फ़िल्टरिंग का उपयोग करते हैं)
    # Firebase में 'starts with' के लिए केवल एक फ़ील्ड पर क्वेरी की अनुमति है।
    # हम यहाँ 'HRMS ID' और 'Employee Name' दोनों को एक साथ खोज नहीं सकते,
    # इसलिए हम HRMS ID की जाँच करने के बाद, नाम से खोजने के लिए 'starts with' का प्रयास करते हैं।
    if not results and len(search_term) >= 3:
        try:
             # नाम से खोज: Firestore में रेंज क्वेरी के रूप में 'starts with'
             start_key = search_term
             end_key = search_term + '\uf8ff' # Unicode trick for 'starts with'

             docs_name = db.collection(EMPLOYEE_COLLECTION)\
                           .where('Employee Name', '>=', start_key)\
                           .where('Employee Name', '<=', end_key)\
                           .limit(20).get() # लिमिट 20 परिणाम

             for doc in docs_name:
                 doc_data = doc.to_dict()
                 # सुनिश्चित करें कि हम पहले से मिले परिणाम को दोबारा न जोड़ें
                 if doc_data not in results:
                     results.append(doc_data)

        except Exception as e:
            st.warning(f"Error during employee name search: {e}")
            
    # परिणामों को DataFrame में बदलें
    if not results:
        return pd.DataFrame()
        
    df = pd.DataFrame(results)
    
    # आवश्यक लुकअप फ़ील्ड सुनिश्चित करें
    lookup_cols = ['HRMS ID', 'Employee Name', 'Designation', 'Unit', 'PF Number']
    for col in lookup_cols:
        if col not in df.columns:
            df[col] = 'N/A'
            
    return df[['HRMS ID', 'Employee Name', 'Designation', 'Unit', 'PF Number']]


@st.cache_data(ttl=3600)
def get_employee_details_by_hrms_id(hrms_id):
    """केवल HRMS ID से एक ही कर्मचारी का विवरण प्राप्त करता है।"""
    if db is None or not hrms_id: return None
    hrms_id_str = clean_data_string(hrms_id)
    
    try:
        docs = db.collection(EMPLOYEE_COLLECTION).where('HRMS ID', '==', hrms_id_str).limit(1).get()
        if not docs: return None
            
        record = docs[0].to_dict()
        
        return {
            'hrms_id': hrms_id_str,
            'employee_name_english': str(record.get('Employee Name', 'NA')),
            'designation_english': str(record.get('Designation', 'NA')),
            # यदि हिंदी फ़ील्ड नहीं है तो अंग्रेजी का उपयोग करें
            'employee_name_hindi': str(record.get('Employee Name in Hindi', str(record.get('Employee Name', 'NA')))),
            'designation_hindi': str(record.get('Designation in Hindi', str(record.get('Designation', 'NA')))),
            'pf_number': str(record.get('PF Number', 'NA')),
            'unit': str(record.get('Unit', 'NA'))
        }
    except Exception as e:
        st.error(f"Error fetching employee {hrms_id} by ID: {e}")
        return None

# ----------------------------------------------------------------------
# 4. WORD FILE GENERATION & CORE LOGIC (थोड़ा संशोधित)
# ----------------------------------------------------------------------

# (यहां `generate_word_file`, `allot_quarter`, और `vacate_quarter` फ़ंक्शंस पिछले वर्शन से रहेंगे,
#  लेकिन `allot_quarter` अब `get_employee_details_by_hrms_id` का उपयोग करेगा)
# ...

# ----------------------------------------------------------------------
# 5. UI CHANGES (अलॉटमेंट सर्च)
# ----------------------------------------------------------------------

# UI में:
# I. Allot Quarter Tab
# ...
            with tab_allot:
                # ... (बाकी UI कोड)
                
                with st.form("allotment_form"):
                    
                    st.subheader("1. कर्मचारी खोजें (नाम या HRMS ID)")
                    search_term = st.text_input("कर्मचारी नाम या HRMS ID दर्ज करें (कम से कम 3 अक्षर)", key='allot_search_term').strip()
                    
                    selected_hrms_id = None
                    employee_details_display = None
                    
                    if len(search_term) >= 3:
                        df_search_results = search_employee_details_from_firebase(search_term)
                        
                        if not df_search_results.empty:
                            df_search_results['Display'] = df_search_results['HRMS ID'] + ' - ' + df_search_results['Employee Name'] + ' (' + df_search_results['Designation'] + ')'
                            selected_display = st.selectbox("परिणामों में से कर्मचारी चुनें", df_search_results['Display'].tolist(), key='allot_q_select_hrms')
                            
                            if selected_display:
                                selected_hrms_id = selected_display.split(' - ')[0].strip()
                                st.info(f"चयनित कर्मचारी HRMS ID: **{selected_hrms_id}**")
                                
                                # चयनित HRMS ID द्वारा विवरण लोड करें
                                employee_details_display = get_employee_details_by_hrms_id(selected_hrms_id)
                                
                        else:
                            st.warning("कोई कर्मचारी विवरण नहीं मिला। Employee Master Collection की जाँच करें।")
                    elif len(search_term) > 0:
                        st.info("खोजने के लिए कम से कम 3 अक्षर दर्ज करें।")

                    st.subheader("2. क्वार्टर और तिथि")
                    
                    if selected_hrms_id:
                        allot_date = st.date_input("अलॉटमेंट तिथि", datetime.date.today(), key='allot_date') 
                        
                        if employee_details_display:
                            st.success("✅ कर्मचारी विवरण सफलतापूर्वक प्राप्त हुआ।")
                            st.json({
                                "Name": employee_details_display['employee_name_english'],
                                "Designation": employee_details_display['designation_english'],
                                "PF No.": employee_details_display['pf_number'],
                                "Unit": employee_details_display['unit']
                            })
                        
                        submitted = st.form_submit_button("🔑 क्वार्टर अलॉट करें और लेटर जेनरेट करें")

                        if submitted:
                            # अलॉटमेंट लॉजिक (यहां allot_quarter को कॉल करें)
                            if selected_hrms_id:
                                # ... (allot_quarter(selected_q_num, selected_station, selected_hrms_id, allot_date) को कॉल करें)
                                pass # (यह लॉजिक पिछले कोड से आएगा)
                            else:
                                st.error("कृपया एक कर्मचारी चुनें।")
# ...
