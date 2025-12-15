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
import re 
import uuid

# ----------------------------------------------------------------------
# 0. कॉन्फ़िगरेशन (Config)
# ----------------------------------------------------------------------

# NOTE: सुनिश्चित करें कि ये फ़ाइलें /data फ़ोल्डर में मौजूद हैं।
INVENTORY_CSV_PATH = "data/Quarter_Register.csv" 
# संग्रह का नाम आपके बताए अनुसार "employees" सेट किया गया है।
EMPLOYEE_COLLECTION = "employees"          
QUARTER_MASTER_COLLECTION = "master_quarters"
QUARTER_HISTORY_COLLECTION = "quarter_history" 

# --- SECURITY CONFIGURATION ---
# पासवर्ड secrets.toml से लोड करें
CORRECT_PASSWORD = st.secrets.get("CORRECT_PASSWORD", "default_password")
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
            if st.secrets.get("firebase_config"):
                service_account_info_attrdict = st.secrets["firebase_config"]
                final_credentials = dict(service_account_info_attrdict)
                
                # Private key में Newline characters को संभालें
                if isinstance(final_credentials.get('private_key'), str):
                     final_credentials['private_key'] = final_credentials['private_key'].replace('\\n', '\n')
                cred = credentials.Certificate(final_credentials)
            
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
    """Firestore ID/पाथ्स के लिए स्ट्रिंग को साफ़ करता है।"""
    if s is None or pd.isna(s):
        return 'NA'
    s = str(s).strip()
    s = re.sub(r'[/\\-]', '_', s)
    s = re.sub(r'\s+', '', s)
    return s.upper()

# ----------------------------------------------------------------------
# 1.2. SINGLE-PASS DATA MIGRATION TO FIRESTORE
# ----------------------------------------------------------------------

def load_initial_data_to_firestore():
    """
    CSV से Master Quarters और History को लोड करता है (एक ही पास में)।
    यह केवल तब चलता है जब Master Collection खाली हो।
    """
    if db is None: return False
    
    try:
        # केवल तभी चलाएँ जब Master Collection खाली हो
        master_count = len(list(db.collection(QUARTER_MASTER_COLLECTION).limit(1).get()))
        
        if master_count == 0:
            st.info(f"CSV ({INVENTORY_CSV_PATH}) से Master Quarters और History को लोड किया जा रहा है...")
            df_inventory = pd.read_csv(INVENTORY_CSV_PATH)
            df_inventory.columns = df_inventory.columns.str.strip().str.upper()
            
            if not all(col in df_inventory.columns for col in ['QUARTER_NUMBER', 'STATION']):
                 st.error("CSV Error: 'QUARTER_NUMBER' और 'STATION' कॉलम आवश्यक हैं।")
                 return False

            batch = db.batch()
            master_count_total = 0
            history_count = 0
            
            for index, row in df_inventory.iterrows():
                clean_station = clean_data_string(row['STATION'])
                clean_quarter = clean_data_string(row['QUARTER_NUMBER'])
                
                q_doc_id = f"{clean_station}_{clean_quarter}"
                q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)
                
                is_occupied_str = str(row.get('IS_OCCUPIED', 'No')).strip().upper()
                is_occupied = is_occupied_str in ('YES', 'TRUE', '1')

                # --- 1. Master Register Creation ---
                master_status = 'Occupied' if is_occupied else 'Vacant'
                last_occupant_id = clean_data_string(row.get('HRMS_ID')) if is_occupied else None
                
                master_data = {
                    'quarter_number': clean_quarter,
                    'station': clean_station,
                    'current_status': master_status, 
                    'last_occupant_id': last_occupant_id,
                    'created_at': firestore.SERVER_TIMESTAMP
                }
                batch.set(q_doc_ref, master_data)
                master_count_total += 1

                # --- 2. History Record Creation (if Occupied) ---
                if is_occupied:
                    try:
                        allot_date_str = str(row.get('ALLOTMENT_DATE')).strip()
                        # FIX: Firestore में लिखने के लिए datetime.datetime का उपयोग करें
                        allot_date_obj = datetime.datetime.strptime(allot_date_str, '%Y-%m-%d') 
                    except (ValueError, TypeError):
                         st.warning(f"Invalid ALLOTMENT_DATE format for quarter {q_doc_id}. Skipping history.")
                         continue
                         
                    history_data = {
                        'quarter_number': clean_quarter, 
                        'station': clean_station, 
                        'hrms_id': clean_data_string(row.get('HRMS_ID')), 
                        'employee_name': str(row.get('EMPLOYEE_NAME', 'NA')).strip(), 
                        'pf_number': str(row.get('PF_NUMBER', 'NA')).strip(), 
                        'designation': str(row.get('DESIGNATION', 'NA')).strip(), 
                        'unit': str(row.get('UNIT', 'NA')).strip(),
                        'allotment_date': allot_date_obj, # अब यह Firestore के लिए मान्य है
                        'vacation_date': None,
                        'is_current': True,
                        'created_at': firestore.SERVER_TIMESTAMP
                    }
                    batch.set(db.collection(QUARTER_HISTORY_COLLECTION).document(), history_data)
                    history_count += 1

            batch.commit()
            st.success(f"🎉 सफलता: {master_count_total} Master Quarters और {history_count} History Records लोड किए गए।")
            st.cache_data.clear()
            
    except FileNotFoundError:
        st.warning(f"Warning: Quarter Register CSV file not found at {INVENTORY_CSV_PATH}. Data not loaded.")
    except Exception as e:
        st.error(f"Error loading initial data to Firestore: {e}")
        return False
    
    return True

# ----------------------------------------------------------------------
# 2. FIREBASE QUARTER DATA ACCESS (Robust)
# ----------------------------------------------------------------------

@st.cache_data(ttl=5) 
def get_all_quarters():
    """Firestore से सभी क्वार्टर और उनके स्टेटस फ़ेच करता है (Robust version)।"""
    if db is None:
        return pd.DataFrame(columns=['quarter_number', 'station', 'current_status'])
    
    required_cols = ['quarter_number', 'station', 'current_status']
    
    try:
        docs = db.collection(QUARTER_MASTER_COLLECTION).stream()
        data = [doc.to_dict() for doc in docs]
        
        if not data:
            return pd.DataFrame(columns=required_cols)
            
        df = pd.DataFrame(data) 
        
        for col in required_cols:
            if col not in df.columns:
                df[col] = 'Vacant' if col == 'current_status' else 'N/A'
                
        df = df.sort_values(by=['station', 'quarter_number']).reset_index(drop=True)
            
        return df[required_cols]
        
    except Exception as e:
        st.error(f"Error fetching quarters from Firestore: {e}")
        return pd.DataFrame(columns=required_cols)


@st.cache_data(ttl=5)
def get_quarter_history_df():
    """Firestore से सभी क्वार्टर इतिहास फ़ेच करता है (Robust version, NaT फिक्स के साथ)।"""
    if db is None: 
        return pd.DataFrame(columns=['quarter_number', 'station', 'hrms_id', 'is_current'])
    
    required_cols = ['quarter_number', 'station', 'hrms_id', 'is_current', 'employee_name', 'allotment_date', 'vacation_date', 'pf_number', 'designation', 'unit']

    try:
        docs = db.collection(QUARTER_HISTORY_COLLECTION).stream()
        data = []
        for doc in docs:
             record = doc.to_dict()
             record['id'] = doc.id
             data.append(record)
             
        if not data:
             return pd.DataFrame(columns=required_cols)
             
        df = pd.DataFrame(data)
        
        if 'is_current' not in df.columns:
             df['is_current'] = False
        
        # DATE फ़ील्ड को स्ट्रिंग में बदलें (रिपोर्टिंग के लिए)
        date_format = '%Y-%m-%d'
        
        # FIX: NaT/None हैंडलिंग के लिए सुरक्षित फ़ंक्शन
        def safe_date_to_str(x):
            # None, np.nan, या pd.NaT के लिए जाँच करें
            if pd.isna(x) or x is None:
                return 'N/A'
            try:
                # यदि यह datetime object है (Firestore Timestamps)
                if isinstance(x, datetime.datetime):
                    return x.strftime(date_format)
                # यदि यह date object है
                elif hasattr(x, 'date'):
                    return x.date().strftime(date_format)
                return str(x)
            except Exception:
                return 'N/A' # यदि कोई और पार्सिंग त्रुटि हो
        
        if 'allotment_date' in df.columns:
             df['allotment_date'] = df['allotment_date'].apply(safe_date_to_str)
             
        if 'vacation_date' in df.columns:
             df['vacation_date'] = df['vacation_date'].apply(safe_date_to_str)
             
        for col in required_cols:
             if col not in df.columns:
                 df[col] = 'N/A'
             
        return df.sort_values(by=['station', 'quarter_number', 'allotment_date'], ascending=[True, True, False])
        
    except Exception as e:
        st.error(f"Error fetching history from Firestore: {e}")
        return pd.DataFrame(columns=required_cols)

# ----------------------------------------------------------------------
# 3. FIREBASE EMPLOYEE DATA LOOKUP (Search and ID)
# ----------------------------------------------------------------------

@st.cache_data(ttl=3600)
def search_employee_details_from_firebase(search_term):
    """Firebase में नाम या HRMS ID द्वारा कर्मचारी खोजता है (Robust version - client-side filtering)."""
    if db is None or not search_term: return pd.DataFrame()
    search_term = str(search_term).strip()
    results = []

    # 1. HRMS ID से exact मैच की जाँच करें (अभी भी इंडेक्स की आवश्यकता है)
    try:
        docs_id = db.collection(EMPLOYEE_COLLECTION)\
                    .where('HRMS ID', '==', search_term)\
                    .limit(1).get()
        if docs_id:
            results.append(docs_id[0].to_dict())
    except Exception:
        pass # अगर HRMS ID क्वेरी इंडेक्स के कारण विफल हो जाती है, तो नाम खोज पर आगे बढ़ें
        
    # 2. नाम से खोज: (Indicated fix: Use client-side filtering to bypass index errors)
    if not results and len(search_term) >= 3:
        try:
             # FIX: Firestore इंडेक्स की आवश्यकता से बचने के लिए, सभी रिकॉर्ड फ़ेच करें
             docs_all = db.collection(EMPLOYEE_COLLECTION).stream() 
             
             search_term_lower = search_term.lower()
             
             for doc in docs_all:
                 doc_data = doc.to_dict()
                 # सुनिश्चित करें कि 'Employee Name' मौजूद है और उसे lowercase में बदलें
                 employee_name = doc_data.get('Employee Name', '').lower()
                 
                 # यदि नाम सर्च टर्म से शुरू होता है (Starts With)
                 if employee_name.startswith(search_term_lower):
                     if doc_data.get('HRMS ID') not in [r.get('HRMS ID') for r in results]:
                         results.append(doc_data)
                         
                 if len(results) >= 20: # केवल पहले 20 परिणाम दिखाएँ
                     break
                     
        except Exception as e:
            st.warning(f"Error during employee name search: {e}")
            
    if not results:
        return pd.DataFrame()
        
    df = pd.DataFrame(results)
    
    lookup_cols = ['HRMS ID', 'Employee Name', 'Designation', 'Unit', 'PF Number']
    for col in lookup_cols:
        if col not in df.columns:
            df[col] = 'N/A'
            
    return df[['HRMS ID', 'Employee Name', 'Designation', 'Unit', 'PF Number']]


@st.cache_data(ttl=3600)
def get_employee_details_by_hrms_id(hrms_id):
    """केवल HRMS ID से एक ही कर्मचारी का विवरण प्राप्त करता है (FIXED for Index Error)."""
    if db is None or not hrms_id: return None
    hrms_id_str = clean_data_string(hrms_id)
    
    try:
        # FIX: इंडेक्सिंग त्रुटि से बचने के लिए, हम stream() का उपयोग करके डेटा खींचते हैं और Python में सटीक HRMS ID मैच के लिए फ़िल्टर करते हैं।
        
        docs = db.collection(EMPLOYEE_COLLECTION).stream()
        record = None

        for doc in docs:
            doc_data = doc.to_dict()
            # clean_data_string का उपयोग करके यह सुनिश्चित करें कि तुलना केस/स्पेस संवेदनशील नहीं है
            if clean_data_string(doc_data.get('HRMS ID')) == hrms_id_str:
                record = doc_data
                break
        
        if record is None:
            return None
            
        # सुनिश्चित करें कि सभी आवश्यक कुंजियाँ मौजूद हैं
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
        st.error(f"Error fetching employee {hrms_id} by ID. Details: {e}")
        return None

# ----------------------------------------------------------------------
# 4. WORD FILE GENERATION
# ----------------------------------------------------------------------

def generate_word_file(template_name, data):
    """Word टेम्पलेट को भरता है और io.BytesIO स्ट्रीम में वापस करता है।"""
    current_date = datetime.date.today()
    current_date_str_letter = current_date.strftime('%d/%m/%Y') 
    
    template_path = f'{template_name}.docx'
    
    # FIX 2: यदि टेम्पलेट फ़ाइल मौजूद नहीं है तो एक स्पष्ट त्रुटि लौटाएँ
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}. Ensure templates are present in the app's root directory.")
        return None

    try:
        document = Document(template_path)
        
        # हिंदी नामों के लिए fallback सहित सभी प्रतिस्थापन
        replacements = {
            '{{DATE}}': current_date_str_letter,
            '{{QUARTER_NUMBER}}': data.get('quarter_number', 'NA'),
            '{{STATION}}': data.get('station', 'NA'),
            '{{EMPLOYEE_NAME}}': data.get('employee_name_hindi', data.get('employee_name_english', 'NA')), 
            '{{DESIGNATION}}': data.get('designation_hindi', data.get('designation_english', 'NA')), 
            '{{HRMS_ID}}': data.get('hrms_id', 'NA'),
            '{{PF_NUMBER}}': data.get('pf_number', 'NA'),
            '{{UNIT}}': data.get('unit', 'NA'),
        }

        # पैराग्राफ में प्रतिस्थापन
        for paragraph in document.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
        
        # टेबल्स में प्रतिस्थापन
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

def allot_quarter(quarter_num, station, hrms_id, allot_date, employee_details):
    """क्वार्टर को अलॉट करता है और Firestore में मास्टर और हिस्ट्री अपडेट करता है।"""
    if db is None: return False, "Database connection failed."
    
    # FIX: Ensure clean strings are used for document IDs
    clean_q_num = clean_data_string(quarter_num)
    clean_station = clean_data_string(station)
    clean_hrms_id = clean_data_string(hrms_id)
    
    q_doc_id = f"{clean_station}_{clean_q_num}"
    q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)
    
    try:
        batch = db.batch()
        
        # 1. Check for Duplicate Allotment (Employee already has a quarter)
        docs_dup = db.collection(QUARTER_HISTORY_COLLECTION)\
                     .where('hrms_id', '==', clean_hrms_id)\
                     .where('is_current', '==', True).limit(1).get()
        if docs_dup:
            # FIX: If another quarter is found, use the correct quarter number in the error message
            dup_q_num = docs_dup[0].to_dict().get('quarter_number', 'N/A')
            dup_station = docs_dup[0].to_dict().get('station', 'N/A')
            return False, f"Error: Employee ({hrms_id}) already occupies quarter {dup_q_num} at {dup_station}."

        # 2. Check quarter status
        q_doc = q_doc_ref.get()
        if not q_doc.exists:
            return False, f"Error: Quarter {quarter_num} at {station} not found in Master Register."
            
        q_data = q_doc.to_dict()
        if q_data.get('current_status') == 'Occupied':
            return False, f"Warning: Quarter {quarter_num} at {station} is already occupied. Status: {q_data.get('current_status')}"

        # A. Master Register Update (Set status to Occupied)
        batch.update(q_doc_ref, {
            'current_status': 'Occupied', 
            'last_occupant_id': clean_hrms_id,
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        # B. History Log Insert
        if isinstance(allot_date, datetime.date) and not isinstance(allot_date, datetime.datetime):
             allot_date_obj = datetime.datetime.combine(allot_date, datetime.time())
        else:
             allot_date_obj = allot_date

        history_data = {
            'quarter_number': quarter_num, 
            'station': station, 
            'hrms_id': hrms_id, 
            'pf_number': employee_details['pf_number'], 
            'designation': employee_details['designation_english'], 
            'unit': employee_details['unit'], 
            'employee_name': employee_details['employee_name_english'], 
            'allotment_date': allot_date_obj,
            'vacation_date': None,
            'is_current': True,
            'created_at': firestore.SERVER_TIMESTAMP
        }
        batch.set(db.collection(QUARTER_HISTORY_COLLECTION).document(), history_data) 
        
        batch.commit()
        st.cache_data.clear()
        
        # C. Generate Word Allotment Letter
        file_stream = generate_word_file("Allotment_Template", employee_details | {"quarter_number": quarter_num, "station": station})
        
        # FIX 2: अगर लेटर जेनरेट नहीं होता है, तो स्पष्ट त्रुटि संदेश दें
        if file_stream is None:
            return False, f"Allotment successful, but **Error generating Allotment Letter** for {quarter_num} at {station}. Check if the 'Allotment_Template.docx' file exists."

        return True, file_stream

    except Exception as e:
        st.error(f"Allotment failed: {e}")
        return False, f"Allotment failed: {e}"


def vacate_quarter(quarter_num, station, vacate_date):
    """क्वार्टर को खाली करता है और Firestore में मास्टर और हिस्ट्री अपडेट करता है।"""
    if db is None: return False, "Database connection failed."

    # FIX: Ensure clean strings are used for document IDs
    clean_q_num = clean_data_string(quarter_num)
    clean_station = clean_data_string(station)

    q_doc_id = f"{clean_station}_{clean_q_num}"
    q_doc_ref = db.collection(QUARTER_MASTER_COLLECTION).document(q_doc_id)

    try:
        batch = db.batch()
        
        # A. Get current occupant history record
        # Note: 'station' और 'quarter_number' के साथ 'is_current' = True पर फ़िल्टर करने के लिए Firestore में Compound Index की आवश्यकता हो सकती है।
        docs_history = db.collection(QUARTER_HISTORY_COLLECTION)\
                         .where('quarter_number', '==', quarter_num)\
                         .where('station', '==', station)\
                         .where('is_current', '==', True).limit(1).get()

        if not docs_history:
            return False, f"Warning: Quarter {quarter_num} at {station} is not currently occupied in the database."

        history_doc = docs_history[0]
        history_data = history_doc.to_dict()
        hrms_id = history_data.get('hrms_id', 'NA')
        
        # B. Look up details for the letter
        employee_details_full = get_employee_details_by_hrms_id(hrms_id)
        if employee_details_full is None:
            employee_details_full = history_data 
            employee_details_full['employee_name_english'] = history_data.get('employee_name', 'NA')
            employee_details_full['designation_english'] = history_data.get('designation', 'NA')
            employee_details_full['hrms_id'] = history_data.get('hrms_id', 'NA')
            employee_details_full['pf_number'] = history_data.get('pf_number', 'NA')
            employee_details_full['unit'] = history_data.get('unit', 'NA')
            st.warning("कर्मचारी विवरण (Firebase) नहीं मिला। मेमो में इतिहास डेटा का उपयोग किया जाएगा।")

        # FIX: Ensure vacate_date is datetime.datetime object for Firebase
        if isinstance(vacate_date, datetime.date) and not isinstance(vacate_date, datetime.datetime):
             vacate_date_obj = datetime.datetime.combine(vacate_date, datetime.time())
        else:
             vacate_date_obj = vacate_date

        # C. History Log Update
        batch.update(history_doc.reference, {
            'vacation_date': vacate_date_obj,
            'is_current': False,
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        # D. Master Register Update
        batch.update(q_doc_ref, {
            'current_status': 'Vacant', 
            'updated_at': firestore.SERVER_TIMESTAMP
        })

        batch.commit()
        st.cache_data.clear()
        
        # E. Generate Word Vacation Memo
        template_data = employee_details_full | {"quarter_number": quarter_num, "station": station}
        file_stream = generate_word_file("Vacation_Template", template_data)
        
        # FIX 2: अगर लेटर जेनरेट नहीं होता है, तो स्पष्ट त्रुटि संदेश दें
        if file_stream is None:
            return False, f"Vacation successful, but **Error generating Vacation Memo** for {quarter_num} at {station}. Check if the 'Vacation_Template.docx' file exists."


        return True, file_stream

    except Exception as e:
        st.error(f"FATAL DB ERROR: Vacation transaction failed. Details: {e}") 
        return False, f"Vacation failed: {e}"

# ----------------------------------------------------------------------
# 6. REPORTING 
# ----------------------------------------------------------------------

def generate_current_status_report():
    df_master = get_all_quarters()
    df_history = get_quarter_history_df()
    
    if df_master.empty: return pd.DataFrame()
    
    # FIX: .copy() का उपयोग करके SettingWithCopyWarning से बचें
    df_current_occupants = df_history[df_history['is_current'] == True].copy().rename(columns={
        'employee_name': 'current_occupant',
        'hrms_id': 'occupant_hrms_id',
        'pf_number': 'pf_number',
        'designation': 'designation',
        'unit': 'unit',
        'allotment_date': 'allotment_date'
    })
    
    df_current_occupants = df_current_occupants[['quarter_number', 'station', 'current_occupant', 
                                                 'occupant_hrms_id', 'pf_number', 'designation', 
                                                 'unit', 'allotment_date']]

    df_report = df_master.merge(
        df_current_occupants, 
        on=['quarter_number', 'station'], 
        how='left'
    ).rename(columns={'occupant_hrms_id': 'hrms_id'})
    
    df_report = df_report.fillna('N/A')
    
    display_cols = ['quarter_number', 'station', 'current_status', 'current_occupant', 
                    'hrms_id', 'pf_number', 'designation', 'unit', 'allotment_date']
    
    return df_report[[col for col in display_cols if col in df_report.columns]]

def generate_full_history_report():
    df_history = get_quarter_history_df()
    if df_history.empty: return pd.DataFrame()
    
    # FIX: .copy() का उपयोग करके SettingWithCopyWarning से बचें
    df_history_copy = df_history.copy()
    
    df_history_copy['vacation_date'] = df_history_copy.apply(
        lambda row: 'CURRENTLY OCCUPIED' if row['is_current'] else row['vacation_date'], axis=1
    )
    
    df_history_copy['record_type'] = df_history_copy['is_current'].apply(
        lambda x: 'Current Occupant' if x else 'History Record'
    )
    
    display_cols = [
        'quarter_number', 'station', 'employee_name', 'hrms_id', 'pf_number', 
        'designation', 'unit', 'allotment_date', 'vacation_date', 'record_type'
    ]
    
    return df_history_copy[[col for col in display_cols if col in df_history_copy.columns]].fillna('N/A')

# ----------------------------------------------------------------------
# 7. AUTHENTICATION & UI
# ----------------------------------------------------------------------

def check_password(password):
    return password == CORRECT_PASSWORD

def authenticate_user():
    
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if 'allot_download_data' not in st.session_state:
        st.session_state.allot_download_data = None
    if 'vacate_download_data' not in st.session_state:
        st.session_state.vacate_download_data = None


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
    
    if st.session_state.authenticated:
        if not db:
            st.error("Firebase DB connection failed. Check your Streamlit Secrets.")
            st.stop()
        
        # CSV से डेटा लोड करने का प्रयास (केवल एक बार जब Master Collection खाली हो)
        load_initial_data_to_firestore() 

        # --- SIDEBAR CONTROLS ---
        st.sidebar.markdown("---")
        if st.sidebar.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.allot_download_data = None
            st.session_state.vacate_download_data = None
            st.rerun()

        # --- UI TABS ---
        tab_allot, tab_vacate, tab_report = st.tabs(["🔑 क्वार्टर अलॉट करें", "🔓 क्वार्टर खाली करें", "📈 रिपोर्ट"])

        # -----------------------------------------------------------
        # I. Allot Quarter Tab
        # -----------------------------------------------------------
        with tab_allot:
            st.header("🔑 नया क्वार्टर अलॉट करें")
            
            quarters_df = get_all_quarters()
            # FIX: .copy() का उपयोग करके SettingWithCopyWarning से बचें
            vacant_quarters = quarters_df[quarters_df['current_status'] == 'Vacant'].copy() 
            
            # डाउनलोड बटन के लिए स्टेट क्लियर करें
            st.session_state.allot_download_data = None 

            if vacant_quarters.empty:
                st.warning("अलॉट करने के लिए कोई खाली क्वार्टर उपलब्ध नहीं है।")
            
            if not vacant_quarters.empty:
                # FIX 1: रोबस्ट सेलेक्शन के लिए इंडेक्स को रीसेट करें
                vacant_quarters = vacant_quarters.reset_index(drop=True).copy()
                vacant_quarters['Display'] = vacant_quarters['quarter_number'] + ' (' + vacant_quarters['station'] + ')'
                display_list = vacant_quarters['Display'].tolist()
                
                with st.form("allotment_form"):
                    
                    # 1. Employee Search Section
                    st.subheader("1. कर्मचारी खोजें (नाम या HRMS ID)")
                    search_term = st.text_input("कर्मचारी नाम या HRMS ID दर्ज करें (कम से कम 3 अक्षर)", key='allot_search_term').strip()
                    
                    selected_hrms_id = None
                    employee_details_full = None
                    
                    if len(search_term) >= 3:
                        df_search_results = search_employee_details_from_firebase(search_term)
                        
                        if not df_search_results.empty:
                            df_search_results['Display'] = df_search_results['HRMS ID'] + ' - ' + df_search_results['Employee Name'] + ' (' + df_search_results['Designation'] + ')'
                            selected_display = st.selectbox("परिणामों में से कर्मचारी चुनें", df_search_results['Display'].tolist(), key='allot_q_select_hrms')
                            
                            if selected_display:
                                selected_hrms_id = selected_display.split(' - ')[0].strip()
                                st.info(f"चयनित कर्मचारी HRMS ID: **{selected_hrms_id}**")
                                
                                # Fetch full details for allotment
                                employee_details_full = get_employee_details_by_hrms_id(selected_hrms_id)
                                
                                if employee_details_full:
                                     st.success("✅ कर्मचारी विवरण सफलतापूर्वक प्राप्त हुआ।")
                                     st.json({
                                        "Name": employee_details_full['employee_name_english'],
                                        "Designation": employee_details_full['designation_english'],
                                        "PF No.": employee_details_full['pf_number'],
                                        "Unit": employee_details_full['unit']
                                     })
                                else:
                                    st.error(f"Error: HRMS ID {selected_hrms_id} के लिए पूर्ण विवरण नहीं मिला। (डेटाबेस में HRMS ID/नाम की वर्तनी जाँचें)")
                        else:
                            st.warning("कोई कर्मचारी विवरण नहीं मिला।")
                    elif len(search_term) > 0:
                        st.info("खोजने के लिए कम से कम 3 अक्षर दर्ज करें।")

                    # 2. Quarter Selection and Date
                    st.subheader("2. क्वार्टर चुनें और तिथि दर्ज करें")
                    
                    # FIX 1: Display list index based selectbox
                    selected_index = st.selectbox("खाली क्वार्टर चुनें", 
                                                range(len(display_list)), 
                                                format_func=lambda i: display_list[i], 
                                                key='allot_q_select_index')
                    
                    allot_date = st.date_input("अलॉटमेंट तिथि", datetime.date.today(), key='allot_date') 
                    
                    st.markdown("---")
                    
                    submitted = st.form_submit_button("🔑 क्वार्टर अलॉट करें और लेटर जेनरेट करें")

                    if submitted:
                        if selected_hrms_id and employee_details_full and selected_index is not None:
                            
                            # FIX 1: इंडेक्स का उपयोग करके quarter_number और station को सीधे DataFrame से प्राप्त करें
                            selected_row = vacant_quarters.iloc[selected_index]
                            selected_q_num = selected_row['quarter_number']
                            selected_station = selected_row['station']

                            with st.spinner(f"क्वार्टर {selected_q_num} अलॉटमेंट संसाधित किया जा रहा है..."):
                                success, result = allot_quarter(selected_q_num, selected_station, selected_hrms_id, allot_date, employee_details_full)
                            
                            if success:
                                st.success("🎉 अलॉटमेंट सफलतापूर्वक पूरा हुआ!")
                                # result अब file_stream है, जो None नहीं हो सकता (क्योंकि allot_quarter में चेक हो गया है)
                                st.session_state.allot_download_data = {
                                    "stream": result,
                                    "filename": f"Allotment_Letter_{selected_q_num}_{selected_hrms_id}.docx"
                                }
                                st.rerun() 
                            else:
                                # result में अब स्पष्ट त्रुटि संदेश है
                                st.error(result) 
                        else:
                            st.error("कृपया एक खाली क्वार्टर चुनें और सुनिश्चित करें कि कर्मचारी विवरण सफलतापूर्वक प्राप्त हुआ है।")

            # डाउनलोड बटन फॉर्म के बाहर!
            if st.session_state.allot_download_data:
                dl_data = st.session_state.allot_download_data
                st.markdown("---")
                st.download_button(
                    label="डाउनलोड अलॉटमेंट लेटर (.docx)",
                    data=dl_data["stream"],
                    file_name=dl_data["filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key='download_allotment_final'
                )
                st.session_state.allot_download_data = None 


        # -----------------------------------------------------------
        # II. Vacate Quarter Tab
        # -----------------------------------------------------------
        with tab_vacate:
            st.header("🔓 क्वार्टर खाली करें")
            
            quarters_df = get_all_quarters()
            # FIX: .copy() का उपयोग करके SettingWithCopyWarning से बचें
            occupied_quarters = quarters_df[quarters_df['current_status'] == 'Occupied'].copy()
            
            if occupied_quarters.empty:
                st.warning("खाली करने के लिए कोई ऑक्यूपाइड क्वार्टर नहीं है।")

            # डाउनलोड बटन के लिए स्टेट क्लियर करें
            st.session_state.vacate_download_data = None
                
            if not occupied_quarters.empty:
                # FIX 1: रोबस्ट सेलेक्शन के लिए इंडेक्स को रीसेट करें
                occupied_quarters = occupied_quarters.reset_index(drop=True).copy()
                occupied_quarters['Display'] = occupied_quarters['quarter_number'] + ' (' + occupied_quarters['station'] + ')'
                display_list_vacate = occupied_quarters['Display'].tolist()
                
                # FIX 1: Display list index based selectbox
                selected_index_vacate = st.selectbox("ऑक्यूपाइड क्वार्टर चुनें", 
                                                range(len(display_list_vacate)), 
                                                format_func=lambda i: display_list_vacate[i], 
                                                key='vacate_q_select_index')


                if selected_index_vacate is not None:
                    # FIX 1: इंडेक्स का उपयोग करके quarter_number और station को सीधे DataFrame से प्राप्त करें
                    selected_row_vacate = occupied_quarters.iloc[selected_index_vacate]
                    selected_q_num = selected_row_vacate['quarter_number']
                    selected_station = selected_row_vacate['station']
                    
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
                                st.session_state.vacate_download_data = {
                                    "stream": result,
                                    "filename": f"Vacation_Memo_{selected_q_num}_{selected_station}.docx"
                                }
                                st.rerun() 
                            else:
                                st.error(result)

            # डाउनलोड बटन फॉर्म के बाहर!
            if st.session_state.vacate_download_data:
                 dl_data = st.session_state.vacate_download_data
                 st.markdown("---")
                 st.download_button(
                     label="डाउनलोड वेकेशन मेमो (.docx)",
                     data=dl_data["stream"],
                     file_name=dl_data["filename"],
                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     key='download_vacation_final'
                 )
                 st.session_state.vacate_download_data = None


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
