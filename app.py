import streamlit as st
import datetime
import os 
import pandas as pd
from docx import Document 
import io
from sqlalchemy.sql import text # For parameterized queries and DDL execution
import sys

# ----------------------------------------------------------------------
# 0. कॉन्फ़िगरेशन (Config)
# ----------------------------------------------------------------------

# NOTE: सुनिश्चित करें कि ये फ़ाइलें 'data' फ़ोल्डर में मौजूद हैं।
HRMS_MASTER_FILE = "data/UNIT_MUSTER_MASTER.xlsx" 
CSV_FILE_PATH = "data/Quarter_Register.csv" 

# --- SECURITY CONFIGURATION ---
CORRECT_PASSWORD = "Sgam@1234" 
# ------------------------------

st.set_page_config(layout="wide", page_title="रेलवे क्वार्टर प्रबंधन (PostgreSQL)")

# ----------------------------------------------------------------------
# 1. DATABASE CONNECTION & SETUP (स्थायी PostgreSQL डेटाबेस)
# ----------------------------------------------------------------------

@st.cache_resource
def get_pg_connection():
    """Streamlit Secretes का उपयोग करके PostgreSQL कनेक्शन स्थापित करता है।"""
    try:
        # 'quarter_db' secrets.toml में कनेक्शन का नाम है
        return st.connection("quarter_db", type="sql") 
    except Exception as e:
        st.error(f"Database Connection Error. Check your Streamlit Secrets: {e}")
        return None

# NOTE: हम कैशिंग हटा रहे हैं ताकि यह सुनिश्चित हो सके कि DB initialization हर बार चलती है।
def initialize_database():
    """PostgreSQL में टेबल्स को बनाता है यदि वे मौजूद नहीं हैं।"""
    conn = get_pg_connection()
    if conn is None:
        return False
    
    try:
        # 1. Master Quarters TABLE (Lowercase for PostgreSQL, using text() for DDL)
        conn.session.execute(text('''
            CREATE TABLE IF NOT EXISTS master_quarters (
                quarter_number TEXT,
                station TEXT,
                current_status TEXT,
                last_occupant_id TEXT,
                PRIMARY KEY (quarter_number, station)
            )
        '''))
        
        # 2. Quarter History TABLE (Lowercase for PostgreSQL, using text() for DDL)
        conn.session.execute(text('''
            CREATE TABLE IF NOT EXISTS quarter_history (
                history_id SERIAL PRIMARY KEY, 
                quarter_number TEXT,
                station TEXT, 
                hrms_id TEXT,
                pf_number TEXT, 
                designation TEXT, 
                unit TEXT, 
                employee_name TEXT,
                allotment_date DATE, 
                vacation_date DATE NULL,
                is_current BOOLEAN
            )
        '''))
        
        conn.session.commit()
        return True
    except Exception as e:
        st.error(f"Error initializing tables: {e}")
        return False

# --- NEW: Inventory Loading Function (Using CSV) ---
def load_quarter_inventory_from_csv(csv_path):
    """
    CSV फ़ाइल से क्वार्टर इन्वेंट्री लोड करता है और उन्हें master_quarters टेबल में डालता है।
    यह फ़ंक्शन केवल तब चलाया जाना चाहिए जब टेबल खाली हो।
    """
    conn = get_pg_connection()
    if conn is None: return False

    try:
        # 1. CSV फ़ाइल लोड करें
        df_inventory = pd.read_csv(csv_path) 
        df_inventory.columns = df_inventory.columns.str.strip().str.upper()
        
        required_cols = ['QUARTER_NUMBER', 'STATION']
        if not all(col in df_inventory.columns for col in required_cols):
             st.error("CSV Error: Quarter Inventory फ़ाइल में 'QUARTER_NUMBER' और 'STATION' कॉलम नहीं मिले।")
             return False

        # 2. डेटाबेस में कुल मौजूदा क्वार्टरों की संख्या जाँचें
        existing_count = conn.query("SELECT COUNT(*) FROM master_quarters").iloc[0, 0]
        
        if existing_count > 0:
            # st.warning(f"मास्टर क्वार्टर टेबल पहले से ही {existing_count} रिकॉर्ड्स से भरी हुई है। इन्वेंट्री लोड करना छोड़ दिया गया।")
            return True 

        st.info(f"CSV से {len(df_inventory)} क्वार्टर रिकॉर्ड्स लोड किए जा रहे हैं...")

        # 3. डेटा को PostgreSQL में INSERT करें
        records_to_insert = []
        for index, row in df_inventory.iterrows():
            records_to_insert.append({
                'q_num': str(row['QUARTER_NUMBER']).strip(),
                'station': str(row['STATION']).strip().upper(),
                'status': 'Vacant', # डिफ़ॉल्ट रूप से Vacant
                'last_id': None
            })

        conn.session.execute(text('''
            INSERT INTO master_quarters (quarter_number, station, current_status, last_occupant_id)
            VALUES (:q_num, :station, :status, :last_id)
        '''), records_to_insert)

        conn.session.commit()
        st.success(f"कुल {len(records_to_insert)} क्वार्टर सफलतापूर्वक मास्टर रजिस्टर में लोड किए गए।")
        return True

    except FileNotFoundError:
        st.error(f"Error: Quarter Register CSV file not found at {csv_path}.")
        return False
    except Exception as e:
        conn.session.rollback()
        st.error(f"Error inserting inventory: {e}")
        return False
# --- END NEW FUNCTION ---


@st.cache_data(ttl=5) # Reduced cache time for fresh data
def get_all_quarters():
    """सभी क्वार्टर और उनके स्टेटस फ़ेच करता है।"""
    conn = get_pg_connection()
    if conn is None:
        return pd.DataFrame()
    
    try:
        # NOTE: Lowercase table name 'master_quarters'
        df = conn.query("SELECT quarter_number, station, current_status FROM master_quarters ORDER BY station, quarter_number")
        return df
    except Exception as e:
        st.error(f"Error fetching quarters: {e}")
        return pd.DataFrame()


# ----------------------------------------------------------------------
# 2. WORD FILE GENERATION (वर्ड फाइल जनरेशन) - Unchanged
# ----------------------------------------------------------------------

def generate_word_file(template_name, data):
    """Word टेम्पलेट को भरता है और io.BytesIO स्ट्रीम में वापस करता है।"""
    current_date = datetime.date.today()
    current_date_str_letter = current_date.strftime('%d/%m/%Y') 
    
    template_path = f'{template_name}.docx'
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}. Please upload it to your GitHub root.")
        return None

    try:
        document = Document(template_path)
        
        replacements = {
            '{{DATE}}': current_date_str_letter,
            '{{QUARTER_NUMBER}}': data['quarter_number'],
            '{{STATION}}': data.get('station', 'NA'),
            '{{EMPLOYEE_NAME}}': data.get('employee_name_hindi', data.get('employee_name_english', 'NA')), 
            '{{DESIGNATION}}': data.get('designation_hindi', data.get('designation_english', 'NA')), 
            '{{HRMS_ID}}': data.get('hrms_id', 'NA'),
            '{{PF_Number}}': data.get('pf_number', 'NA'),
            '{{UNIT}}': data.get('unit', 'NA'),
        }

        for paragraph in document.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
        
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
# 3. CORE LOGIC (PostgreSQL के लिए अपडेटेड - Lowecase Table Names)
# ----------------------------------------------------------------------

def allot_quarter(quarter_num, station, employee_details, allot_date):
    """क्वार्टर को किसी कर्मचारी को अलॉट करता है और PostgreSQL में हिस्ट्री अपडेट करता है।"""
    
    conn = get_pg_connection()
    if conn is None: return False, "Database connection failed."

    hrms_id = employee_details['hrms_id']
    allot_date_str = allot_date.strftime('%Y-%m-%d')
    
    try:
        # 1. CHECK FOR DUPLICATE ALLOTMENT
        df_dup = conn.query("SELECT quarter_number FROM quarter_history WHERE hrms_id = :hrms_id AND is_current = TRUE",
                            params={"hrms_id": hrms_id})
        if not df_dup.empty:
            return False, f"Error: Employee ({hrms_id}) already occupies quarter {df_dup.iloc[0]['quarter_number']}. Please vacate the previous one first."
            
        # 2. Check quarter status
        df_status = conn.query("SELECT current_status FROM master_quarters WHERE quarter_number = :q_num AND station = :station", 
                                params={"q_num": quarter_num, "station": station})
        
        if df_status.empty:
            return False, f"Error: Quarter {quarter_num} at {station} not found in Master Register."
            
        if df_status.iloc[0]['current_status'] == 'Occupied':
            return False, f"Warning: Quarter {quarter_num} at {station} is already occupied. Must vacate first."

        # A. Master Register Update
        conn.session.execute(text('''
            UPDATE master_quarters 
            SET current_status = 'Occupied', last_occupant_id = :hrms_id
            WHERE quarter_number = :q_num AND station = :station
        '''), params={'hrms_id': hrms_id, 'q_num': quarter_num, 'station': station})

        # B. History Log Insert
        conn.session.execute(text('''
            INSERT INTO quarter_history 
            (quarter_number, station, hrms_id, pf_number, designation, unit, employee_name, allotment_date, is_current)
            VALUES (:q_num, :station, :hrms_id, :pf_num, :desig, :unit, :emp_name, :allot_date, TRUE)
        '''), params={
            'q_num': quarter_num, 
            'station': station, 
            'hrms_id': hrms_id, 
            'pf_num': employee_details['pf_number'], 
            'desig': employee_details['designation_english'], 
            'unit': employee_details['unit'], 
            'emp_name': employee_details['employee_name_english'], 
            'allot_date': allot_date_str
        })
        conn.session.commit()
        
        # C. Generate Word Allotment Letter
        file_stream = generate_word_file("Allotment_Template", employee_details | {"quarter_number": quarter_num, "station": station})
        
        return True, file_stream

    except Exception as e:
        conn.session.rollback()
        return False, f"Allotment failed: {e}"


def vacate_quarter(quarter_num, station, vacate_date):
    """क्वार्टर को खाली करता है और PostgreSQL हिस्ट्री अपडेट करता है।"""
    conn = get_pg_connection()
    if conn is None: return False, "Database connection failed."
    vacate_date_str = vacate_date.strftime('%Y-%m-%d')

    try:
        # A. Get current occupant details
        df_history = conn.query('''
            SELECT employee_name, hrms_id, pf_number, designation, unit FROM quarter_history 
            WHERE quarter_number = :q_num AND station = :station AND is_current = TRUE
        ''', params={'q_num': quarter_num, 'station': station})
        
        if df_history.empty:
            return False, f"Warning: Quarter {quarter_num} at {station} is not currently occupied in the database."

        history_data = df_history.iloc[0]
        hrms_id = history_data['hrms_id']
        
        # B. Look up Hindi name/designation from Excel
        employee_details_full = search_employee_and_get_details_by_hrms(hrms_id)
        
        # C. History Log Update
        conn.session.execute(text('''
            UPDATE quarter_history 
            SET vacation_date = :vacate_date, is_current = FALSE
            WHERE quarter_number = :q_num AND station = :station AND is_current = TRUE
        '''), params={'vacate_date': vacate_date_str, 'q_num': quarter_num, 'station': station})

        # D. Master Register Update
        conn.session.execute(text('''
            UPDATE master_quarters 
            SET current_status = 'Vacant', last_occupant_id = :hrms_id
            WHERE quarter_number = :q_num AND station = :station
        '''), params={'hrms_id': hrms_id, 'q_num': quarter_num, 'station': station})

        conn.session.commit()

        # E. Generate Word Vacation Memo (Using merged details)
        template_data = history_data.to_dict() | employee_details_full | {"quarter_number": quarter_num, "station": station}
        file_stream = generate_word_file("Vacation_Template", template_data)

        return True, file_stream

    except Exception as e:
        conn.session.rollback()
        return False, f"Vacation failed: {e}"


# ----------------------------------------------------------------------
# 4. REPORTING (PostgreSQL के लिए अपडेटेड - Lowecase Table Names)
# ----------------------------------------------------------------------

def generate_current_status_report():
    conn = get_pg_connection()
    if conn is None: return pd.DataFrame()
    query = '''
        SELECT 
            MQ.quarter_number, MQ.station, MQ.current_status, 
            COALESCE(QH.employee_name, 'N/A') AS current_occupant,
            COALESCE(QH.hrms_id, 'N/A') AS hrms_id, COALESCE(QH.pf_number, 'N/A') AS pf_number,
            COALESCE(QH.designation, 'N/A') AS designation, COALESCE(QH.unit, 'N/A') AS unit,
            COALESCE(CAST(QH.allotment_date AS TEXT), 'N/A') AS allotment_date
        FROM master_quarters MQ
        LEFT JOIN quarter_history QH 
            ON MQ.quarter_number = QH.quarter_number 
            AND MQ.station = QH.station AND QH.is_current = TRUE 
        ORDER BY MQ.station, MQ.quarter_number
    '''
    df = conn.query(query)
    return df

def generate_full_history_report():
    conn = get_pg_connection()
    if conn is None: return pd.DataFrame()
    query = '''
        SELECT 
            quarter_number, station, employee_name, hrms_id, pf_number, designation, unit,
            CAST(allotment_date AS TEXT) AS allotment_date, 
            COALESCE(CAST(vacation_date AS TEXT), 'CURRENTLY OCCUPIED') as vacation_date,
            CASE WHEN is_current = TRUE THEN 'Current Occupant' ELSE 'History Record' END as record_type
        FROM quarter_history 
        ORDER BY station, quarter_number, allotment_date DESC
    '''
    df = conn.query(query)
    return df


# ----------------------------------------------------------------------
# 5. EMPLOYEE DATA SEARCH & LOOKUP (Excel) - HRMS MASTER FILE
# ----------------------------------------------------------------------

@st.cache_data(ttl=3600)
def load_master_excel():
    """HRMS मास्टर Excel फ़ाइल लोड करता है।"""
    try:
        df_master = pd.read_excel(HRMS_MASTER_FILE, sheet_name='Master sheet')
        df_master.columns = df_master.columns.str.strip().str.upper() 
        hrms_id_col = 'HRMS ID' if 'HRMS ID' in df_master.columns else 'HRMSID'
        df_master[hrms_id_col] = df_master.get(hrms_id_col, pd.Series()).astype(str).str.strip()
        df_master['EMPLOYEE NAME'] = df_master.get('EMPLOYEE NAME', pd.Series()).astype(str).str.strip().str.upper()
        df_master['STATION'] = df_master.get('STATION', pd.Series()).astype(str).str.strip().str.upper()
        return df_master, hrms_id_col
    except FileNotFoundError:
        st.error(f"Error: HRMS Master Excel file not found at {HRMS_MASTER_FILE}. Check the 'data' folder in your GitHub repository.")
        return pd.DataFrame(), None
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame(), None

def search_employee_and_get_details(search_term, station_filter):
    df_master, hrms_id_col = load_master_excel()
    if df_master.empty: return pd.DataFrame()

    search_term = str(search_term).strip().upper()
    station_filter_upper = str(station_filter).strip().upper()

    results_df = df_master[
        (df_master['STATION'] == station_filter_upper) & 
        ((df_master[hrms_id_col].str.contains(search_term, na=False)) |
         (df_master['EMPLOYEE NAME'].str.contains(search_term, na=False)))
    ].head(10).reset_index(drop=True)
    return results_df


def extract_employee_data(selected_row):
    df_master, hrms_id_col = load_master_excel()
    if df_master.empty or selected_row.empty: return None

    unit_val = str(selected_row.get('UNIT', 'NA')) 
    
    return {
        'hrms_id': str(selected_row.get(hrms_id_col, 'NA')),
        'employee_name_english': str(selected_row.get('EMPLOYEE NAME', 'NA')),
        'designation_english': str(selected_row.get('DESIGNATION', 'NA')),
        'employee_name_hindi': str(selected_row.get('EMPLOYEE NAME IN HINDI', str(selected_row.get('EMPLOYEE NAME', 'NA')))),
        'designation_hindi': str(selected_row.get('DESIGNATION IN HINDI', str(selected_row.get('DESIGNATION', 'NA')))),
        'pf_number': str(selected_row.get('PF NUMBER', 'NA')),
        'unit': unit_val[:2] 
    }

def search_employee_and_get_details_by_hrms(hrms_id):
    df_master, hrms_id_col = load_master_excel()
    if df_master.empty: return None
        
    hrms_id_str = str(hrms_id).strip()
    selected_rows = df_master[df_master[hrms_id_col] == hrms_id_str]
    
    if selected_rows.empty: return None

    selected_row = selected_rows.iloc[0]
    
    return {
        'employee_name_hindi': str(selected_row.get('EMPLOYEE NAME IN HINDI', str(selected_row.get('EMPLOYEE NAME', 'NA')))),
        'designation_hindi': str(selected_row.get('DESIGNATION IN HINDI', str(selected_row.get('DESIGNATION', 'NA')))),
    }

# ----------------------------------------------------------------------
# 6. AUTHENTICATION (प्रमाणीकरण) - Unchanged
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
                        st.error("Invalid Password. Please try again.")
        
        return False
    
    return True

# ----------------------------------------------------------------------
# 7. STREAMLIT UI (उपयोगकर्ता इंटरफ़ेस)
# ----------------------------------------------------------------------

def main_streamlit_ui():
    """मुख्य Streamlit इंटरफ़ेस।"""
    
    # 1. Run Authentication Check
    if not authenticate_user():
        return 

    # If authenticated, proceed
    st.title("🏡 रेलवे क्वार्टर प्रबंधन प्रणाली")
    
    # 2. Run DB Initialization
    if not initialize_database():
        st.warning("Database initialization failed. Please check connection secrets.")
        return 
        
    # 3. RUN INVENTORY LOADING ONCE HERE
    if not load_quarter_inventory_from_csv(CSV_FILE_PATH):
        st.error("क्वार्टर इन्वेंट्री लोड नहीं हो सकी। कृपया CSV फ़ाइल और कॉलम नाम जांचें।")
        return 

    # 4. Session State Initialization and Data Load
    if 'quarter_df' not in st.session_state or st.button("Refresh Status", key='refresh_status'):
        st.session_state.quarter_df = get_all_quarters()
        
    if st.session_state.quarter_df.empty and 'current_status' not in st.session_state.quarter_df.columns:
        st.error("Cannot load quarter data. Please check logs for PostgreSQL connection issues.")
        # We skip return here to show the rest of the UI structure

    if 'search_results' not in st.session_state:
        st.session_state.search_results = pd.DataFrame()
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = None


    # Tabs for Navigation
    tab1, tab2, tab3, tab4 = st.tabs(["🏠 होम | स्थिति", "🔑 क्वार्टर आवंटन", "🗑️ क्वार्टर खाली करना", "📊 रिपोर्ट"])

    with tab1:
        st.header("वर्तमान क्वार्टर स्थिति")
        df_status = st.session_state.quarter_df
        
        if df_status.empty or 'current_status' not in df_status.columns:
            st.info("डेटाबेस से कोई क्वार्टर डेटा लोड नहीं हुआ है।")
        else:
            col1, col2, col3 = st.columns(3)
            col1.metric("कुल क्वार्टर", len(df_status))
            col2.metric("आवंटित (Occupied)", len(df_status[df_status['current_status'] == 'Occupied']))
            col3.metric("खाली (Vacant)", len(df_status[df_status['current_status'] == 'Vacant']))
            
            st.subheader("मास्टर क्वार्टर सूची")
            st.dataframe(df_status, use_container_width=True)


    with tab2:
        st.header("🔑 नया क्वार्टर आवंटन")
        
        if 'current_status' not in st.session_state.quarter_df.columns:
             st.warning("क्वार्टर आवंटन के लिए डेटाबेस कनेक्शन स्थापित नहीं है।")
        else:
            stations = st.session_state.quarter_df['station'].unique().tolist()
            stations.insert(0, '--- स्टेशन चुनें ---')
            
            selected_station = st.selectbox("1. स्टेशन चुनें", stations, key='allot_station')
            
            if selected_station and selected_station != '--- स्टेशन चुनें ---':
                
                available_quarters = st.session_state.quarter_df[
                    (st.session_state.quarter_df['station'] == selected_station) & 
                    (st.session_state.quarter_df['current_status'] != 'Occupied')
                ]['quarter_number'].tolist()

                if not available_quarters:
                    st.warning(f"स्टेशन **{selected_station}** पर कोई खाली/उपलब्ध क्वार्टर नहीं है।")
                else:
                    available_quarters.insert(0, '--- क्वार्टर चुनें ---')
                    selected_q_num = st.selectbox("2. खाली क्वार्टर चुनें", available_quarters, key='allot_q_num')

                    if selected_q_num != '--- क्वार्टर चुनें ---':
                        st.subheader(f"कर्मचारी खोजें ({selected_station} के लिए)")

                        search_col, status_col = st.columns([2, 1])

                        with search_col:
                            search_term = st.text_input("3. कर्मचारी का नाम या HRMS ID दर्ज करें (स्टेशन-फ़िल्टर लागू)", key='allot_search_term')
                        
                        if st.button("खोज शुरू करें", key='allot_search_btn'):
                            if search_term:
                                st.session_state.search_results = search_employee_and_get_details(search_term, selected_station)
                                st.session_state.selected_employee = None
                            else:
                                st.warning("कृपया खोज के लिए नाम या HRMS ID दर्ज करें।")

                        # Display Search Results
                        if not st.session_state.search_results.empty:
                            st.subheader("4. खोज परिणाम (कर्मचारी चुनें)")
                            
                            display_options = [
                                f"{row['EMPLOYEE NAME']} ({row.get('DESIGNATION', 'N/A')}, HRMS: {row.get('HRMS ID', 'N/A')})"
                                for index, row in st.session_state.search_results.iterrows()
                            ]
                            
                            selection_index = st.radio("कर्मचारी चुनें:", options=list(range(len(display_options))), format_func=lambda x: display_options[x], key='allot_emp_select')
                            
                            selected_row_data = st.session_state.search_results.iloc[selection_index]
                            st.session_state.selected_employee = extract_employee_data(selected_row_data)

                            st.write(f"**चयनित कर्मचारी:** {st.session_state.selected_employee['employee_name_hindi']} ({st.session_state.selected_employee['hrms_id']})")
                            
                            # Date Picker and Final Button
                            allot_date = st.date_input("5. आवंटन की तिथि चुनें", datetime.date.today(), key='allot_date')
                            
                            if st.session_state.selected_employee and st.button("🔑 आवंटन पूरा करें और पत्र डाउनलोड करें", key='final_allot_btn'):
                                
                                success, result = allot_quarter(
                                    selected_q_num, 
                                    selected_station, 
                                    st.session_state.selected_employee, 
                                    allot_date
                                )
                                
                                # Refresh status manually after transaction
                                st.session_state.quarter_df = get_all_quarters()
                                
                                if success:
                                    st.success(f"**सफलता!** क्वार्टर {selected_q_num} को {st.session_state.selected_employee['employee_name_hindi']} को आवंटित किया गया।")
                                    
                                    if result is not None:
                                        file_name = f"{selected_station}_{selected_q_num.replace('/', '_')}_Allotment_{datetime.date.today().strftime('%Y%m%d')}.docx"
                                        st.download_button(
                                            label="Word पत्र डाउनलोड करें",
                                            data=result,
                                            file_name=file_name,
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                        )
                                    else:
                                        st.warning("Word मेमो जेनरेट नहीं हो सका।")
                                        
                                    st.session_state.search_results = pd.DataFrame()
                                    st.session_state.selected_employee = None
                                    
                                else:
                                    st.error(result)

    with tab3:
        st.header("🗑️ क्वार्टर खाली करना")

        if 'current_status' not in st.session_state.quarter_df.columns:
             st.warning("क्वार्टर खाली करने के लिए डेटाबेस कनेक्शन स्थापित नहीं है।")
        else:
            occupied_quarters_df = st.session_state.quarter_df[st.session_state.quarter_df['current_status'] == 'Occupied']
            
            if occupied_quarters_df.empty:
                st.info("वर्तमान में कोई भी क्वार्टर आवंटित नहीं है।")
            else:
                occupied_quarters_df['display'] = occupied_quarters_df['quarter_number'] + " (" + occupied_quarters_df['station'] + ")"
                
                quarters_to_vacate = occupied_quarters_df['display'].tolist()
                quarters_to_vacate.insert(0, '--- क्वार्टर चुनें ---')
                
                selected_quarter_display = st.selectbox("1. खाली करने के लिए क्वार्टर चुनें", quarters_to_vacate, key='vacate_q_select')
                
                if selected_quarter_display != '--- क्वार्टर चुनें ---':
                    
                    q_num, station = selected_quarter_display.split(' (')
                    station = station.replace(')', '')
                    
                    vacate_date = st.date_input("2. खाली करने की तिथि चुनें", datetime.date.today(), key='vacate_date')

                    if st.button("🗑️ वेकेशन मेमो जनरेट करें और क्वार्टर खाली करें", key='final_vacate_btn'):
                        
                        success, result = vacate_quarter(q_num, station, vacate_date)
                        
                        # Refresh status manually after transaction
                        st.session_state.quarter_df = get_all_quarters() 

                        if success:
                            st.success(f"**सफलता!** क्वार्टर {q_num} ({station}) सफलतापूर्वक खाली कर दिया गया।")
                            
                            if result is not None:
                                file_name = f"{station}_{q_num.replace('/', '_')}_Vacation_{datetime.date.today().strftime('%Y%m%d')}.docx"
                                st.download_button(
                                    label="Word वेकेशन मेमो डाउनलोड करें",
                                    data=result,
                                    file_name=file_name,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            else:
                                st.warning("Word मेमो जेनरेट नहीं हो सका।")

                        else:
                            st.error(result)


    with tab4:
        st.header("📊 रिपोर्ट जेनरेट करें")
        
        if 'current_status' not in st.session_state.quarter_df.columns:
             st.warning("रिपोर्ट जेनरेट करने के लिए डेटाबेस कनेक्शन स्थापित नहीं है।")
        else:
            report_choice = st.radio(
                "रिपोर्ट का प्रकार चुनें:",
                ('वर्तमान स्थिति रिपोर्ट', 'संपूर्ण इतिहास रिपोर्ट'),
                key='report_type_select'
            )
            
            if report_choice == 'वर्तमान स्थिति रिपोर्ट':
                st.subheader("सभी क्वार्टरों की वर्तमान स्थिति")
                df_report = generate_current_status_report()
                st.dataframe(df_report, use_container_width=True)
                
                csv = df_report.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="वर्तमान स्थिति रिपोर्ट (CSV) डाउनलोड करें",
                    data=csv,
                    file_name=f"All_Quarter_Current_Status_{datetime.date.today().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key='download_current_status'
                )
            
            elif report_choice == 'संपूर्ण इतिहास रिपोर्ट':
                st.subheader("सभी क्वार्टरों का सम्पूर्ण इतिहास")
                df_history = generate_full_history_report()
                st.dataframe(df_history, use_container_width=True)
                
                csv = df_history.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="संपूर्ण इतिहास रिपोर्ट (CSV) डाउनलोड करें",
                    data=csv,
                    file_name=f"Full_Quarter_History_{datetime.date.today().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key='download_full_history'
                )

if __name__ == '__main__':
    if not os.path.exists('data'):
        os.makedirs('data')

    main_streamlit_ui()
