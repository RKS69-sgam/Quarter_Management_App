import streamlit as st
import sqlite3
import datetime
import os 
import pandas as pd
from docx import Document 
import io
import base64
import sys

# ----------------------------------------------------------------------
# 0. कॉन्फ़िगरेशन (Config)
# ----------------------------------------------------------------------

# NOTE: सुनिश्चित करें कि ये फ़ाइलें और फ़ोल्डर मौजूद हैं:
# 1. quarter_register.db (यह ऑटो-क्रिएट होता है)
# 2. data/UNIT_MUSTER_MASTER.xlsx (कर्मचारी मास्टर डेटा)
# 3. Allotment_Template.docx and Vacation_Template.docx
EXCEL_FILE_PATH = "data/UNIT_MUSTER_MASTER.xlsx" 
DB_NAME = 'quarter_register.db'

# --- SECURITY CONFIGURATION ---
CORRECT_PASSWORD = "Sgam@1234" # आपका निर्धारित पासवर्ड
# ------------------------------

st.set_page_config(layout="wide", page_title="रेलवे क्वार्टर प्रबंधन")

# ----------------------------------------------------------------------
# 1. DATABASE SETUP (डेटाबेस सेटअप) - Caching resources
# ----------------------------------------------------------------------

@st.cache_resource
def initialize_database():
    """डेटाबेस और टेबल्स को बनाता है यदि वे मौजूद नहीं हैं।"""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Master Quarter Register Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Master_Quarters (
            quarter_number TEXT,
            station TEXT,
            current_status TEXT, -- 'Occupied', 'Vacant', 'Damaged'
            last_occupant_id TEXT,
            PRIMARY KEY (quarter_number, station)
        )
    ''')

    # Quarter History Log Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Quarter_History (
            history_id INTEGER PRIMARY KEY AUTOINCREMENT,
            quarter_number TEXT,
            station TEXT, 
            hrms_id TEXT,
            pf_number TEXT, 
            designation TEXT, 
            unit TEXT, 
            employee_name TEXT,
            allotment_date TEXT,
            vacation_date TEXT NULL,
            is_current BOOLEAN, -- 1 for current, 0 for past
            FOREIGN KEY (quarter_number, station) REFERENCES Master_Quarters(quarter_number, station)
        )
    ''')

    conn.commit()
    conn.close()
    return True

def get_db_connection():
    """SQLite कनेक्शन खोलता है।"""
    return sqlite3.connect(DB_NAME)

def get_all_quarters():
    """सभी क्वार्टर और उनके स्टेटस फ़ेच करता है।"""
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT quarter_number, station, current_status FROM Master_Quarters ORDER BY station, quarter_number", conn)
    conn.close()
    return df

# ----------------------------------------------------------------------
# 2. WORD FILE GENERATION (वर्ड फाइल जनरेशन)
# ----------------------------------------------------------------------

def generate_word_file(template_name, data):
    """Word टेम्पलेट को भरता है और io.BytesIO स्ट्रीम में वापस करता है।"""
    
    current_date = datetime.date.today()
    current_date_str_letter = current_date.strftime('%d/%m/%Y') 
    
    template_path = f'{template_name}.docx'
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}. Please upload it.")
        return None

    try:
        document = Document(template_path)
        
        replacements = {
            '{{DATE}}': current_date_str_letter,
            '{{QUARTER_NUMBER}}': data['quarter_number'],
            '{{STATION}}': data.get('station', 'NA'),
            # Use Hindi data for template
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
# 3. CORE LOGIC (मुख्य लॉजिक)
# ----------------------------------------------------------------------

def allot_quarter(quarter_num, station, employee_details, allot_date):
    """क्वार्टर को किसी कर्मचारी को अलॉट करता है और हिस्ट्री अपडेट करता है।"""
    
    hrms_id = employee_details['hrms_id']
    emp_name_db = employee_details['employee_name_english']
    pf_number = employee_details['pf_number']
    designation_db = employee_details['designation_english']
    unit = employee_details['unit'] 
    
    conn = get_db_connection()
    cursor = conn.cursor()
    allot_date_str = allot_date.strftime('%Y-%m-%d')
    
    try:
        # 1. CHECK FOR DUPLICATE ALLOTMENT
        cursor.execute('''
            SELECT quarter_number, station FROM Quarter_History 
            WHERE hrms_id = ? AND is_current = 1
        ''', (hrms_id,))
        if cursor.fetchone():
            return False, f"Error: Employee ({hrms_id}) already occupies a quarter. Please vacate the previous one first."
            
        # 2. Check quarter status
        cursor.execute("SELECT current_status FROM Master_Quarters WHERE quarter_number = ? AND station = ?", (quarter_num, station))
        master_data = cursor.fetchone()
        
        if not master_data:
            return False, f"Error: Quarter {quarter_num} at {station} not found in Master Register."
            
        if master_data[0] == 'Occupied':
            return False, f"Warning: Quarter {quarter_num} at {station} is already occupied. Must vacate first."

        # A. Master Register Update
        cursor.execute('''
            UPDATE Master_Quarters 
            SET current_status = 'Occupied', last_occupant_id = ?
            WHERE quarter_number = ? AND station = ?
        ''', (hrms_id, quarter_num, station))

        # B. History Log Insert
        cursor.execute('''
            INSERT INTO Quarter_History 
            (quarter_number, station, hrms_id, pf_number, designation, unit, employee_name, allotment_date, is_current)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        ''', (quarter_num, station, hrms_id, pf_number, designation_db, unit, emp_name_db, allot_date_str))

        conn.commit()
        
        # C. Generate Word Allotment Letter
        file_stream = generate_word_file("Allotment_Template", {
            "type": "Allotment",
            "quarter_number": quarter_num,
            "station": station,
            "employee_name_english": emp_name_db, 
            "employee_name_hindi": employee_details['employee_name_hindi'],
            "designation_english": designation_db, 
            "designation_hindi": employee_details['designation_hindi'],
            "date": allot_date_str, 
            "hrms_id": hrms_id,
            "pf_number": pf_number, 
            "unit": unit 
        })
        
        return True, file_stream

    except sqlite3.Error as e:
        conn.rollback()
        return False, f"Allotment failed: {e}"
    finally:
        conn.close()


def vacate_quarter(quarter_num, station, vacate_date):
    """क्वार्टर को खाली करता है और हिस्ट्री अपडेट करता है।"""
    conn = get_db_connection()
    cursor = conn.cursor()
    vacate_date_str = vacate_date.strftime('%Y-%m-%d')

    try:
        # A. Get current occupant details (English)
        cursor.execute('''
            SELECT employee_name, hrms_id, pf_number, designation, unit FROM Quarter_History 
            WHERE quarter_number = ? AND station = ? AND is_current = 1
        ''', (quarter_num, station))
        history_data = cursor.fetchone()
        
        if not history_data:
            return False, f"Warning: Quarter {quarter_num} at {station} is not currently occupied."

        emp_name_db, hrms_id, pf_number, designation_db, unit = history_data
        
        # B. Look up Hindi name/designation from Excel
        employee_details_full = search_employee_and_get_details_by_hrms(hrms_id)
        
        emp_name_template = employee_details_full.get('employee_name_hindi', emp_name_db) if employee_details_full else emp_name_db
        designation_template = employee_details_full.get('designation_hindi', designation_db) if employee_details_full else designation_db

        # C. History Log Update
        cursor.execute('''
            UPDATE Quarter_History 
            SET vacation_date = ?, is_current = 0
            WHERE quarter_number = ? AND station = ? AND is_current = 1
        ''', (vacate_date_str, quarter_num, station))

        # D. Master Register Update
        cursor.execute('''
            UPDATE Master_Quarters 
            SET current_status = 'Vacant', last_occupant_id = ?
            WHERE quarter_number = ? AND station = ?
        ''', (hrms_id, quarter_num, station))

        conn.commit()

        # E. Generate Word Vacation Memo
        file_stream = generate_word_file("Vacation_Template", {
            "type": "Vacation",
            "quarter_number": q_num,
            "station": station,
            "employee_name_english": emp_name_db,
            "employee_name_hindi": emp_name_template,
            "designation_english": designation_db,
            "designation_hindi": designation_template,
            "date": vacate_date_str, 
            "hrms_id": hrms_id,
            "pf_number": pf_number, 
            "unit": unit 
        })

        return True, file_stream

    except sqlite3.Error as e:
        conn.rollback()
        return False, f"Vacation failed: {e}"
    finally:
        conn.close()


# ----------------------------------------------------------------------
# 4. REPORTING (रिपोर्टिंग)
# ----------------------------------------------------------------------

def generate_current_status_report():
    """सभी क्वार्टरों की वर्तमान स्थिति रिपोर्ट जेनरेट करता है।"""
    conn = get_db_connection()
    query = '''
        SELECT 
            MQ.quarter_number, 
            MQ.station, 
            MQ.current_status, 
            COALESCE(QH.employee_name, 'N/A') AS current_occupant,
            COALESCE(QH.hrms_id, 'N/A') AS hrms_id,
            COALESCE(QH.pf_number, 'N/A') AS pf_number,
            COALESCE(QH.designation, 'N/A') AS designation,
            COALESCE(QH.unit, 'N/A') AS unit,
            COALESCE(QH.allotment_date, 'N/A') AS allotment_date
        FROM Master_Quarters MQ
        LEFT JOIN Quarter_History QH 
            ON MQ.quarter_number = QH.quarter_number 
            AND MQ.station = QH.station
            AND QH.is_current = 1 
        ORDER BY MQ.station, MQ.quarter_number
    '''
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def generate_full_history_report():
    """सभी क्वार्टरों की पूरी हिस्ट्री (वर्तमान और पिछली दोनों) रिपोर्ट जेनरेट करता है।"""
    conn = get_db_connection()
    query = '''
        SELECT 
            quarter_number, 
            station, 
            employee_name, 
            hrms_id, 
            pf_number,
            designation, 
            unit,
            allotment_date, 
            COALESCE(vacation_date, 'CURRENTLY OCCUPIED') as vacation_date,
            CASE WHEN is_current = 1 THEN 'Current Occupant' ELSE 'History Record' END as record_type
        FROM Quarter_History 
        ORDER BY station, quarter_number, allotment_date DESC
    '''
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df


# ----------------------------------------------------------------------
# 5. EMPLOYEE DATA SEARCH & LOOKUP (कर्मचारी डेटा खोज और लुकअप) - Caching dataframes
# ----------------------------------------------------------------------

@st.cache_data(ttl=3600) # Cache for 1 hour
def load_master_excel():
    """Excel मास्टर फ़ाइल लोड करता है।"""
    try:
        df_master = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Master sheet')
        df_master.columns = df_master.columns.str.strip().str.upper() 
        hrms_id_col = 'HRMS ID' if 'HRMS ID' in df_master.columns else 'HRMSID'
        df_master[hrms_id_col] = df_master.get(hrms_id_col, pd.Series()).astype(str).str.strip()
        df_master['EMPLOYEE NAME'] = df_master.get('EMPLOYEE NAME', pd.Series()).astype(str).str.strip().str.upper()
        df_master['STATION'] = df_master.get('STATION', pd.Series()).astype(str).str.strip().str.upper() 
        return df_master, hrms_id_col
    except FileNotFoundError:
        st.error(f"Error: Master Excel file not found at {EXCEL_FILE_PATH}. Please check the path and folder structure.")
        return pd.DataFrame(), None
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame(), None

def search_employee_and_get_details(search_term, station_filter):
    """कर्मचारी खोजता है और विवरण लौटाता है (स्टेशन फ़िल्टर के साथ)।"""
    df_master, hrms_id_col = load_master_excel()

    if df_master.empty:
        return pd.DataFrame()

    search_term = str(search_term).strip().upper()
    station_filter_upper = str(station_filter).strip().upper()

    results_df = df_master[
        (df_master['STATION'] == station_filter_upper) & 
        ((df_master[hrms_id_col].str.contains(search_term, na=False)) |
         (df_master['EMPLOYEE NAME'].str.contains(search_term, na=False)))
    ].head(10).reset_index(drop=True)

    return results_df

def extract_employee_data(selected_row):
    """चयनित Pandas रो से आवश्यक डेटा निकालता है।"""
    df_master, hrms_id_col = load_master_excel()
    
    if df_master.empty or selected_row.empty:
        return None

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
    """केवल vacate के दौरान हिंदी नाम और पदनाम lookup करने के लिए।"""
    df_master, hrms_id_col = load_master_excel()
    
    if df_master.empty:
        return None
        
    hrms_id_str = str(hrms_id).strip()
    selected_rows = df_master[df_master[hrms_id_col] == hrms_id_str]
    
    if selected_rows.empty:
        return None

    selected_row = selected_rows.iloc[0]
    
    return {
        'employee_name_hindi': str(selected_row.get('EMPLOYEE NAME IN HINDI', str(selected_row.get('EMPLOYEE NAME', 'NA')))),
        'designation_hindi': str(selected_row.get('DESIGNATION IN HINDI', str(selected_row.get('DESIGNATION', 'NA')))),
    }

# ----------------------------------------------------------------------
# 6. AUTHENTICATION (प्रमाणीकरण)
# ----------------------------------------------------------------------

def check_password(password):
    """दिए गए पासवर्ड की जाँच करता है।"""
    return password == CORRECT_PASSWORD

def authenticate_user():
    """लॉगिन UI प्रदर्शित करता है और प्रमाणीकरण संभालता है।"""
    
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        
        # Centering the login form
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
                        st.rerun() # <--- यहाँ बदलाव किया गया है!
                    else:
                        st.error("Invalid Password. Please try again.")
        
        return False
    
    return True
    
    return True

# ----------------------------------------------------------------------
# 7. STREAMLIT UI (उपयोगकर्ता इंटरफ़ेस)
# ----------------------------------------------------------------------

def main_streamlit_ui():
    """मुख्य Streamlit इंटरफ़ेस।"""
    
    # 1. Run Authentication Check
    if not authenticate_user():
        return # Stop execution if user is not logged in

    # If execution reaches here, the user is authenticated
    st.title("🏡 रेलवे क्वार्टर प्रबंधन प्रणाली")
    
    # Run DB Initialization once
    initialize_database()

    # Session State Initialization and Data Load
    if 'quarter_df' not in st.session_state or st.session_state.quarter_df.empty:
        st.session_state.quarter_df = get_all_quarters()
    if 'search_results' not in st.session_state:
        st.session_state.search_results = pd.DataFrame()
    if 'selected_employee' not in st.session_state:
        st.session_state.selected_employee = None


    # Tabs for Navigation
    tab1, tab2, tab3, tab4 = st.tabs(["🏠 होम | स्थिति", "🔑 क्वार्टर आवंटन", "🗑️ क्वार्टर खाली करना", "📊 रिपोर्ट"])

    with tab1:
        st.header("वर्तमान क्वार्टर स्थिति")
        df_status = st.session_state.quarter_df
        
        col1, col2, col3 = st.columns(3)
        col1.metric("कुल क्वार्टर", len(df_status))
        col2.metric("आवंटित (Occupied)", len(df_status[df_status['current_status'] == 'Occupied']))
        col3.metric("खाली (Vacant)", len(df_status[df_status['current_status'] == 'Vacant']))
        
        st.subheader("मास्टर क्वार्टर सूची")
        st.dataframe(df_status, use_container_width=True)


    with tab2:
        st.header("🔑 नया क्वार्टर आवंटन")
        
        stations = st.session_state.quarter_df['station'].unique().tolist()
        stations.insert(0, '--- स्टेशन चुनें ---')
        
        selected_station = st.selectbox("1. स्टेशन चुनें", stations, key='allot_station')
        
        if selected_station and selected_station != '--- स्टेशन चुनें ---':
            
            # Filter available quarters (Vacant or Damaged)
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
                            # Station Filter is applied here automatically by search_employee_and_get_details
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
                            
                            st.session_state.quarter_df = get_all_quarters() # Update status
                            
                            if success:
                                st.success(f"**सफलता!** क्वार्टर {selected_q_num} को {st.session_state.selected_employee['employee_name_hindi']} को आवंटित किया गया।")
                                
                                # Download Button
                                if result is not None:
                                    file_name = f"{selected_station}_{selected_q_num.replace('/', '_')}_Allotment_{datetime.date.today().strftime('%Y%m%d')}.docx"
                                    st.download_button(
                                        label="Word पत्र डाउनलोड करें",
                                        data=result,
                                        file_name=file_name,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                else:
                                    st.warning("Word मेमो जेनरेट नहीं हो सका। कृपया टेम्पलेट फ़ाइल जाँच करें।")
                                    
                                # Reset for new allotment
                                st.session_state.search_results = pd.DataFrame()
                                st.session_state.selected_employee = None
                                
                            else:
                                st.error(result)

    with tab3:
        st.header("🗑️ क्वार्टर खाली करना")

        occupied_quarters_df = st.session_state.quarter_df[st.session_state.quarter_df['current_status'] == 'Occupied']
        
        if occupied_quarters_df.empty:
            st.info("वर्तमान में कोई भी क्वार्टर आवंटित नहीं है।")
        else:
            occupied_quarters_df['display'] = occupied_quarters_df['quarter_number'] + " (" + occupied_quarters_df['station'] + ")"
            
            quarters_to_vacate = occupied_quarters_df['display'].tolist()
            quarters_to_vacate.insert(0, '--- क्वार्टर चुनें ---')
            
            selected_quarter_display = st.selectbox("1. खाली करने के लिए क्वार्टर चुनें", quarters_to_vacate, key='vacate_q_select')
            
            if selected_quarter_display != '--- क्वार्टर चुनें ---':
                
                # Extract original quarter_number and station
                q_num, station = selected_quarter_display.split(' (')
                station = station.replace(')', '')
                
                vacate_date = st.date_input("2. खाली करने की तिथि चुनें", datetime.date.today(), key='vacate_date')

                if st.button("🗑️ वेकेशन मेमो जनरेट करें और क्वार्टर खाली करें", key='final_vacate_btn'):
                    
                    success, result = vacate_quarter(q_num, station, vacate_date)
                    
                    st.session_state.quarter_df = get_all_quarters() # Update status

                    if success:
                        st.success(f"**सफलता!** क्वार्टर {q_num} ({station}) सफलतापूर्वक खाली कर दिया गया।")
                        
                        # Download Button
                        if result is not None:
                            file_name = f"{station}_{q_num.replace('/', '_')}_Vacation_{datetime.date.today().strftime('%Y%m%d')}.docx"
                            st.download_button(
                                label="Word वेकेशन मेमो डाउनलोड करें",
                                data=result,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        else:
                            st.warning("Word मेमो जेनरेट नहीं हो सका। कृपया टेम्पलेट फ़ाइल जाँच करें।")

                    else:
                        st.error(result)


    with tab4:
        st.header("📊 रिपोर्ट जेनरेट करें")
        
        report_choice = st.radio(
            "रिपोर्ट का प्रकार चुनें:",
            ('वर्तमान स्थिति रिपोर्ट', 'संपूर्ण इतिहास रिपोर्ट'),
            key='report_type_select'
        )
        
        if report_choice == 'वर्तमान स्थिति रिपोर्ट':
            st.subheader("सभी क्वार्टरों की वर्तमान स्थिति")
            df_report = generate_current_status_report()
            st.dataframe(df_report, use_container_width=True)
            
            # Excel Download Button
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
            
            # Excel Download Button
            csv = df_history.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="संपूर्ण इतिहास रिपोर्ट (CSV) डाउनलोड करें",
                data=csv,
                file_name=f"Full_Quarter_History_{datetime.date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                key='download_full_history'
            )

if __name__ == '__main__':
    # Initial setup for data folder if not exists
    if not os.path.exists('data'):
        os.makedirs('data')

    main_streamlit_ui()