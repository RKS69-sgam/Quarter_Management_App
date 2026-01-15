import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
import io
import os
import firebase_admin
from firebase_admin import credentials, firestore

# --- 0. Firebase Setup ---
@st.cache_resource
def init_db():
    if not firebase_admin._apps:
        try:
            if "firebase_config" in st.secrets:
                cred = credentials.Certificate(dict(st.secrets["firebase_config"]))
            else:
                cred = credentials.Certificate('sgamoffice-firebase-adminsdk-fbsvc-253915b05b.json')
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}"); st.stop()
    return firestore.client()

# --- 1. Admin Functions (Clean & Import) ---
def delete_collection(collection_name):
    db = init_db()
    docs = db.collection(collection_name).limit(500).stream()
    deleted = 0
    for doc in docs:
        doc.reference.delete()
        deleted += 1
    return deleted

def import_csv_to_firestore(df):
    db = init_db()
    success_count = 0
    for _, row in df.iterrows():
        # Remark ke aadhar par is_current set karein
        is_occ = True if str(row.get('Remark', '')).lower() == 'occupation' else False
        
        data = {
            "station": str(row.get('Station', '')),
            "quarter_number": str(row.get('Quarter Number', '')),
            "pf_number": str(row.get('PF NUMBER', '')),
            "hrms_id": str(row.get('HRMS ID', '')),
            "employee_name": str(row.get('Employee Name', '')),
            "allotment_date": pd.to_datetime(row['Occupation Date']).to_pydatetime() if pd.notna(row.get('Occupation Date')) else None,
            "vacation_date": pd.to_datetime(row['Vacant Date']).to_pydatetime() if pd.notna(row.get('Vacant Date')) else None,
            "is_current": is_occ,
            "designation": "", # Employees collection se fetch hoga allotment ke waqt
            "unit": "SSE/P.Way/SGAM"
        }
        db.collection("quarter_history").add(data)
        success_count += 1
    return success_count

# --- 2. Main App ---
def main():
    st.set_page_config(layout="wide", page_title="Railway Quarter Management")
    db = init_db()

    # Sidebar for Admin Tools
    st.sidebar.title("⚙️ Admin Tools")
    if st.sidebar.checkbox("Show Data Cleaning Tools"):
        st.sidebar.warning("Savadhan: Yeh action database saaf kar dega!")
        
        uploaded_file = st.sidebar.file_uploader("Upload Quarter Register CSV", type=["csv"])
        
        if st.sidebar.button("🔥 Clean & Import New Data"):
            if uploaded_file is not None:
                with st.spinner("Purana data delete ho raha hai..."):
                    del_count = delete_collection("quarter_history")
                    st.sidebar.success(f"Deteled {del_count} old records.")
                
                with st.spinner("Naya data import ho raha hai..."):
                    new_df = pd.read_csv(uploaded_file)
                    imp_count = import_csv_to_firestore(new_df)
                    st.sidebar.success(f"Imported {imp_count} records successfully!")
                    st.cache_data.clear()
                    st.rerun()
            else:
                st.sidebar.error("Kripya pehle CSV file upload karein.")

    # --- Tabs (Allotment, Vacation, Report) ---
    tab1, tab2, tab3 = st.tabs(["🏠 Allotment", "🗝️ Vacation", "📊 Report"])

    # (Baki allotment aur vacation ka logic pehle jaisa hi rahega...)
    
    with tab3:
        st.header("📊 Quarter Master Report")
        # Data fetch logic (already provided in previous response)
        # ...
