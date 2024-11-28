import streamlit as st
import requests
import pandas as pd
import io

st.set_page_config(page_title="Variance Analysis & One-Time Expense", layout="wide")

# Predefined user credentials
USERNAME = "rpa.uat.bot1@sonata-software.com"
PASSWORD = "sonatabot$123"

# Initialize session state variables
if 'analysis_type' not in st.session_state:
    st.session_state.analysis_type = None
if 'sheets' not in st.session_state:
    st.session_state.sheets = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None
if 'account_type' not in st.session_state:
    st.session_state.account_type = None
if 'sheet_data_content' not in st.session_state:
    st.session_state.sheet_data_content = None  # Store the entire Excel file content
if 'job_id' not in st.session_state:
    st.session_state.job_id = None  # Store the job ID from the prepare step

# Sidebar Navigation
st.sidebar.title("Variance Analysis & One-Time Expense")

# Analysis type selection
st.sidebar.write("Choose your analysis type:")
analysis_type = st.sidebar.selectbox("Select Analysis Type:", ("Select Analysis Type", "One-Time Expense", "Variance Analysis"))
st.session_state.analysis_type = analysis_type

# Dropdown for selecting entity
file_option = st.sidebar.selectbox("Select Entity:", ("Select Entity", "SSL_full", "SSNA"))
file_name = f"{file_option}.xlsx" if file_option else None

# Fetch sheets
if st.sidebar.button("Proceed") and file_name:
    with st.spinner("Fetching Months, please wait..."):
        data = {
            "site_url": "https://o365sonata.sharepoint.com/sites/USECASE-4RMG",
            "folder_url": "/sites/USECASE-4RMG/Shared Documents/D365",
            "file_name": file_name,
            "username": USERNAME,
            "password": PASSWORD
        }
        # response = requests.post("http://localhost:7071/api/fetch_sheets", json=data)
        response = requests.post("https://varianceanalysisonetimeexpense.azurewebsites.net/api/fetch_sheets", json=data)
        if response.status_code == 200:
            st.session_state.sheets = response.json().get("sheets", [])
        else:
            st.sidebar.error("Failed to fetch sheets from backend.")

# Display sheet selection on the sidebar
if st.session_state.sheets:
    st.session_state.selected_sheet = st.sidebar.selectbox("Select a Month:", st.session_state.sheets)

# Account type selection (BS or P&L) on the sidebar
if st.session_state.selected_sheet:
    st.session_state.account_type = st.sidebar.radio("Select Account Type:", ("BS", "P&L"))

# Single Button for Generate Results (Triggers both Prepare and Perform Analysis)
if st.session_state.selected_sheet and st.session_state.account_type:
    if st.sidebar.button("Generate Results"):
        with st.spinner("Preparing data and performing analysis, please wait..."):
            # Step 1: Prepare Data
            prepare_payload = {
                "site_url": "https://o365sonata.sharepoint.com/sites/USECASE-4RMG",
                "folder_url": "/sites/USECASE-4RMG/Shared Documents/D365",
                "file_name": file_name,
                "sheet_name": st.session_state.selected_sheet,
                "username": USERNAME,
                "password": PASSWORD,
                "analysis_type": st.session_state.analysis_type,
                "account_type": st.session_state.account_type
            }
            # prepare_response = requests.post("http://localhost:7071/api/prepare_analysis", json=prepare_payload, timeout=600)
            prepare_response = requests.post("https://varianceanalysisonetimeexpense.azurewebsites.net/api/prepare_analysis", json=prepare_payload, timeout=600)
            
            if prepare_response.status_code == 200:
                job_id = prepare_response.json().get("job_id")
                st.session_state.job_id = job_id
                # st.success(f"Data prepared successfully! Job ID: {job_id}")
                
                # Step 2: Perform Analysis
                perform_payload = {
                    "job_id": st.session_state.job_id,
                    "analysis_type": st.session_state.analysis_type
                }
                # perform_response = requests.post("http://localhost:7071/api/perform_analysis", json=perform_payload, stream=True, timeout=600)
                perform_response = requests.post("https://varianceanalysisonetimeexpense.azurewebsites.net/api/perform_analysis", json=perform_payload, stream=True, timeout=600)
                
                if perform_response.status_code == 200:
                    st.session_state.sheet_data_content = perform_response.content
                    st.success("Analysis completed successfully!")
                else:
                    st.error("Failed to perform analysis.")
            else:
                st.error("Failed to prepare data.")

# Step 3: Display and Download Results
if st.session_state.sheet_data_content:
    st.write(f"### Result for {st.session_state.analysis_type}")
    with io.BytesIO(st.session_state.sheet_data_content) as file:
        excel_file = pd.ExcelFile(file)
        for sheet_name in excel_file.sheet_names:
            st.write(f"### {sheet_name}")
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            st.dataframe(df)
    
    st.download_button(
        label="Download Result",
        data=st.session_state.sheet_data_content,
        file_name=f"{st.session_state.analysis_type}_Result_{st.session_state.account_type}_{st.session_state.selected_sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
