# streamlit_app.py

import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe
from google.oauth2.service_account import Credentials
from google.api_core.exceptions import PermissionDenied, NotFound
from excel_agent_step_A import get_data_from_google_sheet
# --- Function to connect and fetch data from Google Sheets ---
# This is our core function from the previous script, slightly adapted for Streamlit.
# We add the @st.cache_data decorator to prevent re-running the function if the inputs haven't changed.

# --- Main Streamlit App Layout ---

st.set_page_config(layout="wide") # Use a wider layout for better table display

st.title("📊 Google Sheets Data Loader")
st.write("This app connects to a Google Sheet using a service account and displays the data as a table.")

st.sidebar.header("Connection Details")

# Input fields for the user
# We use our previously successful key and sheet name as defaults
sheet_key_input = st.sidebar.text_input(
    "Enter Google Sheet Key", 
    "13eZQCeGZwDJRk8-yQuIQMCtroZ62BCDmSMS7rClkhIU"
)
worksheet_name_input = st.sidebar.text_input(
    "Enter Worksheet Name", 
    "Creating_Tables"
)

# A button to trigger the action
if st.sidebar.button("Load Data", type="primary"):
    
    # Simple validation to ensure inputs are not empty
    if sheet_key_input and worksheet_name_input:
        
        with st.spinner("Connecting to Google Sheets and fetching data..."):
            # Call our function with the user's input
            dataframe, error = get_data_from_google_sheet(sheet_key_input, worksheet_name_input)
        
        # Display the results
        if error:
            st.error(error) # Show a red error box
        elif dataframe is not None:
            st.success("Data loaded successfully!")
            st.write("Here is a preview of your data:")
            st.dataframe(dataframe) # Display the dataframe in an interactive table
    else:
        st.warning("Please provide both a Sheet Key and a Worksheet Name.")

st.sidebar.info(
    "**How to use:**\n"
    "1. Enter the unique 'key' from your Google Sheet's URL.\n"
    "2. Enter the name of the specific tab (worksheet).\n"
    "3. Click 'Load Data'."
)