import streamlit as st
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

st.title("Simple Test App")

try:
    # Test credentials loading
    creds = Credentials.from_service_account_info(
        json.loads(st.secrets["google_credentials"]),
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    st.success("Credentials loaded successfully!")

    # Test Google Sheets connection
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=st.secrets["spreadsheet_id"],
        range=st.secrets["range_name"]
    ).execute()
    
    st.success("Connected to Google Sheets!")
    st.write("First few rows:", result['values'][:5])

except Exception as e:
    st.error(f"Error: {str(e)}")
