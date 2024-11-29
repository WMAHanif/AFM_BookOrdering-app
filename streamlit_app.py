import streamlit as st
import pandas as pd
from tkcalendar import Calendar
from datetime import datetime
import os
import pickle

# Initialize global variables for the files
MASTER_DATA_FILE = "master_data.pkl"
master_data = pd.DataFrame()
uploaded_data = pd.DataFrame()


# Utility Functions
def save_master_data(data):
    try:
        with open(MASTER_DATA_FILE, "wb") as file:
            pickle.dump(data, file)
    except Exception as e:
        st.error(f"Failed to save master data to file: {e}")


def load_master_data():
    global master_data
    if os.path.exists(MASTER_DATA_FILE):
        try:
            with open(MASTER_DATA_FILE, "rb") as file:
                master_data = pickle.load(file)
            st.sidebar.success("Master Data Loaded Successfully")
        except Exception as e:
            st.sidebar.error(f"Failed to load master data: {e}")
    else:
        st.sidebar.warning("No master data file found.")


def save_uploaded_file(df, filename_prefix):
    try:
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        output_dir = os.path.join(desktop_path, "Processed_Files")

        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)

        # Generate a timestamp for the file name to avoid overwrite
        timestamp = datetime.now().strftime('%y%m%d_%H%M%S')
        output_file = os.path.join(output_dir, f"{filename_prefix}_{timestamp}.xlsx")

        # Save the DataFrame to Excel
        df.to_excel(output_file, index=False)

        st.success(f"File saved successfully at: {output_file}")
    except Exception as e:
        st.error(f"Failed to save file: {e}")


# App Pages
def main_page():
    st.title("ERP Order Processor")
    st.write("Welcome! Select an option from the sidebar to get started.")

    if st.sidebar.button("Load Master Data"):
        load_master_data()


def upload_master_data():
    global master_data
    st.header("Upload Master Data")
    file = st.file_uploader("Choose a Master Data Excel File", type=["xlsx"])

    if file:
        try:
            master_data = pd.read_excel(file)
            save_master_data(master_data)
            st.success("Master Data Uploaded and Saved Successfully")
        except Exception as e:
            st.error(f"Failed to process master data: {e}")


def process_files():
    global uploaded_data, master_data

    st.header("Upload and Process Files")
    file_type = st.selectbox("Select File Type", ["BookCapital", "MySiswa", "ERP"])
    shop_name = st.selectbox("Select Shop Name", ["BOOKCAFE", "IMAN OFFLINE", "FIXI"])

    file = st.file_uploader("Choose an Excel File", type=["xlsx", "xls"])
    if file and master_data.empty:
        st.warning("Please upload master data first from the sidebar.")

    if file:
        try:
            if file_type == "BookCapital":
                uploaded_data = pd.read_excel(file)
                # Example processing logic for BookCapital
                uploaded_data['Matched SKU'] = uploaded_data['SKU'].apply(
                    lambda x: master_data.loc[master_data['SKU'] == x, 'SKU'].values[0]
                    if x in master_data['SKU'].values else None
                )
                save_uploaded_file(uploaded_data, "Processed_BookCapital")
                st.success("BookCapital File Processed Successfully")
            elif file_type == "MySiswa":
                uploaded_data = pd.read_excel(file)
                # Example processing logic for MySiswa
                st.write("Processing MySiswa file...")
            elif file_type == "ERP":
                uploaded_data = pd.read_excel(file)
                # Example processing logic for ERP
                st.write("Processing ERP file...")
        except Exception as e:
            st.error(f"Error while processing file: {e}")


# Main App Execution
st.sidebar.title("ERP Order Processor")
st.sidebar.markdown("Navigate through the options below:")

pages = {
    "Main Page": main_page,
    "Upload Master Data": upload_master_data,
    "Process Files": process_files,
}

selected_page = st.sidebar.selectbox("Select Page", list(pages.keys()))
pages[selected_page]()
