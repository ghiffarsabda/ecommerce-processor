import streamlit as st
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re
from typing import Tuple
from dataclasses import dataclass
import json
from io import BytesIO

# Page config
st.set_page_config(
    page_title="E-commerce Order Processor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom styling
st.markdown("""
    <style>
    .main {
        background-color: #F8F9FA;
    }
    .stButton>button {
        background-color: #4285F4;
        color: white;
    }
    .stButton>button:hover {
        background-color: #357ABD;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}

class ProductDatabase:
    def __init__(self):
        try:
            # Read the local Excel file
            df = pd.read_excel('dcw_products.xlsx')
            # Convert first two columns to dictionary
            self.data = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        except Exception as e:
            st.error(f"Failed to load database: {str(e)}")
            self.data = {}

    def get_product_name(self, sku):
        return self.data.get(str(sku), "Unknown")

    def is_valid_sku(self, sku):
        return str(sku) in self.data

# Shopee Module
@dataclass
class ShopeeProduct:
    nama_produk: str
    nama_variasi: str
    jumlah: int
    sku: str

class ShopeeProcessor:
    def __init__(self, database):
        self.database = database

    def parse_product_info(self, text: str) -> list[ShopeeProduct]:
        products = []
        if '[1]' in text:
            items = re.split(r'\[\d+\]', text)
            items = [item.strip() for item in items if item.strip()]
        else:
            items = [text]
            
        for item in items:
            try:
                nama_produk_match = re.search(r'Nama Produk:([^;]+)', item)
                nama_variasi_match = re.search(r'Nama Variasi:([^;]+)', item)
                jumlah_match = re.search(r'Jumlah: (\d+)', item)
                sku_match = re.search(r'Nomor Referensi SKU: ?([^;]*);', item)
                
                if all([nama_produk_match, nama_variasi_match, jumlah_match]):
                    product = ShopeeProduct(
                        nama_produk=nama_produk_match.group(1).strip(),
                        nama_variasi=nama_variasi_match.group(1).strip(),
                        jumlah=int(jumlah_match.group(1)),
                        sku=sku_match.group(1).strip() if sku_match else ''
                    )
                    products.append(product)
            except Exception as e:
                st.error(f"Error parsing Shopee product: {e}")
                continue
        return products

    def process_file(self, file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            df = pd.read_excel(file)
            valid_products = []
            invalid_products = []
            
            for order_text in df.iloc[:, 7]:  # Column H
                if not isinstance(order_text, str):
                    continue
                    
                products = self.parse_product_info(order_text)
                
                for product in products:
                    if not product.sku or not self.database.is_valid_sku(product.sku):
                        invalid_products.append({
                            'Nama Produk': product.nama_produk,
                            'Nama Variasi': product.nama_variasi,
                            'Jumlah': product.jumlah
                        })
                    else:
                        valid_products.append({
                            'Kode SKU': product.sku,
                            'Nama Produk': self.database.get_product_name(product.sku),
                            'Jumlah': product.jumlah
                        })
            
            valid_df = pd.DataFrame(valid_products)
            invalid_df = pd.DataFrame(invalid_products)
            
            if not valid_df.empty:
                valid_df = valid_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                valid_df = valid_df.sort_values('Kode SKU')
            
            if not invalid_df.empty:
                invalid_df = invalid_df.sort_values('Nama Produk')
            
            return valid_df, invalid_df
            
        except Exception as e:
            st.error(f"Error processing Shopee file: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()
# Tokopedia Module
class TokopediaProcessor:
    def __init__(self, database):
        self.database = database

    def process_file(self, file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            # Read Excel file, skip first 4 rows
            df = pd.read_excel(file, skiprows=4)
            
            # Extract relevant columns (B=1, C=2, E=4)
            order_data = df.iloc[:, [1, 2, 4]].copy()
            order_data.columns = ['SKU', 'Nama_Produk', 'Jumlah']
            
            # Clean and convert SKU to string, handle NaN values
            order_data['SKU'] = order_data['SKU'].fillna('')
            order_data['SKU'] = order_data['SKU'].astype(str).str.strip()
            order_data['SKU'] = order_data['SKU'].apply(lambda x: x.split('.')[0] if '.' in x else x)
            
            valid_products = []
            invalid_products = []
            
            # Process each row
            for _, row in order_data.iterrows():
                sku = row['SKU']
                nama_produk = row['Nama_Produk']
                jumlah = row['Jumlah']
                
                if self.database.is_valid_sku(sku):
                    valid_products.append({
                        'Kode SKU': sku,
                        'Nama Produk': self.database.get_product_name(sku),
                        'Jumlah': jumlah
                    })
                else:
                    invalid_products.append({
                        'Nama Produk': nama_produk,
                        'Jumlah': jumlah
                    })
            
            # Create DataFrames
            valid_df = pd.DataFrame(valid_products)
            invalid_df = pd.DataFrame(invalid_products)
            
            # Group and sort valid products
            if not valid_df.empty:
                valid_df = valid_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                valid_df = valid_df.sort_values('Kode SKU')
            
            if not invalid_df.empty:
                invalid_df = invalid_df.sort_values('Nama Produk')
            
            return valid_df, invalid_df
            
        except Exception as e:
            st.error(f"Error processing Tokopedia file: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()

# TikTok Module
class TikTokProcessor:
    def __init__(self, database):
        self.database = database

    def process_file(self, file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            # Read Excel file without headers
            df = pd.read_excel(
                file,
                header=None,
                na_filter=False
            )
            
            # Skip first two rows
            data = df.iloc[2:].copy()
            
            # Create DataFrame with specific columns
            order_data = pd.DataFrame({
                'SKU': data.iloc[:, 6],          # Column G
                'Nama_Produk': data.iloc[:, 7],   # Column H
                'Variasi': data.iloc[:, 8],       # Column I
                'Jumlah': data.iloc[:, 9]         # Column J
            })
            
            # Clean SKUs and handle missing values
            order_data['SKU'] = order_data['SKU'].astype(str)
            order_data['SKU'] = order_data['SKU'].apply(lambda x: str(x).strip().strip("'"))
            order_data['Jumlah'] = pd.to_numeric(order_data['Jumlah'], errors='coerce').fillna(0)
            
            valid_products = []
            invalid_products = []
            
            for _, row in order_data.iterrows():
                sku = row['SKU']
                nama_produk = row['Nama_Produk']
                variasi = row['Variasi']
                try:
                    jumlah = int(float(row['Jumlah']))
                except:
                    jumlah = 0
                
                if self.database.is_valid_sku(sku):
                    valid_products.append({
                        'Kode SKU': sku,
                        'Nama Produk': self.database.get_product_name(sku),
                        'Jumlah': jumlah
                    })
                else:
                    invalid_products.append({
                        'Nama Produk': nama_produk,
                        'Variasi Produk': variasi,
                        'Jumlah': jumlah
                    })
            
            # Create DataFrames
            valid_df = pd.DataFrame(valid_products)
            invalid_df = pd.DataFrame(invalid_products)
            
            # Group and sort valid products
            if not valid_df.empty:
                valid_df = valid_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                valid_df = valid_df.sort_values('Kode SKU')
            
            if not invalid_df.empty:
                invalid_df = invalid_df.sort_values('Nama Produk')
            
            return valid_df, invalid_df
            
        except Exception as e:
            st.error(f"Error processing TikTok file: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()
def main():
    st.title("E-commerce Order Processor")
    
    # Initialize database
    database = ProductDatabase()
    
    # Initialize processors
    processors = {
        'Shopee': ShopeeProcessor(database),
        'Tokopedia': TokopediaProcessor(database),
        'TikTok': TikTokProcessor(database)
    }
    
    # Add session state if not exists
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = {}
    
    # Create two columns
    col1, col2 = st.columns([2, 1])
    
    with col1:
        platform = st.selectbox(
            "Select Platform",
            options=['Select Platform'] + list(processors.keys())
        )
        
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx']
        )
        
        if uploaded_file and platform != 'Select Platform':
            file_key = f"{platform}_{uploaded_file.name}"
            if file_key not in st.session_state.processed_files:
                st.session_state.processed_files[file_key] = {
                    'file': uploaded_file,
                    'platform': platform
                }
    
    # Display uploaded files
    if st.session_state.processed_files:
        st.write("Uploaded Files:")
        for key, info in st.session_state.processed_files.items():
            st.write(f"- {info['file'].name} ({info['platform']})")
        
        if st.button("Process Files"):
            process_files(st.session_state.processed_files, processors)
        
        if st.button("Clear All"):
            st.session_state.processed_files = {}
            st.experimental_rerun()

def process_files(files, processors):
    all_valid_products = []
    
    for file_key, info in files.items():
        with st.expander(f"Results for {info['file'].name}", expanded=True):
            processor = processors[info['platform']]
            valid_df, invalid_df = processor.process_file(info['file'])
            
            if not invalid_df.empty:
                st.subheader("Invalid Products")
                st.dataframe(invalid_df)
            
            if not valid_df.empty:
                st.subheader("Valid Products")
                st.dataframe(valid_df)
                all_valid_products.append(valid_df)
    
    if all_valid_products:
        st.subheader("Combined Summary")
        summary_df = pd.concat(all_valid_products)
        summary_df = summary_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
        summary_df = summary_df.sort_values('Kode SKU')
        
        st.dataframe(summary_df)
        
        # Download buttons
        col1, col2 = st.columns(2)
        with col1:
            csv = summary_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "Download CSV",
                csv,
                "summary.csv",
                "text/csv",
                key='download-csv'
            )
        
        with col2:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                summary_df.to_excel(writer, index=False)
            st.download_button(
                "Download Excel",
                buffer.getvalue(),
                "summary.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='download-excel'
            )

if __name__ == "__main__":
    main()
