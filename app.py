import streamlit as st
import pandas as pd
from dataclasses import dataclass
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A6
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import mm
import io
from typing import Tuple
import base64

# Set page config
st.set_page_config(page_title="E-commerce Order Processor", layout="wide")

class ProductDatabase:
    def __init__(self):
        self.data = self.load_database()
    
    def load_database(self):
        """Load the product database without caching"""
        try:
            # Read the GitHub-hosted Excel file
            df = pd.read_excel('https://raw.githubusercontent.com/your-username/your-repo/main/dcw_products.xlsx')
            # Convert first two columns to dictionary
            return dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        except Exception as e:
            st.error(f"Failed to load database: {str(e)}")
            return {}

    def get_product_name(self, sku):
        return self.data.get(str(sku), "Unknown")

    def is_valid_sku(self, sku):
        return str(sku) in self.data

@st.cache_data
def initialize_database():
    """Initialize the database with caching"""
    return ProductDatabase()

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

class TokopediaProcessor:
    def __init__(self, database):
        self.database = database

    def process_file(self, file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            df = pd.read_excel(file, skiprows=4)
            
            order_data = df.iloc[:, [1, 2, 4]].copy()
            order_data.columns = ['SKU', 'Nama_Produk', 'Jumlah']
            
            order_data['SKU'] = order_data['SKU'].fillna('')
            order_data['SKU'] = order_data['SKU'].astype(str).str.strip()
            order_data['SKU'] = order_data['SKU'].apply(lambda x: x.split('.')[0] if '.' in x else x)
            
            valid_products = []
            invalid_products = []
            
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
            
            valid_df = pd.DataFrame(valid_products)
            invalid_df = pd.DataFrame(invalid_products)
            
            if not valid_df.empty:
                valid_df = valid_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                valid_df = valid_df.sort_values('Kode SKU')
            
            if not invalid_df.empty:
                invalid_df = invalid_df.sort_values('Nama Produk')
            
            return valid_df, invalid_df
            
        except Exception as e:
            st.error(f"Error processing Tokopedia file: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()

class TikTokProcessor:
    def __init__(self, database):
        self.database = database

    def process_file(self, file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            df = pd.read_excel(file, header=None, na_filter=False)
            data = df.iloc[2:].copy()
            
            order_data = pd.DataFrame({
                'SKU': data.iloc[:, 6],
                'Nama_Produk': data.iloc[:, 7],
                'Variasi': data.iloc[:, 8],
                'Jumlah': data.iloc[:, 9]
            })
            
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
            
            valid_df = pd.DataFrame(valid_products)
            invalid_df = pd.DataFrame(invalid_products)
            
            if not valid_df.empty:
                valid_df = valid_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                valid_df = valid_df.sort_values('Kode SKU')
            
            if not invalid_df.empty:
                invalid_df = invalid_df.sort_values('Nama Produk')
            
            return valid_df, invalid_df
            
        except Exception as e:
            st.error(f"Error processing TikTok file: {str(e)}")
            return pd.DataFrame(), pd.DataFrame()

def create_pdf(data, size='A4'):
    """Create PDF from summary data"""
    buffer = io.BytesIO()
    
    # Set page size
    if size == 'A4':
        pagesize = A4
        margin = 25 * mm
    else:  # A6
        pagesize = A6
        margin = 10 * mm

    # Create PDF document
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )

    # Prepare data
    table_data = [['SKU', 'Product Name', 'Quantity']]  # Headers
    table_data.extend(data)

    # Calculate column widths
    available_width = pagesize[0] - 2*margin
    col_widths = [available_width*0.2, available_width*0.6, available_width*0.2]

    # Create table
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12 if size == 'A4' else 8),
        ('FONTSIZE', (0, 1), (-1, -1), 10 if size == 'A4' else 6),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12 if size == 'A4' else 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ]))

    # Build PDF
    doc.build([table])
    buffer.seek(0)
    return buffer

def get_download_link(buffer, filename, text):
    """Generate a download link for a file"""
    b64 = base64.b64encode(buffer.getvalue()).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{text}</a>'

def main():
    st.title("E-commerce Order Processor")
    
    # Initialize database
    database = initialize_database()
    
    # Initialize processors
    processors = {
        'Shopee': ShopeeProcessor(database),
        'Tokopedia': TokopediaProcessor(database),
        'TikTok': TikTokProcessor(database)
    }
    
    # File upload section
    st.subheader("Upload Files")
    
    # Create columns for file upload
    col1, col2 = st.columns(2)
    
    with col1:
        platform = st.selectbox(
            "Select Platform",
            ['Select Platform', 'Shopee', 'Tokopedia', 'TikTok']
        )
    
    with col2:
        uploaded_file = st.file_uploader(
            "Choose Excel file",
            type=['xlsx'],
            key="file_uploader"
        )
    
    # Process file when uploaded
    if uploaded_file is not None and platform != 'Select Platform':
        st.subheader("Processing Results")
        
        # Process the file
        processor = processors[platform]
        valid_df, invalid_df = processor.process_file(uploaded_file)
        
        # Display results in tabs
        tab1, tab2 = st.tabs(["Valid Products", "Invalid Products"])
        
        with tab1:
            if not valid_df.empty:
                st.write("Valid Products:")
                st.dataframe(valid_df)
            else:
                st.info("No valid products found.")
        
        with tab2:
            if not invalid_df.empty:
                st.write("Invalid Products:")
                st.dataframe(invalid_df)
            else:
                st.info("No invalid products found.")
        
        # Export options
        if not valid_df.empty:
            st.subheader("Export Options")
            
            col1, col2, col3 = st.columns(3)
            
            # Export to Excel
            with col1:
                excel_buffer = io.BytesIO()
                valid_df.to_excel(excel_buffer, index=False)
                excel_link = get_download_link(excel_buffer, "summary.xlsx", "Download Excel")
                st.markdown(excel_link, unsafe_allow_html=True)
            
            # Export to PDF (A4)
            with col2:
                pdf_buffer_a4 = create_pdf(valid_df.values.tolist(), 'A4')
                pdf_link_a4 = get_download_link(pdf_buffer_a4, "summary_a4.pdf", "Download PDF (A4)")
                st.markdown(pdf_link_a4, unsafe_allow_html=True)
            
            # Export to PDF (A6)
            with col3:
                pdf_buffer_a6 = create_pdf(valid_df.values.tolist(), 'A6')
                pdf_link_a6 = get_download_link(pdf_buffer_a6, "summary_a6.pdf", "Download PDF (A6)")
                st.markdown(pdf_link_a6, unsafe_allow_html=True)

    # Add footer with instructions
    st.markdown("---")
    st.markdown("""
    ### Instructions:
    1. Select the e-commerce platform (Shopee, Tokopedia, or TikTok)
    2. Upload your Excel file containing order data
    3. View the processed results in the Valid and Invalid Products tabs
    4. Download the summary in your preferred format (Excel, PDF A4, or PDF A6)
    
    **Note:** Make sure your Excel files follow the expected format for each platform.
    """)

if __name__ == "__main__":
    main()
