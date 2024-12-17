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
from datetime import datetime

# Set page config
st.set_page_config(page_title="E-commerce Order Processor", layout="wide")

class ProductDatabase:
    def __init__(self):
        try:
            # Read the GitHub-hosted Excel file with correct raw URL format
            df = pd.read_excel('https://raw.githubusercontent.com/ghiffarsabda/ecommerce-processor/main/dcw_products.xlsx')
            # Convert first two columns to dictionary
            self.data = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        except Exception as e:
            st.error(f"Failed to load database: {str(e)}")
            self.data = {}

    def get_product_name(self, sku):
        return self.data.get(str(sku), "Unknown")

    def is_valid_sku(self, sku):
        return str(sku) in self.data

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
            df = pd.read_excel(file, skiprows=3)
            
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
    database = ProductDatabase()
    
    # Initialize processors
    processors = {
        'Shopee': ShopeeProcessor(database),
        'Tokopedia': TokopediaProcessor(database),
        'TikTok': TikTokProcessor(database)
    }

    # Initialize session state for files if it doesn't exist
    if 'files' not in st.session_state:
        st.session_state.files = {}  # {file_key: {'file': file_obj, 'platform': platform}}

    # File upload section with reorganized layout
    st.header("Upload Files")
    st.subheader("Select Platform")
    platform = st.selectbox(
        label="Choose your e-commerce platform",
        options=['Shopee', 'Tokopedia', 'TikTok']
    )
    
    uploaded_file = st.file_uploader(
        "Browse files",
        type=['xlsx']
    )

    # Show filename if file is uploaded
    if uploaded_file:
        st.text(uploaded_file.name)
    
    # Add file button
    if st.button("Add File") and uploaded_file is not None and platform != 'Select Platform':
        # Generate unique key for the file
        file_key = f"{platform}_{uploaded_file.name}"
        if file_key not in st.session_state.files:
            # Store file in session state
            file_data = uploaded_file.getvalue()
            st.session_state.files[file_key] = {
                'file': file_data,
                'platform': platform,
                'filename': uploaded_file.name
            }
            st.success(f"Added {uploaded_file.name} for {platform}")
        else:
            st.warning("This file has already been added!")

    # Display added files
    if st.session_state.files:
        st.markdown("---")
        st.subheader("Added Files")
        for file_key, file_info in list(st.session_state.files.items()):
            col1, col2 = st.columns([5, 1])
            with col1:
                st.text(f"{file_info['filename']} ({file_info['platform']})")
            with col2:
                if st.button("Remove", key=f"remove_{file_key}"):
                    del st.session_state.files[file_key]
                    st.rerun()

        # Process button
        if st.button("Process All Files"):
            if st.session_state.files:
                all_valid_products = []
                all_invalid_products = []
                all_valid_details = []
                all_invalid_details = []
                
                # Create tabs - one Summary tab, one Joined Details tab, and one tab per file
                tabs = ["Summary", "Joined Details"] + [f"{info['filename']}" for info in st.session_state.files.values()]
                current_tabs = st.tabs(tabs)
                
                # Process each file and show in respective tabs
                for idx, (file_key, file_info) in enumerate(st.session_state.files.items(), start=2):
                    with current_tabs[idx]:
                        st.subheader(f"Results for {file_info['filename']}")
                        
                        try:
                            # Convert bytes back to file-like object
                            file_obj = io.BytesIO(file_info['file'])
                            
                            # Get appropriate processor
                            processor = processors[file_info['platform']]
                            
                            # Process file
                            valid_df, invalid_df = processor.process_file(file_obj)
                            
                            if not valid_df.empty:
                                # Add to summary
                                summary_df = valid_df.copy()
                                summary_df = summary_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                                all_valid_products.append(summary_df)
                                
                                # Add to detailed view with filename
                                detailed_df = valid_df.copy()
                                detailed_df['DateToday'] = datetime.now().strftime('%Y-%m-%d')
                                detailed_df['File Name'] = file_info['filename']
                                all_valid_details.append(detailed_df)
                            
                            # Handle invalid products
                            if not invalid_df.empty:
                                st.error("Invalid Products Found:")
                                st.dataframe(invalid_df)
                                all_invalid_products.append(invalid_df)
                                
                                # Add to detailed invalid view with filename
                                detailed_invalid_df = invalid_df.copy()
                                detailed_invalid_df['DateToday'] = datetime.now().strftime('%Y-%m-%d')
                                detailed_invalid_df['File Name'] = file_info['filename']
                                all_invalid_details.append(detailed_invalid_df)
                            
                            # Display valid products
                            if not valid_df.empty:
                                st.success("Valid Products:")
                                st.dataframe(valid_df)
                            
                        except Exception as e:
                            st.error(f"Error processing {file_info['filename']}: {str(e)}")
                
                # Summary tab content
                with current_tabs[0]:
                    if all_valid_products:
                        # Combine and group all valid products
                        summary_df = pd.concat(all_valid_products)
                        summary_df = summary_df.groupby(['Kode SKU', 'Nama Produk'])['Jumlah'].sum().reset_index()
                        summary_df = summary_df.sort_values('Kode SKU')
                        
                        st.success("Summary of All Valid Products")
                        st.dataframe(summary_df)
                        
                        # Export options
                        st.subheader("Export Options")
                        
                        # Create a container with custom spacing
                        export_container = st.container()
                        with export_container:
                            col1, space1, col2, space2, col3 = st.columns([1, 0.1, 1, 0.1, 1])
                        
                            # Export to Excel
                            with col1:
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    summary_df.to_excel(writer, index=False)
                                excel_data = excel_buffer.getvalue()
                                st.download_button(
                                    label="Download Excel",
                                    data=excel_data,
                                    file_name="summary.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                            
                            # Export to PDF A4
                            with col2:
                                pdf_buffer_a4 = create_pdf(summary_df.values.tolist(), 'A4')
                                st.download_button(
                                    label="Download PDF (A4)",
                                    data=pdf_buffer_a4.getvalue(),
                                    file_name="summary_a4.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                            
                            # Export to PDF A6
                            with col3:
                                pdf_buffer_a6 = create_pdf(summary_df.values.tolist(), 'A6')
                                st.download_button(
                                    label="Download PDF (A6)",
                                    data=pdf_buffer_a6.getvalue(),
                                    file_name="summary_a6.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                
                        if all_invalid_products:
                            st.markdown("---")
                            st.error("Summary of All Invalid Products")
                            invalid_summary = pd.concat(all_invalid_products)
                            st.dataframe(invalid_summary)

                # Joined Details tab content
                with current_tabs[1]:
                    st.subheader("Joined Details")
                    
                    # Display joined valid products
                    if all_valid_details:
                        joined_valid_df = pd.concat(all_valid_details, ignore_index=True)
                        joined_valid_df = joined_valid_df[['DateToday', 'Kode SKU', 'Nama Produk', 'Jumlah', 'File Name']]
                        st.success("All Valid Products (Detailed View)")
                        st.dataframe(joined_valid_df)
                        
                        # Add export buttons for joined details
                        st.subheader("Export Joined Details")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                joined_valid_df.to_excel(writer, index=False, sheet_name='Valid Products')
                                if all_invalid_details:
                                    joined_invalid_df = pd.concat(all_invalid_details, ignore_index=True)
                                    joined_invalid_df.to_excel(writer, index=False, sheet_name='Invalid Products')
                            
                            st.download_button(
                                label="Download Detailed Excel",
                                data=excel_buffer.getvalue(),
                                file_name="detailed_summary.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    
                    # Display joined invalid products
                    if all_invalid_details:
                        st.markdown("---")
                        joined_invalid_df = pd.concat(all_invalid_details, ignore_index=True)
                        st.error("All Invalid Products (Detailed View)")
                        st.dataframe(joined_invalid_df)
            
            else:
                st.warning("No files to process!")

        # Clear all button
        if st.button("Clear All Files"):
            st.session_state.files = {}
            st.rerun()

    # Add footer with instructions
    st.markdown("---")
    st.markdown("""
    ### Instructions:
    1. Select the e-commerce platform (Shopee, Tokopedia, or TikTok)
    2. Upload your Excel file and click "Add File"
    3. Repeat steps 1-2 for all files you want to process
    4. Click "Process All Files" to analyze all uploaded files
    5. View individual results in respective file tabs
    6. View combined summary in Summary tab
    7. View all detailed records in Joined Details tab
    8. Download the results in your preferred format
    
    **Note:** Make sure your Excel files follow the expected format for each platform.
    """)

if __name__ == "__main__":
    main()
