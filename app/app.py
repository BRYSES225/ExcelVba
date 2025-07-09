import streamlit as st
import pandas as pd
import openpyxl
import io
import base64

# Page configuration
st.set_page_config(
    page_title="Excel Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def load_excel_file(uploaded_file):
    """Load Excel file and return workbook and sheet names"""
    try:
        # Load workbook with openpyxl to preserve formulas
        workbook = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet_names = workbook.sheetnames
        
        # Also load with pandas for data manipulation
        excel_data = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
        
        return workbook, sheet_names, excel_data
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None, None, None

def display_formulas_info(workbook, sheet_name):
    """Display formula information for a sheet"""
    try:
        worksheet = workbook[sheet_name]
        formulas = []
        
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f':  # Formula cell
                    formulas.append({
                        'Cell': cell.coordinate,
                        'Formula': cell.value,
                        'Value': cell.displayed_value
                    })
        
        if formulas:
            with st.expander("🔢 Formulas in this sheet", expanded=False):
                st.dataframe(pd.DataFrame(formulas), use_container_width=True)
    except Exception as e:
        st.error(f"Error reading formulas: {str(e)}")

def create_download_link(df, filename):
    """Create a download link for the DataFrame"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def main():
    # Header
    st.title("📊 Interactive Excel Dashboard")
    st.markdown("Upload your Excel file to view and interact with all sheets")
    
    # Sidebar for file upload
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload an Excel file with multiple sheets"
        )
        
        if uploaded_file:
            st.success(f"File uploaded: {uploaded_file.name}")
            st.info(f"File size: {uploaded_file.size / 1024:.1f} KB")
    
    if uploaded_file is not None:
        # Load the Excel file
        with st.spinner("Loading Excel file..."):
            workbook, sheet_names, excel_data = load_excel_file(uploaded_file)
        
        if workbook and sheet_names and excel_data:
            st.success(f"Successfully loaded {len(sheet_names)} sheets")
            
            # Create tabs for each sheet
            tabs = st.tabs(sheet_names)
            
            for i, (tab, sheet_name) in enumerate(zip(tabs, sheet_names)):
                with tab:
                    df = excel_data[sheet_name]
                    
                    # Display sheet statistics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Rows", len(df))
                    with col2:
                        st.metric("Columns", len(df.columns))
                    with col3:
                        st.metric("Non-null cells", df.count().sum())
                    with col4:
                        st.metric("Memory usage", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
                    
                    # Display formulas info
                    display_formulas_info(workbook, sheet_name)
                    
                    # Display the data
                    st.subheader(f"📋 {sheet_name}")
                    if not df.empty:
                        st.dataframe(df, use_container_width=True, height=400)
                        
                        # Download section
                        st.markdown("---")
                        if st.button(f"📥 Download {sheet_name} as Excel", key=f"download_{i}"):
                            download_link = create_download_link(df, f"{sheet_name}.xlsx")
                            st.markdown(download_link, unsafe_allow_html=True)
                    else:
                        st.warning("This sheet is empty")
    
    else:
        # Welcome message
        st.info("👆 Please upload an Excel file using the sidebar to get started")
        
        # Example of what the app can do
        st.markdown("""
        ### Features:
        - 📊 **Multi-sheet support**: View all Excel sheets as separate tabs
        - 🔢 **Formula preservation**: See Excel formulas and their calculated values
        - 🔍 **Interactive tables**: Sort, filter, and select data
        - 📥 **Download capability**: Export individual sheets
        - 📱 **Responsive design**: Works on desktop and mobile
        
        ### Supported file formats:
        - `.xlsx` (Excel 2007+)
        - `.xls` (Excel 97-2003)
        """)

if __name__ == "__main__":
    main()

