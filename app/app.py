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

def clean_dataframe(df):
    """Clean dataframe to handle mixed data types and display issues"""
    try:
        # Make a copy to avoid modifying original
        df_clean = df.copy()
        
        # Handle mixed data types in columns
        for col in df_clean.columns:
            # Convert problematic columns to string to avoid Arrow errors
            if df_clean[col].dtype == 'object':
                # Check for mixed types that cause Arrow issues
                sample_values = df_clean[col].dropna().head(10)
                if len(sample_values) > 0:
                    types = sample_values.apply(type).unique()
                    if len(types) > 1:
                        # Mixed types - convert to string
                        df_clean[col] = df_clean[col].astype(str)
        
        # Replace NaN values with empty strings
        df_clean = df_clean.fillna('')
        
        # Clean column names
        df_clean.columns = [f'Column_{i}' if str(col).startswith('Unnamed:') else str(col) 
                           for i, col in enumerate(df_clean.columns)]
        
        return df_clean
        
    except Exception as e:
        st.warning(f"Error cleaning dataframe: {str(e)}")
        return df

def load_excel_file(uploaded_file):
    """Load Excel file and return workbook and sheet names"""
    try:
        # Load workbook with openpyxl to preserve formulas and macros
        workbook = openpyxl.load_workbook(uploaded_file, data_only=False, keep_vba=True)
        sheet_names = workbook.sheetnames
        
        # Load with pandas and clean data types
        excel_data = {}
        for sheet_name in sheet_names:
            try:
                # Read the sheet
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
                
                # Clean the dataframe for better display
                df = clean_dataframe(df)
                excel_data[sheet_name] = df
                
            except Exception as e:
                st.warning(f"Could not read sheet '{sheet_name}': {str(e)}")
                excel_data[sheet_name] = pd.DataFrame()
        
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
                        'Formula': str(cell.value),
                        'Value': str(cell.displayed_value) if cell.displayed_value else ''
                    })
        
        if formulas:
            with st.expander(f"🔢 Formulas in {sheet_name} ({len(formulas)} found)", expanded=False):
                formula_df = pd.DataFrame(formulas)
                st.dataframe(formula_df, use_container_width=True)
    except Exception as e:
        st.warning(f"Could not analyze formulas: {str(e)}")

def display_macro_info(workbook):
    """Display macro/VBA information if present"""
    try:
        if hasattr(workbook, 'vba_archive') and workbook.vba_archive:
            st.success("🔧 **Macro-enabled Excel file detected!**")
            with st.expander("ℹ️ Macro Information", expanded=False):
                st.info("⚠️ Macros are preserved but won't execute in the browser for security reasons.")
        else:
            st.info("📄 Standard Excel file (no macros detected)")
    except Exception as e:
        st.info("ℹ️ Could not detect macro information")

def create_download_link(df, filename):
    """Create a download link for the DataFrame"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        output.seek(0)
        b64 = base64.b64encode(output.read()).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Download {filename}</a>'
        return href
    except Exception as e:
        st.error(f"Error creating download link: {str(e)}")
        return None

def main():
    # Header
    st.title("📊 Interactive Excel Dashboard")
    st.markdown("Upload your Excel file to view and interact with all sheets")
    
    # Sidebar for file upload
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls', 'xlsm'],
            help="Upload an Excel file with multiple sheets (supports macro-enabled files)"
        )
        
        if uploaded_file:
            st.success(f"✅ File: {uploaded_file.name}")
            st.info(f"📏 Size: {uploaded_file.size / 1024:.1f} KB")
    
    if uploaded_file is not None:
        # Load the Excel file
        with st.spinner("🔄 Loading Excel file..."):
            workbook, sheet_names, excel_data = load_excel_file(uploaded_file)
        
        if workbook and sheet_names and excel_data:
            st.success(f"✅ Successfully loaded {len(sheet_names)} sheets")
            
            # Display macro information
            display_macro_info(workbook)
            
            # Create tabs for each sheet
            tabs = st.tabs(sheet_names)
            
            for i, (tab, sheet_name) in enumerate(zip(tabs, sheet_names)):
                with tab:
                    df = excel_data[sheet_name]
                    
                    if not df.empty:
                        # Display sheet statistics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("📊 Rows", len(df))
                        with col2:
                            st.metric("📋 Columns", len(df.columns))
                        with col3:
                            st.metric("✅ Non-null", df.count().sum())
                        with col4:
                            memory_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
                            st.metric("💾 Memory", f"{memory_mb:.1f} MB")
                        
                        # Display formulas info
                        display_formulas_info(workbook, sheet_name)
                        
                        # Display the data with error handling
                        st.subheader(f"📋 {sheet_name} Data")
                        try:
                            st.dataframe(df, use_container_width=True, height=400)
                        except Exception as e:
                            st.error(f"Error displaying data: {str(e)}")
                            st.markdown("**Raw data preview:**")
                            st.text(str(df.head()))
                        
                        # Download section
                        st.markdown("---")
                        if st.button(f"📥 Download {sheet_name}", key=f"download_{i}"):
                            download_link = create_download_link(df, f"{sheet_name}.xlsx")
                            if download_link:
                                st.markdown(download_link, unsafe_allow_html=True)
                    else:
                        st.warning("⚠️ This sheet is empty or could not be read")
    
    else:
        # Welcome message
        st.info("👆 Please upload an Excel file using the sidebar to get started")
        
        # Feature showcase
        st.markdown("""
        ### 🚀 Features:
        - **📊 Multi-sheet support**: View all Excel sheets as separate tabs
        - **🔢 Formula preservation**: See Excel formulas and their calculated values
        - **🔧 Macro support**: Handle .xlsm files with VBA macros
        - **📥 Download capability**: Export individual sheets
        - **🔒 Data cleaning**: Automatic handling of mixed data types
        
        ### 📁 Supported formats:
        - **`.xlsm`** - Excel with macros
        - **`.xlsx`** - Standard Excel 2007+
        - **`.xls`** - Legacy Excel 97-2003
        """)

if __name__ == "__main__":
    main()
