import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# Set page config
st.set_page_config(
    page_title="Excel Duplicate Checker",
    page_icon="🔍",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {font-size: 2.5rem; color: #1f77b4; text-align: center; margin-bottom: 1rem;}
    .sub-header {font-size: 1.5rem; color: #2c3e50; margin: 1.5rem 0 1rem 0;}
    .success-box {background-color: #e8f5e9; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;}
    .warning-box {background-color: #fff3e0; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;}
    .error-box {background-color: #ffebee; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0;}
    .stDataFrame {border-radius: 0.5rem;}
    .stButton>button {width: 100%; margin: 0.5rem 0;}
    </style>
""", unsafe_allow_html=True)

def load_data(uploaded_file):
    """Load data from uploaded Excel file."""
    try:
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            xls = pd.ExcelFile(uploaded_file)
            sheets = {}
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                sheets[sheet_name] = df
            return sheets
        else:
            st.error("Please upload a valid Excel file (.xlsx or .xls)")
            return None
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def find_duplicates(df, columns):
    """Find duplicate rows based on specified columns."""
    if not columns:
        return pd.DataFrame()
    
    # Find duplicates
    duplicates = df[df.duplicated(subset=columns, keep=False)]
    
    # Sort by the duplicate columns for better visualization
    if not duplicates.empty:
        duplicates = duplicates.sort_values(by=columns)
        
    return duplicates

def highlight_duplicates(s, columns):
    """Highlight duplicate rows in the dataframe."""
    if s.name in columns:
        is_duplicate = s.duplicated(keep=False)
        return ['background-color: #ffeb3b' if v else '' for v in is_duplicate]
    return [''] * len(s)

def main():
    st.markdown("<h1 class='main-header'>🔍 Excel Duplicate Checker</h1>", unsafe_allow_html=True)
    st.markdown("Upload an Excel file to check for duplicate entries based on selected columns.")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Load data
        sheets = load_data(uploaded_file)
        
        if sheets:
            # Sheet selection
            sheet_name = st.selectbox("Select a sheet", list(sheets.keys()))
            df = sheets[sheet_name]
            
            # Display basic info
            st.markdown(f"<div class='success-box'>Loaded sheet: <strong>{sheet_name}</strong> with {len(df)} rows and {len(df.columns)} columns.</div>", unsafe_allow_html=True)
            
            # Column selection for duplicate checking
            st.markdown("<div class='sub-header'>Select columns to check for duplicates:</div>", unsafe_allow_html=True)
            
            # Create two columns for better layout
            col1, col2 = st.columns(2)
            
            with col1:
                selected_columns = st.multiselect(
                    "Select columns",
                    options=df.columns.tolist(),
                    default=[df.columns[0]] if len(df.columns) > 0 else []
                )
            
            # Find and display duplicates
            if selected_columns:
                duplicates = find_duplicates(df, selected_columns)
                
                if not duplicates.empty:
                    st.markdown(f"<div class='warning-box'>Found {len(duplicates)} duplicate rows based on selected columns.</div>", unsafe_allow_html=True)
                    
                    # Display duplicates with highlighting
                    st.dataframe(
                        duplicates.style.apply(lambda x: highlight_duplicates(x, selected_columns), axis=0),
                        use_container_width=True,
                        height=400
                    )
                    
                    # Show duplicate summary
                    st.markdown("<div class='sub-header'>Duplicate Summary:</div>", unsafe_allow_html=True)
                    dup_summary = duplicates[selected_columns].value_counts().reset_index()
                    dup_summary.columns = selected_columns + ['Count']
                    st.dataframe(dup_summary, use_container_width=True)
                    
                    # Export options
                    st.markdown("<div class='sub-header'>Export Options:</div>", unsafe_allow_html=True)
                    
                    # Create download buttons
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        duplicates.to_excel(writer, sheet_name='Duplicates', index=False)
                        df[~df.index.isin(duplicates.index)].to_excel(writer, sheet_name='Unique', index=False)
                    
                    st.download_button(
                        label="📥 Download Duplicates",
                        data=output.getvalue(),
                        file_name=f"duplicates_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                else:
                    st.markdown("<div class='success-box'>No duplicates found based on the selected columns. ✅</div>", unsafe_allow_html=True)
            
            # Display raw data
            with st.expander("View Raw Data"):
                st.dataframe(df, use_container_width=True, height=400)
            
            # Data statistics
            with st.expander("Data Statistics"):
                st.write("### Data Types")
                st.write(df.dtypes)
                
                st.write("### Missing Values")
                missing = df.isnull().sum().reset_index()
                missing.columns = ['Column', 'Missing Values']
                st.dataframe(missing, use_container_width=True)
                
                st.write("### Basic Statistics")
                st.dataframe(df.describe(include='all').T, use_container_width=True)
    
    # Add footer
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #666;'>🔍 Excel Duplicate Checker | Last updated: 2025-09-20</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()