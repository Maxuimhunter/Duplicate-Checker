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

def highlight_blanks(s):
    """Highlight blank/empty cells in the dataframe."""
    is_blank = s.isna() | (s.astype(str).str.strip() == '')
    return ['background-color: #ffcdd2' if v else '' for v in is_blank]

def clean_dataframe(df):
    """Clean the dataframe by removing empty rows and trimming whitespace."""
    try:
        # Make a copy to avoid modifying the original
        df_clean = df.copy()
        
        # Store original dtypes for conversion back later
        original_dtypes = df_clean.dtypes
        
        # Convert all columns to string, clean, then convert back
        for col in df_clean.columns:
            # Convert to string, handle NaN values
            df_clean[col] = df_clean[col].astype(str)
            
            # Trim whitespace
            df_clean[col] = df_clean[col].str.strip()
            
            # Replace empty strings and 'nan' (from numpy.nan) with None
            df_clean[col] = df_clean[col].replace(['^\s*$', 'nan'], [None, None], regex=True)
            
            # Convert back to original dtype if possible
            if original_dtypes[col] != 'object':
                try:
                    df_clean[col] = df_clean[col].astype(original_dtypes[col])
                except (ValueError, TypeError):
                    # If conversion fails, keep as string
                    pass
        
        # Remove completely empty rows
        df_clean = df_clean.dropna(how='all')
        
        return df_clean
    except Exception as e:
        st.error(f"Error cleaning dataframe: {str(e)}")
        return df

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
                    
                    # Create tabs for different views
                    tab1, tab2, tab3 = st.tabs(["📋 Duplicates Preview", "🧹 Data Cleaning", "📊 Raw Data"])
                    
                    with tab1:
                        # Process duplicates data
                        unique_df = df[~df.index.isin(duplicates.index)]
                        
                        st.write(f"Found {len(duplicates)} duplicate rows and {len(unique_df)} unique rows")
                        
                        # Tabs for duplicate and unique data preview
                        subtab1, subtab2 = st.tabs(["📋 Duplicates", "✅ Unique"])
                        
                        with subtab1:
                            st.write(f"Showing first 5 of {len(duplicates)} duplicate rows:")
                            st.dataframe(duplicates.head(5), use_container_width=True)
                        
                        with subtab2:
                            st.write(f"Showing first 5 of {len(unique_df)} unique rows:")
                            st.dataframe(unique_df.head(5), use_container_width=True)
                        
                        # Duplicates summary
                        st.write("### Duplicates Summary")
                        
                        # Display metrics
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Rows", len(df))
                        with col2:
                            st.metric("Duplicate Rows", len(duplicates), 
                                     delta=f"{len(duplicates)/len(df):.1%} of total" if len(df) > 0 else "0%",
                                     delta_color="inverse")
                        with col3:
                            st.metric("Unique Rows", len(unique_df),
                                     delta=f"{len(unique_df)/len(df):.1%} of total" if len(df) > 0 else "0%")
                    
                        # Download button for duplicates
                        output_duplicates = io.BytesIO()
                        with pd.ExcelWriter(output_duplicates, engine='openpyxl') as writer:
                            duplicates.to_excel(writer, sheet_name='Duplicates', index=False)
                            unique_df.to_excel(writer, sheet_name='Unique', index=False)
                        
                        st.download_button(
                            label="📥 Download Duplicates",
                            data=output_duplicates.getvalue(),
                            file_name=f"duplicates_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download duplicates and unique entries in separate sheets"
                        )
                    
                    with tab2:
                        # Process cleaned data
                        cleaned_df = clean_dataframe(df)
                        
                        st.write("### Data Cleaning Preview")
                        st.write("First 10 rows of cleaned data:")
                        st.dataframe(cleaned_df.head(10), use_container_width=True)
                        
                        # Show cleaning summary
                        st.write("### Cleaning Summary")
                        
                        # Count of cleaned values
                        original_count = len(df)
                        cleaned_count = len(cleaned_df)
                        removed_count = original_count - cleaned_count
                        
                        # Count of trimmed whitespace
                        whitespace_count = 0
                        for col in df.select_dtypes(include=['object']).columns:
                            whitespace_count += (df[col].astype(str).str.strip() != df[col].astype(str)).sum()
                        
                        # Display metrics
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Rows Removed", f"{removed_count}", 
                                     help="Rows that were completely empty")
                        with col2:
                            st.metric("Whitespace Cleaned", f"{whitespace_count}",
                                     help="Cells with leading/trailing whitespace")
                        
                        # Download button for cleaned data
                        output_cleaned = io.BytesIO()
                        with pd.ExcelWriter(output_cleaned, engine='openpyxl') as writer:
                            cleaned_df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
                        
                        st.download_button(
                            label="🧹 Download Cleaned Data",
                            data=output_cleaned.getvalue(),
                            file_name=f"cleaned_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Download the cleaned version of your data with trimmed whitespace"
                        )
                    
                    with tab3:
                        # Display raw data with highlighting
                        st.write("### Raw Data")
                        # Add a checkbox to toggle blank cell highlighting
                        highlight_blanks_toggle = st.checkbox("Highlight blank/empty cells", value=True, key="highlight_blanks")
                        
                        # Apply styling based on the toggle
                        if highlight_blanks_toggle:
                            st.dataframe(
                                df.style.apply(highlight_blanks),
                                use_container_width=True, 
                                height=400
                            )
                        else:
                            st.dataframe(df, use_container_width=True, height=400)
                        
                        # Show blank cell summary
                        blank_cells = df.isna().sum().sum() + (df.astype(str).apply(lambda x: x.str.strip() == '')).sum().sum()
                        if blank_cells > 0:
                            st.warning(f"⚠️ Found {blank_cells} blank/empty cells in the dataset")
                        
                        # Data statistics
                        with st.expander("Data Statistics", expanded=False):
                            st.write("### Data Types")
                            st.write(df.dtypes)
                            
                            st.write("### Missing Values")
                            missing = df.isnull().sum().reset_index()
                            missing.columns = ['Column', 'Missing Values']
                            st.dataframe(missing, use_container_width=True)
                            
                            st.write("### Basic Statistics")
                            st.dataframe(df.describe(include='all').T, use_container_width=True)
                    
                else:
                    st.markdown("<div class='success-box'>No duplicates found based on the selected columns. ✅</div>", unsafe_allow_html=True)
            
            # Moved raw data and statistics to the main tabs
    
    # Add footer
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #666;'>🔍 Excel Duplicate Checker | Last updated: 2025-09-20</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()