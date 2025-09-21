import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import datetime, timedelta
import os
import streamlit as st
from io import BytesIO

# Initialize Faker
fake = Faker()

def set_seed(seed):
    """Set random seed for reproducibility"""
    random.seed(seed)
    np.random.seed(seed)
    return seed

# Default constants
DEFAULT_DEPARTMENTS = ['HR', 'Finance', 'IT', 'Marketing', 'Operations', 'Sales', 'R&D', 'Customer Support']
DEFAULT_ROLES = {
    'HR': ['HR Manager', 'Recruiter', 'HR Generalist', 'Training Specialist', 'Compensation Analyst'],
    'Finance': ['CFO', 'Financial Analyst', 'Accountant', 'Auditor', 'Financial Planner'],
    'IT': ['CTO', 'Software Engineer', 'Data Scientist', 'DevOps Engineer', 'IT Support'],
    'Marketing': ['CMO', 'Marketing Specialist', 'Content Writer', 'SEO Specialist', 'Social Media Manager'],
    'Operations': ['COO', 'Operations Manager', 'Project Manager', 'Logistics Coordinator'],
    'Sales': ['Sales Director', 'Account Executive', 'Sales Representative', 'Business Development'],
    'R&D': ['Research Scientist', 'Product Manager', 'UX Designer', 'Research Assistant'],
    'Customer Support': ['Support Manager', 'Customer Service Rep', 'Technical Support']
}
DEFAULT_INSURANCE_PROVIDERS = ['Aetna', 'Blue Cross', 'Cigna', 'UnitedHealthcare', 'Kaiser', 'Humana']
DEFAULT_TEAM_NAMES = [f'Team {chr(65 + i)}' for i in range(26)]  # Team A to Team Z

# Generate random employee IDs with some duplicates
def generate_employee_ids(n, duplicate_percentage=5):
    """Generate employee IDs with a specified percentage of duplicates"""
    num_duplicates = int(n * (duplicate_percentage / 100))
    num_unique = n - num_duplicates
    
    # Generate unique IDs first
    ids = [f'EMP{fake.unique.random_number(digits=6, fix_len=True)}' for _ in range(num_unique)]
    
    # Add duplicates if needed
    if num_duplicates > 0 and ids:  # Make sure we have some IDs to duplicate
        duplicate_ids = random.choices(ids, k=num_duplicates)
        ids.extend(duplicate_ids)
    
    # Shuffle to distribute duplicates
    random.shuffle(ids)
    return ids

def generate_employee_data(num_employees, departments, roles, insurance_providers, team_names, 
                         duplicate_percentage=5, missing_data_percentage=5, **kwargs):
    """Generate employee data with the specified parameters"""
    # Generate employee IDs with some duplicates
    employee_ids = generate_employee_ids(num_employees, duplicate_percentage)
    
    # Generate data with the specified parameters
    data = {
        'Employee_ID': employee_ids,
        'First_Name': [fake.first_name() for _ in range(num_employees)],
        'Middle_Name': [fake.first_name() if random.random() < 0.7 else np.nan 
                       for _ in range(num_employees)],
        'Last_Name': [fake.last_name() for _ in range(num_employees)],
        'Department': [random.choice(departments) for _ in range(num_employees)],
        'Role': [],  # Will be filled based on department
        'Years_In_Company': np.round(np.random.gamma(shape=2, scale=2, size=num_employees), 1),
        'Insurance_Provider': random.choices(insurance_providers, 
                                          k=num_employees),
        'Work_Floor': np.random.randint(1, 11, size=num_employees),
        'Days_In_Office': np.random.choice([1, 2, 3, 4, 5], 
                                         size=num_employees, 
                                         p=[0.1, 0.2, 0.3, 0.25, 0.15]),
        'Team': [random.choice(team_names) for _ in range(num_employees)],
        'Hire_Date': [fake.date_between(start_date='-10y', end_date='today') 
                      for _ in range(num_employees)],
        'Salary': np.round(np.random.lognormal(mean=11, sigma=0.4, size=num_employees), 2)
    }
    
    # Set roles based on department
    data['Role'] = [random.choice(roles.get(dept, ['Not Specified'])) for dept in data['Department']]
    
    # Create some missing values
    if missing_data_percentage > 0:
        for col in data:
            if col not in ['Employee_ID', 'Hire_Date']:  # Don't add missing values to these columns
                mask = np.random.random(size=num_employees) < (missing_data_percentage / 100)
                data[col] = [np.nan if mask[i] else val for i, val in enumerate(data[col])]
    
    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Ensure Employee_ID is the first column
    cols = ['Employee_ID'] + [col for col in df.columns if col != 'Employee_ID']
    df = df[cols]
    
    return df

def create_excel_with_multiple_sheets(num_employees, sheet_names, departments, roles, insurance_providers, team_names,
                                    duplicate_percentage=5, missing_data_percentage=5, output_format='xlsx'):
    """Create an Excel file with multiple sheets of employee data"""
    # Create a BytesIO object to store the file in memory
    output = BytesIO()
    
    # Set the file extension and engine based on the output format
    if output_format == 'xlsx':
        engine = 'xlsxwriter'
    else:
        engine = 'openpyxl'
    
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet_name in sheet_names:
            # Generate data for this sheet
            df = generate_employee_data(
                num_employees=num_employees,
                departments=departments,
                roles=roles,
                insurance_providers=insurance_providers,
                team_names=team_names,
                duplicate_percentage=duplicate_percentage,
                missing_data_percentage=missing_data_percentage
            )
            
            # Write the data to the Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Only apply formatting for xlsx files (xlsxwriter specific)
            if output_format == 'xlsx':
                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # Add a header format
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#4F81BD',
                    'font_color': 'white',
                    'border': 1
                })
                
                # Format the header row
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Set column widths
                for i, col in enumerate(df.columns):
                    # Find the maximum length of the column
                    max_length = max(df[col].astype(str).apply(len).max(), len(str(col))) + 2
                    # Set the column width
                    worksheet.set_column(i, i, min(max_length, 25))
                
                # Add conditional formatting for duplicates in Employee_ID
                if duplicate_percentage > 0:
                    duplicate_format = workbook.add_format({
                        'bg_color': '#FFC7CE',
                        'font_color': '#9C0006'
                    })
                    
                    # Apply conditional formatting to highlight duplicates in Employee_ID
                    worksheet.conditional_format(1, 0, len(df), 0, {
                        'type': 'duplicate',
                        'format': duplicate_format
                    })
                
                # Add a table with autofilter
                worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {
                    'columns': [{'header': col} for col in df.columns],
                    'style': 'Table Style Medium 2',
                    'autofilter': True
                })
    
    # Get the Excel file data
    output.seek(0)
    return output

def main():
    """Main function to run the Streamlit app"""
    # Set page config
    st.set_page_config(
        page_title="Employee Data Generator",
        page_icon="📊",
        layout="wide"
    )
    
    # Custom CSS
    st.markdown("""
    <style>
    .main-header {font-size: 2rem; color: #1f77b4; text-align: center; margin-bottom: 1rem;}
    .sub-header {font-size: 1.25rem; color: #2c3e50; margin: 1.5rem 0 1rem 0;}
    .stButton>button {width: 100%; margin: 0.5rem 0;}
    .stDownloadButton>button {width: 100%; margin: 0.5rem 0;}
    </style>
    """, unsafe_allow_html=True)
    
    # Title and description
    st.markdown("<h1 class='main-header'>📊 Employee Data Generator</h1>", unsafe_allow_html=True)
    st.markdown("""
    Generate realistic employee data with customizable parameters. 
    Perfect for testing, demos, and development purposes.
    """)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Basic settings
        st.subheader("Basic Settings")
        num_employees = st.slider("Number of employees per sheet", 10, 1000, 200)
        num_sheets = st.slider("Number of sheets", 1, 12, 3)
        
        # Data quality settings
        st.subheader("Data Quality")
        duplicate_percentage = st.slider("Percentage of duplicate Employee IDs", 0, 20, 5)
        missing_data_percentage = st.slider("Percentage of missing data", 0, 20, 5)
        
        # Output settings
        st.subheader("Output Settings")
        output_format = st.selectbox("Output format", ["xlsx", "xls"])
        
        # Custom data settings
        st.subheader("Custom Data")
        custom_departments = st.text_area("Departments (one per line)", "\n".join(DEFAULT_DEPARTMENTS))
        custom_teams = st.text_area("Team names (one per line)", "\n".join(DEFAULT_TEAM_NAMES))
        custom_insurance = st.text_area("Insurance providers (comma separated)", 
                                      ", ".join(DEFAULT_INSURANCE_PROVIDERS))
    
    # Process custom data
    departments = [d.strip() for d in custom_departments.split('\n') if d.strip()]
    team_names = [t.strip() for t in custom_teams.split('\n') if t.strip()]
    insurance_providers = [i.strip() for i in custom_insurance.split(',') if i.strip()]
    
    # Use default roles for now (could be made customizable)
    roles = DEFAULT_ROLES
    
    # Generate sheet names based on quarters if multiple sheets
    if num_sheets == 1:
        sheet_names = ["Employee_Data"]
    elif num_sheets == 4:
        sheet_names = [f"Q{i+1}_2024" for i in range(4)]
    elif num_sheets == 12:
        sheet_names = [f"{month}_2024" for month in [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]]
    else:
        sheet_names = [f"Sheet_{i+1}" for i in range(num_sheets)]
    
    # Generate data button
    if st.button("🚀 Generate Employee Data"):
        with st.spinner("Generating data... This may take a moment..."):
            try:
                # Set seed for reproducibility
                set_seed(42)
                
                # Generate the Excel file
                excel_file = create_excel_with_multiple_sheets(
                    num_employees=num_employees,
                    sheet_names=sheet_names,
                    departments=departments,
                    roles=roles,
                    insurance_providers=insurance_providers,
                    team_names=team_names,
                    duplicate_percentage=duplicate_percentage,
                    missing_data_percentage=missing_data_percentage,
                    output_format=output_format
                )
                
                # Create a download button
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Employee_Data_{timestamp}.{output_format}"
                
                st.success("✅ Data generated successfully!")
                
                # Show a preview of the first sheet
                st.subheader("📊 Data Preview (First 5 Rows)")
                preview_df = pd.read_excel(excel_file, sheet_name=sheet_names[0])
                st.dataframe(preview_df.head())
                
                # Show statistics
                st.subheader("📈 Data Statistics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Employees", f"{num_employees:,}")
                with col2:
                    st.metric("Number of Sheets", len(sheet_names))
                with col3:
                    st.metric("Duplicate IDs", f"{int(num_employees * (duplicate_percentage / 100)):,}")
                
                # Download button
                st.download_button(
                    label="💾 Download Excel File",
                    data=excel_file,
                    file_name=filename,
                    mime=f"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
    
    # Add some documentation
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        ### How to use this tool:
        1. **Configure** the data settings in the sidebar
        2. Click the **Generate Employee Data** button
        3. **Download** the generated Excel file
        
        ### Features:
        - **Customizable Data**: Adjust departments, teams, and insurance providers
        - **Control Data Quality**: Set the percentage of duplicates and missing data
        - **Multiple Sheets**: Generate data across multiple sheets (e.g., by quarter or month)
        - **Realistic Data**: Names, roles, and other fields are generated realistically
        
        ### Tips:
        - For large datasets, be patient as generation may take some time
        - The Excel file includes formatting and filtering for easy exploration
        - You can customize the data by editing the text areas in the sidebar
        """)

if __name__ == "__main__":
    main()
