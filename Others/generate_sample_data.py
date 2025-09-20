import pandas as pd
import numpy as np
import random
from faker import Faker
from datetime import datetime, timedelta

def generate_employee_data(num_records=200):
    fake = Faker()
    
    # Set random seed for reproducibility
    np.random.seed(42)
    random.seed(42)
    
    # Generate unique employee IDs with some duplicates
    base_ids = [f"AKS{random.randint(10000, 99999)}" for _ in range(int(num_records * 0.8))]
    employee_ids = base_ids + random.choices(base_ids, k=num_records - len(base_ids))
    random.shuffle(employee_ids)
    
    # Generate first names with some duplicates
    first_names = [fake.first_name() for _ in range(int(num_records * 0.7))]
    first_names = first_names + random.choices(first_names, k=num_records - len(first_names))
    
    # Generate last names with some duplicates
    last_names = [fake.last_name() for _ in range(int(num_records * 0.8))]
    last_names = last_names + random.choices(last_names, k=num_records - len(last_names))
    
    # Other data
    business_areas = ['Sales', 'Marketing', 'Engineering', 'HR', 'Finance', 'Operations', 'IT', 'Customer Support']
    roles = {
        'Sales': ['Sales Rep', 'Account Manager', 'Sales Director', 'Business Development'],
        'Marketing': ['Marketing Specialist', 'Content Writer', 'SEO Analyst', 'Marketing Manager'],
        'Engineering': ['Software Engineer', 'DevOps', 'QA Engineer', 'Engineering Manager'],
        'HR': ['HR Specialist', 'Recruiter', 'HR Manager', 'Training Coordinator'],
        'Finance': ['Accountant', 'Financial Analyst', 'Controller', 'CFO'],
        'Operations': ['Operations Manager', 'Logistics Coordinator', 'Facilities Manager'],
        'IT': ['System Admin', 'Network Engineer', 'Help Desk', 'IT Manager'],
        'Customer Support': ['Support Agent', 'Team Lead', 'Support Manager']
    }
    
    insurance_providers = ['Aetna', 'Blue Cross', 'Cigna', 'Kaiser', 'UnitedHealth', 'Humana']
    team_names = [chr(65 + i) + (chr(65 + j) if i < 26 else '') for i in range(26) for j in range(26)][:52]
    
    data = []
    for i in range(num_records):
        # Randomly decide if this record will have missing values (5% chance)
        has_missing = random.random() < 0.05
        
        # Generate employee data
        emp_id = employee_ids[i]
        first_name = first_names[i] if not (has_missing and random.random() < 0.3) else np.nan
        last_name = last_names[i] if not (has_missing and random.random() < 0.3) else np.nan
        
        # Random business area and role
        area = random.choice(business_areas)
        role = random.choice(roles[area])
        
        # Generate years in business (0-20 years)
        years_in_business = random.randint(0, 20)
        
        # Random insurance provider (some will be the same)
        insurance = random.choice(insurance_providers)
        
        # Random floor (1-10)
        floor = random.randint(1, 10)
        
        # Random office days per week (1-5)
        office_days = random.randint(1, 5)
        
        # Random team
        team = random.choice(team_names)
        
        data.append([
            emp_id,          # Employee ID
            first_name,      # First Name
            fake.first_name() if random.random() < 0.3 else np.nan,  # Middle Name (30% chance)
            last_name,       # Last Name
            area,            # Business Area
            role,            # Current Role
            years_in_business,  # Years in Business
            insurance,       # Insurance Provider
            floor,           # Floor/Level
            office_days,     # Office Days per Week
            team             # Team Name
        ])
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=[
        'Employee_ID', 'First_Name', 'Middle_Name', 'Last_Name', 'Business_Area',
        'Current_Role', 'Years_in_Business', 'Insurance_Provider', 'Floor',
        'Office_Days_Per_Week', 'Team_Name'
    ])
    
    return df

def generate_sample_excel(filename='employee_data.xlsx', num_sheets=3, records_per_sheet=200):
    """Generate an Excel file with sample employee data."""
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for i in range(num_sheets):
            df = generate_employee_data(records_per_sheet)
            sheet_name = f'Employees_{i+1}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Generated Excel file '{filename}' with {num_sheets} sheets, {records_per_sheet} records each.")

if __name__ == "__main__":
    generate_sample_excel('sample_employee_data.xlsx', num_sheets=3, records_per_sheet=200)
