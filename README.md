# Excel Duplicate Checker

A Streamlit-based application to detect, analyze, and manage duplicate entries in Excel files, specifically designed for employee/worker data. The application provides an intuitive interface for data exploration, cleaning, and export.

## ✨ Key Features

### 🔍 Duplicate Detection
- Multi-column duplicate checking
- Visual highlighting of duplicate entries
- Summary statistics of duplicates found
- Export duplicates and unique records separately

### 🧹 Data Cleaning
- Automatic whitespace trimming
- Empty cell detection and highlighting
- Removal of completely empty rows
- Preview changes before downloading

### 📊 Data Exploration
- Interactive data previews
- Column statistics and data types
- Missing value analysis
- Summary metrics and visualizations

## 🚀 Getting Started

### Prerequisites
- Python 3.8+
- pip (Python package manager)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/excel-duplicate-checker.git
   cd excel-duplicate-checker
   ```

2. **Create a virtual environment (recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Generate sample data (optional)**
   ```bash
   python generate_sample_data.py
   ```
   This creates `sample_employee_data.xlsx` with realistic test data.

5. **Run the application**
   ```bash
   streamlit run checker.py
   ```
   The application will open in your default web browser at `http://localhost:8501`

## 🖥️ User Guide

### 1. Uploading Files
- Click "Browse files" or drag-and-drop your Excel file
- Supported formats: `.xlsx`, `.xls`
- The application automatically loads the first sheet

### 2. Analyzing Data
- **Select Columns**: Choose which columns to check for duplicates
- **View Results**:
  - Duplicates are highlighted in the preview
  - Summary statistics show the number of duplicates found
  - Data quality metrics are displayed

### 3. Working with Results
- **Preview Tabs**:
  - 📋 **Duplicates Preview**: View and analyze duplicate entries
  - 🧹 **Data Cleaning**: Preview and download cleaned data
  - 📊 **Raw Data**: Explore the original dataset with highlighting options

- **Export Options**:
  - Download duplicates as Excel with separate sheets
  - Export cleaned data with whitespace trimmed and empty rows removed

## 📊 Data Format

The application works with any Excel file, but is optimized for employee data with these common columns:

| Column Name        | Type      | Description                          |
|--------------------|-----------|--------------------------------------|
| Employee_ID        | Text      | Unique employee identifier           |
| First_Name         | Text      | Employee's first name                |
| Middle_Name        | Text      | Employee's middle name (optional)    |
| Last_Name          | Text      | Employee's last name                 |
| Business_Area      | Text      | Department or business unit          |
| Current_Role       | Text      | Job title or position                |
| Years_in_Business  | Number    | Years of service                     |
| Insurance_Provider | Text      | Health insurance provider            |
| Floor              | Number    | Office floor/level (1-10)            |
| Office_Days_Per_Week| Number   | Days working in office (1-5)         |
| Team_Name          | Text      | Team identifier (A-AZ)               |

## 🛠️ Troubleshooting

### Common Issues and Solutions

#### File Loading Problems
- **Symptom**: File fails to load or shows an error message
  - ✅ **Solution**: 
    - Ensure the file is not open in another program
    - Verify the file is not corrupted
    - Check that the file extension is correct (`.xlsx` or `.xls`)
    - Try saving the file in Excel as `.xlsx` format

#### Performance Issues with Large Files
- **Symptom**: Application is slow or becomes unresponsive
  - ✅ **Solution**:
    - For files with >10,000 rows:
      - Close other applications to free up memory
      - Analyze one sheet at a time
      - Consider splitting large files into smaller chunks
      - Use the "View Raw Data" tab with pagination

#### No Duplicates Found
- **Symptom**: Application reports no duplicates when you expect some
  - ✅ **Solution**:
    - Check for leading/trailing spaces in your data
    - Try selecting different columns for duplicate checking
    - Verify that the selected columns contain the expected data
    - Use the "Data Cleaning" tab to standardize your data first

#### Data Format Issues
- **Symptom**: Numbers or dates not being recognized correctly
  - ✅ **Solution**:
    - Ensure consistent data types in each column
    - Format cells correctly in Excel before importing
    - Check for hidden characters or non-printable characters

#### Blank/Empty Cell Detection
- **Symptom**: Blank cells not being detected as expected
  - ✅ **Solution**:
    - Use the "Highlight blank/empty cells" feature in the Raw Data tab
    - Check for cells with only spaces or non-breaking spaces
    - Look for cells with formulas that return empty strings

#### Export Problems
- **Symptom**: Downloaded file is corrupted or empty
  - ✅ **Solution**:
    - Ensure you have write permissions in the download directory
    - Check if your antivirus is blocking the download
    - Try downloading with a different web browser

### Debugging Tips

1. **Check the Browser Console**
   - Press `F12` to open developer tools
   - Look for any error messages in the Console tab

2. **Verify Data Types**
   - Use the "Data Statistics" tab to check column data types
   - Look for mixed data types in the same column

3. **Test with Sample Data**
   - Generate and test with the sample data to verify if the issue is with your file
   ```bash
   python generate_sample_data.py
   ```

4. **Check Application Logs**
   - Look for error messages in the terminal where Streamlit is running
   - Check for Python tracebacks that might indicate the source of the problem

## 📅 Changelog

### [1.1.0] - 2025-09-20
#### Added
- Tab-based interface for better navigation
- Enhanced data cleaning capabilities
- Improved duplicate detection algorithm
- More detailed statistics and metrics
- Better error handling and user feedback

### [1.0.0] - 2025-09-19
- Initial release with basic functionality

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/)
- Data manipulation with [Pandas](https://pandas.pydata.org/)
- Sample data generation with [Faker](https://faker.readthedocs.io/)
- Icons from [EmojiOne](https://www.joypixels.com/)

---

<div align="center">
  Made with ❤️ by Your Maximilians_XXI
</div>

