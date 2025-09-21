# Excel Sample Data Generator

A Streamlit-based web application that generates realistic employee data and exports it to Excel, complete with AI-powered insights and professional PDF reporting.

## Features

- **Customizable Data Generation**:
  - Generate employee records with realistic names, emails, and job details
  - Customize the number of records (default: 100)
  - Control data distribution across departments and positions

- **AI-Powered Analysis**:
  - Automatic generation of data insights using Ollama's AI
  - Comprehensive analysis of workforce distribution and compensation
  - Data quality assessment and HR recommendations

- **Professional Reporting**:
  - Export data to Excel with formatted worksheets
  - Generate detailed PDF reports with visualizations
  - Clean, professional document formatting

## Prerequisites

- Python 3.8+
- Required Python packages (install via `pip install -r requirements.txt`):
  - streamlit
  - pandas
  - numpy
  - faker
  - openpyxl
  - reportlab
  - python-dotenv
  - ollama

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo/Excel\ Sample\ Data\ Generator
   ```

2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up Ollama (for AI insights):
   - Install Ollama from [ollama.ai](https://ollama.ai/)
   - Pull the required model: `ollama pull gemma3:4b`

## Usage

1. Run the Streamlit app:
   ```bash
   streamlit run generate_employee_data.py
   ```

2. In the web interface:
   - Adjust the number of employees using the slider
  - Click "Generate Employee Data" to create the dataset
  - Use the "Export to Excel" button to download the data
  - Generate a PDF report with the "Generate PDF Report" button

## Performance Considerations

### What Can Make It Slow

1. **Large Datasets**:
   - Generating more than 10,000 records may cause performance issues
   - PDF generation is particularly resource-intensive with large datasets

2. **AI Insights**:
   - The AI analysis can take 30-60 seconds depending on your hardware
   - Internet connection quality affects response time from the Ollama server

3. **System Resources**:
   - Running out of memory with very large datasets
   - Multiple concurrent users may slow down the application

### Potential Crash Scenarios

1. **Missing Dependencies**:
   - Ensure all required Python packages are installed
   - Verify Ollama is running if using AI features

2. **File Permissions**:
   - The app needs write permissions in the current directory
   - Check for sufficient disk space before generating large files

3. **Memory Issues**:
   - Generating very large Excel files (>100MB) may cause memory errors
   - PDF generation with many pages can be memory-intensive

## Best Practices

1. **For Large Datasets**:
   - Generate data in smaller batches
   - Consider using the command-line version for headless operation
   - Close other memory-intensive applications

2. **For Better Performance**:
   - Use the latest version of all dependencies
   - Run the application on a machine with sufficient RAM
   - Consider using a virtual environment

3. **Troubleshooting**:
   - Check the console for error messages
   - Verify that Ollama is running if AI features aren't working
   - Clear the browser cache if the UI becomes unresponsive

## About This Project

> "I wanted to make testing easier, so I created a script that generates realistic test subject data with random information. This tool was born out of the need for quick, realistic datasets for development and testing purposes."

