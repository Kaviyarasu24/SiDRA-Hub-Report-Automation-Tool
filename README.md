# SiDRA-Hub-Report-Automation-Tool

## Overview
SiDRA-Hub-Report-Automation-Tool automates the generation of comprehensive farm field reports from Excel data. It extracts structured data, images, and vegetation indices to create visual HTML reports for precision agriculture monitoring and decision-making.

## Features
- **Comprehensive Report Generation**: Creates complete reports with field information, images, and analysis
- **Multiple Vegetation Indices**: Supports NDVI, NDMI, RECI, MSAVI, and NDRE indices
- **Image Extraction**: Automatically extracts and processes embedded images from Excel
- **PDF Export**: Built-in functionality for exporting reports to PDF
- **Multi-field Support**: Generates individual reports for each field in the Excel data
- **Error Handling**: Robust fallbacks for missing or incorrect data

## Vegetation Indices
The tool supports the following vegetation indices:
1. **NDVI (Green Health Score)** - Page 2
2. **NDMI (Moisture Level Indicator)** - Page 3
3. **RECI (Leaf Freshness Index)** - Page 4
4. **MSAVI (Growth Strength Index)** - Page 5
5. **NDRE (Early Stress Checker)** - Page 6

## Usage

### Generating Complete Reports
To generate complete reports for all fields in the Excel file:

```python
python generate_report.py
```

This will:
1. Process all fields in `demo.xlsx`
2. Create individual page reports
3. Combine them into full reports in the `reports` directory
4. Name them as `full_report_<Field_Name>.html`

### Generating Individual Page Reports
You can also generate reports for specific pages:

```python
# Generate only field information (Page 1)
python page1.py

# Generate NDVI report (Page 2)
python page2.py

# Generate NDMI report (Page 3)
python page3.py

# Generate RECI report (Page 4)
python page4.py

# Generate MSAVI report (Page 5)
python page5.py

# Generate NDRE report (Page 6)
python page6.py
```

## Input Data Format
The tool expects an Excel file (`demo.xlsx`) with the following structure:
- Field information (name, crop type, area, etc.)
- Embedded images for each vegetation index
- Date information for current and previous images
- Index values and advisory text

## Output
Reports are generated in HTML format with:
- Clean, responsive layout using Tailwind CSS
- Field information and metadata
- Current and historical index images
- Value comparisons and analysis
- Download to PDF functionality

## Directory Structure
- `templete/`: HTML templates for each page
- `assest/`: Static assets like logos and icons
- `images/`: Extracted images from Excel
- `reports/`: Generated HTML reports

## Requirements
- Python 3.x
- pandas
- openpyxl
- Pillow (PIL)
- web browser with JavaScript enabled for viewing reports
