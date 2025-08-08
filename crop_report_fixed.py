import pandas as pd
import os
from datetime import datetime

def read_excel_data(excel_path):
    """Read data from Excel file"""
    try:
        # Read the Excel file
        df = pd.read_excel(excel_path)
        print("Excel data structure:")
        print(df.head())
        print("\nColumn names:")
        print(df.columns.tolist())
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_report_html(data, template_path, output_path):
    """Generate report HTML from template and data"""
    
    # Read the template
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    # Default report content
    report_content = template_content
    
    # If we have data, use the row with more information
    if data is not None and len(data) > 1:  # Try to use the second row first
        row = data.iloc[1]  # Use second row of data (index 1)
    elif data is not None and len(data) > 0:  # Fall back to first row if only one row
        row = data.iloc[0]
    else:
        # If no data, use default values
        report_content = template_content.replace(' field images ', '../assest/farmland.png')
        report_content = report_content.replace('src="/assest/', 'src="../assest/')
        report_content = report_content.replace('Current date', datetime.now().strftime('%Y-%m-%d'))
        
        # Write the generated report with defaults
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report_content)
        
        print(f"Report generated with default values: {output_path}")
        return report_content
    
    # Extract actual values from Excel data
    field_name = str(row['Field']) if 'Field' in row else 'Sample Field'
    
    # Get crop name
    crop_name = '-'
    if 'Crop' in row and pd.notna(row['Crop']) and str(row['Crop']) != '-' and str(row['Crop']) != 'nan':
        crop_name = str(row['Crop'])
    
    # Get sowing date
    sowing_date = '2025-08-01'  # Default
    if 'Sowing/planting' in row and pd.notna(row['Sowing/planting']):
        sowing_date = str(row['Sowing/planting']).split(' ')[0]
    elif 'Sowing / Planting' in row and pd.notna(row['Sowing / Planting']):
        sowing_date = str(row['Sowing / Planting']).split(' ')[0]
        if sowing_date == '-' or sowing_date == 'nan':
            sowing_date = '-'
    
    # Get area coverage
    area_coverage = '10.5 acres'  # Default
    if 'area' in row and pd.notna(row['area']):
        area_coverage = str(row['area'])
    elif 'Area' in row and pd.notna(row['Area']):
        area_coverage = str(row['Area'])
    
    # Get growth stage
    growth_stage = 'Mature'  # Default
    if 'maturity' in row and pd.notna(row['maturity']):
        growth_stage = str(row['maturity'])
    elif 'Maturity' in row and pd.notna(row['Maturity']):
        growth_stage = str(row['Maturity'])
        if growth_stage == '-' or growth_stage == 'nan':
            growth_stage = 'Not specified'
    
    # Get current date
    current_date = datetime.now().strftime('%Y-%m-%d')  # Default to today
    if 'current data' in row and pd.notna(row['current data']):
        current_date = str(row['current data']).split(' ')[0]
    elif 'Current Image date' in row and pd.notna(row['Current Image date']):
        current_date = str(row['Current Image date']).split(' ')[0]
        if current_date == 'nan' or current_date == 'NaT':
            current_date = datetime.now().strftime('%Y-%m-%d')
    
    # Get farmland image path
    farmland_image = '../assest/farmland.png'  # Default image path

    # Replace specific placeholders in template
    report_content = template_content
    
    # Replace field information
    report_content = report_content.replace('Field Information', f'{field_name} Information')
    
    # Replace field data with improved spacing
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Field Name:\n      </span>\n      <span class="font-extrabold">\n       Field\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Field Name: </span>\n      <span class="font-extrabold">{field_name}</span>\n     </div>'
    )
    
    # Replace crop name
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Crop Name:\n      </span>\n      <span class="font-extrabold">\n       crop\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Crop Name: </span>\n      <span class="font-extrabold">{crop_name}</span>\n     </div>'
    )
    
    # Replace sowing date
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Sowing Date:\n      </span>\n      <span class="font-extrabold">\n       Sowing/planting\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Sowing Date: </span>\n      <span class="font-extrabold">{sowing_date}</span>\n     </div>'
    )
    
    # Replace report date
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Report Date:\n      </span>\n      <span class="font-extrabold">\n       Current date\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Report Date: </span>\n      <span class="font-extrabold">{current_date}</span>\n     </div>'
    )
    
    # Replace area coverage
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Area Coverage (in Acre):\n      </span>\n      <span class="font-extrabold">\n       Area\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Area Coverage: </span>\n      <span class="font-extrabold">{area_coverage}</span>\n     </div>'
    )
    
    # Replace growth stage
    report_content = report_content.replace(
        '<div class="text-gray-600 font-semibold flex justify-between">\n      <span>\n       Growth Stage:\n      </span>\n      <span class="font-extrabold">\n       Maturity\n      </span>\n     </div>',
        f'<div class="text-gray-600 font-semibold">\n      <span>Growth Stage: </span>\n      <span class="font-extrabold">{growth_stage}</span>\n     </div>'
    )
    
    # Add PDF download functionality to the template
    pdf_script = '''
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>'''
    
    # Insert PDF script into head
    report_content = report_content.replace(
        '<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" rel="stylesheet"/>',
        '<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" rel="stylesheet"/>' + pdf_script
    )
    
    # Add download button and wrap content with improved container
    download_button = '''  <!-- PDF Download Button -->
  <div class="fixed top-4 right-4 z-50">
   <button id="downloadPdf" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-lg shadow-lg flex items-center gap-2 transition-colors">
    <i class="fas fa-download"></i>
    Download PDF
   </button>
  </div>
  
  <div id="reportContent" class="max-w-7xl mx-auto p-6 bg-[#dbe8f2] rounded-3xl relative" style="min-height: 800px; height: auto;">'''
    
    report_content = report_content.replace(
        '<div class="max-w-7xl mx-auto p-6 bg-[#dbe8f2] rounded-3xl relative min-h-[600px]">',
        download_button
    )
    
    # Update container structure for better PDF capture
    report_content = report_content.replace(
        '<img alt="SiRDA_Logo" class="absolute top-6 left-6 rounded-lg w-[300px] h-[60px] object-cover"',
        '<img alt="SiRDA_Logo" class="absolute top-6 left-6 w-[300px] h-[60px] object-contain"'
    )
    
    report_content = report_content.replace(
        '<img alt="KSRCT_Logo" class="absolute top-6 right-6 rounded-lg w-[60px] h-[60px] object-cover"',
        '<img alt="KSRCT_Logo" class="absolute top-6 right-6 w-[60px] h-[60px] object-contain"'
    )
    
    # Add improved spacing and bottom padding
    report_content = report_content.replace(
        '<h1 class="text-[72px] font-sans font-normal text-[#1f2937] text-center mt-20">',
        '<h1 class="text-[72px] font-sans font-normal text-[#1f2937] text-center mt-20 mb-8">'
    )
    
    report_content = report_content.replace(
        '<div class="max-w-[600px] mx-auto mt-8 rounded-3xl overflow-hidden">',
        '<div class="max-w-[600px] mx-auto mt-8 mb-6 rounded-3xl overflow-hidden">'
    )
    
    report_content = report_content.replace(
        '<div class="flex justify-center items-center mt-4 text-black font-extrabold text-lg">',
        '<div class="flex justify-center items-center mt-4 mb-6 text-black font-extrabold text-lg">'
    )
    
    # Set additional info to empty string - removed detailed sections as requested
    additional_info = ''
    
    # No NDVI Information section
    # No Field Management section
    # No Field Analysis section with risks and recommendations
    
    # Fix the farmland image path
    report_content = report_content.replace(' field images ', farmland_image)
    
    # Fix the asset paths to be relative
    report_content = report_content.replace('src="/assest/', 'src="../assest/')
    
    # Add closing div and PDF script with better capture settings
    pdf_footer_script = f'''  {additional_info}
  </div>
  <!-- Extra padding to ensure all content is captured in PDF -->
  <div class="pb-10"></div>

  <script>
    document.getElementById('downloadPdf').addEventListener('click', function() {{
      console.log('PDF download button clicked');
      
      // Hide the download button during PDF generation
      this.style.display = 'none';
      
      const element = document.getElementById('reportContent');
      console.log('Report element found:', element);
      
      if (!element) {{
        alert('Report content not found!');
        this.style.display = 'flex';
        return;
      }}
      
      // Store original styles to restore later
      const originalHeight = element.style.height;
      const originalMinHeight = element.style.minHeight;
      const originalOverflow = element.style.overflow;
      
      // Set styles for PDF generation
      element.style.height = 'auto';
      element.style.minHeight = 'auto';
      element.style.overflow = 'visible';
      
      // Wait for any dynamic content to load
      setTimeout(() => {{
        console.log('Starting PDF generation...');
        
        // Optimized options for better PDF quality and reliability
        const opt = {{
          margin: 5,
          filename: 'crop_report_{field_name.replace(" ", "_")}.pdf',
          image: {{ 
            type: 'jpeg', 
            quality: 0.98
          }},
          html2canvas: {{ 
            scale: 2,
            useCORS: true,
            allowTaint: true,
            letterRendering: true,
            backgroundColor: '#dbe8f2'
          }},
          jsPDF: {{ 
            unit: 'mm', 
            format: 'a4', 
            orientation: 'portrait'
          }}
        }};
        
        // Generate PDF directly from the element
        html2pdf()
          .from(element)
          .set(opt)
          .save()
          .then(() => {{
            console.log('PDF generation successful');
            
            // Show the download button again
            document.getElementById('downloadPdf').style.display = 'flex';
            
            // Reset styles to original values
            element.style.height = originalHeight;
            element.style.minHeight = originalMinHeight || '800px';
            element.style.overflow = originalOverflow;
          }})
          .catch((error) => {{
            console.error('PDF generation failed:', error);
            
            // Show the download button again
            document.getElementById('downloadPdf').style.display = 'flex';
            
            // Reset styles to original values
            element.style.height = originalHeight;
            element.style.minHeight = originalMinHeight || '800px';
            element.style.overflow = originalOverflow;
            
            alert('PDF generation failed: ' + error.message);
          }});
      }}, 500);
    }});
  </script>'''
    
    report_content = report_content.replace('</body>', pdf_footer_script + '\n </body>')
    
    print(f"Data used in report:")
    print(f"Field Name: {field_name}")
    print(f"Crop Name: {crop_name}")
    print(f"Sowing Date: {sowing_date}")
    print(f"Report Date: {current_date}")
    print(f"Area Coverage: {area_coverage}")
    print(f"Growth Stage: {growth_stage}")
    
    # Write the generated report
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(report_content)
    
    print(f"Report generated successfully: {output_path}")
    return report_content

def generate_reports_for_all_rows(excel_path, template_path, output_dir="reports"):
    """Generate individual reports for each row in the Excel file"""
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    # Read Excel data
    data = read_excel_data(excel_path)
    
    if data is not None and len(data) > 0:
        print(f"Generating reports for {len(data)} rows of data...")
        
        # Process each row and generate individual reports
        for index, row in data.iterrows():
            # Create field-specific filename
            field_name = str(row['Field']).replace(' ', '_').replace('/', '_') if 'Field' in row else f"field_{index+1}"
            output_path = os.path.join(output_dir, f"crop_report_{field_name}.html")
            
            # Create single row dataframe with this row
            single_row_data = pd.DataFrame([row])
            
            # Generate report for this row
            generate_report_html(single_row_data, template_path, output_path)
            
        print(f"All reports generated successfully in '{output_dir}' directory.")
    else:
        print("No data found in the Excel file.")

if __name__ == "__main__":
    # Paths
    excel_path = "test.xlsx"  # Using the test Excel file
    template_path = "templete/page1.html"
    
    # Generate individual reports for each row
    generate_reports_for_all_rows(excel_path, template_path)
    
    # Also generate a single report with the second row for backward compatibility
    output_path = "crop_report.html"
    data = read_excel_data(excel_path)
    if data is not None and len(data) > 0:
        generate_report_html(data, template_path, output_path)
