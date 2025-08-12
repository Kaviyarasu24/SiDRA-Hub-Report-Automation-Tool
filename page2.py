import pandas as pd
import os
from pathlib import Path
import openpyxl
from PIL import Image
import io

def extract_images_from_excel(excel_file):
    # Create images directory if it doesn't exist
    if not os.path.exists('images'):
        os.makedirs('images')
    
    # Load the workbook with data_only=False to get images
    wb = openpyxl.load_workbook(excel_file, data_only=False)
    sheet = wb.active
    
    current_image_path = None
    old_image_path = None
    
    # Find column indices
    ndvi_col = None
    old_ndvi_col = None
    for col in range(1, sheet.max_column + 1):
        header = sheet.cell(row=1, column=col).value
        if header == 'NDVI Image date':
            ndvi_col = col
        elif header == 'Old NDVI Image date':
            old_ndvi_col = col
    
    try:
        # Get the drawing objects from the sheet
        for image in sheet._images:
            # Get the column where the image is located
            col = image.anchor._from.col + 1
            
            # Save based on which column the image is in
            img_data = io.BytesIO(image._data())
            img = Image.open(img_data)
            
            if col == ndvi_col:
                current_image_path = 'images/current_ndvi.png'
                img.save(current_image_path)
                print(f"Saved current NDVI image to {current_image_path}")
            elif col == old_ndvi_col:
                old_image_path = 'images/old_ndvi.png'
                img.save(old_image_path)
                print(f"Saved old NDVI image to {old_image_path}")
                
    except Exception as e:
        print(f"Error extracting images: {e}")
    
    return current_image_path, old_image_path

def generate_page2(excel_file, template_file, output_file):
    # Extract images from Excel
    current_image_path, old_image_path = extract_images_from_excel(excel_file)
    
    # Read the Excel file for other data
    df = pd.read_excel(excel_file)
    
    # Read the HTML template
    with open(template_file, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Get values from Excel using the correct column names
    old_ndvi_value = str(df['Old NDVI value'].iloc[0])
    current_ndvi_value = str(df['NDVI value'].iloc[0])
    ndvi_change = str(df['NDVI change'].iloc[0])
    ndvi_advisory = str(df['NDVI ADVISORY'].iloc[0])
    
    # Get dates from Excel
    try:
        old_date = df['Old Date'].iloc[0].strip()  # Remove any whitespace or newlines
        old_image_date = pd.to_datetime(old_date).strftime('%d/%m/%Y')
    except:
        old_image_date = "N/A"
        
    try:
        # Use Current image date for the new image
        current_date = df['Current  image'].iloc[0]
        new_image_date = pd.to_datetime(current_date).strftime('%d/%m/%Y')
    except:
        new_image_date = "N/A"
    
    print("\nDates from Excel:")
    print(f"Old Date: {old_image_date}")
    print(f"New Date: {new_image_date}")
    
    # Replace the date placeholders in the HTML
    html_content = html_content.replace('OLD IMAGE DATE:<br/>IMAGE DATE1', f'OLD IMAGE DATE:<br/>{old_image_date}')
    html_content = html_content.replace('NEW IMAGE DATE<br/>IMAGE DATE2', f'NEW IMAGE DATE<br/>{new_image_date}')
    
    # Update the image sources in the HTML
    if old_image_path:  # First box should have the old NDVI image
        html_content = html_content.replace(
            '<img alt="image 2 " class="w-[220px] h-[220px] object-cover" height="220" src=" " width="220"/>',
            f'<img alt="Old NDVI" class="w-[220px] h-[220px] object-cover" height="220" src="{old_image_path}" width="220"/>'
        )
    
    if current_image_path:  # Second box should have the current NDVI image
        html_content = html_content.replace(
            '<img alt="image 3" class="w-[220px] h-[220px] object-cover" height="220" src="  " width="220"/>',
            f'<img alt="Current NDVI" class="w-[220px] h-[220px] object-cover" height="220" src="{current_image_path}" width="220"/>'
        )
    
    # Update NDVI values and advisory
    html_content = html_content.replace('OLD NDVI VALUE', old_ndvi_value)
    html_content = html_content.replace('NDVI VALUE', current_ndvi_value)
    html_content = html_content.replace('NDVI change', ndvi_change)
    html_content = html_content.replace('NDVI ADVISORY', ndvi_advisory)


    # Save the generated HTML
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Page 2 report generated successfully: {output_file}")

if __name__ == "__main__":
    # File paths
    excel_file = "demo.xlsx"
    template_file = "templete/page2.html"
    output_file = "output_page2.html"
    
    # Generate the report
    generate_page2(excel_file, template_file, output_file)
