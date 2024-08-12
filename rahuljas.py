import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set up the path to your GeckoDriver
service = Service(r'C:\Users\JAS\Downloads\geckodriver-v0.34.0-win32\geckodriver.exe')

# Set up Firefox options
options = Options()
options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'

# Initialize the Firefox driver with the service and options
driver = webdriver.Firefox(service=service, options=options)

# Load the Excel file without headers
excel_path = r'C:\Users\JAS\Downloads\gst details.xlsx'
df = pd.read_excel(excel_path, header=None)

# List to store the results
results = []

# Process all GST numbers in the Excel file
for gst_number in df[0]:  # This will process all GST numbers in the file
    # Open the website
    driver.get('https://cleartax.in/gst-number-search/')
    
    # Find the input field by its ID and enter the GST number
    input_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "input"))
    )
    input_field.send_keys(gst_number)
    
    # Simulate hitting the Enter key
    input_field.send_keys(Keys.RETURN)
    
    try:
        # Wait for the specific element using the absolute XPath
        result_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[4]/div[1]/div[2]/div[2]/div[1]/div/div'))
        )
        
        # Extract the required information
        extracted_info = result_element.text.split('\n')
        
        # Parse the extracted info to allocate it to the correct columns
        info_dict = {
            'GST Number': gst_number,
            'Business Name': extracted_info[extracted_info.index('BUSINESS NAME') + 1] if 'BUSINESS NAME' in extracted_info else '',
            'PAN': extracted_info[extracted_info.index('PAN') + 1] if 'PAN' in extracted_info else '',
            'Address': extracted_info[extracted_info.index('ADDRESS') + 1] if 'ADDRESS' in extracted_info else '',
            'Entity Type': extracted_info[extracted_info.index('ENTITY TYPE') + 1] if 'ENTITY TYPE' in extracted_info else '',
            'Nature of Business': extracted_info[extracted_info.index('NATURE OF BUSINESS') + 1] if 'NATURE OF BUSINESS' in extracted_info else '',
            'Pincode': extracted_info[extracted_info.index('PINCODE') + 1] if 'PINCODE' in extracted_info else '',
            'Department Code': extracted_info[extracted_info.index('DEPARTMENT CODE') + 1] if 'DEPARTMENT CODE' in extracted_info else '',
            'Registration Type': extracted_info[extracted_info.index('REGISTRATION TYPE') + 1] if 'REGISTRATION TYPE' in extracted_info else '',
            'Registration Date': extracted_info[extracted_info.index('REGISTRATION DATE') + 1] if 'REGISTRATION DATE' in extracted_info else '',
        }
        
        # Store the information in the results list
        results.append(info_dict)
        
    except Exception as e:
        print(f"Error occurred for GST Number {gst_number}: {e}")
        continue

# Close the browser
driver.quit()

# Convert the results to a DataFrame
result_df = pd.DataFrame(results)

# Save the results to an Excel file
output_path = r'C:\Users\JAS\Downloads\GST_rahul.xlsx'
result_df.to_excel(output_path, index=False)

print(f"Results saved to {output_path}")
