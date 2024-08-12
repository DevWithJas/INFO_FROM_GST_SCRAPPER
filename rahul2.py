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

# List of GST numbers extracted from the image
gst_numbers = [
    '09AALCS2031F1ZI', '09AASFG2420N1ZZ', '09ABHCS7569K1ZG', '09ABIFM3909M1ZR', '09ABJFS6441B1Z6',
    '09ADHPA9698J1ZV', '09ADKFS1347R1Z9', '09AEZPY2080B1ZY', '09AFPPA9624Q1ZN', '09AILPS3457J1ZL',
    '09ALTPB5433C1Z7', '09ANTPR1156N1Z4', '09APHPA0819H1Z6', '09ARNPA8399N1ZS', '09AWEPM7335A2ZN',
    '09BESPS3181K1ZN', '09BILPS3093K1ZI', '09BIXPS3234E1ZS', '09BMXPA5363G1ZO', '09BMZPK0166P1Z2',
    '09BNHPJ3218G1Z3', '09CBAPS9749H1ZZ', '09CBMPA9603C1ZU', '09CEXPS8071Q1ZY', '09CFYPS4629N1Z5',
    '09DGXPB8136C1Z5', '09DZVPS4905P1Z3', '09ECZPK6312Q1ZI', '16AHXPI5207R1ZJ', '18AACCK5599H1Z1',
    '19AACCK5599H1ZZ', '24AAACT5540K1ZD', '24AACCK5599H1Z8', '29AAJCM6655R1ZD', '36AAFCK4574P1ZU',
    '38AAMCS6730J1ZT'
]

# List to store the results
results = []

# Process each GST number
for gst_number in gst_numbers:
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
output_path = r'C:\Users\JAS\Downloads\GST_provided.xlsx'
result_df.to_excel(output_path, index=False)

print(f"Results saved to {output_path}")
