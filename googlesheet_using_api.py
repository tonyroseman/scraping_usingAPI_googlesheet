import requests
import csv
import json
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build


def fetchCompanyData(url):
    print(url)
    results = []
    try:
        response = requests.get(url)

# Check if the request was successful
        if response.status_code == 200:
            # Parse the XML content
            
            data = json.loads(response.content)
            
        # Extract and print specific information
            for result in data['Results']:
                company_name = result['Name']
                
                contact_name = result['Contact']['FullName'] if "FullName" in result['Contact'] else "N/A"
                # contact_last_name = result['Contact']['LastName']
                
                contact_email = result['Contact']['Email1'] if "Email1" in result['Contact'] else "N/A"
                contact_phone = result['Contact']['Phone1'] if "Phone1" in result['Contact'] else "N/A"
                website = result['Contact']['Website1'] if "Website1" in result['Contact'] else "N/A"
                street1 = result['Contact']['Address']['Street1'] + ", " if "Street1" in result['Contact']['Address'] else ""
                street2 = result['Contact']['Address']['Street2'] + ", " if "Street2" in result['Contact']['Address'] else ""
                City = result['Contact']['Address']['City'] + ", " if "City" in result['Contact']['Address'] else ""
                PostalCode = result['Contact']['Address']['PostalCode'] if "PostalCode" in result['Contact']['Address'] else ""
                Region = result['Contact']['Address']['RegionCode'] if "RegionCode" in result['Contact']['Address'] else ""
                CountryCode = result['Contact']['Address']['CountryCode'] if "CountryCode" in result['Contact']['Address'] else ""

                address = street1 + " "+ street2 + City + Region + " " + CountryCode
                data = [company_name, contact_name, contact_email,contact_phone,website,address ]
                results.append(data)
        else:
            print(f"Error: {response.status_code}")

        

        return results
    except Exception as e:
        print(f"Error fetching company data: {e}")
    return None
def read_csv(file_path):
    with open(file_path, 'r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        return list(reader)
def write_to_google_sheets(service, spreadsheet_id, range_name, values):
    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    print(f'{result.get("updatedCells")} cells updated.')
def main():
    company_data = []
    # URL of the company search page
    for page in range(1, 26):
        url = "https://cyberab.org/DesktopModules/ClarityEcommerce/API-Storefront/Stores/Stores?Size=10&StartIndex="+str(page)+"&TypeKey=C3PAOCandidate&__caller=pagingSettings.change"
        
        
        results = fetchCompanyData(url)
        for result in results:
            company_data.append(result)
            
    
    # print(company_data)
    # Write the data to a CSV file
    with open('temp.csv', 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Write the header row
        writer.writerow(["company_name", "contact_name", "contact_email", "contact_phone", "website", "address"])
        # Write the data row
        writer.writerows(company_data)
    # Define the path to your credentials file and the spreadsheet ID
    SERVICE_ACCOUNT_FILE = 'wide-earth-427419-h4-93a1e0b9863e.json'  # Update this path
    SPREADSHEET_ID = '12ganWXO_eGbUZ9xwMheqeJGe-oojhKzCVJMCAZqx3Jk'
    # Define the range where you want to insert the data
    RANGE_NAME = 'Sheet1!A1'  # Adjust if needed
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    csv_file_path = 'temp.csv'  # Path to your CSV file
    data = read_csv(csv_file_path)

    # Update Google Sheets
    write_to_google_sheets(service, SPREADSHEET_ID, RANGE_NAME, data)
if __name__ == "__main__":
    main()
