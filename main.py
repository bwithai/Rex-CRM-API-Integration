import requests
import pandas as pd

# Replace 'API_KEY' with your actual API key/token
api_key = 'API_KEY'
api_url = 'https://api.rexsoftware.com/v1/rex/'

# Appraisals search API endpoint
search_url = api_url + 'appraisals/search'

# modify search_url according your need like
"""
    Appraisals
    Appointments
    Listings
    Properties
    Commission worksheets
    Feedback
"""

# Set headers with API key
headers = {
    'Authorization': f'Bearer {api_key}',
    'Content-Type': 'application/json'
}

# Define the payload for the request
payload = {
    'limit': 13,
    'offset': 0
}

# Send POST request to search appraisals
response = requests.post(search_url, json=payload, headers=headers)

# Process the response data
if response.status_code == 200:
    search_result = response.json()

    if 'result' in search_result:
        result_data = search_result['result']

        if 'rows' in result_data:
            appraisals = result_data['rows']
            total_appraisals = result_data['total']

            print(f'Total Appraisals: {total_appraisals}\n')

            if appraisals:
                # Create a list to store the appraisal data
                appraisal_data = []

                for appraisal in appraisals:
                    # Extract the desired fields from each appraisal
                    agent_1 = appraisal.get('agent_1')
                    agent_2 = appraisal.get('agent_2')
                    appraisal_date = appraisal.get('appraisal_date')
                    min_price = appraisal.get('price_min')
                    max_price = appraisal.get('price_max')
                    rent_per_week = appraisal.get('price_rent')
                    interest_level = appraisal.get('interest_level')
                    archive_date = appraisal.get('archive_date')
                    archive_reason = appraisal.get('archive_reason')
                    archive_lost_agency = appraisal.get('archive_lost_agency')
                    address = appraisal.get('property', {}).get('address')

                    # Append the appraisal data to the list
                    appraisal_data.append({
                        'Agent 1': agent_1,
                        'Agent 2': agent_2,
                        'Appraisal date': appraisal_date,
                        'Min price': min_price,
                        'Max price': max_price,
                        'Rent (p/w)': rent_per_week,
                        'Interest level': interest_level,
                        'Archive Date': archive_date,
                        'Archive Reason': archive_reason,
                        'Archive Lost Agency': archive_lost_agency,
                        'Address': address
                    })

                # Create a DataFrame from the appraisal data
                df = pd.DataFrame(appraisal_data)
                print(len(appraisal_data))

                # Save the data to an Excel file
                filepath = 'appraisals.xlsx'
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Appraisals', index=False)

                print(f'Appraisals data saved successfully as {filepath}')
            else:
                print('No appraisals found.')
        else:
            print('No "rows" field found in the result data.')
    else:
        print('No "result" field found in the search result.')
else:
    print('Error searching appraisals:', response.text)