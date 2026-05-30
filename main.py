#streamlit run main.py

import streamlit as st
import re
import pycountry
import difflib
import openpyxl as op

def validateState(country, state):
    # Get the country object using pycountry
    countryObj = pycountry.countries.get(alpha_2=country)

    if not countryObj:
        return False
    
    # Get the subdivisions (states/provinces) for the country
    subdivisions = pycountry.subdivisions.get(country_code=country)
    
    # Check if the provided state code is in the list of subdivisions
    for subdivision in subdivisions:
        if subdivision.code.split('-')[1] == state:
            return True
            
    return False

def validatePostalCode(country, postalCode):
    # Define regex patterns for postal codes based on country
    PSPatterns = {
        'US': r'^\d{5}(-\d{4})?$',  # 5 digits or 5 digits + 4 digits for US' building specifiers (less used)
        'CA': r'^[A-Za-z]\d[A-Za-z] \d[A-Za-z]\d$',  # Canada
        'GB': r'^[A-Za-z]{1,2}\d{1,2} \d[A-Za-z]{2}$',  # UK
        # Add more, perhaps with AI
    }
    
    pattern = PSPatterns.get(country)
    if pattern and re.match(pattern, postalCode):
        return True
    return False

def validateSteetType(streetType):
    # Get street types from streetTypes.txt
    if not streetType:
        return False
    with open('streetTypes.txt', 'r') as file:
        validStreetTypes = [line.strip() for line in file.readlines()]
    
    # Use difflib to find the closest match for the street type
    closestMatch = difflib.get_close_matches(streetType, validStreetTypes, n=1)
    
    if closestMatch and closestMatch[0] == streetType:
        return True
    return False

def validateAddress(address):
    # Give address as a dictionary with keys: country, state, postal_code, street_type
    country = getCountryCode(address.get('country')) 
    state = address.get('state')
    postal_code = address.get('postal_code')
    street_type = address.get('street_type')
    
    
    if not validateState(country, state): #and not validateState(getCountryCode(country), state):
        return "Invalid state/province for the specified country" #.{country}{state}"
    
    if not validatePostalCode(country, postal_code):
        return "Invalid postal code format for the specified country."
    
    if not validateSteetType(street_type):
        return "Invalid street type."
    
    return "Address is valid."


def getCountryCode(countryName):
    if not countryName:
        return None
    
    countryName = str(countryName).strip()

    try:
        country = pycountry.countries.lookup(countryName)
        return country.alpha_2
    except LookupError:
        return None

def extractStreetType(addressLine):
    with open('streetTypes.txt', 'r') as file:
        validStreetTypes = [line.strip() for line in file.readlines()]
    for streetType in validStreetTypes:
        if streetType in addressLine:
            return streetType

# Example usage
testAddress = {
    'country': 'CA',
    'state': 'ON',
    'postal_code': 'M5V 3L9',
    'street_type': 'ST'
}


# Streamlit UI
st.title("Address Validator")
uploadedFile = st.file_uploader("Upload a spreadsheet with addresses", type=["xlsx"])
if uploadedFile:
    # Process the uploaded file and validate addresses
    with st.spinner("Processing"):
        # Read the Excel file using pandas
        # Excel rows: record id, Address line 1, City, Province/State, Postal Code, Country0
        sheet = op.load_workbook(uploadedFile).active
        valid = 0
        invalid = 0
        errors = []
        for row in sheet.iter_rows(values_only=True): #min_row=2 to skip header
            #recordId, addressLine, City, State, postal_code, country
            if row is None or all(cell is None for cell in row):
                continue
            addressLine = row[1]
            state = row[3]
            postalCode = row[4]
            country = row[5]
            streetType = extractStreetType(addressLine)
            address = {
                'country': country,
                'state': state,
                'postal_code': postalCode,
                'street_type': streetType
            }
            if validateAddress(address) == "Address is valid.":
                valid += 1
            else:
                invalid += 1
                errors.append((row[0], validateAddress(address))) #recordId, error message
    st.success(f"Validation complete! Valid addresses: {valid}, Invalid addresses: {invalid}")
    if errors:
        st.subheader("Errors:")
        for recordId, error in errors:
            st.write(f"Record ID: {recordId}, Error: {error}")
            
            
