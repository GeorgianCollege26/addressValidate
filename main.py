#streamlit run main.py

import streamlit as st
import re
import pycountry
import difflib
import openpyxl as op
import unicodedata


def validateState(country, state):
    if not country:
        return False
    # Check if the country is valid
    try:
        countryObj = pycountry.countries.lookup(country)
    except LookupError:
        return False

    # Get subdivisions
    subdivisions = pycountry.subdivisions.get(country_code=countryObj.alpha_2)
    for subdivision in subdivisions:
        if subdivision.code.split('-')[1] == state:
            #return code if state is valid, otherwise return False
            return subdivision.code.split('-')[1]
    return False

def formatString(string):
    if not string:
        return None

    string = str(string).strip()

    # Convert accented characters to their base characters
    string = unicodedata.normalize('NFKD', string)
    string = ''.join(
        c for c in string
        if not unicodedata.combining(c)
    )

    # Remove punctuation
    string = re.sub(r'[^\w\s]', '', string)

    # Replace multiple spaces with a single space
    string = re.sub(r'\s+', ' ', string)

    return string

def validatePostalCode(country, postalCode):
    # Define regex patterns for postal codes based on country
    PSPatterns = {
        'US': r'^\d{5}(-\d{4})?$',  # 5 digits or 5 digits + 4 digits for US' building specifiers (less used)
        'CA': r'^[A-Za-z]\d[A-Za-z] \d[A-Za-z]\d$',  # Canada
        'GB': r'^(GIR\s?0AA|[A-Za-z]{1,2}\d[A-Za-z\d]?\s?\d[A-Za-z]{2})$',  # UK
        'FR': r'^\d{5}$',  # France
        'ES': r'^\d{5}$',  # Spain

        # Add more, perhaps with AI
    }
    
    pattern = PSPatterns.get(country)
    if not pattern:
        st.warning(f"No postal code pattern defined for country: {country}. Skipping postal code validation.")
        return True 
    matchCode = re.match(pattern, str(postalCode))
    matchCode2 = re.match(pattern, normalizePostalCode(str(postalCode), country))
    if pattern and matchCode:
        return 
    elif pattern and matchCode2:
        return matchCode2
    return False

def normalizePostalCode(postalCode, country):
    # Use regex to insert the space in the correct position
    patterns = {
        'CA': r'^([A-Za-z]\d[A-Za-z])\s?(\d[A-Za-z]\d)$',  # Canada
        # Add more patterns for other countries that have formatting requirements that have strict formatting rules
        # UK won't work as it can e1g 4rte or sw1a 1aa
    }
    pattern = patterns.get(country)
    if not pattern:
        return postalCode  # Return as-is if no specific pattern for the country
    match = re.match(pattern, postalCode)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    else:
        # If it doesn't match the pattern, return as-is or raise an error
        return postalCode


def validateAddress(address):
    if not address:
        return "No address provided."
    try:
        # Give address as a dictionary with keys: country, state, postal_code, street_type
        try:
            country = pycountry.countries.lookup(address.get('country')) if address.get('country') else None
        except LookupError:
            try:
                countryCode = getCountryCode(address.get('country'))
                country = pycountry.countries.get(alpha_2=countryCode) if countryCode else None
            except LookupError:
                return f"Invalid country: {address.get('country')}"
        countryCode = country.alpha_2 if country else None
        countryName = country.name.upper() if country else None

        stateCode = getState(address.get('state'), countryCode, "code") if address.get('state') else None
        stateName = getState(address.get('state'), countryCode, "name") if address.get('state') else None

        postalCode = address.get('postalCode')
        streetAddress = address.get('streetAddress')
        
        stateValid = validateState(countryCode, stateCode)
        postalCodeValid = validatePostalCode(countryCode, postalCode)
        if stateValid == False: 
            return f"Invalid state/province for the specified country {countryCode}/{stateCode}"

        if postalCodeValid == False:
            return f"Invalid postal code format for the specified country. {postalCodeValid}"
        
        if not streetAddress:
            return "Street type could not be determined from the address."
        if streetAddress[0] == None:
            return "Street type could not be determined from the address."

        # If all validations pass, return the validated address as a dictionary
        # but return state codes instead of names for US and Canada, as they are more commonly used in addresses
        if countryCode in ['US', 'CA']:
            stateName = stateCode
        validatedAddress = { 
            'country': countryName,
            'state': stateName,
            'postalCode': postalCodeValid,
            'streetAddress': streetAddress
        }
        st.success(f"{streetAddress}, {stateName.upper()}, {postalCode}, {countryName} is valid.")
        return validatedAddress
    except Exception as e:
        return f"An error occurred during validation: {str(e)}"


def getCountryCode(countryName):
    if not countryName:
        return None
    
    countryName = str(countryName).strip()

    try:
        country = pycountry.countries.lookup(countryName)
        return country.alpha_2
    except LookupError:
        return None
    
def getState(stateName, countryCode, argumentType='name'):
    #Function that can take a state name or code (Allowing for minor misspellings) and a country code, and return the standardized state name or code based on the argumentType 
    if not stateName or not countryCode:
        return None
    try:
        stateName = formatString(str(stateName).strip())
        subdivisions = pycountry.subdivisions.get(country_code=countryCode)
        #subdivisionNames = [subdivision.name for subdivision in subdivisions]
        for subdivision in subdivisions:
            
            #Word match
            if stateName.lower() in subdivision.name.lower():
                if argumentType == 'name':
                    return subdivision.name
                return subdivision.code.split('-')[1]
            #Allow for minor misspellings using difflib
            elif difflib.SequenceMatcher(None, subdivision.name.lower(), stateName.lower()).ratio() > 0.85:
                if argumentType == 'name':
                    return subdivision.name
                return subdivision.code.split('-')[1]
            
            #Code match
            elif stateName.lower() in subdivision.code.split('-')[1].lower():
                if argumentType == 'name':
                    return subdivision.name
                return subdivision.code.split('-')[1]
            elif difflib.SequenceMatcher(None, subdivision.code.split('-')[1].lower(), stateName.lower()).ratio() > 0.85:
                if argumentType == 'name':
                    return subdivision.name
                return subdivision.code.split('-')[1]
    except Exception as e:
        st.warning(f"An error occurred while validating state: {str(e)}")
        return None


def abbreviateAddress(address):
    # shorten things like south, north, east, west to S, N, E, W
    if not address:
        return None
    dictionary = {
        "south": "S",
        "north": "N",
        "east": "E",
        "west": "W",
        "northwest": "NW",
        "northeast": "NE",
        "southwest": "SW",
        "southeast": "SE",
        "s": "S",
        "n": "N",
        "e": "E",
        "w": "W",
        "nw": "NW",
        "ne": "NE",
        "sw": "SW",
        "se": "SE",
        "appartment": "APT",
        "app": "APT",
        "suite": "STE",
        "ste": "STE",
        "unit": "UNIT",
    }
    tokens = formatString(address).lower().replace(".", "").split()
    for token in tokens:
        if token in dictionary:
            address = re.sub(r'\b' + re.escape(token) + r'\b', dictionary[token], address, flags=re.IGNORECASE)
    return address.upper()

def validateStreetType(address):
    if not address:
        return False
    validTypes = {
    #Generated by chatGPT based on https://www.canadapost-postescanada.ca/cpc/en/support/articles/addressing-guidelines/symbols-and-abbreviations.page
    # A
    "abbey": "ABBEY",
    "abbey": "ABBEY",
    "acres": "ACRES",
    "alley": "ALLEY",
    "alley": "ALLEY",
    "avenue": "AVE",
    "ave": "AVE",

    # B
    "bay": "BAY",
    "beach": "BEACH",
    "bend": "BEND",
    "boulevard": "BLVD",
    "blvd": "BLVD",
    "by-pass": "BYPASS",
    "bypass": "BYPASS",
    "byway": "BYWAY",

    # C
    "campus": "CAMPUS",
    "cape": "CAPE",
    "centre": "CTR",
    "center": "CTR",
    "ctr": "CTR",
    "chase": "CHASE",
    "circle": "CIR",
    "cir": "CIR",
    "circuit": "CIRCT",
    "circt": "CIRCT",
    "close": "CLOSE",
    "common": "COMMON",
    "concession": "CONC",
    "conc": "CONC",
    "corners": "CRNRS",
    "crnrs": "CRNRS",
    "court": "CRT",
    "crt": "CRT",
    "cove": "COVE",
    "crescent": "CRES",
    "cres": "CRES",
    "crossing": "CROSS",
    "cross": "CROSS",
    "cul-de-sac": "CDS",
    "cds": "CDS",

    # D
    "dale": "DALE",
    "dell": "DELL",
    "diversion": "DIVERS",
    "divers": "DIVERS",
    "downs": "DOWNS",
    "drive": "DR",
    "dr": "DR",

    # E
    "end": "END",
    "esplanade": "ESPL",
    "espl": "ESPL",
    "estates": "ESTATE",
    "estate": "ESTATE",
    "expressway": "EXPY",
    "expy": "EXPY",
    "extension": "EXTEN",
    "exten": "EXTEN",

    # F
    "farm": "FARM",
    "field": "FIELD",
    "forest": "FOREST",
    "freeway": "FWY",
    "fwy": "FWY",
    "front": "FRONT",

    # G
    "gardens": "GDNS",
    "gdns": "GDNS",
    "gate": "GATE",
    "glade": "GLADE",
    "glen": "GLEN",
    "green": "GREEN",
    "grounds": "GRNDS",
    "grnds": "GRNDS",
    "grove": "GROVE",

    # H
    "harbour": "HARBR",
    "harbor": "HARBR",
    "harbr": "HARBR",
    "heath": "HEATH",
    "heights": "HTS",
    "hts": "HTS",
    "highlands": "HGHLDS",
    "hghlds": "HGHLDS",

    "highway": "HWY",
    "hwy": "HWY",
    "hill": "HILL",
    "hollow": "HOLLOW",

    # I
    "inlet": "INLET",
    "island": "ISLAND",

    # K
    "key": "KEY",
    "knoll": "KNOLL",

    # L
    "landing": "LANDNG",
    "landng": "LANDNG",
    "lane": "LANE",
    "lane": "LANE",
    "limits": "LMTS",
    "lmts": "LMTS",
    "line": "LINE",
    "link": "LINK",
    "lookout": "LKOUT",
    "lkout": "LKOUT",
    "loop": "LOOP",

    # M
    "mall": "MALL",
    "manor": "MANOR",
    "maze": "MAZE",
    "meadow": "MEADOW",
    "mews": "MEWS",
    "moor": "MOOR",
    "mount": "MOUNT",
    "mountain": "MTN",
    "mtn": "MTN",

    # O
    "orchard": "ORCH",
    "orch": "ORCH",

    # P
    "parade": "PARADE",
    "park": "PK",
    "pk": "PK",
    "parkway": "PKY",
    "pky": "PKY",
    "passage": "PASS",
    "pass": "PASS",
    "path": "PATH",
    "pathway": "PTWAY",
    "ptway": "PTWAY",
    "pines": "PINES",
    "place": "PL",
    "pl": "PL",
    "plateau": "PLAT",
    "plat": "PLAT",
    "plaza": "PLAZA",
    "point": "PT",
    "pt": "PT",
    "port": "PORT",
    "private": "PVT",
    "pvt": "PVT",
    "promenade": "PROM",
    "prom": "PROM",

    # Q
    "quay": "QUAY",

    # R
    "ramp": "RAMP",
    "range": "RG",
    "rg": "RG",
    "ridge": "RIDGE",
    "rise": "RISE",
    "road": "RD",
    "rd": "RD",
    "route": "RTE",
    "rte": "RTE",
    "row": "ROW",
    "run": "RUN",

    # S
    "square": "SQ",
    "sq": "SQ",
    "street": "ST",
    "st": "ST",
    "subdivision": "SUBDIV",
    "subdiv": "SUBDIV",

    # T
    "terrace": "TERR",
    "terr": "TERR",
    "thicket": "THICK",
    "thick": "THICK",
    "towers": "TOWERS",
    "townline": "TLINE",
    "tline": "TLINE",
    "trail": "TRAIL",
    "turnabout": "TRNABT",
    "trnabt": "TRNABT",

    # V
    "vale": "VALE",
    "via": "VIA",
    "view": "VIEW",
    "village": "VILLGE",
    "villge": "VILLGE",
    "villas": "VILLAS",
    "vista": "VISTA",

    # W
    "walk": "WALK",
    "way": "WAY",
    "wharf": "WHARF",
    "wood": "WOOD",
    "wynd": "WYND",

    #french types
    "rue": "RUE",
    "allee": "ALLEE",
}
    # Split into tokens and check for street types, starting from the end of the string
    tokens = formatString(address).lower().replace(".", "").split()

    for token in reversed(tokens):
        if token in validTypes:
            #replace the street type in the address string with the standardized version
            standardizedType = validTypes[token]
            #remove the original street type from the address and replace with standardized version
            streetAddress = re.sub(r'\b' + re.escape(token) + r'\b', standardizedType, address, flags=re.IGNORECASE)
            return streetAddress.upper()

    #st.warning(formatString(address) + " | does not contain a valid street type.")
    return None

def fixSwappedCols(addressLineCol, cityCol, stateCol, postalCodeCol, countryRowCol):
    st.info(f"Attempting to fix swapped columns for the following data: {addressLineCol}, {cityCol}, {stateCol}, {postalCodeCol}, {countryRowCol}")
    foundCountry = None
    foundState = None
    foundPostalCode = None
    foundAddressLine = None
    foundCity = None
    confidence = 1
    
    tempList = [addressLineCol, cityCol, stateCol, postalCodeCol, countryRowCol]
    for i in range(len(tempList)):
        test = tempList[i]
        testCountry = getCountryCode(test)
        if testCountry:
            foundCountry = testCountry
            tempList.pop(i)
            break
            
    for j in range(len(tempList)):
        test = formatString(str(tempList[j]).strip())
        testState = getState(test, foundCountry, "name")
        #st.info(f"Testing {test} for state with country {foundCountry}. Result: {testState}")
        testStreet = validateStreetType(abbreviateAddress(test))
        if testState:
            foundState = testState
        elif testStreet:
            foundAddressLine = testStreet
        #Check if 30 characters or shorter and contains no spaces and digits, it's likely a city name
        elif len(test) <= 30 and not any(char.isdigit() for char in test) and " " not in test:
            foundCity = test
            confidence =0
        #Check if 9 digits or shorter and contains digits, it's likely a postal code
        elif len(test) <= 9 and any(char.isdigit() for char in test):
            foundPostalCode = normalizePostalCode(test, foundCountry)
        #If it contains spaces and digits and is longer than 8 characters, it's likely an address line
        elif len(test) > 8 and any(char.isdigit() for char in test) and " " in test:
            foundAddressLine = test
            confidence = 0
        elif foundCountry and foundState and foundPostalCode and foundAddressLine:
            foundCity = test
            
    if not foundCountry or not foundState or not foundPostalCode or not foundAddressLine or not foundCity:
        st.warning("Could not automatically fix swapped columns. Please ensure the spreadsheet is formatted correctly.")
        st.info(f"After attempting to fix swapped columns, the following data was found: Address line:{foundAddressLine}, City: {foundCity}, State: {foundState}, Postal code: {foundPostalCode}, Country: {foundCountry}")
        return None
    else: 
        address = {
            'country': foundCountry,
            'state': foundState,
            'postalCode': foundPostalCode,
            'streetAddress': foundAddressLine,
            'city': foundCity,
            'confidence': confidence
        }
        return address
    return None
        

#Main program


# Streamlit UI
st.title("Address Validator")
uploadedFile = st.file_uploader("Upload a spreadsheet with addresses", type=["xlsx"])
if uploadedFile:
    # Process the uploaded file and validate addresses
    with st.spinner("Processing"):
        # Read the Excel file using pandas
        # Excel rows: record id, Address line 1, City, Province/State, Postal Code, Country
        sheet = op.load_workbook(uploadedFile).active
        valid = 0
        invalid = 0
        for row in sheet.iter_rows(values_only=True, min_row=2): #min_row=2 to skip header
            #recordId, addressLine, City, State, postal_code, country
            if row is None or all(cell is None for cell in row):
                continue
            addressLine = (row[1])
            state = (row[3])
            city = (row[2])
            postalCode = (row[4])
            country = (row[5])
            streetAddress = validateStreetType(abbreviateAddress(addressLine))
            address = {
                'country': country,
                'state': state,
                'city': city,
                'postalCode': postalCode,
                'streetAddress': streetAddress
            }
            validateResult = validateAddress(address)
            if isinstance(validateResult, dict):
                valid += 1
                #st.write(f"{streetAddress} {streetType}, {state}, {postalCode}, {country} is valid.")
            else:
                tryFix = fixSwappedCols(addressLine, city, state, postalCode, country)
                if tryFix:
                    newResult = validateAddress(tryFix)
                    if isinstance(newResult, dict):
                        valid += 1
                        st.write(f"{address} is invalid. However, after attempting to fix swapped columns, {tryFix['streetAddress']}, {tryFix['state']}, {tryFix['postalCode']}, {tryFix['country']} is valid. Confidence level of fix: {tryFix['confidence']}")
                    else:
                        invalid += 1
                        st.write(f"{address} is invalid. Reason: {validateResult}. Also attempted to fix swapped columns, but it still failed validation. Reason: {newResult}")
                else:
                    invalid += 1
                    st.write(f"{address}is invalid. Reason: {validateResult}")

    st.success(f"Validation complete! Valid addresses: {valid}, Invalid addresses: {invalid}")
    #States of england
    for subdivision in pycountry.subdivisions.get(country_code='FR'):
        st.write(subdivision.name)

