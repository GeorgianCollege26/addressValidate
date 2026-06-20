#streamlit run main.py

import streamlit as st
import re
import pycountry
import difflib
import openpyxl as op
import unicodedata
import time


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

def formatString(string, removeAccents=True, removePunctuation=True, toUpper=False, leaveApostrophes=False, removeSpaces=False):
    if not string:
        return None

    string = str(string).strip()

    if removeAccents:
        # Convert accented characters to their base characters
        string = unicodedata.normalize('NFKD', string)
        string = ''.join(
            c for c in string
            if not unicodedata.combining(c)
        )

    if removePunctuation:
        string = re.sub(r'[\\.,/()*?><;:]', '', string)
        string = re.sub(r'[`"]', "'", string)
        if not leaveApostrophes:
            string = re.sub(r"'", ' ', string)
        # Remove punctuation
        string = re.sub(r'[^\w\s]', ' ', string)

    # Replace multiple spaces with a single space
    string = re.sub(r'\s+', ' ', string)

    if removeSpaces:
        string = string.replace(" ", "")

    return string.upper() if toUpper else string

def validatePostalCode(country, postalCode):
    #format the postal code by removing accents and punctuation, and converting to uppercase, to improve regex matching
    postalCode = formatString(postalCode, toUpper=True)
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
        return matchCode.group(0)
    elif pattern and matchCode2:
        return matchCode2.group(0)
    return False

def normalizePostalCode(postalCode, country):
    if not postalCode or not country:
        return postalCode
    postalCode = formatString(postalCode, toUpper=True)
    # Use regex to insert the space in the correct position
    patterns = {
        'CA': r'^([A-Za-z]\d[A-Za-z])\s?(\d[A-Za-z]\d)$',  # Canada
        # Add more patterns for other countries that have formatting requirements that have strict formatting rules
        # UK won't work as it can e1g 4rte or sw1a 1aa
    }
    pattern = patterns.get(country)
    if not pattern:
        #change all Os to 0s and all Is to 1s
        pattern = r'^[A-Za-z0-9]+$'
    match = re.match(pattern, postalCode)
    if match:
        # If there are mutiple groups, join them with a space
        if len(match.groups()) > 1:
            return ' '.join(match.groups())
        else:
            return match.group(0)
    else:
        if country == 'CA':
            postalCode = formatString(postalCode, toUpper=True, removePunctuation=True)
            for i in range(len(postalCode)):
                if i == 1 or i == 4 or i == 6:
                    char = postalCode[i]
                    if char == 'O':
                        postalCode = postalCode.replace('O', '0')
                    elif char == 'I':
                        postalCode = postalCode.replace('I', '1')
                elif i == 0 or i == 2 or i == 5:
                    char = postalCode[i]
                    if char == '0':
                        postalCode = postalCode.replace('0', 'O')
                    elif char == '1':
                        postalCode = postalCode.replace('1', 'I')
            return postalCode
        else:
            return postalCode


def validateAddress(address):
    # If no address is provided, return an error message
    if not address:
        return "No address provided."
    errorFields = []
    
    # If any of the required fields are missing, return an error message specifying which field is missing
    for field in ['country', 'state', 'city', 'postalCode', 'streetAddress', 'recordId']:
        if address.get(field) == None or str(address.get(field)).strip().lower() == "none":
            errorFields.append(f"{field} is required.")
    if errorFields:
        return errorFields.insert(0, "missing")

    try:
        # Give address as a dictionary with keys: country, state, postal_code, street_type
        try:
            country = pycountry.countries.lookup(formatString(address.get('country'), removeAccents=False, leaveApostrophes=True)) if address.get('country') else None
        except LookupError:
            try:
                countryCode = getCountryCode(address.get('country'))
                country = pycountry.countries.get(alpha_2=countryCode) if countryCode else None
            except LookupError:
                errorFields.append(f"Invalid country: {address.get('country')}")
            if not country:
                try:
                    countryNames = [country.name for country in pycountry.countries]
                    closestCountry = difflib.get_close_matches(formatString(address.get('country'), removeAccents=False, leaveApostrophes=True), countryNames, n=1, cutoff=0.8)
                    country = pycountry.countries.get(name=closestCountry[0]) if closestCountry else None
                except Exception as e:
                    errorFields.append(f"Invalid country: {address.get('country')}")
                
        countryCode = country.alpha_2 if country else None
        countryName = country.name.upper() if country else None
        if not country:
            errorFields.append(f"Invalid country: {address.get('country')}")
        state = address.get('state')
        if not state:
            errorFields.append("State/Province is required.")
        
        stateCode = getState(state, countryCode, "code") if address.get('state') else None
        stateName = getState(state, countryCode, "name") if address.get('state') else None

        city = address.get('city')
        if not city:
            errorFields.append("City is required.")
        postalCode = address.get('postalCode')
        if not postalCode:
            errorFields.append("Postal code is required.")
        streetAddress = address.get('streetAddress')
        if not streetAddress:
            errorFields.append("Street address is required.")
        
        if errorFields:
            return errorFields
        
        streetValid = validateStreetType(abbreviateAddress(streetAddress))
        stateValid = validateState(countryCode, stateCode)
        postalCodeValid = validatePostalCode(countryCode, postalCode)

        if stateValid == False: 
            return f"Invalid state/province for the specified country {countryCode}/{stateCode}"

        if postalCodeValid == False:
            return f"Invalid postal code format for the specified country. {postalCode}"
        
        if not streetValid:
            return "Street type could not be determined from the address."
        if streetValid[0] == None:
            return "Street type could not be determined from the address."
        
        if errorFields:
            return errorFields

        # If all validations pass, return the validated address as a dictionary
        # but return state codes instead of names for US and Canada, as they are more commonly used in addresses
        if countryCode in ['US', 'CA']:
            stateName = stateCode
        validatedAddress = { 
            'country': countryName,
            'state': stateName.upper() if stateName else None,
            'postalCode': formatString(postalCodeValid, toUpper=True, removeAccents=True, leaveApostrophes=True),
            'streetAddress': formatString(streetValid, toUpper=True, removeAccents=True, leaveApostrophes=True),
            'city': formatString(city, toUpper=True, removeAccents=True, leaveApostrophes=True),
            'recordId': str(address.get('recordId'))
        }
        ##st.success(f"{validatedAddress['streetAddress']}, {validatedAddress['city']} {validatedAddress['state']}, {validatedAddress['postalCode']}, {validatedAddress['country']} is valid.")
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
        # Avoid matching Quebec city, Mexico City, etc as states/provinces
        if ' city' in stateName.lower():
            return None
        stateName = formatString(str(stateName).strip())
        subdivisions = pycountry.subdivisions.get(country_code=countryCode)
        subdivisionNames = [subdivision.name for subdivision in subdivisions]
        subdivisionCodes = [subdivision.code.split('-')[1] for subdivision in subdivisions]
        #use difflib to find the closest match for the state name or code, allowing for minor misspellings
        closestName = difflib.get_close_matches(stateName, subdivisionNames, n=1, cutoff=0.8)
        closestCode = difflib.get_close_matches(stateName, subdivisionCodes, n=1, cutoff=0.8)
        if closestName:
            closestSubdivision = next(subdivision for subdivision in subdivisions if subdivision.name == closestName[0])
            return closestSubdivision.code.split('-')[1] if argumentType == 'code' else closestSubdivision.name
        elif closestCode:
            closestSubdivision = next(subdivision for subdivision in subdivisions if subdivision.code.split('-')[1] == closestCode[0])
            return closestSubdivision.code.split('-')[1] if argumentType == 'code' else closestSubdivision.name
        else:
            for sn in subdivisionNames:
                if stateName.lower() in sn.lower():
                    closestSubdivision = next(subdivision for subdivision in subdivisions if subdivision.name == sn)
                    return closestSubdivision.code.split('-')[1] if argumentType == 'code' else closestSubdivision.name
                elif formatString(stateName.lower()) in formatString(sn.lower()):
                    closestSubdivision = next(subdivision for subdivision in subdivisions if subdivision.name == sn)
                    return closestSubdivision.code.split('-')[1] if argumentType == 'code' else closestSubdivision.name
            return None
    except Exception as e:
        st.warning(f"An error occurred while validating state/province: {str(e)}")
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

def fixSwappedCols(addressLineCol, cityCol, stateCol, postalCodeCol, countryRowCol, addressID):
    ##st.info(f"Attempting to fix swapped columns for the following data: {addressLineCol}, {cityCol}, {stateCol}, {postalCodeCol}, {countryRowCol}")
    foundCountry = None
    foundState = None
    foundPostalCode = None
    foundAddressLine = None
    foundCity = None
    confidenceError = []
    unclaimed = []
    
    colList = [addressLineCol, cityCol, stateCol, postalCodeCol, countryRowCol]
    for i in range(len(colList)):
        currentCol = colList[i]
        testCountry = getCountryCode(currentCol)
        if testCountry:
            foundCountry = testCountry
            colList.pop(i)
            break
            
    for j in range(len(colList)):
        currentCol = formatString(str(colList[j]).strip(), removeAccents=True, removePunctuation=False)
        unclaimed.append(currentCol)
        testState = getState(currentCol, foundCountry, "name")
        #st.info(f"Testing {currentCol} for state with country {foundCountry}. Result: {testState}")
        testStreet = validateStreetType(abbreviateAddress(currentCol))


        if testState:
            foundState = testState
            #if foundCity is the same as the state, the city match below likely miscaptured it
            if foundCity and foundCity.lower() == foundState.lower():
                    foundCity = None
                    confidenceError.append("city")
            try:
                unclaimed.remove(currentCol)
            except ValueError:                
                pass

            
        if testStreet:
            try:
                confidenceError.remove("address line")
            except ValueError:
                pass
            foundAddressLine = testStreet
            try:
                unclaimed.remove(currentCol)
            except ValueError:                
                pass

        #Check if 9 digits or shorter and contains digits, it's likely a postal code
        if len(currentCol) <= 9 and any(char.isdigit() for char in currentCol):
            foundPostalCode = normalizePostalCode(currentCol, foundCountry)
            try:
                unclaimed.remove(currentCol)
            except ValueError:                
                pass
        
        # if it contains no digits and is shorter than 12 characters, it's likely a city name
        if len(currentCol) < 12 and not any(char.isdigit() for char in currentCol) and foundState and foundCountry and currentCol.lower() != foundState.lower() and currentCol.lower() != foundCountry.lower():
            foundCity = currentCol
            try:
                unclaimed.remove(currentCol)
            except ValueError:                
                pass

        #If it contains spaces and digits and is longer than 8 characters, it's likely an address line
        if len(currentCol) > 8 and any(char.isdigit() for char in currentCol) and " " in currentCol and not testStreet:
            foundAddressLine = currentCol
            confidenceError.append("address line")
            try:
                unclaimed.remove(currentCol)
            except ValueError:                
                pass
        if foundCountry and foundState and foundPostalCode and foundAddressLine and len(unclaimed) == 1:
            foundCity = unclaimed[0]
            
    if foundCountry and foundState and foundPostalCode and foundAddressLine and foundCity:
        address = {
            'country': foundCountry,
            'state': foundState,
            'postalCode': foundPostalCode,
            'streetAddress': foundAddressLine,
            'city': foundCity,
            'confidenceError': confidenceError,
            'recordId': str(addressID)
        }
        return address
    else:
        st.warning("Could not automatically fix swapped columns. Please ensure the spreadsheet is formatted correctly.")
        st.info(f"After attempting to fix swapped columns, the following data was found: Address line:{foundAddressLine}, City: {foundCity}, State: {foundState}, Postal code: {foundPostalCode}, Country: {foundCountry}")
        return None
    
def displayResults(validList, invalidList): 
    col1, col2 = st.columns(2)
   
    with col1:
        st.header("Valid addresses")
        for valid in validList:
            st.success(f"{valid['recordId']} : {valid['streetAddress']}, {valid['city']} {valid['state']}, {valid['postalCode']}, {valid['country']} is valid.")
   
    with col2:
        st.header("Invalid addresses")

        for invalid in invalidList:
            with st.expander(f"{invalid['recordId']} : {invalid['streetAddress']}, {invalid['city']} {invalid['state']}, {invalid['postalCode']}, {invalid['country']}"):
                if invalid.get('confidenceError'):
                    st.warning(f"Confidence warnings for this address: {', '.join(invalid['confidenceError'])}")

                # Let user edit the fields
                for col in ['recordId', 'streetAddress', 'city', 'state', 'postalCode', 'country']:
                    if invalid.get(col):
                        invalid[col] = st.text_input(f"Edit {col}", value=str(invalid[col]), key=f"{col}_{invalid['recordId']}")
                    else:
                        invalid[col] = st.text_input(f"Edit {col}", value="", key=f"{col}_{invalid['recordId']}")

                # Button to test
                if st.button("Test changes", key=f"test_{invalid['recordId']}"):
                    testResult = validateAddress(invalid)
                    if isinstance(testResult, dict):
                        st.success(f"After changes, address is now valid: {testResult['streetAddress']}, {testResult['city']} {testResult['state']}, {testResult['postalCode']}, {testResult['country']}. Moving to valid list...")
                        st.session_state.validList.append(testResult)
                        st.session_state.invalidList.remove(invalid)
                        
                        time.sleep(2)
                        st.rerun()

                # Manually marking as valid
                if st.button("Mark as valid", key=f"valid_{invalid['recordId']}"):
                    for col in ['streetAddress', 'city', 'state', 'country', 'postalCode']:
                        if invalid.get(col):
                            # Skip for state codes that are 2 letters, as they should be uppercase
                            if col == 'state' and len(invalid['state']) == 2:
                                continue
                            # Remove [] and () and their contents from the string, and convert to title case
                            invalid[col] = re.sub(r'\[.*?\]|\(.*?\)', '', str(invalid[col])).title()
                    st.session_state.validList.append(invalid)
                    st.session_state.invalidList.remove(invalid)
                    st.success(f"Address {invalid['recordId']} marked as valid. Moving...")
                    time.sleep(2)
                    st.rerun()
            #st.error(f"{invalid['recordId']} : {invalid['streetAddress']}, {invalid['city']} {invalid['state']}, {invalid['postalCode']}, {invalid['country']} is invalid.")
    

def saveResults(validList, invalidList):

    for valid in validList:
        #Save valid to excel file
        file = op.Workbook()
        sheet = file.active
        sheet.append(["Record ID", "Street Address", "City", "State/Province", "Postal Code", "Country"])
        for valid in validList:
            # Make title case
            for col in ['streetAddress', 'city', 'state', 'country']:
                if valid.get(col):
                    #Skip for state codes that are 2 letters, as they should be uppercase
                    if col == 'state' and len(valid['state']) == 2:
                        continue
                    #remove [] and () and their contents from the string
                    valid[col] = re.sub(r'\[.*?\]|\(.*?\)', '', valid[col]).title()

            sheet.append([valid['recordId'], valid['streetAddress'], valid['city'], valid['state'] if len(str(valid['state'])) == 2 else valid['state'], valid['postalCode'].title(), valid['country']])
        file.save("valid_addresses.xlsx")

    for invalid in invalidList:
        #save invalid to excel file
        file = op.Workbook()
        sheet = file.active
        sheet.append(["Record ID", "Street Address", "City", "State/Province", "Postal Code", "Country"])
        for invalid in invalidList:
            sheet.append([invalid['recordId'], invalid['streetAddress'], invalid['city'], invalid['state'], invalid['postalCode'], invalid['country']])
        file.save("invalid_addresses.xlsx")

    
    
def refreshPage(function):
    for key in list(st.session_state.keys()):
        if key.startswith(("address_", "city_", "state_", "postal_", "country_", "overwrite_")):
            del st.session_state[key]
    st.rerun()
    function()
    
def mainPage():
    # Streamlit UI
    st.title("Address Validator")


    uploadedFile = st.file_uploader("Upload a spreadsheet with addresses", type=["xlsx"])

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        instructionBtn = st.button("Instructions")
    if instructionBtn:
        st.info("To use this address validator, please upload an Excel spreadsheet (.xlsx) with the following columns: Record ID, Address line 1, City, Province/State, Postal Code, Country. The first row should contain the column headers. After uploading the file, click the 'Validate' button to process the addresses. The results will show which addresses are valid and which are invalid, along with reasons for any invalid addresses. You can also save the results to new Excel files for valid and invalid addresses.")


    if st.session_state.get('validList') and st.session_state.get('invalidList'):
        with col4:
            saveBtn = st.button("Save results to Excel")
        if saveBtn:
            saveResults(st.session_state.validList, st.session_state.invalidList)
            st.success("Results saved as valid_addresses.xlsx and invalid_addresses.xlsx")

        with col3:
            displayBtn = st.button("Display results")
        if displayBtn:
            st.switch_page(reviewPage)

    with col2:
        validateBtn = st.button("Validate") and uploadedFile
    if validateBtn:
        # Process the uploaded file and validate addresses
        with st.spinner("Processing"):
            # Read the Excel file using pandas
            # Excel rows: record id, Address line 1, City, Province/State, Postal Code, Country
            sheet = op.load_workbook(uploadedFile).active
            validList = []
            invalidList = []
            # Iterate through the rows of the spreadsheet and validate each address
            for row in sheet.iter_rows(values_only=True, min_row=2): #min_row=2 to skip header
                #skip empty rows
                if row is None or all(cell is None for cell in row):
                    continue
                # Grab columns from the row
                addressLine = (row[1])
                state = (row[3])
                city = (row[2])
                postalCode = ((row[4]))
                country = row[5]
                addressId = row[0]

                #Assign columns to dict
                address = {
                    'country': country,
                    'state': state,
                    'city': city,
                    'postalCode': postalCode,
                    'streetAddress': addressLine,
                    'recordId': addressId
                }
                validateResult = validateAddress(address)
                if isinstance(validateResult, dict):
                    validList.append(validateResult)
                    ##st.write(f"{streetAddress} {streetType}, {state}, {postalCode}, {country} is valid.")
                else:
                    tryFix = None
                    if isinstance(validateResult, list) and "missing" in validateResult:
                        pass
                    elif isinstance(validateResult, list): # and len(validateResult) == 1:
                        tryFix = fixSwappedCols(addressLine, city, state, postalCode, country, addressId)
                    if tryFix:
                        newResult = validateAddress(tryFix)
                        if isinstance(newResult, dict):
                            validList.append(newResult)
                            ##st.write(f"{address} was invalid because {validateResult}. However, after attempting to fix swapped columns, {tryFix['streetAddress']},{tryFix['city']} {tryFix['state']}, {tryFix['postalCode']}, {tryFix['country']} is valid. Confidence warnings of fix: {tryFix['confidenceError']}")
                        else:
                            invalidList.append(address)
                            ##st.error(f"{address} is invalid. Reason: {validateResult}. Also attempted to fix swapped columns, but it still failed validation. Reason: {newResult}")
                    else:
                        invalidList.append(address)
                        ##st.error(f"{address}is invalid. Reason: {validateResult}")

        # Update session state with results
        st.session_state.validList = validList
        st.session_state.invalidList = invalidList

        # After processing all addresses, display results
        st.success(f"Validation complete! Valid addresses: {len(st.session_state.validList)}, Invalid addresses: {len(st.session_state.invalidList)}")
        
        if st.session_state.get('validList') and st.session_state.get('invalidList'):
            if st.button("Confirm"):
                st.rerun()

def reviewPage():
    st.title("Review and Save Results")
    if st.session_state.get('validList') and st.session_state.get('invalidList'):
        displayResults(st.session_state.validList, st.session_state.invalidList)
    else:
        st.info("No results to display. Please upload a file and validate addresses first.")
        if st.button("Go to Validator"):
            st.switch_page(mainPage)

reviewPage = st.Page(reviewPage, title="Review Results", url_path="review")
mainPage = st.Page(mainPage, title="Validator", url_path="validator", default=True)
site = st.navigation([
    mainPage,
    reviewPage
], expanded=True, position="top" 
)
site.run()
#st.sidebar.selectbox("Select a page", ["Upload and Validate", "Review Results"])