import ezsheets
import re

def parseData(id, month, year):
    ss = ezsheets.Spreadsheet(id)

    rateDictDict = makeRateDictDict(ss)

    datePat = rf"0?{month}/\d+/\d*{year%100}" # Can match 12/1/2024 and 12/1/24
    ACDict = makeACDict(ss, datePat)

    baseInfoDict = {}
    dataDict = {}
    for client, rowNum in ACDict.items():
        sheet = ss[client]
        
        # baseInfoList contains [Name, Address1, Address2, Matter]
        baseInfoDict[client] = makeBaseInfoList(sheet)

        rateName = sheet["C4"]
        rateDict = rateDictDict.get(rateName, rateDictDict["Latest"])
        
        dataDict[client] = []
        # Start from the effective date entry row
        row = sheet.getRow(rowNum)
        while True:
            nameList = row[1].split("/") # Name column: Rudy/Ayaka, blank cell is acceptable
            name = nameList[0].replace(" ", "")
            name = name.upper()

            time = row[3].replace(" ", "")
            if time and time.replace(".", "", 1).isdigit():
                time = float(time)
                rate = rateDict.get(name, 0)
                dataDict[client].append([
                    row[0], # Date
                    row[2], # Activity
                    time, # Time
                    rate # Rate
                    ])
            elif time and time[0].upper() == "F": # In case of Flat Fee
                dataDict[client].append([
                    row[0], # Date
                    row[2], # Activity
                    "Flat Fee",
                    0 # Rate
                    ])
            elif time and time[len(time)- 1] == "?":
                dataDict[client].append([
                    row[0], # Date
                    row[2], # Activity
                    float(time.replace("?", "")),
                    rate # Rate
                    ])
            else: # In case of Not Billed, or not digit, blank or anything else not starting with 'F' 
                dataDict[client].append([
                    row[0], # Date
                    row[2], # Activity
                    "Not Billed", # Time
                    0 # Rate
                    ])
            
            rowNum += 1
            row = sheet.getRow(rowNum)
            if not row[0] and not row[2]: # If both Date and Work Progress are blank
                break

            if row[0] and not re.match(datePat, row[0]): # If there is a row, but that is for next month
                break

    return baseInfoDict, dataDict

def makeACDict(ss, datePat):
    ACDict = {}
    for sheet in ss:
        for rowNum, cell in enumerate(sheet.getColumn(1), 1): # Check until the very end of sheet
            if re.match(datePat, cell):
                ACDict[sheet.title] = rowNum
                break
    
    return ACDict

def makeBaseInfoList(sheet):
    tempStr = sheet["C2"] # Cell C2 must be Name(newline)Address1(newline)Address2
    baseInfoList = tempStr.split("\n")

    matterStr = sheet["C3"] # Cell C3 is reserved for Invoice Matter.
    baseInfoList.append(matterStr)

    return baseInfoList

def makeRateDictDict(ss):
    for sheet in ss:
        sheetName = sheet.title
        sheetName.replace(" ", "")
        if sheetName[:4].upper() == "RATE":
            rateMaster = sheet
            break
    
    nameList = []
    nameRow = rateMaster.getRow(1)
    startCol = 1
    name = nameRow[startCol].upper()
    while name:
        nameList.append(name)
        startCol += 1
        name = nameRow[startCol].upper()

    rateDictDict = {}
    startRow = 2
    rateRow = rateMaster.getRow(startRow)
    rateName = rateRow[0]
    while rateName:
        rateDictDict[rateName] = {}
        for colNum, name in enumerate(nameList, 1):
            rateDictDict[rateName][name] = int(rateRow[colNum])
        
        startRow += 1
        rateRow = rateMaster.getRow(startRow)
        rateName = rateRow[0]
    
    return rateDictDict

def makeRateDict(sheet):
    rateDict = {}
    columnF = sheet.getColumn("F") # Rates must be listed in column F (Name) and G (Rate).
    for rowNum, cell in enumerate(columnF, 1):
        if cell:
            rateDict[cell.upper()] = int(sheet[f"G{rowNum}"])
    
    return rateDict




