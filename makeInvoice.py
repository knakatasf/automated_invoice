import openpyxl
from openpyxl import load_workbook
from openpyxl import drawing
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment

def makeFile(templatePath, outputPath):
    workbook = openpyxl.load_workbook(templatePath)
    workbook.save(outputPath)

def pasteLogo(sheet):
    # Make image object
    logo = drawing.image.Image("/Users/lorh/Desktop/Automated_Invoice/lohkLogo.png")
    logo.width = 850
    logo.height = 135
    sheet.add_image(logo, "A1") # Insert image object (size already adjusted)

def pasteFeeList(sheet):
    # Make image object
    feeList = drawing.image.Image("/Users/lorh/Desktop/Automated_Invoice/feeList.png")
    feeList.width = 320
    feeList.height = 110
    sheet.add_image(feeList, "C9") # Insert image object (size already adjusted)


def fillBaseInfo(sheet, baseInfoList):
    # Matter is stored in the last and fill in cell C8
    matterStr = baseInfoList.pop()
    sheet.cell(row=8, column=3, value=matterStr)

    for rowNum, value in enumerate(baseInfoList, start=7):
        sheet.cell(row=rowNum, column=5, value=value)

def findEndRow(sheet, rowNum):
    columnB = sheet.iter_cols(min_col=2, max_col=2, # Only qualify column B
                              min_row=rowNum, values_only=True)
    
    for column in columnB:
        for value in column:
            if value is not None:
                return rowNum
            rowNum += 1

    return None

# def fillFormula(sheet, rowNum, rateRangeDict): # rowNum is the row of "Attorney Services Rendered:"
#     fromManaging = rateRangeDict["Managing Partner"][0]
#     toManaging = rateRangeDict["Managing Partner"][1]
#     fromPartner = rateRangeDict["Partner"][0]
#     toPartner = rateRangeDict["Partner"][1]
#     fromAssociate = rateRangeDict["Associate Attorney"][0]
#     toAssociate = rateRangeDict["Associate Attorney"][1]
#     fromLC = rateRangeDict["Law Clerk/Paralegal"][0]
#     toLC = rateRangeDict["Law Clerk/Paralegal"][1]
#     fromLA = rateRangeDict["Legal Assistant"][0]
#     toLA = rateRangeDict["Legal Assistant"][1]

#     # Fill out "Rates:" section starting from C9
#     sheet["C9"] = f"Managing Partner: ${fromManaging}.00-{toManaging}.00/h"
#     sheet["C10"] = f"Partner: ${fromPartner}.00-{toPartner}.00/h"
#     sheet["C11"] = f"Associate Attorney: ${fromAssociate}.00-{toAssociate}.00/h"
#     sheet["C12"] = f"Law Clerk/Paralegal: ${fromLC}.00-{toLC}.00/h"
#     sheet["C13"] = f"Legal Assistant: ${fromLA}.00-{toLA}.00/h"

#     rowNum += 1 # Now rowNum is the row of "Managing Partner"

#     # Managing Partner HOURS and AMOUNT
#     sheet[f"D{rowNum}"] = f"=SUMIF(E15:E{rowNum-1}, \">={fromManaging}\", D15:D{rowNum-1})"
#     sheet[f"F{rowNum}"] = f"=SUMIF(E15:E{rowNum-1}, \">={fromManaging}\", F15:F{rowNum-1})"

#     # Partner HOURS and AMOUNT
#     sheet[f"D{rowNum+1}"] = f"=SUMIFS(D15:D{rowNum-1}, E15:E{rowNum-1}, \">={fromPartner}\", E15:E{rowNum-1}, \"<{fromManaging}\")"
#     sheet[f"F{rowNum+1}"] = f"=SUMIFS(F15:F{rowNum-1}, E15:E{rowNum-1}, \">={fromPartner}\", E15:E{rowNum-1}, \"<{fromManaging}\")"

#     # Associate Attorney HOURS and AMOUNT
#     sheet[f"D{rowNum+2}"] = f"=SUMIFS(D15:D{rowNum-1}, E15:E{rowNum-1}, \">={fromAssociate}\", E15:E{rowNum-1}, \"<{fromPartner}\")"
#     sheet[f"F{rowNum+2}"] = f"=SUMIFS(F15:F{rowNum-1}, E15:E{rowNum-1}, \">={fromAssociate}\", E15:E{rowNum-1}, \"<{fromPartner}\")"

#     # Law Clark/Paralegal HOURS and AMOUNT
#     sheet[f"D{rowNum+3}"] = f"=SUMIFS(D15:D{rowNum-1}, E15:E{rowNum-1}, \">={fromLC}\", E15:E{rowNum-1}, \"<{fromAssociate}\")"
#     sheet[f"F{rowNum+3}"] = f"=SUMIFS(F15:F{rowNum-1}, E15:E{rowNum-1}, \">={fromLC}\", E15:E{rowNum-1}, \"<{fromAssociate}\")"

#     # Legal Assistant HOURS and AMOUNT
#     sheet[f"D{rowNum+4}"] = f"=SUMIF(E15:E{rowNum-1}, \"<{fromLC}\", D15:D{rowNum-1})"
#     sheet[f"F{rowNum+4}"] = f"=SUMIF(E15:E{rowNum-1}, \"<{fromLC}\", F15:F{rowNum-1})"

#     # Total Attorneys' Fees:
#     sheet[f"F{rowNum+5}"] = f"=SUM(F{rowNum}:F{rowNum+4})"

#     # TOTAL AMOUNT CURRENTLY DUE:
#     sheet[f"F{rowNum+13}"] = f"=SUM(F{rowNum+5}, F{rowNum+8}:F{rowNum+8}, F{rowNum+11}:F{rowNum+11})"

def fillServicesRendered(sheet, rowNum, posRateRangeDict): # rowNum is the row of "Attorney Services Rendered:"
    rateSecRow = 9
    renderedRow = rowNum + 1
    travelTimeRows = 0

    for i, (position, fromToRateList) in enumerate(posRateRangeDict.items()):
        fromRate = fromToRateList[0]
        toRate = fromToRateList[1]
        if i < 5: # Only accept Managing Partner, Partner, Associate Attorney, LC/Paralegal, LA; No Travel Time!
            sheet[f"C{renderedRow + i}"] = f"{position}:"
            sheet[f"C{renderedRow + i}"].alignment = Alignment(horizontal="right", vertical="center")
            if fromRate == toRate:
                # Rate List Section
                sheet[f"C{rateSecRow + i}"] = f"{position}: ${fromRate}.00/hour"

                # Services Rendered Section
                sheet[f"D{renderedRow + i}"] = (
                    f"=SUMIFS(D15:D{rowNum},"
                    f"E15:E{rowNum},{fromRate},"
                    f"C15:C{rowNum},\"<>travel time*\")"
                )
                sheet[f"F{renderedRow + i}"] = (
                    f"=SUMIFS(F15:F{rowNum},"
                    f"E15:E{rowNum},{fromRate},"
                    f"C15:C{rowNum},\"<>travel time*\")"
                )

            else:
                # Rate List Section
                sheet[f"C{rateSecRow + i}"] = f"{position}: ${fromRate}.00-{toRate}.00/hour"

                # Services Rendered Section
                sheet[f"D{renderedRow + i}"] = (
                    f"=SUMIFS(D15:D{rowNum},"
                    f"E15:E{rowNum},\">={fromRate}\","
                    f"E15:E{rowNum},\"<={toRate}\","
                    f"C15:C{rowNum},\"<>travel time*\")"
                )

                sheet[f"F{renderedRow + i}"] = (
                    f"=SUMIFS(F15:F{rowNum},"
                    f"E15:E{rowNum},\">={fromRate}\","
                    f"E15:E{rowNum},\"<={toRate}\","
                    f"C15:C{rowNum},\"<>travel time*\")"
                )
        
        else:
            for row in sheet.iter_rows(min_row=15, max_row=rowNum-1, min_col=2, max_col=5, values_only=True):
                firstPartText = row[1][:11].lower()
                rate = row[3]
                if firstPartText.startswith("travel time") and rate >= fromRate and rate <= toRate:
                    travelTimeRows += 1
                    sheet[f"C{renderedRow + 4 + travelTimeRows}"] = f"{position}:"
                    sheet[f"C{renderedRow + 4 + travelTimeRows}"].alignment = Alignment(horizontal="right", vertical="center")

                    # Services Rendered Section
                    sheet[f"D{renderedRow + i}"] = (
                        f"=SUMIFS(D15:D{rowNum},"
                        f"E15:E{rowNum},\">={fromRate}\","
                        f"E15:E{rowNum},\"<={toRate}\","
                        f"C15:C{rowNum},\"travel time*\")"
                    )

                    sheet[f"F{renderedRow + i}"] = (
                        f"=SUMIFS(F15:F{rowNum},"
                        f"E15:E{rowNum},\">={fromRate}\","
                        f"E15:E{rowNum},\"<={toRate}\","
                        f"C15:C{rowNum},\"travel time*\")"
                    )
                    break
    
    for _ in range(2 - travelTimeRows):
        sheet.delete_rows(renderedRow + 5 + travelTimeRows)
    
    # Total Attorneys' Fees:
    sheet[f"F{rowNum + 6 + travelTimeRows}"] = f"=SUM(F{rowNum}:F{rowNum + 5 + travelTimeRows})"

    # TOTAL AMOUNT CURRENTLY DUE:
    sheet[f"F{rowNum + 14 + travelTimeRows}"] = f"=SUM(F{rowNum + 6 + travelTimeRows}, F{rowNum + 9 + travelTimeRows}:F{rowNum + 9 + travelTimeRows}, F{rowNum + 12 + travelTimeRows}:F{rowNum + 12 + travelTimeRows})"

    

def makeInvoice(outputPath, baseInfoDict, dataDict, invDate, posRateRangeDict):
    workbook = load_workbook(outputPath)
    templateSheet = workbook.active

    for client, entryList in dataDict.items():
        newSheet = workbook.copy_worksheet(templateSheet)

        pasteLogo(newSheet)
        #pasteFeeList(newSheet)

        # Name sheet with client name as appears in ACR
        newSheet.title = client

        # Fill invoice issue date
        newSheet.cell(row=7, column=3, value=invDate)

        # baseInfo is Bill To and Matter section. Matter is the last in the list.
        baseInfoList = baseInfoDict[client]
        fillBaseInfo(newSheet, baseInfoList)

        lightGrey = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")

        rowNum = 15 # Start row. Need to specify!
        for rowData in entryList:
            for colNum, value in enumerate(rowData, 2): # Start column. Need to specify!
                cell = newSheet.cell(row=rowNum, column=colNum, value=value)
                if rowNum%2 == 0: # If row # is even, fill lightgrey in cells
                    cell.fill = lightGrey
            
            timeValue = newSheet[f"D{rowNum}"].value
            if timeValue != "Flat Fee" and timeValue != "Not Billed":
                cell = newSheet.cell(row=rowNum, column=6, value=f"=D{rowNum}*E{rowNum}")
            else:
                cell = newSheet.cell(row=rowNum, column=6, value=0)

            # if newSheet[f"C{rowNum}"].value.lstrip().lower().startswith("travel time"):
            #     cell = newSheet.cell(row=rowNum, column=6, value="Calculate travel time!")
            
            if rowNum%2 == 0:
                cell.fill = lightGrey

            rowNum += 1
        
        endRow = findEndRow(newSheet, rowNum) # Want to find the first row which is not empty
        numToDelete = endRow - rowNum

        for _ in range(numToDelete):
             newSheet.delete_rows(rowNum)

        # fillFormula(newSheet, rowNum, rateRangeDict)
        fillServicesRendered(newSheet, rowNum, posRateRangeDict)

    workbook.save(outputPath)
