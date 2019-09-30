#!python3
# HangerCalculator.py
# Calculates hanger data and prepares label for mail merge

import openpyxl, pprint, time, math
from openpyxl.styles import Alignment
from fractions import Fraction

# This function is needed to sum allthread and strut mixed number strings to floats
def convert_to_float(frac_str):
    try:
        return float(frac_str)
    except ValueError:
        num, denom = frac_str.split('/')
        try:
            leading, num = num.split(' ')
            whole = float(leading)
        except ValueError:
            whole = 0
        frac = float(num) / float(denom)
        return whole - frac if whole < 0 else whole + frac

def roundUp(x, decimal = 0):
    multiplier = 10 ** decimal
    return math.ceil(x * multiplier) / multiplier

while True:
    try:
        print("""\nReadMe:\nColumn A: Level\nColumn B: Area ID\nColumn C: Family and Type\nColumn D: Material\nColumn E: Hanger ID\n
Column F: Support Span\nColumn G: Support 1 Cut Length\nColumn H: Attachment 1 Elevation\nColumn I: Attachment 1 Facing down\n
Column J: Attachment 2 Elevation\nColumn K: Attachment 2 Facing down\nColumn L: Attachment 3 Elevaton\nColumn M: Attachment 3 Facing down\n
Column N: Distance Between Tier 1 & 2\nColumn O: Distance between Tier 2 & 3\n\n""")
        excelSheetName = input('Enter Name of Excel Sheet: ')
        jobNumber = input('Enter the job number: ')
        jobName = input('Enter the job name: ')
        wb = openpyxl.load_workbook(excelSheetName + '.xlsx')
        sheet = wb.get_sheet_by_name(excelSheetName)
        break
    except:
        print("""\n--------------------------------------\n\nSomething went wrong.\nCheck workbook spelling. Your input is case sensitive.
\nEnsure first sheet is the same name as the workbook.\nEnsure this program is in the same folder as the workbook.
\n\n--------------------------------------""")


AssemblyListsSheet = wb.create_sheet()
AssemblyListsSheet.title = 'Assembly Lists'
createdConcatSheet = wb.create_sheet()
createdConcatSheet.title = 'Print Me'

hangerListTotal = []
hangerList = []
strutLengthList = []
strutLengthListTotal = []
strutTypeList = []
strutTypeListTotal = []
allthreadList = []
allthreadListTotal = []
allthreadLengthList = []
strutTypeLengthList = []
strutNameList = []
deepStrut = []
deepStrutLength = []
B2BdeepStrut = []
otherMaterial = []
B2BdeepStrutLength = []
shallowStrut = []
shallowStrutLength = []
B2BshallowStrut = []
B2BshallowStrutLength = []
singleTierDeep = []
doubleTierDeep = []
tripleTierDeep = []
singleTierB2BDeep = []
doubleTierB2BDeep = []
tripleTierB2BDeep = []
singleTierShallow = []
doubleTierShallow = []
tripleTierShallow = []
singleTierB2BShallow = []
doubleTierB2BShallow = []
tripleTierB2BShallow = []

next_strut_length_row = 4
next_assembly_row = 4
next_allthread_row = 4
next_concat_row = 2

# below creates title blocks for each created sheet

AssemblyListsSheet.merge_cells('A1:B2')
title = AssemblyListsSheet.cell(row = 1, column = 1)
title.value = 'Total Strut'
title.alignment = Alignment(horizontal = 'center', vertical = 'center')
AssemblyListsSheet.cell(row = 3, column = 1).value = "Strut Type"
AssemblyListsSheet.cell(row = 3, column = 2).value = "Quantity"

AssemblyListsSheet.merge_cells('D1:E2')
title = AssemblyListsSheet.cell(row = 1, column = 4)
title.value = 'Total Allthread'
title.alignment = Alignment(horizontal = 'center', vertical = 'center')
AssemblyListsSheet.cell(row = 3, column = 4).value = "Allthread Length"
AssemblyListsSheet.cell(row = 3, column = 5).value = "Quantity"

AssemblyListsSheet.merge_cells('G1:H2')
title = AssemblyListsSheet.cell(row = 1, column = 7)
title.value = 'Assembly Name and Quantity'
title.alignment = Alignment(horizontal = 'center', vertical = 'center')
AssemblyListsSheet.cell(row = 3, column = 7).value = "Assembly Name"
AssemblyListsSheet.cell(row = 3, column = 8).value = "Quantity"

createdConcatSheet.cell(row = 1, column = 1).value = "PRINT_ME"

# Below generates 2 lists of values
for row in range(4, sheet.max_row+1):
    try:

       strutCut = sheet{'E' + str(row).value + sheet{'Material': 'D' + str(row).value + sheet{'Cut Length': 'F' + str(row)}}.value

       level = sheet['A' + str(row)].value

       area = sheet['B' + str(row)].value

       material = sheet['D' + str(row)].value
       

       hangerID = sheet['E' + str(row)].value
       hangerList.append(hangerID)  # Used for counting the total unique hanger ID's. All hanger ID's are in this list
       
       support_span = sheet['F' + str(row)].value

       strutType = sheet['D' + str(row)].value + ': ' + sheet['F' + str(row)].value
       strutTypeList.append(strutType)  # Used for tracking each instance of strutType. This is the list that is counted against for every time it appears.
       if 'DEEP' in strutType and 'B2B' not in strutType:
           deepStrut.append(strutType)
       elif 'DEEP' in strutType and 'B2B' in strutType:
           B2BdeepStrut.append(strutType)
       elif 'SHALLOW' in strutType and 'B2B' not in strutType:
           shallowStrut.append(strutType)
       elif 'SHALLOW' in strutType and 'B2B' in strutType:
           B2BshallowStrut.append(strutType)
       else:
           otherMaterial.append(strutTupe)
       
       allthreadLength = sheet['G' + str(row)].value
       allthreadList.append(allthreadLength)

       attach1_elev = sheet['H' + str(row)].value

       attach1_face = sheet['I' + str(row)].value

       attach2_elev = sheet['J' + str(row)].value

       attach2_face = sheet['K' + str(row)].value

       attach3_elev = sheet['L' + str(row)].value

       attach3_face = sheet['M' + str(row)].value

       dist1_2 = sheet['N' + str(row)].value

       dist2_3 = sheet['O' + str(row)].value

                
       # althreadListTotal calculates the total unique allthread cuts. Only 1 unique value in this list
       if allthreadLength not in allthreadListTotal:
           allthreadListTotal.append(allthreadLength)

       # strutTypeListTotal calculates the total unique strut cuts. Only 1 unique value in this list
       if strutType not in strutTypeListTotal:
           strutTypeListTotal.append(strutType)

       # hangerListTotal calculates the total unique hanger ID's. Only 1 unique value in this list
       if hangerID not in hangerListTotal:
           hangerListTotal.append(hangerID)
    except:

        continue

    # Below creates the concatenation to be used in mail merge
    label = jobNumber + "-" + jobName + '  ' + level + '         ' + area + " Tag: " + hangerID + "                                      TOU: " + attach1_elev + "                                                                    " + strutType + "                                                               Allthread Length: " + allthreadLength
    createdConcatSheet.cell(column = 1, row = next_concat_row, value = label)
    next_concat_row += 1

    if attach2_elev:
        label = jobNumber + "-" + jobName + '  ' + level + '         ' + area + " Tag: " + hangerID + "                                 TOU: " + attach2_elev + "                                                                    " + strutType + "  TIER #2"
        createdConcatSheet.cell(column = 1, row = next_concat_row, value = label)
        next_concat_row += 1
    
    if attach3_elev:
        label = jobNumber + "-" + jobName + '  ' + level + '         ' + area + " Tag: " + hangerID + "                                 TOU: " + attach3_elev + "                                                                    " + strutType + "  TIER #3"
        createdConcatSheet.cell(column = 1, row = next_concat_row, value = label)
        next_concat_row += 1
    

allthreadListTotal.sort()
strutTypeListTotal.sort()
hangerListTotal.sort()

# Below prints total counts on each created sheet
for x in hangerListTotal:
    AssemblyListsSheet.cell(column = 7, row = next_assembly_row, value = x)
    AssemblyListsSheet.cell(column = 8, row = next_assembly_row, value = hangerList.count(x))
    next_assembly_row += 1

for x in strutTypeListTotal:    # Contains each strut type once
    AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = x)
    AssemblyListsSheet.cell(column = 2, row = next_strut_length_row, value = strutTypeList.count(x))
    next_strut_length_row += 1

for x in allthreadListTotal:
    AssemblyListsSheet.cell(column = 4, row = next_allthread_row, value = x)
    AssemblyListsSheet.cell(column = 5, row = next_allthread_row, value = (allthreadList.count(x) * 2))
    next_allthread_row += 1

# Creating sum of allthread length
for x in allthreadList:
    x = x.replace('"','')
    x = convert_to_float(x)
    allthreadLengthList.append(x)

# Creating sum of strut length
for x in strutTypeList:
    x = x.split(': ')
    x0 = x[0]
    x1 = x[1]
    x1 = x1.replace('"','')
    x1 = convert_to_float(x1)
    if '(2)' in x0:
        strutTypeLengthList.append((2 * x1))
    elif '(3)' in x0:
        strutTypeLengthList.append((3 * x1))
    else:
        strutTypeLengthList.append(x1)

for x in deepStrut:
    x = x.split(': ')
    x0 = x[0]
    x1 = x[1]
    x1 = x1.replace('"','')
    x1 = convert_to_float(x1)
    if '(2)' in x0:
        deepStrutLength.append((2 * x1))
        doubleTierDeep.append(x0)
    elif '(3)' in x0:
        deepStrutLength.append((3 * x1))
        tripleTierDeep.append(x0)
    else:
        deepStrutLength.append(x1)
        singleTierDeep.append(x0)

for x in B2BdeepStrut:
    x = x.split(': ')
    x0 = x[0]
    x1 = x[1]
    x1 = x1.replace('"','')
    x1 = convert_to_float(x1)
    if '(2)' in x0:
        B2BdeepStrutLength.append((2 * x1))
        doubleTierB2BdeepStrut.append(x0)
    elif '(3)' in x0:
        B2BdeepStrutLength.append((3 * x1))
        tripleTierB2BDeep.append(x0)
    else:
        B2BdeepStrutLength.append(x1)
        singleTierB2BDeep.append(x0)

for x in shallowStrut:
    x = x.split(': ')
    x0 = x[0]
    x1 = x[1]
    x1 = x1.replace('"','')
    x1 = convert_to_float(x1)
    if '(2)' in x0:
        shallowStrutLength.append((2 * x1))
        doubleTierShallow.append(x0)
    elif '(3)' in x0:
        shallowStrutLength.append((3 * x1))
        tripleTierShallow.append(x0)
    else:
        shallowStrutLength.append(x1)
        singleTierShallow.append(x0)

for x in B2BshallowStrut:
    x = x.split(': ')
    x0 = x[0]
    x1 = x[1]
    x1 = x1.replace('"','')
    x1 = convert_to_float(x1)
    if '(2)' in x0:
        B2BshallowStrutLength.append((2 * x1))
        doubleTierB2BShallow.append(x0)
    elif '(3)' in x0:
        B2BshallowStrutLength.append((3 * x1))
        tripleTierB2BShallow.append(x0)
    else:
        B2BshallowStrutLength.append(x1)
        singleTierB2BShallow.append(x0)

# Summing lists and printing in feet
totalAllthreadLength = sum(allthreadLengthList) * 2
totalAlltrheadLengthFeet = totalAllthreadLength / 12
totalStrutLength = sum(strutTypeLengthList)
totaldeepStrutLength = sum(deepStrutLength)
totalB2BdeepStrutLength = sum(B2BdeepStrutLength)
totalshallowStrutLength = sum(shallowStrutLength)
totalB2BshallowStrutLength = sum(B2BshallowStrutLength)
totalStrutLengthFeet = totalStrutLength / 12
totaldeepStrutLengthFeet = totaldeepStrutLength / 12
totalB2BdeepStrutLengthFeet = totalB2BdeepStrutLength / 12
totalshallowStrutLengthFeet = totalshallowStrutLength / 12
totalB2BshallowStrutLengthFeet = totalB2BshallowStrutLength / 12
totalAlltrheadLengthFeet = roundUp(totalAlltrheadLengthFeet, -1)
totalStrutLengthFeet = roundUp(totalStrutLengthFeet, -1)
totaldeepStrutLengthFeet = roundUp(totaldeepStrutLengthFeet, -1)
totalB2BdeepStrutLengthFeet = roundUp(totalB2BdeepStrutLengthFeet, -1)
totalshallowStrutLengthFeet = roundUp(totalshallowStrutLengthFeet, -1)
totalB2BshallowStrutLengthFeet = roundUp(totalB2BshallowStrutLengthFeet, -1)


AssemblyListsSheet.cell(column = 4, row = next_allthread_row, value = "Total allthread length = " + str(totalAlltrheadLengthFeet) + " ft")
AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total strut length = " + str(totalStrutLengthFeet) + " ft")
next_strut_length_row +=2

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  Deep strut length = " + str(totaldeepStrutLengthFeet) + " ft.")
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  B2B Deep strut length = " + str(totalB2BdeepStrutLengthFeet) + " ft.")
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  Shallow strut length = " + str(totalshallowStrutLengthFeet) + " ft.")
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  B2B Shallow strut length = " + str(totalB2BshallowStrutLengthFeet) + " ft.")
next_strut_length_row +=2

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  single tier Deep Strut racks = "  + str(len(singleTierDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  single tier Back to Back Deep Strut racks = "  + str(len(singleTierB2BDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  single tier Shallow Strut racks = "  + str(len(singleTierShallow)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  single tier Back to Back Shallow Strut racks = "  + str(len(singleTierB2BShallow)))
next_strut_length_row +=2

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  double tier Deep Strut racks = "  + str(len(doubleTierDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  double tier Back to Back Deep Strut racks = "  + str(len(doubleTierB2BDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  double tier Shallow Strut racks = "  + str(len(doubleTierShallow)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  double tier Back to Back Shallow Strut racks = "  + str(len(doubleTierB2BShallow)))
next_strut_length_row +=2

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  triple tier Deep Strut racks = "  + str(len(tripleTierDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  triple tier Back to Back Deep Strut racks = "  + str(len(tripleTierB2BDeep)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  triple tier Shallow Strut racks = "  + str(len(tripleTierShallow)))
next_strut_length_row +=1

AssemblyListsSheet.cell(column = 1, row = next_strut_length_row, value = "Total  triple tier Back to Back Shallow Strut racks = "  + str(len(tripleTierB2BShallow)))

wb.save(excelSheetName + '.xlsx')
print('------------------------Done------------------------')
time.sleep(3)

# Needs total quantities for the different type of strut racks to enter into the pdf calculator
# Need to output label values into new cells on the row for a better future label Word template. (Top cells will need headers)
# Need: To catch and report if any struts are not on an even 2" cut length. (Undecided if this calc should auto round to the nearest 2" cut length or not.
#       As we still want the revit model to be accurate.