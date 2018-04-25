import xlrd
import sys


########################################

from xlrd import open_workbook

global b_loc

def main_questionair():
    print('WELCOME TO THE DFASTENERS PROGRAM\n')

    print('Note: All firewalls will be treated as 3 hour firewalls. So adjust gypboard fasteners and compound accordingly ')
   # try:
    global b_loc
    path = "C:\Downloads\\1.xls"#input("PLEASE COPY AND PAST THE FILE PATH TO THE TAKEOFF\n")
    bldg = "bldg 1"#input('WHAT IS THE NAME OF THE BUILDING ON THE SPREADSHEET?\n').lower()

    workbook = open_workbook(path)

    global sheet
    sheet = workbook.sheet_by_index(0)


    for c in range(sheet.ncols):
        global b_loc
        global target
        cell = sheet.cell(1, c)
        value = cell.value
        #print(value)

        if str(bldg) == str(value).lower():
            #print (c)
            b_loc = c
            excell_pull_routine()
            target = 1
    if target != 1:
        print("YOU SEEM TO HAVE TYPED THE NAME WRONG")
        main_questionair()

   # except:
     #   print("YOU TYPED THE FILE PATH INCORRECTLY \n")
      #  main_questionair()


def excell_pull_routine():
    parts = {}

#DOWNSPOUTS
    parts["downspouts"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "TDS" in value:
            parts["downspouts"] += sheet.cell(r, b_loc).value
    print (parts["downspouts"], "DOWNSPOUTS")

#OVERFLOWS
    parts["overflows"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "T31283" in value:
            parts["overflows"] += sheet.cell(r, b_loc).value
    print (parts["overflows"], "overflows")

#316 roof panels
    parts["316-panels"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "B4" in value or "BDS" in value:
            parts["316-panels"] += sheet.cell(r, b_loc).value
    print (parts["316-panels"], "316 roof panels")

# SA-3A top angle
    parts["topangle"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "S50131" in value:
            parts["topangle"] += sheet.cell(r, b_loc).value
    print (parts["topangle"], "Top angle")

# 316 closures
    parts["roof-panel-closure"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "R20005" in value:
            parts["roof-panel-closure"] += sheet.cell(r, b_loc).value

# gutter clips
    parts["gutterclips"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "T30003" in value:
            parts["gutterclips"] += sheet.cell(r, b_loc).value


# exterior headers
    parts["ext-headers"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "HEAA" in value:
            parts["ext-headers"] += sheet.cell(r, b_loc).value

# control joint columns
    parts["cjc"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "S20138" in value:
            parts["cjc"] += sheet.cell(r, b_loc).value

# our roll up doors
    parts["rudoors"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "D7" in value:
            parts["rudoors"] += sheet.cell(r, b_loc).value

# gyp board 5/8" x 10
    parts["gboard-625-10"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "M60072" in value:
            parts["gboard-625-10"] += sheet.cell(r, b_loc).value

# gyp board 5/8" x 12
    parts["gboard-625-12"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "M60073" in value:
            parts["gboard-625-12"] += sheet.cell(r, b_loc).value


# sign
    parts["sign"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "M61308" in value:
            parts["sign"] += sheet.cell(r, b_loc).value


# Trap-roof
    parts["TrapRoof"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "trap roof part number" in value:
            parts["TrapRoof"] += sheet.cell(r, b_loc).value

# 5' base angles
    parts["5'nominal_base"] = 0
    parts["10'nominal_base"] = 0
    parts["15'nominal_base"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "SBA" in value:
            print("here")
            if "0" in value:
                for x in range(10,80):
                    if str(x) in value:
                        parts["5'nominal_base"] += sheet.cell(r, b_loc).value

            for x in range(90,120):
                if str(x) in value:
                    parts["10'nominal_base"] += sheet.cell(r, b_loc).value

            for x in range(150,200):
                if str(x) in value:
                    parts["15'nominal_base"] += sheet.cell(r, b_loc).value


    print(parts["5'nominal_base"])
    print(parts["10'nominal_base"])
    print(parts["15'nominal_base"])

main_questionair()
