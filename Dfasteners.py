#<<<<<<< HEAD
import xlrd
import sys

parts = {}
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
            PartsSearch()
            target = 1
    if target != 1:
        print("YOU SEEM TO HAVE TYPED THE NAME WRONG")
        main_questionair()

   # except:
     #   print("YOU TYPED THE FILE PATH INCORRECTLY \n")
      #  main_questionair()

###########################################################
def excell_pull_routine(key,strings):
    try:
        parts[key] = 0
        for r in range(3,sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value
            for string in strings:
                #print (value)
                if string in value:
                    parts[key] += sheet.cell(r, b_loc).value
        print (parts[key], "")

    except:
        pass
###########################################################



def PartsSearch():
    collection = [["downspouts",["TDS"]],["overflows",["T31283"]],["316Panels",["B4","BDS"]],["TopAngles",["S50131"]],["RoofPanelClosures",["R20005"]],
                 ["GutterClips", ["T30003"]],["ExtHeaders",["HEAA"]],["cjc",["S20138"]],["RollUpDoors",["D7"]],["Gboard-625-10",["M60072"]],["Gboard-625-12",
                 ["M60073"]],["Sign",["M61308"]]]
    for arguments in collection:
        key, string = arguments
        excell_pull_routine(key,string)

#  base angles
    parts["5'nominal_base"] = 0
    parts["10'nominal_base"] = 0
    parts["15'nominal_base"] = 0
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value
        #print (value)
        if "SBA" in value:
            if "0" in value:
                for x in range(20,80):
                    if str(x) in value:
                        parts["5'nominal_base"] += sheet.cell(r, b_loc).value

            for x in range(90,120):
                if str(x) in value:
                    parts["10'nominal_base"] += sheet.cell(r, b_loc).value

            for x in range(150,199):
                if str(x) in value:
                    parts["15'nominal_base"] += sheet.cell(r, b_loc).value


    print(parts["5'nominal_base"])
    print(parts["10'nominal_base"])
    print(parts["15'nominal_base"])

main_questionair()