
import re,math

info = {}
parts = {}
fastners = {}
roofpanelsnumber = []
roofpanelslength = []
cjcheighths = []
cjccount = []
########################################

from xlrd import open_workbook


def main_questionair():
    print('WELCOME TO THE DFASTENERS PROGRAM\n')
    print('Note: All firewalls will be treated as 3 hour firewalls. So adjust gypboard fasteners and compound accordingly ')
   # try:
    global b_loc
    global bldg
    path = "C:/Downloads/1.xls"#input("PLEASE COPY AND PAST THE FILE PATH TO THE TAKEOFF\n")
    bldg = "bldg 1"#input('WHAT IS THE NAME OF THE BUILDING ON THE SPREADSHEET?\n').lower()
    sqft = int(input('What is the SQFT of ' + bldg))
    info["sqft"] = sqft
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
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value
            for string in strings:
                #print (value)
                if string in value:
                    parts[key] += sheet.cell(r, b_loc).value


    except:
        pass
###########################################################



def PartsSearch():
    collection = [["downspouts",["TDS"]],["overflows",["T31283"]],["316Panels",["B4","BDS"]],
                  ["TopAngles",["S50131"]],["RoofPanelClosures",["R20005"]],["GutterClips",
                  ["T30003"]],["ExtHeaders",["HEAA"]],["cjc",["S20138"]],["RollUpDoors",["D7"]],
                  ["Gboard-625-10",["M60072"]],["Gboard-625-12",["M60073"]],["Sign",["M61308"]],
                  ["16WallColumn",["CEBA"]],["18WallColumn",["CEBB"]],["CornerColumn",["CEAA"]],
                  ["24CornerColumn",["CEBC"]],["InsideCorner",["CEAF"]],["22gaColumn",["CIAD"]],
                  ["7'ScrewGuard", ["S50180"]],["7'10ScrewGuard",["S50049Gl"]],["JambChanngel",["S46185"]],
                  ["20'gutter",["T30218"]],["20'3gutter",["T30219"]],["RidgeCap",["R30025GL"]], ["RoofClips",["R20059"]]]
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
            num = re.findall(r'\d{3}', value)
            num = num[0]
            if num:
                if num in range(80):
                    parts["5'nominal_base"] += sheet.cell(r, b_loc).value
                if num in range(80,140):
                    parts["10'nominal_base"] += sheet.cell(r, b_loc).value
                if num in range(140,200):
                    parts["15'nominal_base"] += sheet.cell(r, b_loc).value


    for r in range(sheet.nrows):

        cell = sheet.cell(r,0)
        value = cell.value
        if "B4" in value or "BDS" in value:
            num = re.findall(r'\d{3}', value)
            num = num[0]
            roofpanelslength.append(float(num)/12)
            roofpanelsnumber.append((sheet.cell(r, b_loc).value))


    for r in range(sheet.nrows):
        cell = sheet.cell(r,0)
        value = cell.value
        if "CJAA" in value:
            num = re.findall(r'\d{3}', value)
            num = num[0]
            cjcheighths.append(float(num)/12)
            cjccount.append((sheet.cell(r, b_loc).value) / 2)









#########################################################################################
def Add_Fastner(key,value):

    fastners[key] += value


def Fastner_Calcs():
    global SidewallLength
    if parts["RidgeCap"] == 0:
        SidewallLength = (int(parts["316Panels"]) * 16 / 12)
    else:
        SidewallLength = input(("What is the total length of (Just give the number in feet)", bldg,"\n"))


    #M20050
    fastners["M20050"] = 0
    M20050 = [((parts['downspouts'] + parts['overflows']) / 10),
            ((parts["7'10ScrewGuard"] + parts["7'ScrewGuard"]) / 8)]


    if parts["20'gutter"] == 0:
        M20050.append(SidewallLength * 2/40)

    for x in M20050:
        Add_Fastner("M20050",x)


    #########
    #M20055
    fastners["M20055"] = 0
    M20055 = [(SidewallLength * 2/ 40),
             ((parts["7'10ScrewGuard"] + parts["7'ScrewGuard"]) / 8 * 0.2),
              (SidewallLength * 2 / 30),
             (parts["RoofPanelClosures"] * 10 / 30)]
    if parts["20'gutter"] != 0:
        M20055.append(SidewallLength * 2/40)
    for x in M20055:
        Add_Fastner("M20055",x)

    #M30001
    fastners["M30001"] = 0
    PanelsxLength = 0
    for x in range(len(roofpanelslength)):
        PanelsxLength += roofpanelslength[x] * roofpanelsnumber[x]
    M30001 = [((PanelsxLength + (parts["7'10ScrewGuard"] + parts["7'ScrewGuard"]) * 5) / 50),
              (parts["RoofClips"] / 80),
              (parts["GutterClips"] / 150)]
    for x in M30001:
        Add_Fastner("M30001",x)



    #M30005
    fastners["M30005"] = 0
    for x in range(len(cjccount)):
        cjcheightxcjccount = 0
        cjcheightxcjccount += cjccount[x] * cjcheighths[x]
        print(parts["cjc"])
    M30005 = [parts["ExtHeaders"] * 4, (cjcheightxcjccount)]
    for x in M30005:
        Add_Fastner("M30005",x)


    #M30020
    fastners["M30020"] = 0
    if parts["20'gutter"] != 0:
       M30020 = [SidewallLength * 2 / 50,
                 (parts["RoofPanelClosures"] * 16 / 12 / 25)]
       for x in M30020:
           Add_Fastner("M30020",x)


    #M40030
    fastners["M40030"] = 0
    fastners["M6919"] = 0
    M40030 = [info["sqft"] / 10000]
    for x in M40030:
        Add_Fastner("M40030",x)
        Add_Fastner("M6919", x)

    #M60905
    fastners["M60905"] = 0
    M60905 = [parts["RollUpDoors"] / 50]
    for x in M60905:
        Add_Fastner("M60905",x)


    #M60936
    fastners["M60936"] = 0
    M60936 = [parts["RollUpDoors"] / 10]
    for x in M60936:
        Add_Fastner("M60936",x)


    #M61308
    fastners["M61308"] = 0
    Add_Fastner("M61308", 1)

    # M20053
    fastners["M20053"] = 0
    M20053 = [parts["Gboard-625-10"] / 24,
              parts["Gboard-625-12"] / 20]
    for x in M20053:
        Add_Fastner("M20053", x)

    # M20054
    fastners["M20054"] = 0
    M20054 = [parts["Gboard-625-10"] / 30,
              parts["Gboard-625-12"] / 30]
    for x in M20054:
        Add_Fastner("M20054", x)

    # M20054
    fastners["M20054"] = 0
    M20054 = [parts["Gboard-625-10"] / 30,
              parts["Gboard-625-12"] / 30]
    for x in M20054:
        Add_Fastner("M20054", x)


    # F88057
    fastners["F88057"] = 0
    F88057 = [parts["Gboard-625-10"] / 30,
              parts["Gboard-625-12"] / 30]
    for x in M20054:
        Add_Fastner("M20054", x)



    for key in list(fastners):
        fastners[key] = math.ceil(fastners[key])





main_questionair()
Fastner_Calcs()



print(fastners)





##########################NOTES################################
#FIGURE OUT 1 PER 40FT OF FLASHING/COUNTER FLASHING AT MASONRY M20055
#HANDLE M40040(get insulation part nums)
#do tri bead mastic,(get all part numbers associated with trap roof)