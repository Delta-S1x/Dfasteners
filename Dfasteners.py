
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
    path = "/root/1.xls"#input("PLEASE COPY AND PAST THE FILE PATH TO THE TAKEOFF\n")
    bldg = "bldg 1"#input('WHAT IS THE NAME OF THE BUILDING ON THE SPREADSHEET?\n').lower()
    sqft = int(input('What is the SQFT of ' + bldg + "\n"))
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
                  ["T30003"]],["ExtHeaders",["HEAA"]],["cjc",["S20139"]],["RollUpDoors",["D7"]],
                  ["Gboard-625-10",["M60072"]],["Gboard-625-12",["M60073"]],["Sign",["M61308"]],
                  ["16WallColumn",["CEBA"]],["18WallColumn",["CEBB"]],["CornerColumn",["CEAA"]],
                  ["24WallColumn",["CEBC"]],["InsideCorner",["CEAF"]],["22gaColumn",["CIAD"]],
                  ["7'ScrewGuard", ["S50180"]],["7'10ScrewGuard",["S50049Gl"]],["JambChanngel",["S46185"]],
                  ["20'gutter",["T30218"]],["20'3gutter",["T30219"]],["RidgeCap",["R30025GL"]], ["RoofClips",["R20059"]],
                  ["StandOffRoofClips", ["R20061"]],["7'10-236panels", ["A6093M"]],["7'0-236panels", ["A6093M"]],["24Jam", ["DJBA096XBW"]],
                  ["12Jam", ["DJBB096XBW"]],["10Jam", ["DJBC096XBW"]], ["8Jam", ["DJBJ096XBW"]], ["8Jam", ["DJBJ096XBW"]],
                  ["Studtrack", ["S362T12543120X3"]],["MasonryFlashing1", ["T4065"]],["FirewallTrimAngle", ["S50024"]],
                  ["RmaxBoard", ["M40263B"]],["1.5RmaxBoard", ["M40264B"]],["2RmaxBoard", ["M40265B"]],["ClipAngle", ["S64508","S66726"]],
                  ["OutsideAngle", ["T50149"]]
                ]



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

##Roof panels, and roof panel lengths
    for r in range(sheet.nrows):
        cell = sheet.cell(r,0)
        value = cell.value
        if "B4" in value or "BDS" in value:
            num = re.findall(r'\d{3}', value)
            num = num[0]
            roofpanelslength.append(float(num)/12)
            roofpanelsnumber.append((sheet.cell(r, b_loc).value))

##CJC amount of lengths
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


    # F88057
    vc = input("Number of Vertical Cee's at Masonry wall?\n")

    fastners["F88057"] = 0
    F88057 = [parts["5'nominal_base"] * 3,
              parts["10'nominal_base"] * 4,
              parts["15'nominal_base"] * 5,
              parts["16WallColumn"] * 3,
              parts["18WallColumn"] * 3,
              parts["CornerColumn"] * 3,
              parts["cjc"] * 6,
              parts["24WallColumn"] * 4,
              parts["InsideCorner"] * 6,
              parts["22gaColumn"] * 2,
              parts["7'ScrewGuard"] / 2,        ###this will get the same results as 4 per 10' wall panel and 3 per 5' wall panel
              parts["7'10ScrewGuard"] / 2,
              parts["7'10-236panels"],
              parts["7'0-236panels"],
              parts["24Jam"] * 4,                ################# Need to ad 3 per corner column extension
              (parts["12Jam"] + parts["10Jam"] + parts["8Jam"]) * 2,
              parts["Studtrack"] / 1.125 / 2 * 6,
              parts["MasonryFlashing1"] *10,
              int(vc) * 4]

    for x in F88057:
        Add_Fastner("F88057", x)

    # F10039
    vc = input("Number of Vertical Cee's at Gyp Board wall?\n")
    fastners["F10039"] = 0
    F10039 = [parts["FirewallTrimAngle"] * 10,
              int(vc) * 4]
    for x in F10039:
        Add_Fastner("F10039", x)


    # F10020
    fastners["F10020"] = 0
    F10020 = [parts["RmaxBoard"] * 15]
    for x in F10020:
        Add_Fastner("F10020", x)


    # F10017
    fastners["F10017"] = 0
    F10017 = [parts["1.5RmaxBoard"] * 15]
    for x in F10017:
        Add_Fastner("F10017", x)






    # F10028
    fastners["F10028"] = 0
    F10028 = [parts["StandOffRoofClips"] * 3]
    for x in F10028:
        Add_Fastner("F10028", x)



    # F10008
    fastners["F10008"] = 0
    F10008 = [parts["RoofClips"] * 3,
              parts["ClipAngle"] * 3,
              parts["OutsideAngle"] * 7]

    for x in F10008:
        Add_Fastner("F10008", x)






main_questionair()
Fastner_Calcs()



print(fastners)





##########################NOTES################################
# FIGURE OUT 1 PER 40FT OF FLASHING/COUNTER FLASHING AT MASONRY M20055
# HANDLE M40040(get insulation part nums)
# do tri bead mastic,(get all part numbers associated with trap roof)
#F10017 what 1 1/2" insulation