import xlrd
import sys


##########################################
#variables
global eheaders, downspouts, overflows, roof_panels, top_angle, roof_closure, roofclip, gutterclip
global cjc, rdoors, Gboard_625_12, Gboard_625_10
downspouts, overflows, roof_panels, top_angle, roof_closure, roofclip, gutterclip =0,0,0,0,0,0,0

eheaders, cjc, rdoors, Gboard_625_10, Gboard_625_12 = 0,0,0,0,0


global script
script = 0
global target
target = 0








from xlrd import open_workbook

global b_loc

def main_questionair():
    print('WELCOME TO THE DFASTENERS PROGRAM\n')

    print('Note: All firewalls will be treated as 3 hour firewalls. So adjust gypboard fasteners and compound accordingly ')
    try:
        global b_loc
        path = input("PLEASE COPY AND PAST THE FILE PATH TO THE TAKEOFF\n")
        bldg = input('WHAT IS THE NAME OF THE BUILDING ON THE SPREADSHEET?\n').lower()

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

    except:
        print("YOU TYPED THE FILE PATH INCORRECTLY \n")
        main_questionair()


def excell_pull_routine():


    try: #DOWNSPOUTS
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            TDS = "TDS"   #this is the parameter to search for downspouts
            value = cell.value
            #print (value)
            if "TDS" in value:
                global downspouts

                predownspouts = sheet.cell(r, b_loc)


                downspouts = predownspouts.value + downspouts
                int(downspouts)
                #print (downspouts)
        print (downspouts, "DOWNSPOUTS")
    except:
        pass


    try: #OVERFLOW ASSYMBLY'S
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value
            #print (value)
            if "T31283" in value:
                global overflows

                preoverflows = sheet.cell(r, b_loc)


                overflows = preoverflows.value + overflows
                #print (downspouts)
        print (overflows, "overflows")
    except:
        pass

    try: # 316 roof panels
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value
            #print (value)
            if "B4" in value or "BDS" in value:
                global roof_panels

                prepanels = sheet.cell(r, b_loc)

                roof_panels = prepanels.value + roof_panels

        print (roof_panels, "roof panels")
    except:
        pass

    try: # SA-3A TOP ANGLE
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value
            #print (value)
            if "S50131" in value:
                global top_angle

                preangle = sheet.cell(r, b_loc)

                top_angle = preangle.value + top_angle

        print (top_angle, "top angles")
    except:
        pass

    try:  # roof panel closures
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "R20005" in value:
                global roof_closure

                preclosure = sheet.cell(r, b_loc)

                roof_closure = preclosure.value + roof_closure

        print(roof_closure, "roof closures")
    except:
        pass

    try:  # roofclips
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "R20059" in value:
                global roofclip

                preclip = sheet.cell(r, b_loc)

                roofclip = preclip.value + roofclip

        print(roof_clip, "roof clips")
    except:
        pass

    try:  # gutter clips
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "T30003" in value:
                global gutterclip

                pgclip = sheet.cell(r, b_loc)

                gutterclip = pgclip.value + gutterclip

        print(gutterclip, "gutter clips")
    except:
        pass

    #try:  # extheaders
    for r in range(sheet.nrows):
        cell = sheet.cell(r, 0)
        value = cell.value

        if "HEAA" in value or "HEAB" in value:
            global eheaders

            #preeheaders = sheet.cell(r, b_loc)

            eheaders += sheet.cell(r, b_loc).value

    print(eheaders, "external headers")
    #except:
        #pass

    try:  # control join columns
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "S20138" in value:
                global cjc

                precjc = sheet.cell(r, b_loc)

                cjc = precjc.value + cjc

        print(cjc, "CJC's")
    except:
        pass

    try:  # our roll up doors
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "D7" in value:
                global rdoors

                prerdoors = sheet.cell(r, b_loc)

                rdoors = prerdoors.value + rdoors

        print(rdoors, "Roll up doors by us")
    except:
        pass


    try:  # gyp board 5/8" x 10
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "M60072" in value:
                global Gboard_625_10

                preGboard_625_10 = sheet.cell(r, b_loc)

                Gboard_625_10= preGboard_625_10.value + Gboard_625_10

        print (Gboard_625_10, '5/8)"x 10\'')
    except:
        pass

    try:  # gyp board 5/8" x 12
        for r in range(sheet.nrows):
            cell = sheet.cell(r, 0)
            value = cell.value

            if "M60073" in value:
                global Gboard_625_12

                preGboard_625_12 = sheet.cell(r, b_loc)

                Gboard_625_12 = preGboard_625_12.value + Gboard_625_12

        print (Gboard_625_12, '5/8)"x 12\'')
    except:
        pass


main_questionair()
