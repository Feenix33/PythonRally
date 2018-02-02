"""
lrnXL Excel experiments tied to Test1.xlsx
01 - Read the sheet and convert to list
02 - See if we can change the value
03 - Iterate through ws and convert to list
04 - Sum a column with filtering
05 - Sum a column by team names using dictionary
06 - Summarize an iteration output
07 - Add Summary sheet to existing XL (at least 27 times ;-)
08 - Write summary (Combine 06 and 07) but using helper functions
"""

from openpyxl import Workbook
from openpyxl import load_workbook
import pprint
import datetime

#fname = "Test.xlsx"
#fname = "Test1.xlsx"
#fname = "Iter.xlsx"
fname = "IterA.xlsx"

attribs = ['NumStories', 'TotalPoints', 'NumDefects','ZeroStories']
############
gFieldnames = ['FormattedID', 'Name', 'PlanEstimate', 'TeamFeature.FormattedID', 
        'TeamFeature.Name', 'ScheduleState', 'Project.Name', 'Iteration.Name', 
        'Owner.Name', 'CreationDate']
data = []
anly = {}

def readWSData(ws):
    ldata = []
    for row in ws.values:
        ldata.append(list(row))
    return ldata

def analyzeData(pastDat):
    colTeam = pastDat[0].index("Project.Name")
    teamNames = list(set(row[colTeam] for row in pastDat[1:]))

    empty = dict.fromkeys(attribs, 0)
    lanly = dict.fromkeys(teamNames, None)
    for key in lanly.keys():
        lanly[key] = empty.copy()

    #
    colID = pastDat[0].index('FormattedID')
    colPts = pastDat[0].index('PlanEstimate')
    for datrow in pastDat[1:]:
        if datrow[colID][0] == "U":
            lanly[datrow[colTeam]]['NumStories'] += 1
            try:
                lanly[datrow[colTeam]]['TotalPoints'] += datrow[colPts]
            except TypeError:
                ptErrors += 1
                print("PtError ", ptErrors, datrow[colPts])
            if datrow[colPts] == 0:
                lanly[datrow[colTeam]]['ZeroStories'] += 1
        elif datrow[colID][0] == "D":
            lanly[datrow[colTeam]]['NumDefects'] += 1
    return lanly

def writeAnalysis(wBk, aDict):
    # create the sheet
    today = datetime.date.today()
    wsNameBase = ("Summary {:04d}.{:02d}.{:02d}").format(today.year, today.month, today.day)
    suffix = 'A'
    wsName = wsNameBase
    print (wBk.sheetnames)
    while wsName in wBk.sheetnames:
        wsName = wsNameBase + suffix
        suffix = chr(ord(suffix)+1)
    ws = wBk.create_sheet(wsName)

    ws.append( ["Project.Name"] + attribs )
    for key in aDict.keys():
        #ws.append([key] + [aDict[key][bkey] for bkey in aDict[key].keys()])
        # do it this way, not sure if above retains order
        ws.append([key] + [aDict[key][atr] for atr in attribs])

def Test08(wb):
    print ("Test 08 ", "-"*40)
    ws = wb.active #assume data on first sheet
    pp = pprint.PrettyPrinter(indent=4)
    data08 = readWSData(ws)

    #print ("Test 08 Input Data", "-"*40)
    #pp.pprint (data08)
    anly08 = analyzeData(data08)
    writeAnalysis(wb, anly08)
    print ("Test 08 Analyzed Data", "-"*40)
    #pp.pprint(anly08)
    """
    """

def Test07(wb):
    # chr(ord('A')+1)
    today = datetime.date.today()
    wsNameBase = ("Summary {:04d}.{:02d}.{:02d}").format(today.year, today.month, today.day)
    suffix = 'A'
    wsName = wsNameBase
    print (wb.sheetnames)
    while wsName in wb.sheetnames:
        wsName = wsNameBase + suffix
        suffix = chr(ord(suffix)+1)
    ws = wb.create_sheet(wsName)
    print (wb.sheetnames)
    wb.save(fname)

def Test06():
    ptErrors = 0
    print ("Test 06 ", "-"*40)
    pp = pprint.PrettyPrinter(indent=4)

    colTeam = data[0].index("Project.Name")
    teamNames = list(set(row[colTeam] for row in data[1:]))

    empty = dict.fromkeys(attribs, 0)
    anly = dict.fromkeys(teamNames, None)
    for key in anly.keys():
        anly[key] = empty.copy()

    #for attrib in attribs:
    #    print (attrib, anly['Sabre Cruises'][attrib])
    colID = data[0].index('FormattedID')
    colPts = data[0].index('PlanEstimate')
    for datrow in data[1:]:
        if datrow[colID][0] == "U":
            anly[datrow[colTeam]]['NumStories'] += 1
            try:
                anly[datrow[colTeam]]['TotalPoints'] += datrow[colPts]
            except TypeError:
                ptErrors += 1
                print("PtError ", ptErrors, datrow[colPts])
            if datrow[colPts] == 0:
                anly[datrow[colTeam]]['ZeroStories'] += 1
        elif datrow[colID][0] == "D":
            anly[datrow[colTeam]]['NumDefects'] += 1
    pp.pprint(anly)
    print("SPt errors = ", ptErrors)


def Test04():
    print ("Test 04 ", "-"*40)
    colTeam = data[0].index("Team")
    value = sum(row[0] for row in data[1:] if row[colTeam] == "Red" )
    print (value)

def Test05():
    print ("Test 05 ", "-"*40)
    #for row in data: print(row[0])
    colTeam = data[0].index("Team")
    teamNames = list(set(row[colTeam] for row in data[1:]))
    print (teamNames)
    value = dict.fromkeys(teamNames, 0)
    print (value)

    for row in data[1:]:
        value[row[colTeam]] += row[0]
    print (value)

def main():
    wb = load_workbook(fname)
    #ws = wb.active
    #
    #for row in ws.values:
    #    data.append(list(row))

    Test08(wb)
    #Test07(wb)
    #Test06()
    #Test05()
    #Test04()
    wb.save(fname)

if __name__ == '__main__':
    main()
    print ("Fini ", "-"*40)
