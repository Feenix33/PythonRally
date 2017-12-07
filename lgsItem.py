from __future__ import print_function
import os
import sys
import csv
import getopt
import datetime
import argparse

from pyral import Rally, rallySettings, rallyWorkset
from openpyxl import Workbook

"""
---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
.Process FEA TF US
.Output the hierarchy
.User Story and defect processing
.Generic query
.Query from input file
.Input file with ignore
.Restructure output to have separate modules to write rather than as you go
.Add excel output
.what globals should be global
.should globals be in a structure
.excel page per FEA
.Consolidate query routines

 Get tasks if asked for user stories and defects
 Commnand line switches
 ini/yaml config file
---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
.input file
.screen status or silent
.excel or csv
.output filename
.not found string
 tasks or not for all or only explicit us/de
---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
"""


gFieldnames = ['FormattedID', 'Name', 'PlanEstimate', 'TeamFeature.FormattedID', 'TeamFeature.Name',
                        'ScheduleState', 'Project.Name', 'Iteration.Name', 'Owner.Name', 'CreationDate']
gPPMFieldnames = ['FormattedID', 'Name', 'LeafStoryPlanEstimateTotal', 'Parent.FormattedID', 'Parent.Name', 
                        'LeafStoryCount', 'AcceptedLeafStoryPlanEstimateTotal', 'AcceptedLeafStoryCount', 'Owner.Name', 'CreationDate']



class theConfig:
    def __init__(self):
        self.outToConsole = False
        self.writeCSV = True
        self.writeExcel = True
        self.fnameCSV = "lgsItem.csv"
        self.fnameExcel = "lgsItem.xlsx"
        self.fnameExceloutType = 0
        self.fnameInput = 'in.txt'
        self.strNoExist = "--"


gConfig = theConfig()

def processRecord(record):
    return  [getattrError(record, attr) for attr in gFieldnames] 

 
def getattrError(n, attr):
    try:
        return getattr(n, attr)
    except AttributeError:
        dot = attr.find(".")
        if dot == -1 :
            return gConfig.strNoExist
        try:
            return getattr( getattr(n, attr[:dot]), attr[(dot+1):])
        except AttributeError:
            return gConfig.strNoExist

def processFEA(instRally, feaID):
    entityName = 'PortfolioItem'
    queryString = 'FormattedID = "' + feaID + '"'
    response = instRally.get(entityName, fetch=True, projectScopeUp=True, query=queryString)
    records = []
    for ppmFeature in response:
        records.append( [getattrError(ppmFeature, attr) for attr in gPPMFieldnames] )
        if ppmFeature.DirectChildrenCount > 0 and hasattr(ppmFeature, 'Children'):
            for teamfeature in ppmFeature.Children:
                # add the team feature to the data table
                records.append( [getattrError(teamfeature, attr) for attr in gPPMFieldnames] )
                # add the subsequent children
                records.extend( processStory(instRally, teamfeature.FormattedID))
    return records 

def processStory(instRally, storyID): # story is user story or defect or TF
    records = []
    if storyID[0] == "U":
        entity_name = 'HierarchicalRequirement'
        queryString = 'FormattedID = "' + storyID + '"'
    elif storyID[0] == "D":
        entity_name = 'Defect'
        queryString = 'FormattedID = "' + storyID + '"'
    else:
        entity_name = 'HierarchicalRequirement'
        queryString = 'Feature.FormattedID = "' + storyID + '"'
    response = instRally.get(entity_name, fetch=True, projectScopeDown=True, query=queryString)
    consoleStatus ("   " + queryString + "    Count = " + str(response.resultCount))
    for story in response:
        records.append( processRecord( story ))
    return records


def tokens(fileobj):
    for line in fileobj:
        if line[0] != '#':
            for word in line.split():
                yield word


def writeCSV(listData, fname):
    fout = open(fname, 'wb')
    fwriter = csv.writer(fout, dialect='excel')
    fwriter.writerow( gFieldnames )
    fwriter.writerows(listData)
    fout.close()


def writeExcel(listData, fname, outType=None):
    consoleStatus("Writing Excel")
    wb = Workbook()
    ws = wb.active
    ws.title = "Rally Output"
    #wsB = wb.create_sheet("Bravo")

    if not outType or outType <= 0 or outType > 2:
        ws.append(gFieldnames)
        for row in listData:
            ws.append(row)
    elif outType == 1:
        ws.title = "User Stories"
        wsPPM = wb.create_sheet("PPM Items")
        ws.append(gFieldnames)
        wsPPM.append(gPPMFieldnames)
        for row in listData:
            if row[0][0] == "U" or row[0][0] == "D":
                ws.append(row)
            else:
                wsPPM.append(row)
    else: #outType == 2: new sheet for each FEA encountered
        ws.title = "User Stories"
        ws.append(gFieldnames)
        wsUsing = ws
        for row in listData:
            if row[0][0] == "F":
                wsUsing = wb.create_sheet(row[0])
                wsUsing.append(gFieldnames)
            wsUsing.append(row)

    wb.save(fname)

def consoleStatus(message):
    if gConfig.outToConsole: print (message)

def buildInputParser(parser):
    #parser.add_argument("-v", "--verbose", action="store_true", help="Print status messages to console")
    parser.add_argument("infile", help="Input file to process")

def mapArgsserToGlobal(args):
    #gConfig.outToConsole = args.verbose
    gConfig.fnameInput = arg.infile

def main():
    inParser = argparse.ArgumentParser()
    buildInputParser(inParser)
    args = inParser.parse_args()
    mapArgsserToGlobal(args)

    dataHR = []
    queryList = []

    #try:
    #    finputfile = sys.argv[1]
    #except IndexError:
    #    finputfile = gConfig.fnameInput
    #    consoleStatus ("Trying to use " + str(finputfile))

    finputfile = gConfig.fnameInput
    try:
        infile = open(finputfile, 'r')
        queryList = tokens(infile)


        consoleStatus('Logging in...')
        rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

        consoleStatus('Query execution...')

        for queryItem in queryList:
            if queryItem[:2] == "FE":
                dataHR.extend (processFEA(rally, queryItem))
            elif queryItem[:2] == "TF":
                dataHR.extend (processStory(rally, queryItem))
            elif queryItem[:2] == "US":
                #dataHR.extend (processUS(rally, queryItem))
                dataHR.extend (processStory(rally, queryItem))
            elif queryItem[:2] == "DE":
                #dataHR.extend (processDE(rally, queryItem))
                dataHR.extend (processStory(rally, queryItem))
            else:
                print ("Error query for " + queryItem)

        if gConfig.writeCSV: writeCSV(dataHR, gConfig.fnameCSV)
        if gConfig.writeExcel: writeExcel(dataHR, gConfig.fnameExcel, outType=gConfig.fnameExceloutType)

    except IOError:
        print ("IOError: Cannot open", sys.argv[0], "<input file>")

    consoleStatus('Fini')



if __name__ == '__main__':
    execfile("apifig.py")
    main()
