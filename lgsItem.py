from __future__ import print_function
import os
import sys
import csv
import getopt
import datetime
from pyral import Rally, rallySettings, rallyWorkset
from openpyxl import Workbook
"""
xProcess FEA TF US
xOutput the hierarchy
xUser Story and defect processing
xGeneric query
xQuery from input file
xInput file with ignore
Restructure output to have separate modules to write rather than as you go
Get tasks if asked for user stories and defects
Add excel output
excel page per FEA
Add file input
Commnand line switches
ini/yaml config file
----
input file
screen status or silent
excel or csv
output filename
not found string
tasks or not for all or only explicit us/de
"""


gQueryList = []

gFilename = 'lgsout.csv'
gFieldnames = ['FormattedID', 'Name', 'PlanEstimate', 'TeamFeature.FormattedID', 'TeamFeature.Name',
                        'ScheduleState', 'Project.Name', 'Iteration.Name', 'Owner.Name', 'CreationDate']
gPPMFieldnames = ['FormattedID', 'Name', 'LeafStoryPlanEstimateTotal', 'Parent.FormattedID', 'Parent.Name', 
                        'LeafStoryCount', 'AcceptedLeafStoryPlanEstimateTotal', 'AcceptedLeafStoryCount', 'Owner.Name', 'CreationDate']

gWriter = None
strNoExist = "--"
featureHierarchy = [gPPMFieldnames]

def openCSV(fname):
    fout = open(fname, 'wb')
    writer = csv.writer(fout, dialect='excel')
    return writer

def NEWwriteHeader(writer, worksheet):
    if writer: writer.writerow(gFieldnames)
    if worksheet: worksheet.append(gFieldnames)

def writeHeader(writer):
    writer.writerow(gFieldnames)


def processUserStory(writer, record):
    writer.writerow ( [getattrError(record, attr) for attr in gFieldnames] )

 
def getattrError(n, attr):
    try:
        return getattr(n, attr)
    except AttributeError:
        dot = attr.find(".")
        if dot == -1 :
            return strNoExist
        try:
            return getattr( getattr(n, attr[:dot]), attr[(dot+1):])
        except AttributeError:
            return strNoExist

def processFEA(instRally, feaID, writer):
    entityName = 'PortfolioItem'
    queryString = 'FormattedID = "' + feaID + '"'
    response = instRally.get(entityName, fetch=True, projectScopeUp=True, query=queryString)
    for ppmFeature in response:
        if ppmFeature.DirectChildrenCount > 0 and hasattr(ppmFeature, 'Children'):
            for teamfeature in ppmFeature.Children:
                #writer.writerow ( [getattrError(teamfeature, attr) for attr in gPPMFieldnames] )
                featureHierarchy.append( [getattrError(teamfeature, attr) for attr in gPPMFieldnames] )
                processTF(instRally, teamfeature.FormattedID, writer)
            featureHierarchy.append( [getattrError(ppmFeature, attr) for attr in gPPMFieldnames] )
    writer.writerows(featureHierarchy)

def processTF(instRally, formID, writer):
    queryString = 'Feature.FormattedID = "' + formID + '"'
    entity_name = 'HierarchicalRequirement'
    response = instRally.get(entity_name, fetch=True, projectScopeDown=True, query=queryString)
    print ("   " + queryString + "    Count = " + str(response.resultCount))
    for userstory in response:
        processUserStory(writer, userstory)

def processUS(instRally, usID, writer):
    queryString = 'FormattedID = "' + usID + '"'
    entity_name = 'HierarchicalRequirement'
    response = instRally.get(entity_name, fetch=True, projectScopeDown=True, query=queryString)
    print ("   " + queryString + "    Count = " + str(response.resultCount))
    for userstory in response:
        processUserStory(writer, userstory)

def processDE(instRally, defectID, writer):
    queryString = 'FormattedID = "' + defectID + '"'
    entity_name = 'Defect'
    response = instRally.get(entity_name, fetch=True, projectScopeDown=True, query=queryString)
    print ("   " + queryString + "    Count = " + str(response.resultCount))
    for defect in response:
        processUserStory(writer, defect)

def tokens(fileobj):
    for line in fileobj:
        if line[0] != '#':
            for word in line.split():
                yield word

def main():

    try:
        infile = open(sys.argv[1], 'r')
        gQueryList = tokens(infile)

        writer = workbk = None
        print('Create File...')
        writer = openCSV( 'lgsout.csv' )
        #workbk = Workbook()
        #wsA = workbk.active

        #writeHeader(writer, wsA)
        writeHeader(writer)

        #if workbk: workbk.save('lgsout.xlsx')

        print('Logging in...')
        rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

        print('Query execution...')

        for queryItem in gQueryList:
            if queryItem[:2] == "FE":
                processFEA(rally, queryItem, writer)
            elif queryItem[:2] == "TF":
                processTF(rally, queryItem, writer)
            elif queryItem[:2] == "US":
                processUS(rally, queryItem, writer)
            elif queryItem[:2] == "DE":
                processDE(rally, queryItem, writer)
            else:
                print ("Error query for " + queryItem)
    except IOError:
        print ("IOError: ", sys.argv[0], "<input file>")
    except IndexError:
        print ("IndexError: ", sys.argv[0], "<input file>")
    print('Fini')



if __name__ == '__main__':
    execfile("apifig.py")
    main()
