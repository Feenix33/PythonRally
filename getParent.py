#!C:\Python27\python.exe
from __future__ import print_function
from pyral import Rally, rallySettings, rallyWorkset
import sys
import argparse
import csv

"""
getParent.py
Identify the parent for the passed in items
"""

errout = sys.stderr.write

def getTokens(infile):
    with open(infile) as f:
        lines = [line.rstrip('\n') for line in f]
    return lines

def buildInputParser(parser):
    parser.add_argument("IDs", nargs='*', help="Items to find a parent")
    parser.add_argument("-i", "--infile", nargs='?', const='in.txt', default=None, help="Input file name")
    #parser.add_argument("-o", "--outfile", nargs='?', const='out.csv', default=None, help="Output file name")

def main():
    # Parse command line
    inParser = argparse.ArgumentParser()
    buildInputParser(inParser)
    args = inParser.parse_args()

    if args.infile:
        tokens = args.IDs + getTokens(args.infile)
    else:
        tokens = args.IDs

    # Rally setup
    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    #print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)
    #print ('Start queries...')

    for id in tokens:
        if id[0] == "U": entityName = 'HierarchicalRequirement'
        elif id[0] == "T": entityName = 'PortfolioItem'
        elif id[0] == "F": entityName = 'PortfolioItem'
        else: continue

        queryString = 'FormattedID = "%s"' % id
        response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)
        if response.resultCount > 0:
            for item in response:
                if id[0] == "U": 
                    if item.TeamFeature == None:
                        outstring = '%s, %s, "%s"' % (item.FormattedID, 'None', 'None')
                    else:
                        outstring = '%s, %s, "%s"' % (item.FormattedID, item.TeamFeature.FormattedID, item.TeamFeature.Name.encode('ascii', 'replace')) #good for US
                else:
                    if item.Parent == None:
                        outstring = '%s, %s, "%s"' % (item.FormattedID, 'None', 'None')
                    else:
                        outstring = '%s, %s, "%s"' % (item.FormattedID, item.Parent.FormattedID, item.Parent.Name.encode('ascii','replace')) #good for TF and FEA
                print (outstring)

if __name__ == '__main__':
    main()
