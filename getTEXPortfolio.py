"""
getTEXPortfolio.py
Tries to get the entire TEX portfolio by looking at the parent 
and then recursively looking at the children
Hardcoded to pull the root TEX project then finds all children recursively

"""
import sys
import csv
import pyral

from pyral import Rally, rallySettings, rallyWorkset
from collections import defaultdict
  
queryList = [ 'FormattedID = "PGM155"']

def writePortfolioHeader(writer):
  writer.writerow(   (\
      "FormattedID", 
      "Name", 
      "ParentID",
      "Parent",
      "DirectChildrenCount",
      "Owner",
      "CreationDate",
  ) )



def writePortfolioItem(writer, item, parentID):
  if hasattr(item, 'Owner') and hasattr(item.Owner, 'Name'):
    ownerName = item.Owner.Name
  else:
    ownerName = ""
  if hasattr(item, 'Parent') and hasattr(item.Parent, 'Name'):
    parentName = item.Parent.Name
  else:
    parentName = ""
  writer.writerow(   (\
      item.FormattedID, item.Name, \
      parentID,
      parentName,
      item.DirectChildrenCount,
      ownerName,
      item.CreationDate,
  ) )


def processChildren(writer, item, parentID):
  writePortfolioItem(writer, item, parentID)
  if item.DirectChildrenCount > 0 and hasattr(item, 'Children'):
    for child in item.Children:
      processChildren(writer, child, item.FormattedID)


def main():
  if len(sys.argv) > 1:
    fout = open(sys.argv[1], 'wb')
  else:
    fout = open("out.csv", 'wb')
  
  print "Logging in ..."
  server, user, password, apikey, workspace, project = 'rally1.rallydev.com', 'christopher.eberhardt@sabre.com', 'Unif0rm.21', '', 'Sabre Production Workspace', 'Traveler Experience' 
  apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
  #rally = Rally(server, user, password, workspace=workspace, project=project)
  rally = Rally(server, apikey=apikey, workspace=workspace, project=project)
  
  writer = csv.writer(fout, dialect='excel')
  
  for queryString in queryList:
    print "Processing " + queryString
    #response = rally.get('PortfolioItem', fetch=True, query=queryString)
    #response = rally.get('PortfolioItem', fetch=True, projectScopeDown=True, query=queryString)
    response = rally.get('PortfolioItem', fetch=True, projectScopeUp=True, query=queryString)
    writePortfolioHeader(writer)

    for pgm in response:
      strID = pgm.FormattedID
      if strID[:3] == "PGM": # query pulls in PRJ155 too for some reason
        processChildren(writer, pgm, "Root")
  
  fout.close()


if __name__ == '__main__':
  execfile("config.py")
  main()
