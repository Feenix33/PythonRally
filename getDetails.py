#!C:\Python27\python.exe
from __future__ import print_function
from pyral import Rally, rallySettings, rallyWorkset
import sys


"""
"""

errout = sys.stderr.write

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]

    #if len(args) != 1:
    #    errout('ERROR: Wrong number of arguments\n')
    #    sys.exit(3)

    args = ["DE77086"]
    args = ["TF20404"]
    args = ["FEA10264"]
    args = ["US371475"]

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)
    #print (rally.apikey)

    print ('Query execution...')
    queryString = 'FormattedID = "%s"' % args[0]

    entityName = 'Defect'
    entityName = 'HierarchicalRequirement'
    entityName = 'PortfolioItem'
    queryString = 'FormattedID = "' + args[0] + '"'
    print ("Query = ", queryString)
    response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

    if response.resultCount == 0:
        errout('No item found for %s %s\n' % (entityName, args[0]))
        sys.exit(4)
    for item in response:
        #print (item.details())
        #print (item.TeamFeature.FormattedID, item.TeamFeature.Name) #good for US
        print (item.Parent.FormattedID, item.Parent.Name) #good for TF and FEA


if __name__ == '__main__':
    main(sys.argv[1:])
