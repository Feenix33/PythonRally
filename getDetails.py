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

    #args = ["DE77086"]
    #args = ["TF20404"]
    #args = ["US371475"]
    #args = ["FEA10264"]

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    print ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    print ('Query execution...')
    for arg in args:

        if arg[0] == "D":
            entityName = 'Defect'
        elif arg[0] == "U":
            entityName = 'HierarchicalRequirement'
        else:
            entityName = 'PortfolioItem'

        queryString = 'FormattedID = "%s"' % arg
        #queryString = 'FormattedID = "' + arg + '"'
        #print ("Query = ", queryString)
        response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

        if response.resultCount == 0:
            errout('No item found for %s %s\n' % (entityName, arg))
        else:
            for item in response:
                print (item.details())


if __name__ == '__main__':
    main(sys.argv[1:])
