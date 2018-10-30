from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv
import operator



gFieldnames = ['FormattedID', 'PlanEstimate', 'ScheduleState',\
                        'Iteration.Name', 'Project.Name', 'Project.Parent.Name', \
                        'TeamFeature.FormattedID', 'TeamFeature.Parent.FormattedID']

gData = []

def processRecord(record):
    return  [getattrAgain(record, attr) for attr in gFieldnames] 

 
def getattrAgain(n, attr):
    try:
        return operator.attrgetter(attr)(n)
    except AttributeError:
        return ""

def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    #project = 'LGS Lodging'
    project = 'LGS Titans (BLR)'
    eprint ('Logging in...')
    rally = Rally(server, apikey=apikey, workspace=workspace, project=project)

    eprint ('Query execution...')

    queryString = 'Iteration.Name contains "S19#2018"'      # iteration query
    queryString = 'Iteration.Name contains "2018"'      # iteration query
    entityName = 'HierarchicalRequirement'
    eprint ("Query = ", queryString)
    response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)
    eprint ('Processing responses...')

    if response.resultCount == 0:
        eprint('No item found for %s %s\n' % (entityName, arg))
    else:
        gData.append(gFieldnames)
        for item in response:
            gData.append (processRecord(item))
        #print (gData)

        # write csv file
        fout = open('out.csv', 'w', newline='')
        outputWriter = csv.writer(fout, dialect='excel')
        outputWriter.writerows (gData)
        fout.close()


if __name__ == '__main__':
    main(sys.argv[1:])
