from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv


"""
"""

gAttribs = [
        "FormattedID", #"Name",
        "ScheduleState",
        "Project.Name",
        "Project.Parent.Name",
        "PlanEstimate",
        "Iteration.Name",
        "PortfolioItem.FormattedID", #"PortfolioItem.Name",
        "PortfolioItem.Parent.FormattedID", #"PortfolioItem.Parent.Name",
        ]

def get_deep_attr(obj, attrs):
    for attr in attrs.split("."):
        try:
            obj = getattr(obj, attr)
        except AttributeError:
            return ""
    return obj

def has_deep_attr(obj, attrs):
    try:
        get_deep_attr(obj, attrs)
        return True
    except AttributeError:
        return False

errout = sys.stderr.write

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]

    #if len(args) != 1:
    #    errout('ERROR: Wrong number of arguments\n')
    #    sys.exit(3)
    #args = ["US504765"] # no TF
    #args = ["US487422"]  #for Titans

    server = 'rally1.rallydev.com'
    apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ'
    workspace = 'Sabre Production Workspace'
    project = 'Sabre' 
    project = 'LGS Titans (BLR)' 
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

        #queryString = 'FormattedID = "%s"' % arg
        queryString = '(Iteration.StartDate > "2017-12-31")'
        entityName = 'HierarchicalRequirement'

        print ("Query = ", queryString)
        response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

        if response.resultCount == 0:
            errout('No item found for %s %s\n' % (entityName, arg))
        else:
            fileName = 'out.csv'
            with open (fileName, 'w', newline='') as csvfile:
                outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                outrow = [field for field in gAttribs] 
                outfile.writerow(outrow)
                for item in response:
                    outfile.writerow( [get_deep_attr(item, param) for param in gAttribs])


if __name__ == '__main__':
    main(sys.argv[1:])
