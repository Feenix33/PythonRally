from pyral import Rally, rallySettings, rallyWorkset
import sys
import csv


"""
"""

errout = sys.stderr.write

def printHelp():
    print ("USAGE: getMilestone [-s] [-h] <milestone>")
    print ("    <milestone>   Use only the number 'TN-Hotel' is added")
    print ("    -s    Print stories")
    print ("    -h    Help")
            

def main(args):
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]

    if len(args) != 1:
        errout('ERROR: Wrong number of arguments\n')
        printHelp()
        sys.exit(3)

    #args = ["TN-Hotel 18.7"]
    #args = ["18.7"]

    bPrintStories = True
    for opt in options:
        if opt[1] == "n":
            bPrintStories = False
        elif opt[1] == "h":
            printHelp()
            sys.exit(0)

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

        #queryString = 'FormattedID = "%s"' % arg
        queryString = 'Milestones.Name contains "%s"' % ("TN-Hotel "+arg)
        print ("Query = ", queryString)
        response = rally.get(entityName, fetch=True, projectScopeDown=True, query=queryString)

        if response.resultCount == 0:
            errout('No item found for %s %s\n' % (entityName, arg))
        else:
            fileName = 'Mile' + arg + '.csv'
            print ("Printing to '%s'" % fileName)
            with open (fileName, 'w', newline='') as csvfile:
                outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                outfile.writerow(["Type", "FEA", "TF", "US", "Name", "Project",
                            "LeafStoryCount",
                            "AcceptedLeafStoryCount",
                            "UnEstimatedLeafStoryCount",
                            "LeafStoryPlanEstimateTotal",
                            "AcceptedLeafStoryPlanEstimateTotal",
                            "PlanEstimate",
                            "ScheduleState",
                            "Iteration",
                            "Tags"
                            ])
                for item in response:
                    outfile.writerow(["FEA", item.FormattedID, "", "", item.Name, "",
                        item.LeafStoryCount,
                        item.AcceptedLeafStoryCount,
                        item.UnEstimatedLeafStoryCount,
                        item.LeafStoryPlanEstimateTotal,
                        item.AcceptedLeafStoryPlanEstimateTotal,
                        "",
                        "",
                        " ".join(tag.Name for tag in item.Tags),
                        ])
                    if hasattr(item, "Children"):
                        for child in item.Children:
                            outfile.writerow(["TF",item.FormattedID, child.FormattedID, "", child.Name, 
                                child.Project.Name,
                                child.LeafStoryCount,
                                child.AcceptedLeafStoryCount,
                                child.UnEstimatedLeafStoryCount,
                                child.LeafStoryPlanEstimateTotal,
                                child.AcceptedLeafStoryPlanEstimateTotal,
                                "",
                                "",
                                " ".join(tag.Name for tag in child.Tags),
                                ])
                            if bPrintStories and hasattr(child, "UserStories"):
                                for story in child.UserStories:
                                    sIteration = ""
                                    if hasattr(story, "Iteration"):
                                        if hasattr(story.Iteration, "Name"):
                                            sIteration = story.Iteration.Name
                                    outfile.writerow(["US",item.FormattedID, child.FormattedID, story.FormattedID, story.Name,
                                        story.Project.Name,
                                        "","","","","",
                                        story.PlanEstimate,
                                        story.ScheduleState,
                                        sIteration,
                                        " ".join(tag.Name for tag in story.Tags),
                                        ])


if __name__ == '__main__':
    main(sys.argv[1:])
