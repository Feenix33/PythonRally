"""
"""

import argparse
import csv

def getTokens(infile):
    with open(infile) as f:
        lines = [line.rstrip('\n') for line in f]
    return lines

def buildQuery(tokens):
    strOut = "(("
    for token in tokens:
        if token[0] == 'T':
            strOut += 'TeamFeature.FormattedID = "' + str(token) + '") OR ('
    strOut = strOut[:-5] + ")"
    return strOut

#--------------------------

def buildInputParser(parser):
    parser.add_argument("infile", help="Input file of tokens to build query")
    parser.add_argument("-o", "--outfile", nargs='?', const=1, type=str, default='out.csv', help="Output file name")

def main():
    inParser = argparse.ArgumentParser()
    buildInputParser(inParser)
    args = inParser.parse_args()
    #print (args)
    tokens = getTokens(args.infile)
    #print (tokens)
    query = buildQuery(tokens)
    print (query)
    
#--------------------------

if __name__ == '__main__':
    main()
