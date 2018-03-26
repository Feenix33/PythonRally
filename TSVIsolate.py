from __future__ import print_function
"""
TSVIsolate.py2
Counts phrases around a specified list of words
If the token word is found in the TSV file, then the 2 words before and after are concatenated 
and the resulting phrase is counted (the middle word in each phrase is the target word).
All words are converted to lower case.

Inputs:
    -t token_file   List of interested words, one word per line
    -i input_file   The input TSV file; minimal error detection if not in TSV format
    -o output_file  Output in CSV format

Future / Defects:
    Hardcoded to +/-2 words, probably not too hard to make it +/-x words
    Doesn't strip away double quotes

Versions:

"""

from collections import Counter
import argparse
import csv
#import re
#import pprint

def getTokens(infile):
    tokens =  open(infile,"r").readlines()
    return [aToken[:-1].lower() for aToken in tokens]

def findTokensInTSV(tokenList, infile, outfile):
    cnt = Counter()
    queue = ["","","","",""]
    with open(infile,"r") as fin:
        for line in fin:
            wordList = line.split()
            if len(wordList) == 2:
                aWord = wordList[0].lower()
                #add the word to the queue
                queue.pop(0)
                queue.append(aWord)

                #check if middle of queue is an interesting word
                if queue[2] in tokenList:
                    phrase = "_".join(queue)
                    cnt[phrase] += 1


    with open(outfile,"w") as fout:
        for key, value in cnt.items():
            fout.write ("{},{}\n".format(key, value))

#--------------------------
def writeCSV(fname, listData):
    with open(fname, "w") as f:
        cw = csv.writer(f)
        cw.writerows(listData)

def buildInputParser(parser):
    parser.add_argument("-t", "--token", nargs='?', const=1, type=str, default='token.txt', 
            help="tokens to count")
    parser.add_argument("-i", "--infile", nargs='?', const=1, type=str, default='in.tsv', 
            help="input file with search parameters")
    parser.add_argument("-o", "--outfile", nargs='?', const=1, type=str, default='out.csv', help="Output file name")

def main():
    inParser = argparse.ArgumentParser()
    buildInputParser(inParser)
    args = inParser.parse_args()
    tokenList = getTokens(args.token)
    findTokensInTSV(tokenList, args.infile, args.outfile)
    
#--------------------------

if __name__ == '__main__':
    main()
