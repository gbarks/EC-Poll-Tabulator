#!/usr/bin/env python

# ==========================================================
#  ElloCoaster poll tabulator
#  Contributions from Jim Winslett, Dave Wong, Grant Barker
# ==========================================================

from __future__ import print_function # for Python 2.x users

import os
import sys
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from phicolor import phicolor, divcolor
    from spinner import Spinner
except:
    print('Could not find "phicolor.py" and/or "spinner.py"; exiting...')
    sys.exit()

# global strings for parsing ballots
commentStr = "* "
blankUserField = "-Replace "
startLine = "! DO NOT CHANGE OR DELETE THIS LINE !"

# command line arguments
parser = argparse.ArgumentParser(description='Process Mitch Hawker-style coaster poll.')

parser.add_argument("-b", "--blankBallot", default="blankballot2017.txt",
                    help="specify blank ballot file")
parser.add_argument("-f", "--ballotFolder", default="ballots2017",
                    help="specify folder containing filled ballots")
parser.add_argument("-m", "--minRiders", type=int, default=6,
                    help="specify minimum number of riders for a coaster to rank")
parser.add_argument("-o", "--outfile", default="Poll Results.xlsx",
                    help="specify name of output .xlsx file")
parser.add_argument("-c", "--colorize", action="store_true",
                    help="color coaster labels by make in output spreadsheet")
parser.add_argument("-i", "--includeVoterInfo", action="store_true",
                    help="include sensitive voter data in output spreadsheet")
parser.add_argument("-r", "--botherRCDB", action="store_true",
                    help="bother RCDB to grab metadata from links in blankBallot")
parser.add_argument("-v", "--verbose", action="store_true",
                    help="print data as it's processed")

args = parser.parse_args()

if not os.path.isfile(args.blankBallot):
    print('Blank ballot source "{0}" is not a file; exiting...'.format(args.blankBallot))
    sys.exit()

if not os.path.isdir(args.ballotFolder) or len(os.listdir(args.ballotFolder)) < 1:
    print('Ballot folder "{0}" does not exist or is empty; exiting...'.format(args.ballotFolder))
    sys.exit()

if args.outfile[-5:] != ".xlsx":
    args.outfile += ".xlsx"

if args.botherRCDB and sys.version_info >= (3,0):
    import lxml
    from bs4 import BeautifulSoup
    from urllib.request import urlopen
elif args.botherRCDB:
    print("Bothering RCDB requires Python 3; ignoring '-r' flag...")
    args.botherRCDB = False



# ==================================================
#  onto main()!
# ==================================================

def main():
    # create Excel workbook
    xlout = Workbook()
    xlout.active.title = "Coaster Masterlist"

    # preferred fixed-width font
    menlo = Font(name="Menlo")

    light = 186
    multi = 1
    offset = 0.018034
    makeColors = { # fill colors for certain roller coaster manufacturers/designers
        "Custom Coasters International, Inc."   : PatternFill("solid", fgColor=phicolor( 0, light, multi, offset)),
        "Dinn Corporation"                      : PatternFill("solid", fgColor=phicolor( 1, light, multi, offset)),
        "Gravitykraft Corporation"              : PatternFill("solid", fgColor=phicolor( 2, light, multi, offset)),
        "The Gravity Group, LLC"                : PatternFill("solid", fgColor=phicolor( 2, light, multi, offset)),
        "Great Coasters International"          : PatternFill("solid", fgColor=phicolor( 3, light, multi, offset)),
        "Intamin Amusement Rides"               : PatternFill("solid", fgColor=phicolor( 4, light, multi, offset)),
        "Martin & Vleminckx"                    : PatternFill("solid", fgColor=phicolor( 5, light, multi, offset)),
        "National Amusement Device Company"     : PatternFill("solid", fgColor=phicolor( 6, light, multi, offset)),
        "Philadelphia Toboggan Coasters, Inc."  : PatternFill("solid", fgColor=phicolor( 7, light, multi, offset)),
        "Rocky Mountain Construction"           : PatternFill("solid", fgColor=phicolor( 8, light, multi, offset)),
        "Roller Coaster Corporation of America" : PatternFill("solid", fgColor=phicolor( 9, light, multi, offset)),
        "S&S Worldwide"                         : PatternFill("solid", fgColor=phicolor(10, light, multi, offset)),
        "Vekoma"                                : PatternFill("solid", fgColor=phicolor(11, light, multi, offset)),
        "Other Known Manufacturer"              : PatternFill("solid", fgColor="cccccc"),
        ""                                      : PatternFill("solid", fgColor="eeeeee")
    }

    # list of tuples of the form (fullCoasterName, abbreviatedCoasterName)
    coasterDict = getCoasterDict(xlout.active, menlo, makeColors)

    # create color key for makeColors
    if args.colorize:
        coastermakews = xlout.create_sheet("Coaster Make Color Key")
        i = 1
        for make in makeColors.keys():
            if make:
                coastermakews.append([make])
            else:
                coastermakews.append(["Other [Unknown]"])
            coastermakews.cell(row=i, column=1).fill = makeColors[make]
            i += 1
        coastermakews.column_dimensions['A'].width = 30.83

    # list of ballot filepaths
    ballotList = getBallotFilepaths()

    # for each pair of coasters, a list of numbers of the form [wins, losses, ties, winPercent]
    winLossMatrix = createMatrix(coasterDict)

    # loop through all the ballot filenames and process each ballot
    if args.includeVoterInfo:
        voterinfows = xlout.create_sheet("Voter Info (SENSITIVE)")
        voterinfows.append(["Ballot Filename","Name","Email","City","State/Province","Country","Coasters Ridden"])
        voterinfows.column_dimensions['A'].width = 24.83
        voterinfows.column_dimensions['B'].width = 16.83
        voterinfows.column_dimensions['C'].width = 24.83
        for col in ['D','E','F','G']:
            voterinfows.column_dimensions[col].width = 12.83
    for filepath in ballotList:
        voterInfo = processBallot(filepath, coasterDict, winLossMatrix)
        if args.includeVoterInfo and voterInfo:
            voterinfows.append(voterInfo)
    if args.includeVoterInfo:
        voterinfows.freeze_panes = voterinfows['A2']

    calculateResults(coasterDict, winLossMatrix)

    # sorted lists of tuples of the form (rankedCoaster, relevantNumbers)
    finalResults, finalPairs, finalRiders = sortedLists(coasterDict, winLossMatrix)

    # write worksheets related to finalResults, finalPairs, finalRiders, and winLossMatrix
    printToFile(xlout, finalResults, finalPairs, finalRiders, winLossMatrix, coasterDict, menlo, makeColors)

    # save the Excel file
    print("Saving...", end=" ")
    spinner = Spinner()
    spinner.start()
    xlout.save(args.outfile)
    spinner.stop()
    print('output saved to "Poll Results.xlsx".')



# ==================================================
#  populate dictionary of coasters in the poll
# ==================================================

def getCoasterDict(masterlistws, preferredFixedWidthFont, makeColors):
    if not args.botherRCDB:
        print("Creating list of every coaster on the ballot...", end=" ")
        spinner = Spinner()
        spinner.start()

    # set up Coaster Masterlist worksheet
    headerRow = ["Full Coaster Name", "Abbrev.", "Name", "Park", "State", "RCDB Link"]
    if args.botherRCDB:
        headerRow.extend(["Make", "Year"])
    masterlistws.append(headerRow)
    masterlistws.column_dimensions['A'].width = 45.83
    masterlistws.column_dimensions['B'].width = 12.83
    masterlistws.column_dimensions['C'].width = 25.83
    masterlistws.column_dimensions['D'].width = 25.83
    masterlistws.column_dimensions['E'].width = 6.83
    masterlistws.column_dimensions['F'].width = 16.83
    masterlistws['B1'].font = preferredFixedWidthFont
    masterlistws['E1'].font = preferredFixedWidthFont
    if args.botherRCDB:
        masterlistws.column_dimensions['G'].width = 25.83
        masterlistws.column_dimensions['H'].width = 4.83

    coasterDict = {} # return value

    #open the blank ballot file
    with open(args.blankBallot) as f:
        lineNum = 0
        startProcessing = False

        # begin going through the blank ballot line by line
        for line in f:

            sline = line.strip() # strip whitespace from start and end of line
            lineNum += 1

            # skip down the file to the coasters
            if startProcessing == False and sline == startLine:
                startProcessing = True

            # add the coasters to coasterDict and the masterlist worksheet
            elif startProcessing == True:

                if commentStr in sline: # skip comment lines (begin with "* ")
                    continue

                elif sline == "": # skip blank lines
                    continue

                else:
                    # break the line into its components: rank, full coaster name, abbreviation
                    words = [x.strip() for x in sline.split(',')]

                    # make sure there are at least 3 'words' in each line (rank, fullName, abbrName)
                    if len(words) < 3:
                        print("Error in {0}, Line {1}: {2}".format(args.blankBallot, lineNum, line))

                    else:
                        fullName = words[1]
                        abbrName = words[2]

                        coasterMake = ""

                        # list of strings that will form a row in the spreadsheet
                        rowVals = [fullName, abbrName]

                        # add the coaster to the dictionary of coasters on the ballot
                        coasterDict[fullName] = {}

                        # extract park and state/country information from fullName
                        subwords = [x.strip() for x in fullName.split('-')]
                        if len(subwords) != 3:
                            rowVals.extend(["", "", ""])
                            coasterDict[fullName]["Name"] = ""
                            coasterDict[fullName]["Park"] = ""
                            coasterDict[fullName]["State"] = ""
                        else:
                            rowVals.extend([subwords[0], subwords[1], subwords[2]])
                            coasterDict[fullName]["Name"] = subwords[0]
                            coasterDict[fullName]["Park"] = subwords[1]
                            coasterDict[fullName]["State"] = subwords[2]

                        coasterDict[fullName]["RCDB"] = ""
                        coasterDict[fullName]["Make"] = ""
                        coasterDict[fullName]["Year"] = ""
                        if len(words) > 3:
                            coasterDict[fullName]["RCDB"] = words[3]
                            rowVals.append('=HYPERLINK("{0}", "{1}")'.format(words[3], words[3][8:]))

                            if args.botherRCDB:
                                response = urlopen(coasterDict[fullName]["RCDB"])
                                html = response.read()
                                soup = BeautifulSoup(html, 'lxml')

                                for x in soup.body.findAll('div', attrs={'class':'scroll'}):
                                    if "Make: " in x.text:
                                        subtext = x.text.split("Make: ", 1)[1]
                                        if "Model: " in subtext:
                                            subtext = subtext.split("Model: ", 1)[0]
                                        coasterDict[fullName]["Make"] = subtext
                                        break
                                rowVals.append(coasterDict[fullName]["Make"])

                                for x in soup.body.findAll(True):
                                    if "Operating since " in x.text:
                                        subtext = x.text.split("Operating since ", 1)[1][:10].split('/')[-1][:4]
                                        coasterDict[fullName]["Year"] = int(subtext)
                                        break
                                    elif "Operated from " in x.text:
                                        subtext = x.text.split("Operated from ", 1)[1].split(' ')[0].split('/')[-1]
                                        coasterDict[fullName]["Year"] = int(subtext)
                                        break
                                rowVals.append(coasterDict[fullName]["Year"])

                                print("{0},  \t{1},\t{2}".format(
                                    abbrName, coasterDict[fullName]["Year"], coasterDict[fullName]["Make"]))

                        # final values associated with this coaster
                        coasterDict[fullName]["Abbr"] = abbrName

                        # variable values associated with this coaster
                        coasterDict[fullName]["Riders"] = 0
                        coasterDict[fullName]["Total Wins"] = 0
                        coasterDict[fullName]["Total Losses"] = 0
                        coasterDict[fullName]["Total Ties"] = 0
                        coasterDict[fullName]["Total Win Percentage"] = 0.0
                        coasterDict[fullName]["Overall Rank"] = 0
                        coasterDict[fullName]["Ridership Rank"] = 0
                        coasterDict[fullName]["Tied Coasters"] = []

                        masterlistws.append(rowVals)
                        masterlistws.cell(row=len(coasterDict)+1, column=5).font = preferredFixedWidthFont
                        masterlistws.cell(row=len(coasterDict)+1, column=2).font = preferredFixedWidthFont
                        if coasterDict[fullName]["RCDB"]:
                            masterlistws.cell(row=len(coasterDict)+1, column=6).style = "Hyperlink"

                        if args.colorize:
                            makeCell = masterlistws.cell(row=len(coasterDict)+1, column=7)
                            nameCell = masterlistws.cell(row=len(coasterDict)+1, column=1)
                            abbrCell = masterlistws.cell(row=len(coasterDict)+1, column=2)
                            if coasterDict[fullName]["Make"]:
                                if coasterDict[fullName]["Make"] in makeColors.keys():
                                    makeCell.fill = makeColors[coasterDict[fullName]["Make"]]
                                    nameCell.fill = makeColors[coasterDict[fullName]["Make"]]
                                    abbrCell.fill = makeColors[coasterDict[fullName]["Make"]]
                                else:
                                    makeCell.fill = makeColors["Other Known Manufacturer"]
                                    nameCell.fill = makeColors["Other Known Manufacturer"]
                                    abbrCell.fill = makeColors["Other Known Manufacturer"]
                            else:
                                makeCell.fill = makeColors[""]
                                nameCell.fill = makeColors[""]
                                abbrCell.fill = makeColors[""]

    masterlistws.freeze_panes = masterlistws['A2']
    if not args.botherRCDB:
        spinner.stop()
    print("{0} coasters on the ballot.".format(len(coasterDict)))
    return coasterDict



# ==================================================
#  import filepaths of ballots
# ==================================================

def getBallotFilepaths():
    print("Getting the filepaths of submitted ballots...", end=" ")
    spinner = Spinner()
    spinner.start()

    ballotList = []
    for file in os.listdir(args.ballotFolder):
        if file.endswith(".txt"):
            ballotList.append(os.path.join(args.ballotFolder, file))

    spinner.stop()
    print("{0} ballots submitted.".format(len(ballotList)))
    return ballotList



# ==================================================
#  create win/loss matrix
# ==================================================

def createMatrix(coasterDict):
    print("Creating the win/loss matrix...", end=" ")
    spinner = Spinner()
    spinner.start()

    winLossMatrix = {}
    for coasterA in coasterDict.keys():
        for coasterB in coasterDict.keys():

            # can't compare a coaster to itself
            if coasterA != coasterB:
                winLossMatrix[coasterA, coasterB] = {}
                winLossMatrix[coasterA, coasterB]["Wins"] = 0
                winLossMatrix[coasterA, coasterB]["Losses"] = 0
                winLossMatrix[coasterA, coasterB]["Ties"] = 0
                winLossMatrix[coasterA, coasterB]["Win Percentage"] = 0.0
                winLossMatrix[coasterA, coasterB]["Pairwise Rank"] = 0

    spinner.stop()
    print("{0} pairings.".format(len(winLossMatrix)))
    return winLossMatrix



# ================================================================
#  read a ballot (just ONE ballot)
#
#  you need a loop to call this function for each ballot filename
# ================================================================

def processBallot(filepath, coasterDict, winLossMatrix):
    filename = os.path.basename(filepath)
    print("Processing ballot: {0}".format(filename))

    voterInfo = [filename, "", "", "", "", ""] # return value
    coasterAndRank = {}
    creditNum = 0
    error = False

    # open the ballot file
    with open(filepath) as f:
        infoField = 1
        lineNum = 0
        startProcessing = False

        for line in f:
            sline = line.strip()
            lineNum += 1

            # begin at top of ballot and get the voter's info first
            if startProcessing == False and infoField <= 5 and not commentStr in sline and len(sline) != 0:

                # if the line begins with "-Replace" then record a non-answer
                if blankUserField in sline:
                    voterInfo[infoField] = ""
                    infoField += 1
                elif not startLine in sline:
                    voterInfo[infoField] = sline.strip('-')
                    infoField += 1

            # skip down the file to the coasters
            if startProcessing == False and sline == startLine:
                startProcessing = True

            elif startProcessing == True:

                # break the line into its components: rank, name
                words = [x.strip() for x in sline.split(',')]

                if commentStr in sline: # skip comment lines (begin with "* ")
                    continue

                elif sline == "": # skip blank lines
                    continue

                # make sure there are at least 2 'words' in each line
                elif len(words) < 2:
                    print("Error in {0}, Line {1}: {2}".format(args.blankBallot, lineNum, line))

                # make sure the ranking is a number
                elif not words[0].isdigit():
                    print("Error in reading {0}, Line {1}: Rank must be an int.".format(filename, lineNum))
                    error = True

                else:
                    coasterName = words[1]
                    coasterRank = int(words[0])

                    # skip coasters ranked zero or less (those weren't ridden)
                    if coasterRank <= 0:
                        continue

                    # check to make sure the coaster on the ballot is legit
                    if coasterName  in coasterDict.keys():
                        creditNum += 1
                        coasterDict[coasterName]["Riders"] += 1

                        # add this voter's ranking of the coaster
                        coasterAndRank[coasterName] = coasterRank

                    else: # it's not a legit coaster!
                        print("Error in reading {0}, Line {1}: Unknown coaster {2}".format(filename, lineNum, coasterName))
                        error = True

    # don't tally the ballot if there were any errors, don't return voter info
    if error:
        print("Error encountered. File {0} not added.".format(filename))
        return []

    # cycle through each pair of coasters this voter ranked
    for coasterA in coasterAndRank.keys():
        for coasterB in coasterAndRank.keys():

            # can't compare a coaster to itself
            if coasterA != coasterB:

                # if the coasters have the same ranking, call it a tie
                if coasterAndRank[coasterA] == coasterAndRank[coasterB]:
                    winLossMatrix[coasterA, coasterB]["Ties"] += 1
                    coasterDict[coasterA]["Total Ties"] += 1

                # if coasterA outranks coasterB (the rank's number is lower), call it a win for coasterA
                elif coasterAndRank[coasterA] < coasterAndRank[coasterB]:
                    winLossMatrix[coasterA, coasterB]["Wins"] += 1
                    coasterDict[coasterA]["Total Wins"] += 1

                # if not a tie nor a win, it must be a loss
                else:
                    winLossMatrix[coasterA, coasterB]["Losses"] += 1
                    coasterDict[coasterA]["Total Losses"] += 1

    print(" ->", end=" ")

    for i in range(1,len(voterInfo)):
        if voterInfo[i] != "":
            print("{0},".format(voterInfo[i]), end=" ")

    print("CC: {0}".format(creditNum))

    voterInfo.append(creditNum)

    return voterInfo



# ========================================================
#  calculate results
#
#  no need to loop through this, since it calculates with
#    numbers gathered when the ballots were processed
# ========================================================

def calculateResults(coasterDict, winLossMatrix):
    print("Calculating results...", end=" ")
    spinner = Spinner()
    spinner.start()

    if args.verbose:
        print("")

    # iterate through all the pairs in the matrix
    for coasterA in coasterDict.keys():
        for coasterB in coasterDict.keys():

            # can't compare a coaster to itself
            if coasterA != coasterB:
                pairWins = winLossMatrix[coasterA, coasterB]["Wins"]
                pairLoss = winLossMatrix[coasterA, coasterB]["Losses"]
                pairTies = winLossMatrix[coasterA, coasterB]["Ties"]
                pairContests = pairWins + pairLoss + pairTies

                if pairContests > 0:
                    winLossMatrix[coasterA, coasterB]["Win Percentage"] = (((pairWins + float(pairTies / 2)) / pairContests)) * 100

                    if args.verbose:
                        print("{0},{1},\tWins: {2},\tTies: {3},\t#Con: {4},\tWin%: {5}".format(
                            coasterDict[coasterA]["Abbr"], coasterDict[coasterB]["Abbr"],
                            pairWins, pairTies, pairContests, winLossMatrix[coasterA, coasterB]["Win Percentage"]))

    for x in coasterDict.keys():
        numWins = coasterDict[x]["Total Wins"]
        numLoss = coasterDict[x]["Total Losses"]
        numTies = coasterDict[x]["Total Ties"]
        numContests = numWins + numLoss + numTies

        if  numContests > 0:
            coasterDict[x]["Total Win Percentage"] = ((numWins + float(numTies/2)) / numContests) * 100

            if args.verbose:
                print("{0},\tWins: {1},\tTies: {2},\t#Con: {3},\tWin%: {4}".format(
                    coasterDict[x]["Abbr"], numWins, numTies, numContests,
                    coasterDict[x]["Total Win Percentage"]))

    spinner.stop()
    print(" ")



# ==================================================
#  add to "Tied Coasters" variable and print
# ==================================================

def markTies(coasterDict, winLossMatrix, tiedCoasters):
    for coasterA in tiedCoasters:
        coastersTiedWithA = []
        for coasterB in tiedCoasters:
            if coasterA != coasterB:
                coastersTiedWithA.append(coasterB)
        coasterDict[coasterA]["Tied Coasters"] = coastersTiedWithA
    if args.verbose: # print Mitch Hawker-style pairwise matchups between tied coasters
        print("  ===Tied===", end="\t")
        for coaster in tiedCoasters:
            print(" {0} ".format(coasterDict[coasterB]["Abbr"]), end="\t")
        print("")
        for coasterA in tiedCoasters:
            print("  {0}".format(coasterDict[coasterA]["Abbr"]), end="\t")
            for coasterB in tiedCoasters:
                cellStr = " "
                if coasterA != coasterB:
                    if winLossMatrix[coasterA, coasterB]["Wins"] > winLossMatrix[coasterA, coasterB]["Losses"]:
                        cellStr += "W "
                    elif winLossMatrix[coasterA, coasterB]["Wins"] < winLossMatrix[coasterA, coasterB]["Losses"]:
                        cellStr += "L "
                    else:
                        cellStr += "T "
                    cellStr += str(winLossMatrix[coasterA, coasterB]["Wins"]) + "-"
                    cellStr += str(winLossMatrix[coasterA, coasterB]["Losses"]) + "-"
                    cellStr += str(winLossMatrix[coasterA, coasterB]["Ties"])
                else:
                    cellStr += "       "
                print("{0}".format(cellStr), end="\t")
            print("")



# ==================================================
#  create sorted list of coasters by win pct and
#    sorted list of coasters by pairwise win pct
# ==================================================

def sortedLists(coasterDict, winLossMatrix):
    print("Sorting the results...", end=" ")
    spinner = Spinner()
    spinner.start()

    results = []
    numRiders = []
    pairPercents = []

    # iterate through coasterDict by coasters
    for coasterName in coasterDict.keys():
        numRiders.append((coasterName, coasterDict[coasterName]["Riders"]))
        if int(coasterDict[coasterName]["Riders"]) >= int(args.minRiders):
            results.append((coasterName, coasterDict[coasterName]["Total Win Percentage"]))

    # iterate through winLossMatrix by coaster pairings
    for coasterPair in winLossMatrix.keys():      
        pairPercents.append((coasterPair, winLossMatrix[coasterPair]["Win Percentage"]))

    # sort lists by win percentages and ridership
    sortedResults = sorted(results, key=lambda x: x[1], reverse=True)
    sortedPairs = sorted(pairPercents, key=lambda x: x[1], reverse=True)
    sortedRiders = sorted(numRiders, key=lambda x: x[1], reverse=True)

    if args.verbose:
        print("")

    # determine rankings including ties for all three lists
    overallRank = 0
    curRank = 0
    curValue = 0.0
    tiedCoasters = []
    for x in sortedResults:
        overallRank += 1
        if x[1] != curValue:
            if len(tiedCoasters) > 1: # do stuff on complete list of tied coasters
                markTies(coasterDict, winLossMatrix, tiedCoasters)
            curRank = overallRank
            curValue = x[1]
            tiedCoasters = []
        tiedCoasters.append(x[0])
        coasterDict[x[0]]["Overall Rank"] = curRank
        if args.verbose:
            print("Rank: {0},\tVal: {1},\tCoaster: {2}".format(coasterDict[x[0]]["Overall Rank"], x[1], x[0]))
    if len(tiedCoasters) > 1: # in case last few coasters were tied
        markTies(coasterDict, winLossMatrix, tiedCoasters)

    overallRank = 0
    curRank = 0
    curValue = 0.0
    for x in sortedPairs:
        overallRank += 1
        if x[1] != curValue:
            curRank = overallRank
            curValue = x[1]
        winLossMatrix[x[0][0], x[0][1]]["Pairwise Rank"] = curRank

    overallRank = 0
    curRank = 0
    curValue = 0
    for x in sortedRiders:
        overallRank += 1
        if x[1] != curValue:
            curRank = overallRank
            curValue = x[1]
        coasterDict[x[0]]["Ridership Rank"] = curRank


    spinner.stop()
    print(" ")

    return sortedResults, sortedPairs, sortedRiders



# ==================================================
#  print everything to a file
# ==================================================

def printToFile(xl, results, pairs, riders, winLossMatrix, coasterDict, preferredFixedWidthFont, makeColors):
    print("Writing the results...", end=" ")
    spinner = Spinner()
    spinner.start()

    # create and write primary results worksheet
    resultws = xl.create_sheet("Ranked Results")
    resultws.append(["Rank","Coaster","Total Win Percentage","Total Wins","Total Losses","Total Ties"])
    resultws.column_dimensions['A'].width = 4.83
    resultws.column_dimensions['B'].width = 45.83
    resultws.column_dimensions['C'].width = 16.83
    resultws.column_dimensions['D'].width = 8.83
    resultws.column_dimensions['E'].width = 9.83
    resultws.column_dimensions['F'].width = 7.83
    i = 2
    for x in results:
        resultws.append([coasterDict[x[0]]["Overall Rank"], x[0],
                         coasterDict[x[0]]["Total Win Percentage"],
                         coasterDict[x[0]]["Total Wins"],
                         coasterDict[x[0]]["Total Losses"],
                         coasterDict[x[0]]["Total Ties"]])
        if args.colorize:
            if coasterDict[x[0]]["Make"]:
                if coasterDict[x[0]]["Make"] in makeColors.keys():
                    resultws.cell(row=i, column=2).fill = makeColors[coasterDict[x[0]]["Make"]]
                else:
                    resultws.cell(row=i, column=2).fill = makeColors["Other Known Manufacturer"]
            else:
                resultws.cell(row=i, column=2).fill = makeColors[""]
        i += 1
    resultws.freeze_panes = resultws['A2']

    # create and write pairwise result worksheet
    pairws = xl.create_sheet("Ranked Pairs")
    pairws.append(["Rank","Primary Coaster","Rival Coaster","Win Percentage","Wins","Losses","Ties"])
    pairws.column_dimensions['A'].width = 4.83
    pairws.column_dimensions['B'].width = 45.83
    pairws.column_dimensions['C'].width = 45.83
    pairws.column_dimensions['D'].width = 12.83
    pairws.column_dimensions['E'].width = 4.5
    pairws.column_dimensions['F'].width = 5.5
    pairws.column_dimensions['G'].width = 3.83
    i = 2
    for x in pairs:
        pairws.append([winLossMatrix[x[0][0], x[0][1]]["Pairwise Rank"], x[0][0], x[0][1],
                       winLossMatrix[x[0][0], x[0][1]]["Win Percentage"],
                       winLossMatrix[x[0][0], x[0][1]]["Wins"],
                       winLossMatrix[x[0][0], x[0][1]]["Losses"],
                       winLossMatrix[x[0][0], x[0][1]]["Ties"]])
        if args.colorize:
            if coasterDict[x[0][0]]["Make"]:
                if coasterDict[x[0][0]]["Make"] in makeColors.keys():
                    pairws.cell(row=i, column=2).fill = makeColors[coasterDict[x[0][0]]["Make"]]
                else:
                    pairws.cell(row=i, column=2).fill = makeColors["Other Known Manufacturer"]
            else:
                pairws.cell(row=i, column=2).fill = makeColors[""]
            if coasterDict[x[0][1]]["Make"]:
                if coasterDict[x[0][1]]["Make"] in makeColors.keys():
                    pairws.cell(row=i, column=3).fill = makeColors[coasterDict[x[0][1]]["Make"]]
                else:
                    pairws.cell(row=i, column=3).fill = makeColors["Other Known Manufacturer"]
            else:
                pairws.cell(row=i, column=3).fill = makeColors[""]
        i += 1
    pairws.freeze_panes = pairws['A2']

    # create and write ridership worksheet
    riderws = xl.create_sheet("Number of Riders")
    riderws.append(["Rank","Coaster","Number of Riders"])
    riderws.column_dimensions['A'].width = 4.83
    riderws.column_dimensions['B'].width = 45.83
    riderws.column_dimensions['C'].width = 13.83
    i = 2
    for x in riders:
        riderws.append([coasterDict[x[0]]["Ridership Rank"], x[0],
                        coasterDict[x[0]]["Riders"]])
        if args.colorize:
            if coasterDict[x[0]]["Make"]:
                if coasterDict[x[0]]["Make"] in makeColors.keys():
                    riderws.cell(row=i, column=2).fill = makeColors[coasterDict[x[0]]["Make"]]
                else:
                    riderws.cell(row=i, column=2).fill = makeColors["Other Known Manufacturer"]
            else:
                riderws.cell(row=i, column=2).fill = makeColors[""]
        i += 1
    riderws.freeze_panes = riderws['A2']

    # create and write Mitch Hawker-style mutual rider comparison worksheet
    hawkerWLTws = xl.create_sheet("Coaster vs Coaster Win-Loss-Tie")
    headerRow = ["Rank",""]
    for coaster in results:
        headerRow.append(coasterDict[coaster[0]]["Abbr"])
    hawkerWLTws.append(headerRow)
    hawkerWLTws.column_dimensions['A'].width = 4.83
    hawkerWLTws.column_dimensions['B'].width = 45.83
    for col in range(3, len(results)+3):
        hawkerWLTws.column_dimensions[get_column_letter(col)].width = 12.83
        if args.colorize:
            if coasterDict[results[col-3][0]]["Make"]:
                if coasterDict[results[col-3][0]]["Make"] in makeColors.keys():
                    hawkerWLTws.cell(row=1, column=col).fill = makeColors[coasterDict[results[col-3][0]]["Make"]]
                else:
                    hawkerWLTws.cell(row=1, column=col).fill = makeColors["Other Known Manufacturer"]
            else:
                hawkerWLTws.cell(row=1, column=col).fill = makeColors[""]
    for i in range(0, len(results)):
        resultRow = [coasterDict[results[i][0]]["Overall Rank"], results[i][0]]
        for j in range(0, len(results)):
            coasterA = results[i][0]
            coasterB = results[j][0]
            cellStr = ""
            if coasterA != coasterB:
                if winLossMatrix[coasterA, coasterB]["Wins"] > winLossMatrix[coasterA, coasterB]["Losses"]:
                    cellStr += "W "
                elif winLossMatrix[coasterA, coasterB]["Wins"] < winLossMatrix[coasterA, coasterB]["Losses"]:
                    cellStr += "L "
                else:
                    cellStr += "T "
                cellStr += str(winLossMatrix[coasterA, coasterB]["Wins"]) + "-"
                cellStr += str(winLossMatrix[coasterA, coasterB]["Losses"]) + "-"
                cellStr += str(winLossMatrix[coasterA, coasterB]["Ties"])
            resultRow.append(cellStr)
        hawkerWLTws.append(resultRow)
        if args.colorize:
            if coasterDict[results[i][0]]["Make"]:
                if coasterDict[results[i][0]]["Make"] in makeColors.keys():
                    hawkerWLTws.cell(row=i+2, column=2).fill = makeColors[coasterDict[results[i][0]]["Make"]]
                else:
                    hawkerWLTws.cell(row=i+2, column=2).fill = makeColors["Other Known Manufacturer"]
            else:
                hawkerWLTws.cell(row=i+2, column=2).fill = makeColors[""]
    hawkerWLTws.freeze_panes = hawkerWLTws['C2']
    for col in hawkerWLTws.iter_cols(min_col=3):
        for cell in col:
            cell.font = preferredFixedWidthFont

    spinner.stop()
    print(" ")



# ==================================================
#  OK, let's do this!
# ==================================================

if __name__ == "__main__": # allows us to put main at the beginning
    main()



# ==================================================
#  still to do
# ==================================================

# handle ties: decide which one wins, if possible
# make subsets: rankings of gigas, hypers, types, parks, etc
