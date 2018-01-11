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

useSpinner = True
try:
    from spinner import Spinner
except:
    useSpinner = False

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
parser.add_argument("-m", "--minRiders", type=int, default=10,
                    help="specify minimum number of riders for a coaster to rank")
parser.add_argument("-o", "--outfile", default="Poll Results.xlsx",
                    help="specify name of output .xlsx file")
parser.add_argument("-c", "--colorize", action="store_true",
                    help="color coaster labels by designer in output spreadsheet")
parser.add_argument("-i", "--includeVoterInfo", action="store_true",
                    help="include sensitive voter data in output spreadsheet")
parser.add_argument("-r", "--botherRCDB", action="store_true",
                    help="bother RCDB to grab metadata from links in blankBallot")
parser.add_argument("-v", "--verbose", action="count", default=0,
                    help="print data as it's processed; duplicate for more detail")

args = parser.parse_args()

if not os.path.isfile(args.blankBallot):
    print('Blank ballot source "{0}" is not a file; exiting...'.format(args.blankBallot))
    sys.exit()

if not os.path.isdir(args.ballotFolder) or len(os.listdir(args.ballotFolder)) < 1:
    print('Ballot folder "{0}" does not exist or is empty; exiting...'.format(args.ballotFolder))
    sys.exit()

if args.outfile[-5:] != ".xlsx":
    args.outfile += ".xlsx"

# only import HTML tools if using '-r' flag and running in Python 3
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

    manuColors = { # fill colors for certain roller coaster manufacturers/designers
        "Custom Coasters International, Inc."   : PatternFill("solid", fgColor="fdb2b3"), # red
        "Dinn Corporation"                      : PatternFill("solid", fgColor="fed185"), # orange
        "The Gravity Group, LLC"                : PatternFill("solid", fgColor="fffd87"), # yellow
        "Great Coasters International"          : PatternFill("solid", fgColor="cde4cd"), # green
        "Intamin Amusement Rides"               : PatternFill("solid", fgColor="b2b4fd"), # blue
        "National Amusement Device Company"     : PatternFill("solid", fgColor="c9b3d8"), # purple
        "Philadelphia Toboggan Coasters, Inc."  : PatternFill("solid", fgColor="fecdfe"), # pink
        "Rocky Mountain Construction"           : PatternFill("solid", fgColor="ceffff"), # cyan
        "Roller Coaster Corporation of America" : PatternFill("solid", fgColor="e8d9c6"), # light brown
        "S&S Worldwide"                         : PatternFill("solid", fgColor="cb999a"), # dark brown
        "Vekoma"                                : PatternFill("solid", fgColor="fefe9e"), # amber
        "Other Known Manufacturer"              : PatternFill("solid", fgColor="cccccc"), # dark gray
        ""                                      : PatternFill("solid", fgColor="eeeeee") # light gray
    }

    # list of tuples of the form (fullCoasterName, abbreviatedCoasterName)
    coasterDict = getCoasterDict(xlout.active, menlo, manuColors)

    # create color key for manuColors
    if args.colorize:
        coasterdesignerws = xlout.create_sheet("Coaster Designer Color Key")
        i = 1
        for designer in manuColors.keys():
            if designer:
                coasterdesignerws.append([designer])
            else:
                coasterdesignerws.append(["Other [Unknown]"])
            coasterdesignerws.cell(row=i, column=1).fill = manuColors[designer]
            i += 1
        coasterdesignerws.column_dimensions['A'].width = 30.83

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
    finalResults, finalPairs = sortedLists(coasterDict, winLossMatrix)

    # write worksheets related to finalResults, finalPairs, and winLossMatrix
    printToFile(xlout, finalResults, finalPairs, winLossMatrix, coasterDict, menlo, manuColors)

    # save the Excel file
    print("Saving...", end=" ")
    if useSpinner:
        spinner = Spinner()
        spinner.start()
    xlout.save(args.outfile)
    if useSpinner:
        spinner.stop()
    print('output saved to "Poll Results.xlsx".')



# ==================================================
#  function for getting manufacturer's color
# ==================================================

def colorizeRow(worksheet, rowNum, colList, coasterDict, coaster, colorDict):
    if args.colorize:
        if coasterDict[coaster]["Manufacturer"]:
            if coasterDict[coaster]["Manufacturer"] in colorDict.keys():
                for c in colList:
                    worksheet.cell(row=rowNum, column=c).fill = colorDict[coasterDict[coaster]["Manufacturer"]]
            else:
                for c in colList:
                    worksheet.cell(row=rowNum, column=c).fill = colorDict["Other Known Manufacturer"]
        else:
            for c in colList:
                    worksheet.cell(row=rowNum, column=c).fill = colorDict[""]



# ==================================================
#  populate dictionary of coasters in the poll
# ==================================================

def getCoasterDict(masterlistws, preferredFixedWidthFont, manuColors):
    if not args.botherRCDB:
        print("Creating list of every coaster on the ballot...", end=" ")
        if useSpinner:
            spinner = Spinner()
            spinner.start()

    # set up Coaster Masterlist worksheet
    headerRow = ["Full Coaster Name", "Abbrev.", "Name", "Park", "State", "RCDB Link"]
    if args.botherRCDB:
        headerRow.extend(["Designer/Manufacturer", "Year"])
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

                        # set default RCDB-pulled info
                        coasterDict[fullName]["RCDB"] = ""
                        coasterDict[fullName]["Manufacturer"] = ""
                        coasterDict[fullName]["Year"] = ""

                        # check if an RCDB link is provided
                        if len(words) > 3:
                            coasterDict[fullName]["RCDB"] = words[3]
                            rowVals.append('=HYPERLINK("{0}", "{1}")'.format(words[3], words[3][8:]))

                            # open URL if '-r' flag is used
                            if args.botherRCDB:
                                response = urlopen(coasterDict[fullName]["RCDB"])
                                html = response.read()
                                soup = BeautifulSoup(html, 'lxml')

                                # find a "Designer/Manufacturer" if available
                                for x in soup.body.findAll('div', attrs={'class':'scroll'}):

                                    # try the "Make" field at the top of the page
                                    if "Make: " in x.text:
                                        subtext = x.text.split("Make: ", 1)[1]
                                        if "Model: " in subtext:
                                            subtext = subtext.split("Model: ", 1)[0]
                                        coasterDict[fullName]["Manufacturer"] = subtext
                                        break

                                # if the "Make" field didn't exist or linked to an unknown manufacturer, try "Designer" field
                                if (coasterDict[fullName]["Manufacturer"] == "" or
                                    coasterDict[fullName]["Manufacturer"] not in manuColors.keys()):
                                    for x in soup.body.findAll('table', attrs={'class':'objDemoBox'}):
                                        if "Designer:" in x.text:
                                            subtext = x.text.split("Designer:", 1)[1]
                                            if "Installer:" in subtext:
                                                subtext = subtext.split("Installer:", 1)[0]
                                            if "Musical Score:" in subtext:
                                                subtext = subtext.split("Musical Score:", 1)[0]
                                            if "Construction Supervisor:" in subtext:
                                                subtext = subtext.split("Construction Supervisor:", 1)[0]

                                            # if a known manufacturer is a substring of subtext, use that
                                            alreadyKnownManu = next((y for y in manuColors.keys() if y in subtext), False)
                                            if alreadyKnownManu:
                                                coasterDict[fullName]["Manufacturer"] = alreadyKnownManu

                                            # otherwise, use the provided "Designer"
                                            elif coasterDict[fullName]["Manufacturer"] == "" and subtext != "":
                                                coasterDict[fullName]["Manufacturer"] = subtext
                                            break

                                # exception for Gravity Group, who has two names on RCDB for some reason
                                if coasterDict[fullName]["Manufacturer"] == "Gravitykraft Corporation":
                                    coasterDict[fullName]["Manufacturer"] = "The Gravity Group, LLC"

                                rowVals.append(coasterDict[fullName]["Manufacturer"])

                                # find an opening year if available
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

                                # print RCDB-pulled info as it's acquired
                                print("{0},   \t{1},\t{2}".format(
                                    abbrName, coasterDict[fullName]["Year"], coasterDict[fullName]["Manufacturer"]))

                        # final values associated with this coaster
                        coasterDict[fullName]["Abbr"] = abbrName

                        # variable values associated with this coaster
                        coasterDict[fullName]["Riders"] = 0
                        coasterDict[fullName]["Total Wins"] = 0
                        coasterDict[fullName]["Total Losses"] = 0
                        coasterDict[fullName]["Total Ties"] = 0
                        coasterDict[fullName]["Total Win Percentage"] = 0.0
                        coasterDict[fullName]["Average Win Percentage"] = 0.0
                        coasterDict[fullName]["Win Percentages"] = []
                        coasterDict[fullName]["Overall Rank"] = 0
                        coasterDict[fullName]["Tied Coasters"] = []

                        # append the row values and set styles
                        masterlistws.append(rowVals)
                        masterlistws.cell(row=len(coasterDict)+1, column=5).font = preferredFixedWidthFont
                        masterlistws.cell(row=len(coasterDict)+1, column=2).font = preferredFixedWidthFont
                        if coasterDict[fullName]["RCDB"]:
                            masterlistws.cell(row=len(coasterDict)+1, column=6).style = "Hyperlink"

                        colorizeRow(masterlistws, len(coasterDict)+1, [1,2,7], coasterDict, fullName, manuColors)

    masterlistws.freeze_panes = masterlistws['A2']
    if useSpinner and not args.botherRCDB:
        spinner.stop()
    print("{0} coasters on the ballot.".format(len(coasterDict)))
    return coasterDict



# ==================================================
#  import filepaths of ballots
# ==================================================

def getBallotFilepaths():
    print("Getting the filepaths of submitted ballots...", end=" ")
    if useSpinner:
        spinner = Spinner()
        spinner.start()

    ballotList = []
    for file in os.listdir(args.ballotFolder):
        if file.endswith(".txt"):
            ballotList.append(os.path.join(args.ballotFolder, file))

    if useSpinner:
        spinner.stop()
    print("{0} ballots submitted.".format(len(ballotList)))
    return ballotList



# ==================================================
#  create win/loss matrix
# ==================================================

def createMatrix(coasterDict):
    print("Creating the win/loss matrix...", end=" ")
    if useSpinner:
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

    if useSpinner:
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
    if useSpinner:
        spinner = Spinner()
        spinner.start()

    if args.verbose > 0:
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
                    coasterDict[coasterA]["Win Percentages"].append(winLossMatrix[coasterA, coasterB]["Win Percentage"])

                    # only print pairwise results with '-vv' flag
                    if args.verbose > 1:
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
            coasterDict[x]["Average Win Percentage"] = sum(coasterDict[x]["Win Percentages"]) / len(coasterDict[x]["Win Percentages"])

            # print singular results with just a '-v' flag
            if args.verbose > 0:
                print("{0},\tWins: {1},\tTies: {2},\t#Con: {3},\tWin%: {4}, \tAvgWin%: {5}".format(
                    coasterDict[x]["Abbr"], numWins, numTies, numContests,
                    coasterDict[x]["Total Win Percentage"], coasterDict[x]["Average Win Percentage"]))

    if useSpinner:
        spinner.stop()
    print(" ")



# ==================================================
#  add to "Tied Coasters" variable in coasterDict
# ==================================================

def markTies(coasterDict, winLossMatrix, tiedCoasters):
    for coasterA in tiedCoasters:
        coastersTiedWithA = []
        for coasterB in tiedCoasters:
            if coasterA != coasterB:
                coastersTiedWithA.append(coasterB)
        coasterDict[coasterA]["Tied Coasters"] = coastersTiedWithA

    # print Mitch Hawker-style pairwise matchups between tied coasters with '-v' flag
    if args.verbose > 0:
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
    if useSpinner:
        spinner = Spinner()
        spinner.start()

    results = []
    pairPercents = []

    # iterate through coasterDict by coasters
    for coasterName in coasterDict.keys():
        if int(coasterDict[coasterName]["Riders"]) >= int(args.minRiders):
            results.append((coasterName,
                            coasterDict[coasterName]["Total Win Percentage"],
                            coasterDict[coasterName]["Average Win Percentage"]))

    # iterate through winLossMatrix by coaster pairings
    for coasterPair in winLossMatrix.keys():      
        pairPercents.append((coasterPair, winLossMatrix[coasterPair]["Win Percentage"]))

    # sort lists by win percentages
    sortedResults = sorted(results, key=lambda x: x[1], reverse=True)
    sortedPairs = sorted(pairPercents, key=lambda x: x[1], reverse=True)

    if args.verbose > 0:
        print("")

    # determine rankings including ties for overall rank
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
        if args.verbose > 0:
            print("Rank: {0},\tVal: {1},  \tCoaster: {2}".format(coasterDict[x[0]]["Overall Rank"], x[1], x[0]))
    if len(tiedCoasters) > 1: # in case last few coasters were tied
        markTies(coasterDict, winLossMatrix, tiedCoasters)

    # determine rankings including ties for pairwise ranks
    overallRank = 0
    curRank = 0
    curValue = 0.0
    for x in sortedPairs:
        overallRank += 1
        if x[1] != curValue:
            curRank = overallRank
            curValue = x[1]
        winLossMatrix[x[0][0], x[0][1]]["Pairwise Rank"] = curRank

    if useSpinner:
        spinner.stop()
    print(" ")

    return sortedResults, sortedPairs



# ==================================================
#  print everything to a file
# ==================================================

def printToFile(xl, results, pairs, winLossMatrix, coasterDict, preferredFixedWidthFont, manuColors):
    print("Writing the results...", end=" ")
    if useSpinner:
        spinner = Spinner()
        spinner.start()

    # create and write primary results worksheet
    resultws = xl.create_sheet("Ranked Results")
    headerRow = ["Rank","Coaster","Total Win Percentage","Average Win Percentage",
                 "Total Wins","Total Losses","Total Ties","Number of Riders"]
    if args.botherRCDB:
        headerRow.extend(["Designer/Manufacturer", "Year"])
    resultws.append(headerRow)
    resultws.column_dimensions['A'].width = 4.83
    resultws.column_dimensions['B'].width = 45.83
    resultws.column_dimensions['C'].width = 16.83
    resultws.column_dimensions['D'].width = 18.83
    resultws.column_dimensions['E'].width = 8.83
    resultws.column_dimensions['F'].width = 9.83
    resultws.column_dimensions['G'].width = 7.83
    resultws.column_dimensions['H'].width = 13.83
    resultws.column_dimensions['I'].width = 23.83
    resultws.column_dimensions['J'].width = 8.83
    i = 2
    for x in results:
        resultws.append([coasterDict[x[0]]["Overall Rank"], x[0],
                         coasterDict[x[0]]["Total Win Percentage"],
                         coasterDict[x[0]]["Average Win Percentage"],
                         coasterDict[x[0]]["Total Wins"],
                         coasterDict[x[0]]["Total Losses"],
                         coasterDict[x[0]]["Total Ties"],
                         coasterDict[x[0]]["Riders"],
                         coasterDict[x[0]]["Manufacturer"],
                         coasterDict[x[0]]["Year"]])
        colorizeRow(resultws, i, [2,9], coasterDict, x[0], manuColors)
        i += 1
    resultws.freeze_panes = resultws['A2']

    # append coasters that weren't ranked to the bottom of results worksheet
    for x in coasterDict.keys():
        if x not in [y[0] for y in results] and coasterDict[x]["Riders"] > 0:
            resultws.append(["N/A", x,
                             "Insufficient Riders, {0}".format(coasterDict[x]["Total Win Percentage"]),
                             "Insufficient Riders, {0}".format(coasterDict[x]["Average Win Percentage"]),
                             coasterDict[x]["Total Wins"],
                             coasterDict[x]["Total Losses"],
                             coasterDict[x]["Total Ties"],
                             coasterDict[x]["Riders"],
                             coasterDict[x]["Manufacturer"],
                             coasterDict[x]["Year"]])
            colorizeRow(resultws, i, [2,9], coasterDict, x, manuColors)
            i += 1

    # append coasters that weren't ridden to the bottom of results worksheet
    for x in coasterDict.keys():
        if x not in [y[0] for y in results] and coasterDict[x]["Riders"] == 0:
            resultws.append(["N/A", x, "No Riders", "No Riders",
                             coasterDict[x]["Total Wins"],
                             coasterDict[x]["Total Losses"],
                             coasterDict[x]["Total Ties"],
                             coasterDict[x]["Riders"],
                             coasterDict[x]["Manufacturer"],
                             coasterDict[x]["Year"]])
            colorizeRow(resultws, i, [2,9], coasterDict, x, manuColors)
            i += 1

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
        colorizeRow(pairws, i, [2], coasterDict, x[0][0], manuColors)
        colorizeRow(pairws, i, [3], coasterDict, x[0][1], manuColors)
        i += 1
    pairws.freeze_panes = pairws['A2']

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
        colorizeRow(hawkerWLTws, 1, [col], coasterDict, results[col-3][0], manuColors)
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
        colorizeRow(hawkerWLTws, i+2, [2], coasterDict, results[i][0], manuColors)
    hawkerWLTws.freeze_panes = hawkerWLTws['C2']
    for col in hawkerWLTws.iter_cols(min_col=3):
        for cell in col:
            cell.font = preferredFixedWidthFont

    # create and write Mitch Hawker-style mutual rider comparison worksheet sorted by Avg Win Percentage
    resortedResults = sorted(results, key=lambda x: x[2], reverse=True)
    hawkerWLT2 = xl.create_sheet("CvC Win-Loss-Tie by AvgWin%")
    headerRow = ["Rank",""]
    for coaster in resortedResults:
        headerRow.append(coasterDict[coaster[0]]["Abbr"])
    hawkerWLT2.append(headerRow)
    hawkerWLT2.column_dimensions['A'].width = 4.83
    hawkerWLT2.column_dimensions['B'].width = 45.83
    for col in range(3, len(resortedResults)+3):
        hawkerWLT2.column_dimensions[get_column_letter(col)].width = 12.83
        colorizeRow(hawkerWLT2, 1, [col], coasterDict, resortedResults[col-3][0], manuColors)
    for i in range(0, len(resortedResults)):
        resultRow = [coasterDict[resortedResults[i][0]]["Overall Rank"], resortedResults[i][0]]
        for j in range(0, len(resortedResults)):
            coasterA = resortedResults[i][0]
            coasterB = resortedResults[j][0]
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
        hawkerWLT2.append(resultRow)
        colorizeRow(hawkerWLT2, i+2, [2], coasterDict, resortedResults[i][0], manuColors)
    hawkerWLT2.freeze_panes = hawkerWLT2['C2']
    for col in hawkerWLT2.iter_cols(min_col=3):
        for cell in col:
            cell.font = preferredFixedWidthFont

    # create and write sheet to compare where coasters would have ranked if sorted by AvgWin%
    comparisonws = xl.create_sheet("TotalWin% vs AvgWin% Rankings")
    comparisonws.append(["Coaster","TotalWin% Rank","AvgWin% Rank","Difference"])
    comparisonws.column_dimensions['A'].width = 45.83
    comparisonws.column_dimensions['B'].width = 12.83
    comparisonws.column_dimensions['C'].width = 12.83
    for i in range(0, len(resortedResults)):
        coaster = resortedResults[i][0]
        oldRank = coasterDict[coaster]["Overall Rank"]
        newRank = i+1
        diff = oldRank - newRank
        if diff == 0:
            comparisonws.append([coaster, oldRank, newRank, ""])
        else:
            comparisonws.append([coaster, oldRank, newRank, diff])
            if diff <= -16:
                diffColor = "ff0000" # maximum red
            elif diff < 0:
                diffColor = "ff{0}{0}".format(hex(256 + 16 * diff)[2:]) # gradual red
            elif diff < 16:
                diffColor = "{0}ff{0}".format(hex(256 - 16 * diff)[2:]) # gradual green
            else:
                diffColor = "00ff00" # maximum green
            comparisonws.cell(row=i+2, column=4).fill = PatternFill("solid", fgColor=diffColor)
        colorizeRow(comparisonws, i+2, [1], coasterDict, coaster, manuColors)
    comparisonws.freeze_panes = comparisonws['A2']

    if useSpinner:
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
