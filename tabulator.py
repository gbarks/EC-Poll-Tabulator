#!/usr/bin/env python

# ==========================================================
#  ElloCoaster poll tabulator
#  Contributions from Jim Winslett, Dave Wong, Grant Barker
# ==========================================================

from __future__ import print_function # for Python 2.x users

import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from phicolor import phicolor
    from spinner import Spinner
except:
    print('Could not find "phicolor.py" and/or spinner.py"; exiting...')
    sys.exit()

# global strings for parsing ballots
commentStr = "* "
blankUserField = "-Replace "
startLine = "! DO NOT CHANGE OR DELETE THIS LINE !"

def main():

    # variable defaults that can be set by command line arguments
    minRiders = 6
    blankBallot = "blankballot2017.txt"
    ballotFolder = "ballots2017"

    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        minRiders = int(sys.argv[1])
    if len(sys.argv) > 2:
        blankBallot = sys.argv[2]
    if len(sys.argv) > 3:
        ballotFolder = sys.argv[3]

    if not os.path.isfile(blankBallot):
        print('Blank ballot source "{0}" is not a file; exiting...'.format(blankBallot))
        sys.exit()

    if not os.path.isdir(ballotFolder) or len(os.listdir(ballotFolder)) < 1:
        print('Ballot folder "{0}" does not exist or is empty; exiting...'.format(ballotFolder))
        sys.exit()

    # for each coaster, the number of people who rode it
    riders = {}

    # for each coaster, a list of numbers of the form [wins, losses, ties, totalContests]
    totalWLT = {}

    # create Excel workbook
    xlout = Workbook()
    xlout.active.title = "Coaster Masterlist"

    # preferred fixed-width font
    menlo = Font(name="Menlo")

    light = 186
    multi = 1
    offset = 0.4
    manColors = { # fill colors for certain roller coaster manufacturers/designers
        "CCI"            : PatternFill("solid", fgColor=phicolor(1, light, multi, offset)),
        "GG"             : PatternFill("solid", fgColor=phicolor(2, light, multi, offset)),
        "GCI"            : PatternFill("solid", fgColor=phicolor(3, light, multi, offset)),
        "Intamin"        : PatternFill("solid", fgColor=phicolor(4, light, multi, offset)),
        "PTC"            : PatternFill("solid", fgColor=phicolor(5, light, multi, offset)),
        "Prior & Church" : PatternFill("solid", fgColor=phicolor(6, light, multi, offset)),
        "RMC"            : PatternFill("solid", fgColor=phicolor(7, light, multi, offset)),
        "Locally built"  : PatternFill("solid", fgColor=phicolor(8, light, multi, offset)),
        "Other"          : PatternFill("solid", fgColor="bababa")
    }

    # list of tuples of the form (fullCoasterName, abbreviatedCoasterName)
    coasterList = getCoasterList(blankBallot, riders, totalWLT, xlout.active, menlo, manColors)

    # list of ballot filepaths
    ballotList = getBallotFilepaths(ballotFolder)

    # for each pair of coasters, a list of numbers of the form [wins, losses, ties, winPercent]
    winLossMatrix = createMatrix(coasterList)

    # loop through all the ballot filenames and process each ballot
    voterinfows = xlout.create_sheet("Voter Info (SENSITIVE)")
    voterinfows.append(["Ballot Filename","Name","Email","City","State/Province","Country","Coasters Ridden"])
    voterinfows.column_dimensions['A'].width = 24.83
    voterinfows.column_dimensions['B'].width = 16.83
    voterinfows.column_dimensions['C'].width = 24.83
    for col in ['D','E','F','G']:
        voterinfows.column_dimensions[col].width = 12.83
    for filepath in ballotList:
        voterInfo = processBallot(filepath, coasterList, riders, totalWLT, winLossMatrix)
        if voterInfo:
            voterinfows.append(voterInfo)
    voterinfows.freeze_panes = voterinfows['A2']

    # for each coaster, its win percentage across all pairings
    winPercentage = calculateResults(coasterList, totalWLT, winLossMatrix)

    # sorted lists of tuples of the form (rankedCoaster, relevantNumber)
    finalResults, finalPairs, finalRiders = sortedLists(riders, minRiders, totalWLT, winLossMatrix, winPercentage)

    # write worksheets related to finalResults, finalPairs, finalRiders, and winLossMatrix
    printToFile(xlout, finalResults, finalPairs, finalRiders, winLossMatrix, coasterList, menlo)

    # save the Excel file
    print("Saving...", end=" ")
    spinner = Spinner()
    spinner.start()
    xlout.save("Poll Results.xlsx")
    spinner.stop()
    print('output saved to "Poll Results.xlsx".')



# ==================================================
#  populate list of coasters in the poll
# ==================================================

def getCoasterList(blankBallot, riders, totalWLT, masterlistws, preferredFixedWidthFont, manColors):
    print("Creating list of every coaster on the ballot...", end=" ")
    spinner = Spinner()
    spinner.start()

    # set up Coaster Masterlist worksheet
    masterlistws.append(["Full Coaster Name","Abbrev.","Name","Park","State","Designer"])
    masterlistws.column_dimensions['A'].width = 45.83
    masterlistws.column_dimensions['B'].width = 12.83
    masterlistws.column_dimensions['C'].width = 25.83
    masterlistws.column_dimensions['D'].width = 25.83
    masterlistws.column_dimensions['E'].width = 6.83
    masterlistws.column_dimensions['F'].width = 11.83
    masterlistws['B1'].font = preferredFixedWidthFont
    masterlistws['E1'].font = preferredFixedWidthFont

    coasterList = [] # return value

    #open the blank ballot file
    with open(blankBallot) as f:
        lineNum = 0
        startProcessing = False

        # begin going through the blank ballot line by line
        for line in f:

            sline = line.strip() # strip whitespace from start and end of line
            lineNum += 1

            # skip down the file to the coasters
            if startProcessing == False and sline == startLine:
                startProcessing = True

            # add the coasters to coasterList and the masterlist worksheet
            elif startProcessing == True:

                if commentStr in sline: # skip comment lines (begin with "* ")
                    continue

                elif sline == "": # skip blank lines
                    continue

                else:
                    # break the line into its components: rank, full coaster name, abbreviation
                    words = [x.strip() for x in sline.split(',')]

                    # make sure there are 3 'words' in each line
                    if len(words) != 3:
                        print("Error in {0}, Line {1}: {2}".format(blankBallot, lineNum, line))

                    else:
                        fullName = words[1]
                        abbrName = words[2]

                        # add an entry for the coaster in the dicts
                        riders[fullName] = 0
                        totalWLT[fullName] = [0, 0, 0, 0]

                        manufacturerName = ""
                        if len(words) > 3:
                            manufacturerName = words[3]

                        # add the coaster to the list of coasters on the ballot
                        coasterList.append((fullName, abbrName, manufacturerName))

                        # extract park and state/country information from fullName to write to worksheet
                        subwords = [x.strip() for x in fullName.split('-')]
                        if len(subwords) != 3:
                            masterlistws.append([fullName,abbrName,manufacturerName])
                        else:
                            masterlistws.append([fullName,abbrName,subwords[0],subwords[1],subwords[2],manufacturerName])
                            masterlistws.cell(row=len(coasterList)+1, column=5).font = preferredFixedWidthFont
                        masterlistws.cell(row=len(coasterList)+1, column=2).font = preferredFixedWidthFont

                        if manufacturerName and manufacturerName in manColors.keys():
                            masterlistws.cell(row=len(coasterList)+1, column=6).fill = manColors[manufacturerName]
                        else:
                            masterlistws.cell(row=len(coasterList)+1, column=6).fill = manColors["Other"]

    masterlistws.freeze_panes = masterlistws['A2']
    spinner.stop()
    print("{0} coasters on the ballot.".format(len(coasterList)))
    return coasterList



# ==================================================
#  import filepaths of ballots
# ==================================================

def getBallotFilepaths(ballotFolder):
    print("Getting the filepaths of submitted ballots...", end=" ")
    spinner = Spinner()
    spinner.start()

    ballotList = []
    for file in os.listdir(ballotFolder):
        if file.endswith(".txt"):
            ballotList.append(os.path.join(ballotFolder, file))

    spinner.stop()
    print("{0} ballots submitted.".format(len(ballotList)))
    return ballotList



# ==================================================
#  create win/loss matrix
# ==================================================

def createMatrix(coasterList):
    print("Creating the win/loss matrix...", end=" ")
    spinner = Spinner()
    spinner.start()

    winLossMatrix = {}
    for row in coasterList:
        for col in coasterList:
            winLossMatrix[row[0],col[0]] = [0, 0, 0, 0.0]

    spinner.stop()
    print("{0} pairings.".format(len(winLossMatrix)))
    return winLossMatrix



# ================================================================
#  read a ballot (just ONE ballot)
#
#  you need a loop to call this function for each ballot filename
# ================================================================

def processBallot(filepath, coasterList, riders, totalWLT, winLossMatrix):
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
                    print("Error in {0}, Line {1}: {2}".format(blankBallot, lineNum, line))

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
                    if [coasterName in x[0] for x in coasterList]:
                        creditNum += 1
                        riders[coasterName] += 1

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
                totalWLT[coasterA][3] += 1 # increment number of contests

                # if the coasters have the same ranking, call it a tie
                if coasterAndRank[coasterA] == coasterAndRank[coasterB]:
                    winLossMatrix[coasterA, coasterB][2] += 1
                    totalWLT[coasterA][2] += 1

                # if coasterA outranks coasterB (the rank's number is lower), call it a win for coasterA
                elif coasterAndRank[coasterA] < coasterAndRank[coasterB]:
                    winLossMatrix[coasterA, coasterB][0] += 1
                    totalWLT[coasterA][0] += 1

                # if not a tie nor a win, it must be a loss
                else:
                    winLossMatrix[coasterA, coasterB][1] += 1
                    totalWLT[coasterA][1] += 1

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

def calculateResults(coasterList, totalWLT, winLossMatrix):
    print("Calculating results...", end=" ")
    spinner = Spinner()
    spinner.start()

    # iterate through all the pairs in the matrix
    for coasterA in coasterList:
        for coasterB in coasterList:

            x = coasterA[0]
            y = coasterB[0]

            if x == y: # skip a coaster paired with itself
                continue

            wins = winLossMatrix[x, y][0]
            loss = winLossMatrix[x, y][1]
            ties = winLossMatrix[x, y][2]
            numContests = wins + loss + ties

            # if this pair of coasters had mutual riders and there were ties, calculate with this formula
            if ties != 0 and numContests > 0:
                # formula: wins + half the ties divided by the number of times they competed against each other
                # Multiply that by 100 to get the percentage, then round to three digits after the decimal
                winLossMatrix[x, y][3] = round((((wins + (ties / 2)) / numContests)) * 100, 3)
            # if this pair had mutual riders, but there were no ties, use this formula
            elif numContests > 0:
                winLossMatrix[x, y][3] = round(((wins / numContests)) * 100, 3)

    winPercentage = {} # return value

    # all those calculations we just did for each pair of coasters, now do for each coaster by itself
    # tallying up ALL the contests it had, not just the pairwise contests
    # this will give the total overall win percentage for each coaster, which will be used to determine
    # the final ranking of all the coasters
    for coaster in coasterList:

        x = coaster[0]

        if totalWLT[x][2] > 0 and totalWLT[x][3] > 0: # if numTies and numContests > 0
            winPercentage[x] = round((((totalWLT[x][0] + (totalWLT[x][2]/2)) / totalWLT[x][3])) * 100, 3)

        elif totalWLT[x][3] > 0: # if numTies == 0 and numContests > 0
            winPercentage[x] = round(((totalWLT[x][0] / totalWLT[x][3])) * 100, 3)

    spinner.stop()
    print(" ")

    return winPercentage



# ==================================================
#  create sorted list of coasters by win pct and
#    sorted list of coasters by pairwise win pct
# ==================================================

def sortedLists(riders, minRiders, totalWLT, winLossMatrix, winPercentage):
    print("Sorting the results...", end=" ")
    spinner = Spinner()
    spinner.start()

    results = []
    numRiders = []
    pairPercents = []

    # iterate through the winPercentage dict by coasters
    for i in winPercentage.keys():
        numRiders.append((i, riders[i]))
        if int(riders[i]) >= int(minRiders):

            # values are: "Rank", "Coaster", "Win %", "Total Wins", "Total Losses", "Total Ties"
            results.append((i, winPercentage[i], totalWLT[i][0], totalWLT[i][1], totalWLT[i][1]))

    # iterate through the winLossMatrix dict by coaster pairings
    for i in winLossMatrix.keys():

        # values are: "Rank", "coasterA", "coasterB", "Win %", "Wins", "Losses", "Ties"        
        pairPercents.append((i, winLossMatrix[i][3], winLossMatrix[i][0], winLossMatrix[i][1], winLossMatrix[i][2]))

    # sort lists by win percentages and ridership
    sortedResults = sorted(results, key=lambda x: x[1], reverse=True)
    sortedPairs = sorted(pairPercents, key=lambda x: x[1], reverse=True)
    sortedRiders = sorted(numRiders, key=lambda x: x[1], reverse=True)

    spinner.stop()
    print(" ")

    return sortedResults, sortedPairs, sortedRiders



# ==================================================
#  print everything to a file
# ==================================================

def printToFile(xl, results, pairs, riders, winLossMatrix, coasterList, preferredFixedWidthFont):
    print("Writing the results...", end=" ")
    spinner = Spinner()
    spinner.start()

    # create and write primary results worksheet
    resultws = xl.create_sheet("Ranked Results")
    resultws.append(["Rank","Coaster","Win Percentage","Total Wins","Total Losses","Total Ties"])
    resultws.column_dimensions['A'].width = 4.83
    resultws.column_dimensions['B'].width = 45.83
    resultws.column_dimensions['C'].width = 12.83
    resultws.column_dimensions['D'].width = 8.83
    resultws.column_dimensions['E'].width = 9.83
    resultws.column_dimensions['F'].width = 7.83
    for i in range(0, len(results)):
        resultws.append([i+1, results[i][0], results[i][1], results[i][2], results[i][3], results[i][4]])
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
    for i in range(0, len(pairs)):
        pairws.append([i+1, pairs[i][0][0], pairs[i][0][1], pairs[i][1], pairs[i][2], pairs[i][3], pairs[i][4]])
    pairws.freeze_panes = pairws['A2']

    # create and write ridership worksheet
    riderws = xl.create_sheet("Number of Riders")
    riderws.append(["Rank","Coaster","Number of Riders"])
    riderws.column_dimensions['A'].width = 4.83
    riderws.column_dimensions['B'].width = 45.83
    riderws.column_dimensions['C'].width = 13.83
    for i in range(0, len(riders)):
        riderws.append([i+1, riders[i][0], riders[i][1]])
    riderws.freeze_panes = riderws['A2']

    # create and write Mitch Hawker-style mutual rider comparison worksheet
    hawkerWLTws = xl.create_sheet("Coaster vs Coaster Win-Loss-Tie")
    headerRow = ["Rank",""]
    orderAbbr = []
    for coaster in results:
        for abbr in coasterList:
            if coaster[0] == abbr[0]:
                headerRow.append(abbr[1])
                orderAbbr.append((abbr[0], abbr[1])) # coasterList sorted by win% (as in resultws)
                break
    hawkerWLTws.append(headerRow)
    hawkerWLTws.column_dimensions['A'].width = 4.83
    hawkerWLTws.column_dimensions['B'].width = 45.83
    for col in range(3, len(orderAbbr)+3):
        hawkerWLTws.column_dimensions[get_column_letter(col)].width = 12.83
    for i in range(0, len(orderAbbr)):
        resultRow = [i+1, orderAbbr[i][0]]
        for j in range(0, len(orderAbbr)):
            coasterA = orderAbbr[i][0]
            coasterB = orderAbbr[j][0]
            cellStr = ""
            if coasterA != coasterB:
                if winLossMatrix[coasterA, coasterB][0] > winLossMatrix[coasterA, coasterB][1]:
                    cellStr += "W "
                elif winLossMatrix[coasterA, coasterB][0] < winLossMatrix[coasterA, coasterB][1]:
                    cellStr += "L "
                else:
                    cellStr += "T "
                cellStr += str(winLossMatrix[coasterA, coasterB][0]) + "-"
                cellStr += str(winLossMatrix[coasterA, coasterB][1]) + "-"
                cellStr += str(winLossMatrix[coasterA, coasterB][2])
            resultRow.append(cellStr)
        hawkerWLTws.append(resultRow)
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

# handle ties: decide which one wins, if possible. If still tied, rank them the same
# convert pairwise results into numpy array or pandas dataframe
# make subsets: rankings of gigas, hypers, types, parks, etc
