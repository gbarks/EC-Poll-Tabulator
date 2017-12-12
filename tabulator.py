# ==================================================
#  ElloCoaster poll tabulator
#  Contributions from Jim Winslett, Grant Barker
# ==================================================

import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font

# exit if the first arg isn't an integer
if len(sys.argv) < 2 or  not sys.argv[1].isdigit():
    print "Please specify the minimum number of riders per comparison as the 1st arg."
    sys.exit()

# global strings for parsing ballots
commentStr = "* "
blankUserField = "-Replace "
startLine = "! DO NOT CHANGE OR DELETE THIS LINE !"

def main():
    xlout = Workbook()
    menlo = Font(name="Menlo")

    blankBallot = "blankballot2017.txt"
    ballotFolder = "ballots2017"

    # list of tuples representing every coaster on the ballot
    coasterList = []

    # for each coaster on the ballot, the number of people who rode that coaster
    riders = {}

    # unsorted list of the coaster and the number of riders it had
    numRiders = []

    # the number of total credits for all the voters
    totalCredits = 0

    # sorted version of the above
    sortedRiders = []

    # the total number of wins, losses, ties, and totalContest in a list for each coaster
    totalWLT = {}

    # unsorted list of coasters and their total win percentages
    results = []

    # sorted version of above
    sortedResults = []

    # unsorted list of every pair of coasters with the pair's win percentage
    pairsList = []

    # sorted version of above
    sortedPairs = []


    xlout.active.title = "Coaster Masterlist"
    getCoasterList(blankBallot, coasterList, riders, totalWLT, xlout.active, menlo)

    ballotList = getBallotFilepaths(ballotFolder)

    # for each pair of coasters, a string containing w, l, or t
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
        voterInfo = processBallot(filepath, coasterList, riders, totalCredits, totalWLT, winLossMatrix)
        if voterInfo:
            voterinfows.append(voterInfo)
    voterinfows.freeze_panes = voterinfows['A2']

    winPercentage = calculateResults(coasterList, totalWLT, winLossMatrix)

    for coaster in winLossMatrix.keys():
        if "Lake Compounce" in coaster[0] and winLossMatrix[coaster][0] > 0:
            print coaster, winLossMatrix[coaster]

    # for coaster in winLossMatrix.keys():
    #     if winLossMatrix[coaster][0] > 1:
    #         print coaster, winLossMatrix[coaster]

    # for coaster in winPercentage.keys():
    #     if winPercentage[coaster] > 0:
    #         print " ->", coaster, winPercentage[coaster]

    minRiders = int(sys.argv[1])

    # sortedLists()

    # printToFile()

    xlout.save("Poll Results.xlsx")
    print 'Output saved to "Poll Results.xlsx".'



# ==================================================
#  populate list of coasters in the poll
# ==================================================

def getCoasterList(blankBallot, coasterList, riders, totalWLT, masterlistws, preferredFixedWidthFont):
    print "Creating list of every coaster on the ballot...",

    masterlistws.append(["Full Coaster Name","Abbrev.","Name","Park","State"])
    masterlistws.column_dimensions['A'].width = 45.83
    masterlistws.column_dimensions['B'].width = 12.83
    masterlistws.column_dimensions['C'].width = 25.83
    masterlistws.column_dimensions['D'].width = 25.83
    masterlistws.column_dimensions['E'].width = 6.83
    masterlistws['B1'].font = preferredFixedWidthFont
    masterlistws['E1'].font = preferredFixedWidthFont

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

                        # add the coaster to the list of coasters on the ballot
                        coasterList.append((fullName, abbrName))
                        subwords = [x.strip() for x in fullName.split('-')]
                        if len(subwords) != 3:
                            masterlistws.append([fullName,abbrName])
                        else:
                            masterlistws.append([fullName,abbrName,subwords[0],subwords[1],subwords[2]])
                            masterlistws.cell(row=len(coasterList)+1, column=5).font = preferredFixedWidthFont
                        masterlistws.cell(row=len(coasterList)+1, column=2).font = preferredFixedWidthFont

    masterlistws.freeze_panes = masterlistws['A2']
    print len(coasterList), "coasters on the ballot."



# ==================================================
#  import filepaths of ballots
# ==================================================

def getBallotFilepaths(ballotFolder):
    print "Getting the filepaths of submitted ballots...",
    ballotList = []
    for file in os.listdir(ballotFolder):
        if file.endswith(".txt"):
            ballotList.append(os.path.join(ballotFolder, file))
    print len(ballotList), "ballots submitted."
    return ballotList



# ==================================================
#  create win/loss matrix
# ==================================================

def createMatrix(coasterList):
    print "Creating the win/loss matrix...",

    # create a matrix of blank strings for each pair of coasters
    # these strings will later contain w, l, t for each matchup
    #   followed by the respective w, l, t numbers
    winLossMatrix = {}
    for row in coasterList:
        for col in coasterList:
            winLossMatrix[row[0],col[0]] = [0, 0, 0, 0.0]
    print len(winLossMatrix), "pairings."
    return winLossMatrix



# ================================================================
#  read a ballot (just ONE ballot)
#
#  you need a loop to call this function for each ballot filename
# ================================================================

def processBallot(filepath, coasterList, riders, totalCredits, totalWLT, winLossMatrix):
    filename = os.path.basename(filepath)
    print("Processing ballot: {0}".format(filename))

    voterInfo = [filename, "", "", "", "", ""]
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
                    voterInfo[infoField] = sline
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

    # don't tally the ballot if there were any errors
    if error:
        print "Error encountered. File {0} not added.".format(filename)
        return []

    # add this voter's credit count to the total credits
    totalCredits += creditNum

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

    print " ->",

    for i in range(1,len(voterInfo)):
        if voterInfo[i] != "":
            print "{0},".format(voterInfo[i]),

    print "CC: {0}".format(creditNum)

    voterInfo.append(creditNum)

    return voterInfo



# ========================================================
#  calculate results
#
#  no need to loop through this, since it calculates with
#    numbers gathered when the ballots were processed
# ========================================================

def calculateResults(coasterList, totalWLT, winLossMatrix):
    print("Calculating results...")

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

    winPercentage = {}

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

    return winPercentage



# ==================================================
#  create sorted list of coasters by win pct and
#    sorted list of coasters by pairwise win pct
# ==================================================

def sortedLists():
    print("Sorting the results")

    global results
    global pairsList
    global sortedResults
    global sortedPairs
    global riders
    global sortedRiders

    # iterate through the winPercentage dict by keys
    for i in winPercentage.keys():
        # pull out just the single-coaster keys for the total win percentage and number of riders
        if i in coasterList:
            numRiders.append((i, riders[i]))
            if int(riders[i]) >= int(minRiders):
                results.append((i, winPercentage[i]))
            else:
                continue

        # the rest are pairs keys and pairwise win percentages, they go in their own list
        else:
            pairsList.append((i, winPercentage[i]))
    # now sort both lists by the win percentages, highest numbers first
    sortedResults = sorted(results, key=lambda x: x[1], reverse=True)
    sortedPairs = sorted(pairsList, key=lambda x: x[1], reverse=True)
    sortedRiders = sorted(numRiders, key=lambda x: x[1], reverse=True)

    return sortedRiders, sortedPairs, sortedResults



# ==================================================
#  print everything to a file
# ==================================================

def printToFile():

    with open("numriders2017.txt", "w") as f:

        f.write("Coasters by number of riders\n")
        for i in range(0, len(sortedRiders)):

            f.write(str(i+1) + ": " + str(sortedRiders[i]) + "\n")
        f.write("\n")

    with open("rankedresults2017.txt", "w") as f:
        f.write("Total number of valid ballots received:" + str(len(ballotList)) + "\n")
        f.write("total number of coasters on ballot:" + str(len(coasterList))+ "\n")
        f.write("average number of coasters ridden by each voter:" + str(int(totalCredits/len(ballotList)))+ "\n")
        f.write("Coasters by ranking\n")
        for i in range(0,len(sortedResults)):
            f.write(str(i+1) + ": " + str(sortedResults[i]) + "\n")

        f.write("\n")

    with open("pairsrank2017.txt", "w") as f:
        f.write("Pairs by ranking\n")
        for i in range(0, len(sortedPairs)):

            f.write(str(i+1) + ": " + str(sortedPairs[i]) + "\n")

    with open("stuff.txt", "w") as f:
        f.write("Pairs\n")
        for i in range(0, len(pairsList)):
            f.write(str(i+1) + ": " + str(pairsList[i]) + "\n")



# ==================================================
#  OK, let's do this!
# ==================================================

if __name__ == "__main__": # allows us to put main at the beginning
    main()

# <<<<<<< Updated upstream
# =======
# getCoasterList(blankBallot)

# getBallotFilenames(ballotFolder)

# createDict()

# createMatrix()

# minRiders = input("Minimum number of riders to qualify? ")

# runTheContest()

# calculateResults()

# for coaster in winLossMatrix.keys():
#     if len(winLossMatrix[coaster]) > 0:
#         print coaster, winLossMatrix[coaster]

# # for coaster in winPercentage.keys():
# #     if winPercentage[coaster] > 0:
# #         print " ->", coaster, winPercentage[coaster]

# # sortedLists()

# # printToFile()
# >>>>>>> Stashed changes


# ==================================================
#  still to do
# ==================================================

# handle ties: decide which one wins, if possible. If still tied, rank them the same
# convert pairwise results into numpy array or pandas dataframe
# make subsets: rankings of gigas, hypers, types, parks, etc
