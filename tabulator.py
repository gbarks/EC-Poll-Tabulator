# ==================================================
#  ElloCoaster poll tabulator
#  Contributions from Jim Winslett, Grant Barker
# ==================================================

import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font

# Excel file to save all results to
xlout = Workbook()

# font style for openpyxl
preferredFixedWidthFont = Font(name="Menlo")

# lines in a ballot beginning with this will be ignored
commentStr = "* "

# ballot field still in place if voter didn't fill in the info
blankUserField = "-Replace "

# ballot line that separates voter info from coaster list
startLine = "! DO NOT CHANGE OR DELETE THIS LINE !"

# name of blank ballot file
blankBallot = "blankballot2017.txt"

 # folder where ballots are contained
ballotFolder = "ballots2017"

# list of tuples representing every coaster on the ballot
coasterList = []

# list of ballot filenames
ballotList = []

# dict that assigns a number to each coaster on the ballot
# potentially useful for chart printouts, using an int rather than the whole name of coaster
coasterDict = {}

# for each pair of coasters, a string containing w, l, or t representing every contest between that pair
winLossMatrix = {}

# list of strings containing name, email, city, state/prov, country
voterInfo = []

# for each ballot, the ranking given to each coaster voted on
coasterAndRank = {}

# for each pair of coasters, the % of times coasterA beat coasterB
winPercentage = {}

# for each coaster on the ballot, the number of people who rode that coaster
riders = {}

# unsorted list of the coaster and the number of riders it had
numRiders = []

# the number of total credits for all the voters
totalCredits = 0

# sorted version of the above
sortedRiders = []

# the number of times each coaster was paired up against another coaster
totalContests = {}

# the total number of wins for each coaster
totalWins = {}

# the total number of ties for each coaster
totalTies = {}

# the total number of losses for each coaster
totalLosses = {}

# unsorted list of coasters and their total win percentages
results = []

# sorted version of above
sortedResults = []

# unsorted list of every pair of coasters with the pair's win percentage
pairsList = []

# sorted version of above
sortedPairs = []

# minRiders
# minimum number of riders a coaster must have before being included in the results



def main():
    xlout.active.title = "Coaster Masterlist"

    getCoasterList(blankBallot, xlout.active)

    # getBallotFilenames(ballotFolder)

    # createDict()

    # createMatrix()

    # minRiders = input("Minimum number of riders to qualify? ")

    # runTheContest()

    # calculateResults()

    # sortedLists()

    # printToFile()

    xlout.save("Poll Results.xlsx")
    print 'Output saved to "Poll Results.xlsx".'



# ==================================================
#  populate list of coasters in the poll
# ==================================================

def getCoasterList(blankBallot, masterlistws):
    print "Creating list of every coaster on the ballot..."

    global coasterList
    global riders

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
                        fullCoasterName = words[1]
                        abbrCoasterName = words[2]

                        # add a fullCoasterName entry to the dicts
                        riders[fullCoasterName] = 0
                        totalContests[fullCoasterName] = 0
                        totalWins[fullCoasterName] = 0
                        totalTies[fullCoasterName] = 0
                        totalLosses[fullCoasterName] = 0

                        # add the coaster to the list of coasters on the ballot
                        coasterList.append((fullCoasterName, abbrCoasterName))
                        subwords = [x.strip() for x in fullCoasterName.split('-')]
                        if len(subwords) != 3:
                            masterlistws.append([fullCoasterName,abbrCoasterName])
                        else:
                            masterlistws.append([fullCoasterName,abbrCoasterName,subwords[0],subwords[1],subwords[2]])
                            masterlistws.cell(row=len(coasterList)+1, column=5).font = preferredFixedWidthFont
                        masterlistws.cell(row=len(coasterList)+1, column=2).font = preferredFixedWidthFont

    return coasterList, riders



# ==================================================
#  import filenames of ballots
# ==================================================

def getBallotFilenames(ballotFolder):
    print("Getting the filenames of submitted ballots")

    global ballotList

    # iterate through the list of files in the ballot folder
    for file in os.listdir(ballotFolder):

        # only pull out text files
        if file.endswith(".txt"):
            # add the filename to the list of ballot files
            ballotList.append(file)
    return ballotList



# =====================================================
#  create dictionary of coaster names paired with nums
#
#  might make it easier to print charts and such later
#    since a grid with coaster names as rows and cols
#    could be unruly
# =====================================================

def createDict():
    print("Creating the coaster dictionary")

    global coasterDict
    coasterNumber = 0
    for coaster in coasterList:
        coasterDict[coaster] = coasterNumber
        coasterNumber += 1

    return coasterDict



# ==================================================
#  create win/loss matrix
# ==================================================

def createMatrix():
    print("Creating the win/loss matrix")

    # create a matrix of blank strings for each pair of coasters
    # these strings will later contain w, l, t for each matchup
    global winLossMatrix
    for row in coasterList:
        for col in coasterList:
            winLossMatrix[row,col] = ''


    return winLossMatrix



# ================================================================
#  read a ballot (just ONE ballot)
#
#  you need a loop to call this function for each ballot filename
# ================================================================

def processBallot(filename):
    print("Processing ballot:")

    global voterInfo
    error = False
    global creditNum
    global totalCredits
    global riders

    # open the ballot file
    with open(filename) as f:
        # get the voter info
        infoField = 1
        lineNum = 0
        creditNum = 0
        voterInfo = [filename, "", "", "", "", ""]
        coasterName = ''
        coasterRank = 0
        startProcessing = False
        error = False
        coasterAndRank = {}

        for line in f:
            sline = line.strip()
            lineNum += 1

            # begin at top of ballot and get the voter's info first
            if startProcessing == False and infoField <= 5 and not commentStr in sline and len(sline) != 0:

                # if the line begins with "-Replace" then record a non-answer
                if blankUserField in sline:
                    voterInfo[infoField] = "(no answer)"
                    infoField += 1
                elif not startLine in sline:
                    voterInfo[infoField] = sline
                    infoField += 1




            # get the list of coasters this voter has ridden
            # check for the ballot line indicating that the coasters follow it
            if startProcessing == False and sline == startLine:
                startProcessing = True

            # break the line into its components: rank, name
            elif startProcessing == True:

                # strip away any blank space, save just the text, look for the comma to split words
                words = [x.strip() for x in sline.split(',')]

                # skip comment lines (begin with * )
                if commentStr in sline:
                    continue

                # skip blank lines
                elif sline == "":
                    continue

                # make sure there are 2 'words' in each line
                elif len(words) != 2:
                    print("Error in {0}, Line {1}: {2}".format(blankBallot, lineNum, line))

                # make sure the ranking is a number
                elif not words[0].isdigit():
                    print("Error in reading {0}, Line {1}: Rank must be an int.".format(filename, lineNum))
                    error = True

                    # Everything good? do this
                else:

                    # pull out the coaster name
                    coasterName = words[1]
                    # pull out the coaster's rank
                    coasterRank = int(words[0])

                    # skip coasters ranked zero or less (those weren't ridden)
                    if coasterRank <= 0:
                        continue


                    # check to make sure the coaster on the ballot is legit
                    if coasterName in coasterList:
                        # it is! Add to this voter's credit count
                        creditNum += 1
                        # add one to the number of riders this coaster has
                        riders[coasterName] += 1
                        # add this voter's ranking of the coaster
                        coasterAndRank[coasterName] = coasterRank

                    else:
                        # it's not a legit coaster!
                        print("Error in reading {0}, Line {1}: Unknown coaster {2}".format(filename, lineNum, coasterName))
                        error = True

    # no errors? Tally the ballot!
    if not error:

        # add this voter's credit count to the total credits
        totalCredits = totalCredits + creditNum

        # cycle through each pair of coasters this voter ranked
        for coasterA in coasterAndRank.keys():
            for coasterB in coasterAndRank.keys():
                # you can't compare a coaster to itself, so skip those pairs
                if coasterA == coasterB:
                    continue

                # if the coasters have the same ranking, call it a tie
                elif coasterAndRank[coasterA] == coasterAndRank[coasterB]:
                    # add a 't' to this pair's cell on the winLossMatrix
                    winLossMatrix[coasterA, coasterB] = winLossMatrix[coasterA, coasterB] + ("t")
                    # add one to the total contests that coasterA has had
                    totalContests[coasterA] += 1
                    # add one to the total ties coasterA has had
                    totalTies[coasterA] += 1

                # if coasterA outranks coasterB (the rank's number is lower), call it a win for coasterA
                elif coasterAndRank[coasterA] < coasterAndRank[coasterB]:
                    # add a 'w' to this pair's cell on the winLossMatrix
                    winLossMatrix[coasterA, coasterB] = winLossMatrix[coasterA, coasterB] + ("w")
                    # add one to the total contests coasterA has had
                    totalContests[coasterA] += 1
                    # add one to the total wins coasterA has had
                    totalWins[coasterA] += 1

                # if not a tie nor a win, it must be a loss
                else:
                    # if coasterB outranks coasterA (A's rank is a larger number), call it a loss for coasterA
                    # add an 'l' to this pair's cell on the winLossMatrix
                    winLossMatrix[coasterA, coasterB] = winLossMatrix[coasterA, coasterB] + ("l")
                    # add one to the total contests for coasterA
                    totalContests[coasterA] += 1
                    # add one to the total losses for coasterA
                    totalLosses[coasterA] += 1


    # if none of the above conditions were met, there must've been an error
    else:
        print("Errors. File {0} not added.".format(filename))

    return winLossMatrix



# ========================================================
#  calculate results
#
#  no need to loop through this, since it calculates with
#    numbers gathered when the ballots were processed
# ========================================================

def calculateResults():
    print("Calculating results")

    # initialize/reset the number of pairwise contests for each pair
    contestsHead2Head = 0
    # initialize/reset the number of wins for each pair of coasters
    wins = 0
    # initialize/reset the number of losses for each pair of coasters
    losses = 0
    # initialize/reset the number of ties for each pair of coasters
    ties = 0

    global winPercentage
    global totalContests

    # iterate through all the pairs in the matrix
    for row in coasterList:
        for col in coasterList:
            # there will be no info for a coaster paired with itself, so skip it
            if row == col:
                continue

            # look at the pair of coasters
            # and calculate the win percentage for coasterA(row) vs coasterB (col)
            else:
                # see how many times this pair went head-to-head
                contestsHead2Head = len(winLossMatrix[row,col])
                # count the number of wins for coasterA (row)
                wins = winLossMatrix[row,col].count("w")
                # count the number of losses for coasterA (row)
                losses = winLossMatrix[row,col].count("l")
                # count the number of ties for coasterA (row)
                ties = winLossMatrix[row,col].count("t")


                # if this pair of coasters had mutual riders and there were ties, calculate with this formula
                if ties != 0 and contestsHead2Head > 0:
                    # formula: wins + half the ties divided by the number of times they competed against each other
                    # Multiply that by 100 to get the percentage, then round to three digits after the decimal
                    winPercentage[row,col] = round((((wins + (ties / 2)) / contestsHead2Head)) * 100, 3)
                # if this pair had mutual riders, but there were no ties, use this formula
                elif contestsHead2Head > 0:
                    winPercentage[row,col] = round(((wins / contestsHead2Head)) * 100, 3)
                # if there were no mutual riders for this pair, skip it
                else:
                    continue




    # all those calculations we just did for each pair of coasters, now do for each coaster by itself
    # tallying up ALL the contests it had, not just the pairwise contests
    # this will give the total overall win percentage for each coaster, which will be used to determine
    # the final ranking of all the coasters
    for row in coasterList:
        if ties != 0 and totalContests[row] > 0:
            winPercentage[row] = round((((totalWins[row] + (totalTies[row]/2)) / totalContests[row])) * 100, 3)

        elif totalContests[row] > 0:
            winPercentage[row] = round(((totalWins[row] / totalContests[row])) * 100, 3)


    return winPercentage, totalContests



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
#  cycle through all the ballots and tabulate them
# ==================================================

def runTheContest():
    # loop through all the ballot filenames and process each ballot
    for filename in ballotList:
        print("Processing ballot: {0}".format(filename))
        processBallot("ballots2017/" + filename)

        print("=========================================================")
        for i in range(0,len(voterInfo)):
            if i == 0:
                print("Ballot: {0}".format(voterInfo[i]))

            elif i == 1:
                print("Name: {0}".format(voterInfo[i]))

            elif i == 2:
                print("Email: {0}".format(voterInfo[i]))

            elif i == 3:
                print("City: {0}".format(voterInfo[i]))

            elif i == 4:
                print("State/Province: {0}".format(voterInfo[i]))

            elif i == 5:
                print("Country: {0}".format(voterInfo[i]))


        print("Coasters ridden: {0}".format(creditNum))
        print("=========================================================")
        print("=========================================================")

        print()



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



# ==================================================
#  still to do
# ==================================================

# handle ties: decide which one wins, if possible. If still tied, rank them the same
# convert pairwise results into numpy array or pandas dataframe
# make subsets: rankings of gigas, hypers, types, parks, etc
