#!/usr/bin/env python

# ==========================================================
#  ElloCoaster poll tabulator: coaster class definitions
#  Author: Grant Barker
# ==========================================================

# these HTML tools don't work in Python 2.x
import sys
if sys.version_info >= (3,0):
    import lxml
    from bs4 import BeautifulSoup
    from urllib.request import urlopen

class Coaster:
    name = ""
    park = ""
    location = ""

    rcdb = ""
    designer = ""
    # descolor = "ffffff"
    year = ""

    riders = 0
    totalWins = 0
    totalLosses = 0
    totalTies = 0
    totalWinPercentage = 0.0
    pairwiseWins = 0
    pairwiseLosses = 0
    pairwiseTies = 0
    pairwiseWinPercentage = 0.0
    overallRank = 0
    tiedCoasters = []

    def __init__(self, fullName, abbrName, rcdblink=None, designerSet=None):
        self.uniqueID = fullName
        self.abbr = abbrName

        # extract park and state/country information from fullUniqueCoasterName
        subwords = [x.strip() for x in fullName.split('-')]
        if len(subwords) == 3:
            self.name = subwords[0]
            self.park = subwords[1]
            self.location = subwords[2]

        # open URL if rcdblink is provided (time consuming)
        if rcdblink is not None and designerSet is not None:
            self.rcdb = rcdblink

            # these HTML tools don't work in Python 2.x; return with default values
            if sys.version_info < (3,0):
                return

            response = urlopen(rcdblink)
            html = response.read()
            soup = BeautifulSoup(html, 'lxml')

            # scan page for a "Make" field, usually at the top
            for x in soup.body.findAll('div', attrs={'class':'scroll'}):
                if "Make: " in x.text:
                    subtext = x.text.split("Make: ", 1)[1]

                    # strip out the "Model" field from the string, if neccessary
                    if "Model: " in subtext:
                        subtext = subtext.split("Model: ", 1)[0]
                    self.designer = subtext
                    break

            # if the "Make" field didn't exist or used an unknown manufacturer, try "Designer" field
            if self.designer == "" or self.designer not in designerSet:
                for x in soup.body.findAll('table', attrs={'class':'objDemoBox'}):
                    if "Designer:" in x.text:
                        subtext = x.text.split("Designer:", 1)[1]

                        # strip out other fields from the string, if neccessary
                        if "Installer:" in subtext:
                            subtext = subtext.split("Installer:", 1)[0]
                        if "Musical Score:" in subtext:
                            subtext = subtext.split("Musical Score:", 1)[0]
                        if "Construction Supervisor:" in subtext:
                            subtext = subtext.split("Construction Supervisor:", 1)[0]

                        # if a known manufacturer is a substring of subtext, use that
                        alreadyKnownManu = next((y for y in designerSet if y in subtext), False)
                        if alreadyKnownManu:
                            self.designer = alreadyKnownManu

                        # otherwise, use the provided "Designer"
                        elif self.designer == "" and subtext != "":
                            self.designer = subtext
                        break

            # exception for Gravity Group, who has two names on RCDB for some reason
            if self.designer == "Gravitykraft Corporation":
                self.designer = "The Gravity Group, LLC"

            # find an opening year if available
            for x in soup.body.findAll(True):
                if "Operating since " in x.text:
                    subtext = x.text.split("Operating since ", 1)[1][:10].split('/')[-1][:4]
                    self.year = int(subtext)
                    break
                elif "Operated from " in x.text:
                    subtext = x.text.split("Operated from ", 1)[1].split(' ')[0].split('/')[-1]
                    self.year = int(subtext)
                    break
