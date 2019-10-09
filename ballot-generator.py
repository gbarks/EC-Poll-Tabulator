#!/usr/bin/env python3

# ==========================================================
#  ElloCoaster ballot generator
#  Author: Grant Barker
#
#  Requires Python 3
# ==========================================================

from __future__ import print_function

import sys

if sys.version_info[0] < 3:
    print("Program requires Python 3; running {0}.{1}".format(sys.version_info[0], sys.version_info[1]))
    sys.exit()

import re
import lxml
import argparse
import datetime
from bs4 import BeautifulSoup
from urllib.request import urlopen

# command line arguments
parser = argparse.ArgumentParser(description='Pull coaster info from RCDB list into .csv ballot')

def valid_year(s):
    try:
        return datetime.datetime.strptime(s, "%Y").strftime("%Y")
    except ValueError:
        msg = "Not a valid year: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)

parser.add_argument("-i", "--rcdblink", action="append", required=True,
                    help="RCDB input url (required) - can use multiple -i args")
parser.add_argument("-o", "--outballot", default="rcdb_ballot.csv",
                    help="specify name of output [ballot].csv file")
parser.add_argument("-O", "--outdetails", default="rcdb_ballot_details.csv",
                    help="specify name of output [details].csv file")
parser.add_argument("-s", "--sortbydate", action="store_true",
                    help="ensure RCDB pages are sorted chronologically")
parser.add_argument("-c", "--combineTracks", action="store_true",
                    help="don't make separate coaster entries for multi-tracks")
parser.add_argument("-u", "--skipunknown", action="store_true",
                    help="skip all coasters named 'unknown'")
parser.add_argument("-d", "--skipnodate", action="store_true",
                    help="skip all coasters with nonspecific open/close date")
parser.add_argument("-y", "--skipwrongyear", action="store_true",
                    help="skip all coasters that did not operate in given year")
parser.add_argument("-Y", "--setyear", default=str(datetime.datetime.now().year),
                    help="set year for -y; defaults to current year", type=valid_year)
parser.add_argument("-k", "--skipkiddie", action="store_true",
                    help="skip all kiddie coasters")
parser.add_argument("-v", "--verbose", action="count", default=0,
                    help="print data as it's processed; duplicate for more detail")

args = parser.parse_args()

# format output filenames
if args.outballot != "rcdb_ballot.csv":
    if args.outballot[-4:] != ".csv":
        args.outballot += ".csv"
    if args.outdetails == "rcdb_ballot_details.csv":
        args.outdetails = args.outballot[:-4] + "_details.csv"
if args.outdetails[-4:] != ".csv":
    args.outdetails += ".csv"

def main():
    coasters = []
    rcdblink = args.rcdblink

    # counters for while loop
    i = 0
    j = len(rcdblink)

    while i < j:
        if args.verbose > 0:
            print("<<< Now checking " + rcdblink[i] + " >>>")

        # handle pages that are lists of coasters
        if is_list_page(rcdblink[i]):
            if args.sortbydate:
                if rcdblink[i][-8:] != "&order=8" and rcdblink[i][-7:-1] != "&page=":
                    rcdblink[i] += "&order=8"

            response = urlopen(rcdblink[i])
            html = response.read()
            soup = BeautifulSoup(html, 'lxml')

            table = soup.find('tbody')

            # iterate over all coasters listed on the page
            for tr in table.find_all('tr'):
                td = tr.find_all('td')[1]
                url = "https://rcdb.com" + td.find('a').get('href')

                # name = td.text

                # if args.skipunknown:
                #     if name == "unknown":
                #         continue
                # park = tr.find_all('td')[2].text
                # date = tr.find_all('td')[6].text

                # c = {}
                # c["name"] = "\"" + name + "\""
                # c["park"] = "\"" + park + "\""
                # c["date"] = date
                # c["url"] = url

                c = parse_rcdb_page(url)
                if c is not None:
                    if isinstance(c, list):
                        for item in c:
                            coasters.append(item)
                    else:
                        coasters.append(c)

            # check if there's another list page to scrape in the footer
            rfoot = soup.find('div', attrs={'id':'rfoot'})
            if rfoot is not None:
                for a in rfoot.find_all('a'):
                    if a.text == ">>":
                        rcdblink.insert(i+1, "https://rcdb.com" + a.get('href'))

        # handle pages that are (presumably) individual coasters
        else:
            c = parse_rcdb_page(rcdblink[i])
            if c is not None:
                if isinstance(c, list):
                    for item in c:
                        coasters.append(item)
                else:
                    coasters.append(c)

        # increment while loop counter
        i += 1
        j = len(rcdblink)

    # file = open(args.outfile, "w")
    # file.write("Name,Park,Location,Opening Date,Length,Height,Drop,Speed,Inversions,Vertical Angle,Duration,RCDB Link\n")

    # def none_to_blank(key, c, csvline):
    #     if key not in c or c[key] is None:
    #         return csvline + ","
    #     else:
    #         return csvline + c[key] + ","

    # for c in coasters:
    #     csvline = c["name"] + "," + c["park"] + "," + c["location"] + ","
    #     csvline = none_to_blank("date", c, csvline)
    #     csvline = none_to_blank("length", c, csvline)
    #     csvline = none_to_blank("height", c, csvline)
    #     csvline = none_to_blank("drop", c, csvline)
    #     csvline = none_to_blank("speed", c, csvline)
    #     csvline = none_to_blank("inver", c, csvline)
    #     csvline = none_to_blank("vert", c, csvline)
    #     csvline = none_to_blank("dur", c, csvline)
    #     csvline = csvline + c["url"] + "\n"
    #     file.write(csvline)

    # file.close()

def is_list_page(url):
    substring = url.split("rcdb.com/", 1)[1]
    if substring[:5] == "r.htm":
        return True
    else:
        return False

def parse_rcdb_page(url):
    c = {}
    c["url"] = url

    cresponse = urlopen(url)
    chtml = cresponse.read()
    csoup = BeautifulSoup(chtml, 'lxml')

    # get name, alt name (for native language users), park, and location
    feature = csoup.find('div', attrs={'id':'feature'})
    title = feature.find('div', attrs={'class':'scroll'})
    name = title.find('h1').text
    if " / " in name:
        altname = name.split(" / ", 1)[1]
        name = name.split(" / ", 1)[0]
        c["altname"] = "\"" + altname + "\""
    c["name"] = "\"" + name + "\""
    park = title.find_all('a')[0].text
    c["park"] = "\"" + park + "\""
    location = title.text[title.text.find("(")+1:title.text.find(")")]
    c["location"] = "\"" + location + "\""

    # get individual city name from the location string
    c["country"] = None
    c["fullcity"] = None
    c["state"] = None
    c["city"] = None
    if "," in location:
        country = location[location.rindex(",")+1:].strip()
        c["country"] = "\"" + country + "\""
        fullcity = location[:location.rindex(",")].strip()
        c["fullcity"] = "\"" + fullcity + "\""
        if "," in fullcity:
            state = fullcity[fullcity.rindex(",")+1:].strip()
            c["state"] = "\"" + state + "\""
            city = fullcity[:fullcity.rindex(",")].strip()
            c["city"] = "\"" + city + "\""

    # skip coasters named "Unknown" (-u arg)
    if args.skipunknown is True and name == "unknown":
        if args.verbose > 0:
            print("--Skipping \"unknown\" at " + park + " - " + location)
        return None

    # get opening date and closing date (and extract coaster type from date string)
    dates = feature.find_all('time')
    datestr = feature.text[feature.text.find(")")+1:]
    if "Mountain Coaster" in datestr:
        datestr = datestr[:datestr.find("Mountain Coaster")]
        c["type"] = "Mountain Coaster"
    elif "Powered Coaster" in datestr:
        datestr = datestr[:datestr.find("Powered Coaster")]
        c["type"] = "Powered Coaster"
    else:
        datestr = datestr[:datestr.find("Roller Coaster")]
        c["type"] = "Roller Coaster"

    c["opendate"] = None
    c["closedate"] = None

    if "Operating" in datestr:
        c["status"] = "Operating"
        if "Operating since" in datestr:
            if "?" in datestr and args.skipnodate:
                if args.verbose > 0:
                    print("--Skipping " + name + " at " + park + " (operating, '?' opening date)")
                return None
            elif " - " in datestr:
                c["opendate"] = dates[0]['datetime'] + " - " + dates[1]['datetime']
            elif "≤" in datestr:
                c["opendate"] = "≤ " + dates[0]['datetime']
            elif "≥" in datestr:
                c["opendate"] = "≥ " + dates[0]['datetime']
            else:
                c["opendate"] = dates[0]['datetime']
            if args.skipwrongyear and len(dates) > 0:
                if int(dates[0]['datetime'][:4]) > int(args.setyear):
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (opened after " + args.setyear + ")")
                    return None
        elif args.skipnodate:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (operating, unknown opening date)")
            return None

    elif "Removed" in datestr:
        c["status"] = "Removed"
        if "Operated from" in datestr:

            # apologies for the spaghetti code; too many edge cases to check
            minopendate = None
            maxclosedate = None
            if "from ? to" in datestr:
                if args.skipnodate:
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (removed, '?' opening date)")
                    return None
                c["opendate"] = "?"
                if " - " in datestr:
                    c["closedate"] =  dates[0]['datetime'] + " - " + dates[1]['datetime']
                    maxclosedate = int(dates[1]['datetime'][:4])
                elif "≤" in datestr:
                    c["closedate"] = "≤ " + dates[0]['datetime']
                elif "≥" in datestr:
                    c["closedate"] = "≥ " + dates[0]['datetime']
                else:
                    c["closedate"] = dates[0]['datetime']
                if maxclosedate is None:
                    maxclosedate = int(dates[0]['datetime'][:4])
            elif "to ?" in datestr:
                c["closedate"] = "?"
                if " - " in datestr:
                    c["opendate"] = dates[0]['datetime'] + " - " + dates[1]['datetime']
                elif "≤" in datestr:
                    c["opendate"] = "≤ " + dates[0]['datetime']
                elif "≥" in datestr:
                    c["opendate"] = "≥ " + dates[0]['datetime']
                else:
                    c["opendate"] = dates[0]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
            elif "from  -  to" in datestr:
                c["opendate"] = dates[0]['datetime'] + " - " + dates[1]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
                if "to  - " in datestr:
                    c["closedate"] =  dates[2]['datetime'] + " - " + dates[3]['datetime']
                    maxclosedate = int(dates[3]['datetime'][:4])
                elif "≤" in datestr:
                    c["closedate"] = "≤ " + dates[2]['datetime']
                elif "≥" in datestr:
                    c["closedate"] = "≥ " + dates[2]['datetime']
                else:
                    c["closedate"] = dates[2]['datetime']
                if maxclosedate is None:
                    maxclosedate = int(dates[2]['datetime'][:4])
            elif "from ≤  to" in datestr:
                c["opendate"] = "≤ " + dates[0]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
                if " - " in datestr:
                    c["closedate"] =  dates[1]['datetime'] + " - " + dates[2]['datetime']
                    maxclosedate = int(dates[2]['datetime'][:4])
                elif "≤" in datestr:
                    c["closedate"] = "≤ " + dates[1]['datetime']
                elif "≥" in datestr:
                    c["closedate"] = "≥ " + dates[1]['datetime']
                else:
                    c["closedate"] = dates[1]['datetime']
                if maxclosedate is None:
                    maxclosedate = int(dates[1]['datetime'][:4])
            elif "from ≥  to" in datestr:
                c["opendate"] = "≥ " + dates[0]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
                if " - " in datestr:
                    c["closedate"] =  dates[1]['datetime'] + " - " + dates[2]['datetime']
                    maxclosedate = int(dates[2]['datetime'][:4])
                elif "≤" in datestr:
                    c["closedate"] = "≤ " + dates[1]['datetime']
                elif "≥" in datestr:
                    c["closedate"] = "≥ " + dates[1]['datetime']
                else:
                    c["closedate"] = dates[1]['datetime']
                if maxclosedate is None:
                    maxclosedate = int(dates[1]['datetime'][:4])
            elif "from  to" in datestr:
                c["opendate"] = dates[0]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
                if " - " in datestr:
                    c["closedate"] =  dates[1]['datetime'] + " - " + dates[2]['datetime']
                    maxclosedate = int(dates[2]['datetime'][:4])
                elif "≤" in datestr:
                    c["closedate"] = "≤ " + dates[1]['datetime']
                elif "≥" in datestr:
                    c["closedate"] = "≥ " + dates[1]['datetime']
                else:
                    c["closedate"] = dates[1]['datetime']
                if maxclosedate is None:
                    maxclosedate = int(dates[1]['datetime'][:4])
            else:
                c["opendate"] = dates[0]['datetime']
                minopendate = int(dates[0]['datetime'][:4])
                if len(dates) > 1:
                    print("Something went wrong in the open/close date parsing...")

            # determine if the ride operated in the given year (if -y arg is used)
            if args.skipwrongyear:
                if minopendate is not None and minopendate > int(args.setyear):
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (opened after " + args.setyear + ")")
                    return None
                if maxclosedate is not None and maxclosedate < int(args.setyear):
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (removed before " + args.setyear + ")")
                    return None
                if minopendate is None:
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (removed, unknown opening year)")
                    return None
                if maxclosedate is None:
                    if args.verbose > 0:
                        print("--Skipping " + name + " at " + park + " (removed, unknown closing year)")
                    return None
        elif args.skipnodate:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (removed, unknown opening date)")
            return None

    elif "SBNO" in datestr:
        c["status"] = "SBNO"
        for tr in csoup.find_all('table', attrs={'class':'objDemoBox'})[-1].find_all('tr'):
            if "Former status" in tr.text:
                td = tr.find_all('td')[-1]
                tdates = td.find_all('time')
                tdtext = re.split('Operate|SBN', td.text)[1:] # hacky text operations incoming
                minopendate = None
                maxclosedate = None

                # check first line of "Operated from" for closing date
                if "d" is tdtext[0][0]:
                    if "to ?" in tdtext[0]:
                        c["closedate"] = "?"
                    elif "to" in tdtext[0] and tdtext[0].count('-') > 1:
                        c["closedate"] = tdates[2]['datetime'] + " - " + tdates[3]['datetime']
                        maxclosedate = int(tdates[3]['datetime'][:4])
                    elif "to" in tdtext[0] and tdtext[0].count('-') > 0:
                        if " -  to" in tdtext[0]:
                            if "to ≤" in tdtext[0]:
                                c["closedate"] = "≤ " + tdates[2]['datetime']
                            elif "to ≥" in tdtext[0]:
                                c["closedate"] = "≥ " + tdates[2]['datetime']
                            else:
                                c["closedate"] = tdates[2]['datetime']
                        else:
                            c["closedate"] = tdates[1]['datetime'] + " - " + tdates[2]['datetime']
                        maxclosedate = int(tdates[2]['datetime'][:4])
                    elif "to ≤" in tdtext[0]:
                        c["closedate"] = "≤ " + tdates[1]['datetime']
                        maxclosedate = int(tdates[1]['datetime'][:4])
                    elif "to ≥" in tdtext[0]:
                        c["closedate"] = "≥ " + tdates[1]['datetime']
                        maxclosedate = int(tdates[1]['datetime'][:4])
                    elif "to" in tdtext[0]:
                        c["closedate"] = tdates[1]['datetime']
                        maxclosedate = int(tdates[1]['datetime'][:4])

                # check last line of "Operated from" for opening date
                if "d" is tdtext[-1][0]:
                    if "from ? to" in tdtext[-1]:
                        if args.skipnodate:
                            if args.verbose > 0:
                                print("--Skipping " + name + " at " + park + " (SBNO, '?' opening date)")
                            return None
                        c["opendate"] = "?"
                    elif "from  -  to" in tdtext[-1]:
                        if "to  - " in tdtext[-1]:
                            c["opendate"] = tdates[-4]['datetime'] + " - " + tdates[-3]['datetime']
                            minopendate = int(tdates[-4]['datetime'][:4])
                        else:
                            c["opendate"] = tdates[-3]['datetime'] + " - " + tdates[-2]['datetime']
                            minopendate = int(tdates[-3]['datetime'][:4])
                    elif "from ≤" in tdtext[-1]:
                        if "to  - " in tdtext[-1]:
                            c["opendate"] = "≤ " + tdates[-3]['datetime']
                            minopendate = int(tdates[-3]['datetime'][:4])
                        else:
                            c["opendate"] = "≤ " + tdates[-2]['datetime']
                            minopendate = int(tdates[-2]['datetime'][:4])
                    elif "from ≥" in tdtext[-1]:
                        if "to  - " in tdtext[-1]:
                            c["opendate"] = "≥ " + tdates[-3]['datetime']
                            minopendate = int(tdates[-3]['datetime'][:4])
                        else:
                            c["opendate"] = "≥ " + tdates[-2]['datetime']
                            minopendate = int(tdates[-2]['datetime'][:4])
                    elif "to" in tdtext[-1]:
                        if "to  - " in tdtext[-1]:
                            c["opendate"] = tdates[-3]['datetime']
                            minopendate = int(tdates[-3]['datetime'][:4])
                        else:
                            c["opendate"] = tdates[-2]['datetime']
                            minopendate = int(tdates[-2]['datetime'][:4])
                    else:
                        if "-" in tdtext[-1]:
                            c["opendate"] = tdates[-2]['datetime'] + " - " + tdates[-1]['datetime']
                            minopendate = int(tdates[-2]['datetime'][:4])
                        else:
                            c["opendate"] = tdates[-1]['datetime']
                            minopendate = int(tdates[-1]['datetime'][:4])

                # determine if the ride operated in the given year (if -y arg is used)
                if args.skipwrongyear:
                    if minopendate is not None and minopendate > int(args.setyear):
                        if args.verbose > 0:
                            print("--Skipping " + name + " at " + park + " (opened after " + args.setyear + ")")
                        return None
                    if maxclosedate is not None and maxclosedate < int(args.setyear):
                        if args.verbose > 0:
                            print("--Skipping " + name + " at " + park + " (SBNO before " + args.setyear + ")")
                        return None
                    if minopendate is None:
                        if args.verbose > 0:
                            print("--Skipping " + name + " at " + park + " (SBNO, unknown opening year)")
                        return None
                    if maxclosedate is None:
                        if args.verbose > 0:
                            print("--Skipping " + name + " at " + park + " (SBNO, unknown closing year)")
                        return None

                break

    elif "In Storage" in datestr:
        c["status"] = "In Storage"
        return None # don't care about coasters in storage, for now
                    # very few are listed, see https://rcdb.com/r.htm?st=312&ot=2

    elif "Under Construction" in datestr:
        c["status"] = "Under Construction"
        return None # don't care about rides that have yet to be constructed

    # get thrill scale
    linkrow = feature.find('span', attrs={'class':'link_row'})
    if "Kiddie" in linkrow.text:
        if args.skipkiddie:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (\"Kiddie\" designation)")
            return None
        c["scale"] = "Kiddie"
    elif "Family" in linkrow.text:
        c["scale"] = "Family"
    elif "Thrill" in linkrow.text:
        c["scale"] = "Thrill"
    elif "Extreme" in linkrow.text:
        c["scale"] = "Extreme"

    # get make and model
    makemodellist = feature.find_all('div', attrs={'class':'scroll'})
    if len(makemodellist) > 1:
        makemodel = makemodellist[1].text
        if "Make: " in makemodel:
            make = makemodel.split("Make: ", 1)[1]
            if "Model: " in makemodel:
                model = make.split("Model: ", 1)[1]
                make = make.split("Model: ", 1)[0]
                if " / " in model:
                    c["submodel"] = model.split(" / ", 1)[1]
                    model = model.split(" / ", 1)[0]
                c["model"] = model
            c["make"] = make
        elif "Model: " in makemodel:
            print("WTF THIS DOESN'T MAKE SENSE!!!!")
            print(c)

    # get number of tracks; write new tracks to separate coaster entries
    stats = csoup.body.find('table', attrs={'id':'statTable'})
    numTracks = len(stats.find('tr').find_all('td'))
    c["tracks"] = str(numTracks)
    trackNames = []
    d = []
    if numTracks > 1 and args.combineTracks is False:
        if stats.find('th').text == "Name":
            for x in stats.find('tr').find_all('td'):
                trackNames.append(x.text)
        for i in range(numTracks):
            c["id"] = chr(ord('a')+i) + url.split("/")[-1].split(".")[0]
            if len(trackNames) > 1:
                c["name"] = "\"" + name + " (" + trackNames[i] + ")\""
            else:
                c["name"] = "\"" + name + " (" + chr(ord('a')+i) + ")\""
            d.append(c.copy())
    else:
        c["id"] = "r" + url.split("/")[-1].split(".")[0]

    # verbose print info per coaster
    if args.verbose > 1: # -vv = coaster name
        print(name, end="")
        if "altname" in c:
            print(" (AKA " + altname + ")", end="")
        print(", " + park, end="")
        if args.verbose > 2: # -vvv = coaster location (same line as coaster name)
            print(" - " + location, end="")
        print("")
        if args.verbose > 3: # -vvvv = coaster status + opening/closing dates
            if "status" in c:
                print("    Status: " + c["status"], end="")
                if "opendate" in c and c["opendate"] is not None:
                    if c["status"] is "Operating":
                        print(" since " + c["opendate"], end="")
                    elif c["status"] is "Removed" or "SBNO":
                        print(", Operated from " + c["opendate"], end="")
                    else:
                        print(", ", end="")
                    if "closedate" in c and c["closedate"] is not None:
                        print(" to " + c["closedate"], end="")
                print("")
                if numTracks > 1:
                    print("    Tracks: " + str(numTracks), end="")
                    if len(trackNames) > 1:
                        print(" - ", end="")
                        for i in range(len(trackNames)):
                            print(trackNames[i], end="")
                            if i < len(trackNames) - 1:
                                print(", ", end="")
                    print("")
            if args.verbose > 4: # -vvvvv = coaster scale + make/model
                if "scale" in c:
                    print("    Scale:  " + c["scale"])
                if "make" in c:
                    print("    Make:   " + c["make"])
                if "model" in c:
                    if "submodel" in c:
                        print("    Model:  " + c["model"] + " / " + c["submodel"])
                    else:
                        print("    Model:  " + c["model"])

    if numTracks == 1 or args.combineTracks is True:
        return c
    else:
        return d

if __name__ == "__main__": # allows us to put main at the beginning
    main()
