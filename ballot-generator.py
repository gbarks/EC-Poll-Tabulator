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

parser.add_argument("-i", "--rcdblink", action="append", required=True,
                    help="RCDB input url (required) - can use multiple -i args")
parser.add_argument("-o", "--outballot", default="rcdb_ballot.csv",
                    help="specify name of output [ballot].csv file")
parser.add_argument("-O", "--outdetails", default="rcdb_ballot_details.csv",
                    help="specify name of output [details].csv file")
parser.add_argument("-s", "--sortbydate", action="store_true",
                    help="ensure RCDB pages are sorted chronologically")
parser.add_argument("-u", "--skipunknown", action="store_true",
                    help="skip all coasters named 'unknown'")
parser.add_argument("-d", "--skipnodate", action="store_true",
                    help="skip all coasters with nonspecific opening date")
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

this_year = str(datetime.datetime.now().year)

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

    # get name, alt name, park, and location
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
    country = location[location.rindex(",")+1:].strip()
    c["country"] = "\"" + country + "\""
    state = location[:location.rindex(",")].strip()
    city = state[:state.rindex(",")].strip()
    state = state[state.rindex(",")+1:].strip()
    c["state"] = "\"" + state + "\""
    c["city"] = "\"" + city + "\""

    if args.skipunknown is True and name == "unknown":
        if args.verbose > 0:
            print("--Skipping \"unknown\" at " + park + " - " + location)
        return None

    # get opening date and closing date
    datestr = feature.text[feature.text.find(")")+1:]
    if "Mountain Coaster" in datestr:
        datestr = datestr[:datestr.find("Mountain Coaster")]
    elif "Powered Coaster" in datestr:
        datestr = datestr[:datestr.find("Powered Coaster")]
    else:
        datestr = datestr[:datestr.find("Roller Coaster")]
    datestr = datestr.replace("\n", "")
    if datestr == "Operating" or datestr == "Removed" or "SBNO" in datestr or "In Storage" in datestr:
        if args.skipnodate:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (unknown opening date)")
            return None
    elif " or earlier" in datestr:
        if args.skipnodate:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (nonspecific opening date)")
            return None
    elif "Operating since " in datestr:
        datestr = datestr.split("Operating since ", 1)[1]
        datestr = re.sub(r'[^\d/]+', '', datestr)
        c["date"] = datestr
    else:
        if datestr[-4:] == this_year:
            if "Removed, Operated from " in datestr:
                datestr = datestr.split("Removed, Operated from ", 1)[1]
                closing = datestr.split(" ")[-1]
                closing = re.sub(r'[^\d/]+', '', closing)
                c["closing"] = closing
                datestr = datestr.split(" ")[0]
                datestr = re.sub(r'[^\d/]+', '', datestr)
                c["date"] = datestr
        else:
            if args.verbose > 0:
                print("--Skipping " + name + " at " + park + " (removed before " + this_year + ")")
            return None

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
    elif "Extreme" in linkrow.text:
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


    # def get_coaster_name(text):
    #     return substring

    # def get_park_name(text):
    #     return substring

    # def get_location(text):
    #     return substring

    # def get_opening_date(text):
    #     return substring

    # def get_make(text):
    #     return substring

    # def get_model(text):
    #     return substring

    # def get_submodel(text):
    #     return substring

    if args.verbose > 1:
        # print(name + ", " + park + ", " + date, end="")
        print(name, end="")
        if "altname" in c:
            print(" (AKA " + altname + ")", end="")
        print(", " + park, end="")
        if args.verbose > 2:
            print(" - " + location)
        else:
            print("")
        if args.verbose > 3:
            if "date" in c:
                print("    Opened: " + c["date"], end="")
                if "closing" in c:
                    print(" (closed " + c["closing"] + ")")
                else:
                    print("")
        if args.verbose > 4:
            if "scale" in c:
                print("    Scale:  " + c["scale"])
            if "make" in c:
                print("    Make:   " + c["make"])
            if "model" in c:
                if "submodel" in c:
                    print("    Model:  " + c["model"] + " / " + c["submodel"])
                else:
                    print("    Model:  " + c["model"])

    def get_stat_val(stat, unit, text):
        if stat in text:
            substring = text.split(stat, 1)[1]
            substring = substring.split(unit, 1)[0]
            substring = substring.replace(',', '')
            return substring

    def get_inver_val(text):
        if "Inversions" in text:
            substring = text.split("Inversions", 1)[1][:2]
            substring = re.sub(r'[^\d]+', '', substring)
            return substring

    def get_dur_val(text):
        if "Duration" in text:
            substring = text.split("Duration", 1)[1][:5]
            substring = re.sub(r'[^\d:]+', '', substring)
            return substring

    # scrape stats data
    for x in csoup.body.find_all('table', attrs={'id':'statTable'}):
        c["length"] = get_stat_val("Length", " ft", x.text)
        c["height"] = get_stat_val("Height", " ft", x.text)
        c["drop"]   = get_stat_val("Drop", " ft", x.text)
        c["speed"]  = get_stat_val("Speed", " mph", x.text)
        c["vert"]   = get_stat_val("Vertical Angle", "°", x.text)
        c["inver"] = get_inver_val(x.text)
        c["dur"] = get_dur_val(x.text)

    return c

if __name__ == "__main__": # allows us to put main at the beginning
    main()
