#!/usr/bin/env python3

# Experiments in pulling length, height, and speed data for generic RCDB lists
# Author: Grant Barker

import re
import sys
import lxml
import argparse
from bs4 import BeautifulSoup
from urllib.request import urlopen

# command line arguments
parser = argparse.ArgumentParser(description='Pull coaster stats from RCDB list into .csv')

parser.add_argument("-i", "--rcdblink", action="append", required=True,
                    help="RCDB input url (required) - can use multiple -i args")
parser.add_argument("-o", "--outfile", default="rcdb_stats.csv",
                    help="specify name of output .csv file")
parser.add_argument("-d", "--sortbydate", action="store_true",
                    help="ensure RCDB pages are sorted chronologically")
parser.add_argument("-u", "--skipunknown", action="store_true",
                    help="skip all coasters named 'unknown'")
parser.add_argument("-v", "--verbose", action="count", default=0,
                    help="print data as it's processed; duplicate for more detail")

args = parser.parse_args()

if args.outfile[-4:] != ".csv":
    args.outfile += ".csv"

rcdblink = args.rcdblink

if args.sortbydate:
    for i in range(0, len(rcdblink)):
        if rcdblink[i][-8:] != "&order=8":
            rcdblink[i] += "&order=8"

coasters = []

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

# counters for while loop
i = 0
j = len(rcdblink)

while i < j:
    if args.verbose > 1:
        print("<<< Now checking " + rcdblink[i] + " >>>")

    response = urlopen(rcdblink[i])
    html = response.read()
    soup = BeautifulSoup(html, 'lxml')

    table = soup.find('tbody')

    # iterate over all coasters listed on the page
    for tr in table.find_all('tr'):
        td = tr.find_all('td')[1]
        name = td.text

        if args.skipunknown:
            if name == "unknown":
                continue

        url = "https://rcdb.com" + td.find('a').get('href')
        park = tr.find_all('td')[2].text
        date = tr.find_all('td')[6].text

        c = {}
        c["name"] = "\"" + name + "\""
        c["park"] = "\"" + park + "\""
        c["date"] = date
        c["url"] = url

        # scrape individual coasters' pages for info
        cresponse = urlopen(url)
        chtml = cresponse.read()
        csoup = BeautifulSoup(chtml, 'lxml')

        # scrape location
        title = csoup.find('div', attrs={'class':'scroll'}).text
        location = title[title.find("(")+1:title.find(")")]
        c["location"] = "\"" + location + "\""

        if args.verbose > 0:
            print(name + ", " + park + ", " + date, end="")
            if args.verbose > 2:
                print(" - " + location)
            else:
                print("")

        # scrape stats data
        for x in csoup.body.find_all('table', attrs={'id':'statTable'}):
            c["length"] = get_stat_val("Length", " ft", x.text)
            c["height"] = get_stat_val("Height", " ft", x.text)
            c["drop"]   = get_stat_val("Drop", " ft", x.text)
            c["speed"]  = get_stat_val("Speed", " mph", x.text)
            c["vert"]   = get_stat_val("Vertical Angle", "°", x.text)
            c["inver"] = get_inver_val(x.text)
            c["dur"] = get_dur_val(x.text)

        coasters.append(c)

    # check if there's another list page to scrape in the footer
    rfoot = soup.find('div', attrs={'id':'rfoot'})
    if rfoot is not None:
        for a in rfoot.find_all('a'):
            if a.text == ">>":
                rcdblink.insert(i+1, "https://rcdb.com" + a.get('href'))

    # increment while loop counter
    i += 1
    j = len(rcdblink)

file = open(args.outfile, "w")
file.write("Name,Park,Location,Opening Date,Length,Height,Drop,Speed,Inversions,Vertical Angle,Duration,RCDB Link\n")

def none_to_blank(key, c, csvline):
    if key not in c or c[key] is None:
        return csvline + ","
    else:
        return csvline + c[key] + ","

for c in coasters:
    csvline = c["name"] + "," + c["park"] + "," + c["location"] + ","
    csvline = none_to_blank("date", c, csvline)
    csvline = none_to_blank("length", c, csvline)
    csvline = none_to_blank("height", c, csvline)
    csvline = none_to_blank("drop", c, csvline)
    csvline = none_to_blank("speed", c, csvline)
    csvline = none_to_blank("inver", c, csvline)
    csvline = none_to_blank("vert", c, csvline)
    csvline = none_to_blank("dur", c, csvline)
    csvline = csvline + c["url"] + "\n"
    file.write(csvline)

file.close()
