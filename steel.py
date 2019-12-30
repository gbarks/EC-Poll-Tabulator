#!/usr/bin/env python3

# ==========================================================
#  ElloCoaster poll tabulator: steel coaster designers
#  Author: Grant Barker
# ==========================================================

from openpyxl.styles import PatternFill

designers = { # fill colors for prominent steel roller coaster manufacturers/designers
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
    "Vekoma"                                : PatternFill("solid", fgColor="8fc6d3"), # amber
    "Other Known Manufacturer"              : PatternFill("solid", fgColor="cccccc"), # dark gray
    ""                                      : PatternFill("solid", fgColor="eeeeee") # light gray
}
