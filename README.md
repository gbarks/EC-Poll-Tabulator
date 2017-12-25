# ElloCoaster Poll Tabulator

Tool for conducting roller coaster polls inspired by Mitch Hawker's ["Best Roller Coaster Poll"](http://ushsho.com/bestrollercoasterpoll.htm). Ballots rank certain coasters, and pairwise matchups of coasters are assigned a "Win", "Loss", or "Tie" depending on their ranking. Once all ballots and W-L-T's are tallied up, roller coasters are ranked by order of win percentage. The method is not dissimilar to how a [ranked pairs](https://en.wikipedia.org/wiki/Ranked_pairs) election is conducted.

## How to Use

To run an election where:

* The minimum number of a riders a coaster needs in order to be ranked is `10`
* The default ballot is named `generic ballot.txt`
* The folder containing submitted ballots is named `ballots`
* The resulting output file should be named `Coaster Poll 20XX`

Simply run:

`python tabulator.py -m 10 -b "generic ballot.txt" -f ballots -o "Coaster Poll 20XX"`

And output will be saved to `Coaster Poll 20XX.xlsx`

Alternatively, the command line argument defaults are:

* `-m 6`
* `-b blankballot2017.txt`
* `-f ballots2017`
* `-o "Poll Results.xlsx"`

Additional command line flags include:

* `-c` sets fill color of certain cells to reflect the make (manufacturer) of the coaster
* `-i` includes sensitive voter data in a spreadsheet in the output file
* `-r` bothers [rcdb.com](https://rcdb.com/) with requests to fill in coaster details
* `-v` prints data as it's processed; `-vv` prints even more

## Dependencies

The script should work with Python 2 or 3 and requires [openpyxl](https://openpyxl.readthedocs.io/en/default/).

Scraping data from [rcdb.com](https://rcdb.com/) with the `-r` flag requires Python 3, [lxml](http://lxml.de/), and [beautifulsoup4](https://www.crummy.com/software/BeautifulSoup/bs4/doc/).

## More Info

Poll designed and hosted by ElloCoaster. Check out our [Wood Coaster Poll](http://www.ellocoaster.com/wood-coaster-poll)!
