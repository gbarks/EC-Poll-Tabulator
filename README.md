# ElloCoaster Poll Tabulator

Tool for conducting roller coaster polls inspired by Mitch Hawker's ["Best Roller Coaster Poll"](http://ushsho.com/bestrollercoasterpoll.htm). Ballots rank certain coasters, and pairwise matchups of coasters are assigned a "Win", "Loss", or "Tie" depending on their ranking. Once all ballots and W-L-T's are tallied up, roller coasters are ranked by order of win percentage. The method is not dissimilar to how a [ranked pairs](https://en.wikipedia.org/wiki/Ranked_pairs) election is conducted.

## How to Use

To run an election where:

* The minimum number of a riders a coaster needs in order to be ranked is `10`
* The default ballot is named `generic ballot.txt`
* The folder containing submitted ballots is named `ballots`

Simply run:

`python tabulator.py 10 "generic ballot.txt" ballots`

And output will be saved to `Poll Results.xlsx`

Alternatively, the command line argument defaults are:

* `6`
* `blankballot2017.txt`
* `ballots2017`

## Dependencies

The script should work with Python 2 or 3 and requires [openpyxl](https://openpyxl.readthedocs.io/en/default/).

## More Info

Poll designed and hosted by ElloCoaster. Check out our [Wood Coaster Poll](http://www.ellocoaster.com/wood-coaster-poll)!
