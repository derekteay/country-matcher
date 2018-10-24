# country-matcher
Match a list of free-form text city input against an ISO list of countries ([pycountry](https://pypi.org/project/pycountry/)) using a fuzzy matcher ([fuzzywuzzy](https://github.com/seatgeek/fuzzywuzzy)) and graph the results using [matplotlib](https://github.com/matplotlib/matplotlib)

This program assumes that any match with an 80% or greater confidence is a valid match.
