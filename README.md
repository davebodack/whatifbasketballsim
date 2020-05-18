# whatifbasketballsim
A repository to hold code that uses Excel to quickly interface with whatifsports.com and run matchups between historic NBA teams.

Running this code requires the use of Python and openpyxl, which can be downloaded [here](https://pypi.org/project/openpyxl/#files). Click [here](https://stackoverflow.com/questions/38364404/how-to-install-openpyxl-with-pip) to read about how to install the downloaded openpyxl files if you're not sure about it.

Ever heard of whatifsports.com? It's a cool website. It lets you simulate matchups between almost any historical basketball team, but the problem is the website only lets you do them one game at a time. Well, no longer! This program simulates postseason tournaments between the 48 [best teams in basketball history](https://www.espn.com/nba/story/_/id/13000418/where-golden-state-warriors-rank-50-greatest-nba-teams). The reason it's 48 and not 16 is because the program simulates three postseasons at a time, using a relegation system like whatever those Brits do across the pond with their soccer leagues. The top two teams in the second and third leagues are promoted, and the bottom two teams in the first and second league, based on number of wins and point differential, are demoted.

The Excel file in this repo can be used to track the results of your simulations each time you do it. Just remove the "Start" in the name of the Excel file, or change the filename referenced in the code. If you want to simulate Year 2, for example, you have to have a sheet named Year 3 Bracket. All you need to do is copy the Year 2 Bracket sheet and rename it, everything in the Year 3 bracket will change once the code runs.

You have to put the year you're trying to run as a command-line argument. So, for example, to run the program for Year 2 you'd just type "py alltimersnba.py 2" in the terminal. 

Fair warning: you'd obviously expect some teams to do better than others, but it seems like whatifsports.com really needs to work on their balancing. I've run the program for 21 years, and the 2016-17 Warriors have won 12 Gold League championships and the 17-18 Warriors have won 6. That's why I'm working on a mode that establishes a handicap for teams, like in golf, based on their average point differential that will give greater parity. Stay tuned.
