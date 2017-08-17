import sys
from helper import *

def main(args):

    path = 'Analytics_Attachment.xlsx'

    print "Reading Excel file..."

    team_data = read_sheet(path, 0)

    scores = read_sheet(path, 1)

    print "Initializing data..."

    teams = initialize_team_data(team_data)

    print "Determining elimination dates..."

    current_date = ''
    highest_game_total = 0
    tiebreak_check = False

    for score in scores:

        # Grabbing the home and away teams from the score result
        home = score[1]
        away = score[2]

        # Check for new date
        if current_date != score[0]:
            # If it's a new date, check for eliminations on previous date
            if tiebreak_check:
                teams = elimination_check(teams, scores, current_date)
            # Update the date
            current_date = score[0]
        
        # Increment games
        teams[home]["Games"] += 1
        teams[home]["Schedule"][away]["Games"] += 1
        teams[away]["Games"] += 1
        teams[away]["Schedule"][home]["Games"] += 1

        # Record winner and loser
        if score[5] == "Home":      #Home team won
            teams[home]["Wins"] += 1
            teams[home]["Schedule"][away]["Wins"] += 1
            teams[away]["Losses"] += 1
            teams[away]["Schedule"][home]["Losses"] += 1
               
        else:                       #Away team won
            teams[away]["Wins"] += 1
            teams[away]["Schedule"][home]["Wins"] += 1
            teams[home]["Losses"] += 1
            teams[home]["Schedule"][away]["Losses"] += 1
        
        # Only check for elimination tiebreakers after a team has played 41 games in the season, to limit unnecesary checks.
        if not tiebreak_check:
            highest_game_total = max(teams[home]["Games"], teams[away]["Games"])
            if highest_game_total >= 41: # Start checking for tiebreakers
                tiebreak_check = True

    # One last
    teams = elimination_check(teams, scores, current_date)
    # todo: figure out how to write to xlsx file
    output_eliminated_teams(teams)



if __name__ == "__main__":
    main(sys.argv[0:])