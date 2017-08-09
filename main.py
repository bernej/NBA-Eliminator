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
        # Check for new date
        if current_date != score[0]:
            current_date = score[0]
        # Record result
        if score[5] == "Home":
            teams[score[1]]["Wins"] += 1
            teams[score[2]]["Losses"] += 1
            if not teams[score[2]]["Eliminated"] and tiebreak_check:
                teams = elimination_check(score[2], teams)                    
        else:
            teams[score[2]]["Wins"] += 1
            teams[score[1]]["Losses"] += 1
            if not teams[score[1]]["Eliminated"] and tiebreak_check:
                teams = elimination_check(score[1], teams)     

        teams[score[1]]["Games"] += 1
        teams[score[2]]["Games"] += 1
        # Check for elimination
        if not tiebreak_check:
            # print highest_game_total
            highest_game_total = max(teams[score[1]]["Games"], teams[score[2]]["Games"])
            if highest_game_total >= 41: # Start checking for tiebreakers
                tiebreak_check = True

        # break
        # print "u"
        # print tiebreak_check


    # for team in teams:
    #     print team + " wins: " + str(teams[team]["Wins"])


    # todo: figure out how to write to xlsx file

if __name__ == "__main__":
    main(sys.argv[0:])