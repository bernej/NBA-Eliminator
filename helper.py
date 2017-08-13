import datetime, xlrd, operator

def read_sheet(path, index):
    path = 'Analytics_Attachment.xlsx'

    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_index(index)

    offset = 0

    rows = []
    for i, row in enumerate(range(worksheet.nrows)):
        if i <= offset:  # (Optionally) skip headers
            continue
        r = []
        for j, col in enumerate(range(worksheet.ncols)):
            if index and col == 0:
                a1 = worksheet.cell_value(rowx=row, colx=col)
                a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, workbook.datemode))
                date = str(a1_as_datetime)
                r.append(date.split()[0])
            elif type(worksheet.cell_value(i, j)) is not float:
                r.append(worksheet.cell_value(i, j).encode("utf-8"))
            else:
                r.append(worksheet.cell_value(i, j))
        rows.append(r)

    return rows

def initialize_team_data(team_data):
    teams = {}
    team_names = []

    for team in team_data:
        team_names.append(team[0])
        teams[team[0]] = {
            "Division": team[1],
            "Conference": team[2],
            "Games": 0,
            "Wins": 0,
            "Losses": 0,
            "Eliminated": False,
            "Schedule": {}
        }

    # Keep a record of games against each other.
    for team in teams:
        # Get 29 other teams
        other_teams = (other_team for other_team in teams if team != other_team)
        # Keep track of team's record against every other team
        for other_team in other_teams:
            teams[team]["Schedule"][other_team] = {
                "Games": 0,
                "Wins": 0,
                "Losses": 0
            }

    return teams

def loser_elimination_check(losing_team, teams, date):
    playoff_teams = 0
    tiebreak_teams = []

    # Break it down by conference
    conf = []
    for team in teams:
        if teams[team]["Conference"] == teams[losing_team]["Conference"]:
            conf.append(team)

    worst_team_alive, win_perc = determine_8th_place(conf, teams)

    for team in conf:
        # print 'yo'
        games_left = 82 - teams[losing_team]["Games"]
        if teams[team]["Wins"] == games_left + teams[losing_team]["Wins"]:
            tiebreak_teams += team
        elif teams[team]["Wins"] > games_left + teams[losing_team]["Wins"]:
            playoff_teams += 1
        if playoff_teams == 8:
            print losing_team + " eliminated from playoff contention on " + date
            teams[losing_team]["Eliminated"] = True
            return teams

    return teams

# Helper function to determine the 8th place team. This team will then be cross-examined
# against the non-eliminated, non-playoff teams to see if it owns tiebreakers
def determine_8th_place(conf, teams, date):
    
    # Find the lowest winning non-eliminated team
    worst_non_elim_team = ""
    non_elim_win_perc = 1.0

    playoff_teams = []

    for team in conf:
        if not teams[team]["Eliminated"]:
            win_perc = teams[team]["Wins"] / float(teams[team]["Games"])
            playoff_teams.append((team,win_perc))

    # Sort the playoff teams by descending record
    playoff_teams = sorted(playoff_teams, key=lambda x: x[1], reverse=True)
    
    # There could exist a tie for 8th place. So, there needs to be a tiebreaker to determine whose 8th.
    # Playing around with the results, I found that there never existed a 3-way tie for 6th or 8th and a 4-way tie for 6th.
    # However, there did exist cases with 2-way ties for 7th, 8th, and a 3-way tie for 7th.
    # In order to handle these scenarios, I will consider 9 'playoff' teams in case one of these ties exist.
    playoff_teams = playoff_teams[:9]

    # In all of these scenarios, the 8th seed must be determined.
    # There exists a 3-way tie for 7th.
    if len(playoff_teams) == 9 and playoff_teams[6][1] == playoff_teams[7][1] == playoff_teams[8][1]:
        # Playing around with the results, there was never a division leader involved in a 3-way tie for 7th. Onto the next tiebreak.
        # Best winning percentage in all games among the tied teams
        tied_teams = {
            playoff_teams[6][0]: 1.0, 
            playoff_teams[7][0]: 1.0, 
            playoff_teams[8][0]: 1.0
        }

        for team in tied_teams:
            other_teams = (other_team for other_team in tied_teams if team != other_team)
            wins = 0
            games = 0

            for other_team in other_teams:
                # print team + " games against " + other_team + " :: " + str(teams[team]["Schedule"][other_team]["Games"])
                wins += teams[team]["Schedule"][other_team]["Wins"]
                games += teams[team]["Schedule"][other_team]["Games"]

            tied_teams[team] = float(wins) / games

        # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
        tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
        # The middle element in tied_teams will be the eight seed
        return tied_teams[1][0]

    # There exists a tie for 8th.
    elif len(playoff_teams) == 9 and playoff_teams[7][1] == playoff_teams[8][1]:
        # There was never a division leader involved in a tie for 8th. Onto the next tiebreak.
        # Best winning percentage in all games among the tied teams
        win_perc = float(teams[playoff_teams[7][0]]["Schedule"][playoff_teams[8][0]]["Wins"]) / \
                    teams[playoff_teams[7][0]]["Schedule"][playoff_teams[8][0]]["Games"]
        if win_perc > 0.5:
            return playoff_teams[7][0]
        elif win_perc < 0.5:
            return playoff_teams[8][0]
        else:
            # Printing the results, this case was reached once between the Bulls and Hornets
            # Neither of them were division leaders not in the same division. Onto the next tiebreak: conference record
            # print "Next tiebreak between " + playoff_teams[7][0] + " " + playoff_teams[8][0]
            tied_teams = {
                playoff_teams[7][0]: 1.0,
                playoff_teams[8][0]: 1.0
            }

            for team in tied_teams:
                other_teams = (other_team for other_team in conf if other_team != team)
                wins = 0
                games = 0

                for other_team in other_teams:
                    wins += teams[team]["Schedule"][other_team]["Wins"]
                    games += teams[team]["Schedule"][other_team]["Games"]

                tied_teams[team] = float(wins) / games

            # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
            tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
            return tied_teams[0][0]


        # print "tie for 8th between: " + playoff_teams[7][0] + " and " + playoff_teams[8][0]
    # There exists a tie for 7th.
    elif playoff_teams[6][1] == playoff_teams[7][1]:
        win_perc = float(teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Wins"]) / \
                   teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Games"]
        if win_perc > 0.5:
            return playoff_teams[6][0]
        elif win_perc < 0.5:
            return playoff_teams[7][0]
        else:

            tied_teams = {
                playoff_teams[6][0]: 1.0,
                playoff_teams[7][0]: 1.0
            }

            if teams[playoff_teams[6][0]]["Division"] == teams[playoff_teams[7][0]]["Division"]:
                # Teams are in the same division. Determine who has better divisional record
                division = teams[playoff_teams[6][0]]["Division"]

                for team in tied_teams:
                    wins = 0
                    games = 0

                    division_teams = (other_team for other_team in conf if teams[other_team]["Division"] == division and other_team != team)
                    # print team
                    for div_team in division_teams:
                        # print div_team + "   " + str(teams[team]["Schedule"][div_team])
                        wins += teams[team]["Schedule"][div_team]["Wins"]
                        games += teams[team]["Schedule"][div_team]["Games"]

                    tied_teams[team] = float(wins) / games

                # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
                tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)             
                    
            else:
                for team in tied_teams:
                    other_teams = (other_team for other_team in conf if other_team != team)
                    wins = 0
                    games = 0

                    for other_team in other_teams:
                        wins += teams[team]["Schedule"][other_team]["Wins"]
                        games += teams[team]["Schedule"][other_team]["Games"]

                    tied_teams[team] = float(wins) / games

                # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
                tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)

            # tied_teams[0] is 7th seed. tied_teams[1] is 8th seed
            return tied_teams[1][0] 

    return playoff_teams[7][0]


# todo: instead of checking for elimination after every result, check after every date. less work. this will be one function instead of two.
# combine the two elimination check functions and call the new function when the date changes

def elimination_check(teams, date):

    # Break it down by conference
    east = []
    west = []

    for team in teams:
        if teams[team]["Conference"] == "East":
            east.append(team)
        else:
            west.append(team)

    eigth_seed = determine_8th_place(east, teams, date)

    eigth_seed = determine_8th_place(west, teams, date)

    # todo: 2 teams could get eliminated from 1 win. possibly find lowest win totals

    return teams





