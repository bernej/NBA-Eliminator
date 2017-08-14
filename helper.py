# Helper module to main script
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


def break_3way_tie(tied_teams, teams):
    # Go through each tied team
    for team in tied_teams:
        # Determine their win % against the other tied teams
        other_teams = (other_team for other_team in tied_teams if team != other_team)
        wins = 0
        games = 0
        for other_team in other_teams:
            wins += teams[team]["Schedule"][other_team]["Wins"]
            games += teams[team]["Schedule"][other_team]["Games"]

        tied_teams[team] = float(wins) / games
    
    # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied teams sorted by their win % against each other
    return tied_teams


def rank_conf_record(tied_teams, conf, teams):
    # Go through each tied team
    for team in tied_teams:
        # Determine their win % against the conference
        other_teams = (other_team for other_team in conf if other_team != team)
        wins = 0
        games = 0
        for other_team in other_teams:
            wins += teams[team]["Schedule"][other_team]["Wins"]
            games += teams[team]["Schedule"][other_team]["Games"]

        tied_teams[team] = float(wins) / games

    # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied teams sorted by their win % against the conference
    return tied_teams


def rank_div_record(tied_teams, conf, division, teams):
    # Go through each tied team
    for team in tied_teams:
        # Determine their win % against their division
        division_teams = (other_team for other_team in conf if teams[other_team]["Division"] == division and other_team != team)
        wins = 0
        games = 0
        for div_team in division_teams:
            wins += teams[team]["Schedule"][div_team]["Wins"]
            games += teams[team]["Schedule"][div_team]["Games"]

        tied_teams[team] = float(wins) / games

    # Printing the results here, there was never a need to go to the next tiebreak as each win percentage was different
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied teams sorted by their win % against the division
    return tied_teams

def determine_div_tiebreak(teams, conf, eigth_seed, team, date):
    division = teams[eigth_seed]["Division"]

    tied_teams = {
        eigth_seed: {
            "Wins": 0,
            "Games": 0
        },
        team: {
            "Wins": 0,
            "Games": 0
        }
    }

    for tied_team in tied_teams:
        # print tied_team
        division_teams = (other_team for other_team in conf if teams[other_team]["Division"] == division and other_team != tied_team)

        for div_team in division_teams:

            tied_teams[tied_team]["Wins"] += teams[tied_team]["Schedule"][div_team]["Wins"]
            tied_teams[tied_team]["Games"] += teams[tied_team]["Schedule"][div_team]["Games"]
    print tied_teams
    div_games_left = 16 - tied_teams[team]["Games"]
    possible_div_wins = div_games_left + tied_teams[team]["Wins"]

    if tied_teams[eigth_seed]["Wins"] > possible_div_wins:
        teams[team]["Eliminated"] = True
        print team + " eliminated using division record tiebreak against " + eigth_seed + " on " + date
    else:
        teams = determine_conf_tiebreak(teams, conf, eigth_seed, team, date)

    return teams

def determine_conf_tiebreak(teams, conf, eigth_seed, team, date):

    tied_teams = {
        eigth_seed: {
            "Wins": 0,
            "Games": 0
        },
        team: {
            "Wins": 0,
            "Games": 0
        }
    }
    for tied_team in tied_teams:
        conf_teams = (other_team for other_team in conf if other_team != tied_team)
        for conf_team in conf_teams:
            tied_teams[tied_team]["Wins"] += teams[tied_team]["Schedule"][conf_team]["Wins"]
            tied_teams[tied_team]["Games"] += teams[tied_team]["Schedule"][conf_team]["Games"]

    conf_games_left = 52 - tied_teams[team]["Games"]
    possible_conf_wins = conf_games_left + tied_teams[team]["Wins"]
    
    print tied_teams

    if tied_teams[eigth_seed]["Wins"] > possible_conf_wins:
        teams[team]["Eliminated"] = True
        print team + " eliminated using conference record tiebreak against " + eigth_seed + " on " + date
    else:
        print "FML on " + date

    return teams

# Helper function to determine the 8th place team. This team will then be cross-examined
# against the non-eliminated, non-playoff teams to see if it owns tiebreakers
def determine_8th_place(conf, teams):
    
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

        tied_teams = break_3way_tie(tied_teams, teams)
        # tied_teams[0] is the 7th seed, tied_teams[1] is the 8th seed based on the win % among tied teams
        return tied_teams[1][0]

    # There exists a 2-way tie for 8th.
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
            # Neither of them were division leaders or in the same division. Onto the next tiebreak: conference record
            tied_teams = {
                playoff_teams[7][0]: 1.0,
                playoff_teams[8][0]: 1.0
            }

            tied_teams = rank_conf_record(tied_teams, conf, teams)
            # tied_teams[0] is the 8th seed based on the conference record tiebreak
            return tied_teams[0][0]

    # There exists a 2-way tie for 7th.
    elif playoff_teams[6][1] == playoff_teams[7][1]:
        win_perc = float(teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Wins"]) / \
                   teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Games"]
        if win_perc > 0.5:
            return playoff_teams[6][0]
        elif win_perc < 0.5:
            return playoff_teams[7][0]
        else:
            # The tied teams are .500 against each other. Onto divisional record if they're in the same division, otherwise conf record.
            tied_teams = {
                playoff_teams[6][0]: 1.0,
                playoff_teams[7][0]: 1.0
            }
            # Teams are in the same division. Determine who has better divisional record
            if teams[playoff_teams[6][0]]["Division"] == teams[playoff_teams[7][0]]["Division"]:
                division = teams[playoff_teams[6][0]]["Division"]
                tied_teams = rank_div_record(tied_teams, conf, division, teams)          
            # Otherwise, teams aren't in the same division. Determine who has the better conference record
            else:
                tied_teams = rank_conf_record(tied_teams, conf, teams)

            # tied_teams[0] is 7th seed. tied_teams[1] is 8th seed based on division record or conference record tiebreak
            return tied_teams[1][0] 

    # There was no tie involving the 8th seed. Return the 8th place team based on record
    return playoff_teams[7][0]


def eliminate(conf, teams, eigth_seed, scores, date):
    
    win_perc = float(teams[eigth_seed]["Wins"]) / teams[eigth_seed]["Games"]

    losing_non_elim_teams = (team for team in conf if eigth_seed != team and win_perc > float(teams[team]["Wins"]) / teams[team]["Games"] and not teams[team]["Eliminated"])

    for team in losing_non_elim_teams:
        games_left = 82 - teams[team]["Games"]
        possible_wins = games_left + teams[team]["Wins"]

        # Extreme scenario: current eigth seed could lose out and current non playoff team could win out
        if teams[eigth_seed]["Wins"] == possible_wins:

            division_flag = ( teams[eigth_seed]["Division"] == teams[team]["Division"] )

            # If team is in the same division as the eigth seed, then they play 4 times.
            if division_flag and teams[eigth_seed]["Schedule"][team]["Wins"] >= 3:
                teams[team]["Eliminated"] = True
                print team + " eliminated from losing 3 or more games to eigth seed in same division on " + date
            
            # Eigth seed and losing team are in same division, but eigth seed has not beaten them the majority of times yet.
            elif division_flag and teams[eigth_seed]["Schedule"][team]["Wins"] <= 2:
                teams = determine_div_tiebreak(teams, conf, eigth_seed, team, date)

            # Teams aren't in the same division
            else:
                # Loop through the scores and determine how many times these teams play each other.
                # (Some non-division, conference opponents are played 4 times a season while others only 3)
                games_against_eachother = 0
                for score in scores:
                    if (score[1] == eigth_seed or score[1] == team) and (score[2] == eigth_seed or score[2] == team):
                        games_against_eachother += 1
                if games_against_eachother == 4 and teams[eigth_seed]["Schedule"][team]["Wins"] >= 3:
                    teams[team]["Eliminated"] = True
                    print team + " eliminated from losing majority of games to " + eigth_seed + " on " + date
                elif games_against_eachother == 3 and teams[eigth_seed]["Schedule"][team]["Wins"] >= 2:
                    teams[team]["Eliminated"] = True
                    print team + " eliminated from losing majority of games to " + eigth_seed + " on " + date
                else:
                    teams = determine_conf_tiebreak(teams, conf, eigth_seed, team, date)

        elif teams[eigth_seed]["Wins"] > possible_wins:
            teams[team]["Eliminated"] = True
            print team + " eliminated on " + date

    return teams


def elimination_check(teams, scores, date):

    # Break it down by conference
    east = []
    west = []
    conferences = []

    for team in teams:
        if teams[team]["Conference"] == "East":
            east.append(team)
        else:
            west.append(team)

    conferences.append(east)
    conferences.append(west)

    for conf in conferences:
        # todo: when determining 8th place, also return the rest of the playoff teams for tiebreak scenarios 5 and 6
        eigth_seed = determine_8th_place(conf, teams)
        teams = eliminate(conf, teams, eigth_seed, scores, date)
        # print eigth_seed

    # todo: look at 9th seeds or more if they are less than 4 wins behind 8th seed or cross-check scores

    # eigth_seed = determine_8th_place(west, teams)
    # print eigth_seed
    # print '\n'
    # # todo: 2 teams could get eliminated from 1 win. possibly find lowest win totals
    # for team in teams[eigth_seed]["Schedule"]:
    #     if teams[eigth_seed]["Schedule"][team]["Games"] == 4:
    #         print team
    return teams





