# Helper module to the main script -- contains the functionality to determine & break tie-breakers and eliminate teams
import datetime, xlrd, operator

TOTAL_DIVISIONAL_GAMES = 16
TOTAL_CONFERENCE_GAMES = 52
TOTAL_GAMES = 82

# Reads the given input xlsx file, and returns a list of the data contained on an excel sheet
# file  -- excel file to investigate
# index -- index of the sheet on the excel file to read 
def read_sheet(file, index):
    file = 'Analytics_Attachment.xlsx'
    workbook = xlrd.open_workbook(file)
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


# Takes in the raw data read from the Division_Info sheet and initalizes a dictionary corresponding to this data
# team data -- list of data from the Division_Info sheet
def initialize_team_data(team_data):
    # The dictionary data structure that will be used throughout the module
    teams = {}
    # Initialize appropriate data for each team
    for team in team_data:
        teams[team[0]] = {
            "Division": team[1],
            "Conference": team[2],
            "Games": 0,
            "Wins": 0,
            "Losses": 0,
            "Eliminated": False,
            "Elimination Date": "",
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
    # Return the initialized teams dictionary data structure
    return teams


# Output the eliminated teams and their elimination dates to the output csv file
# teams -- overarching dictionary of data for every team
def output_eliminated_teams(teams):
    non_playoff_teams = (team for team in teams if teams[team]["Eliminated"])
    eliminated_teams = {}
    # Format the dates to MM/DD/YYYY instead of YYYY-MM-DD
    for team in non_playoff_teams:
        date = teams[team]["Elimination Date"].split("-")
        date = date[1] + "/" + date[2] + "/" + date[0]
        eliminated_teams[team] = date
    # Sort the eliminated teams alphabetically
    eliminated_teams = sorted(eliminated_teams.items(), key=lambda x: x[0])
    # Write to the output file
    with open('output.csv', 'w') as output:
        output.write(','.join(["Team", "Date Eliminated"]) + '\n')
        for team in eliminated_teams:
            output.write(','.join(team) + '\n')


# Takes in three teams with identical win %, determines their win % against each other, & returns the three teams sorted based on that
# tied_teams -- three teams with identical win %, mapped to their win % against each other
# teams      -- overarching dictionary of data for every team
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
        # Calculate the win % against the other tied teams
        tied_teams[team] = float(wins) / games
    # Sort the tied_teams by win % against each other
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied teams sorted by their win % against each other
    return tied_teams


# Takes in two teams, determines their win % against conference opponents, & returns the two teams sorted based on that
# tied_teams -- two teams, mapped to their win % against conference opponents
# conf       -- teams in the tied_teams' conference
# teams      -- overarching dictionary of data for every team
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
        # Calculate their win % against the conference
        tied_teams[team] = float(wins) / games
    # Sort the tied_teams by conference win %
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied teams sorted by their win % against the conference
    return tied_teams


# Takes in two teams, determines their win % against divisional opponents, & returns the two teams sorted based on that
# tied_teams -- two teams in the same division, mapped to their divisional win %
# conf       -- teams in the tied_teams' conference
# division   -- name of the tied_teams' division
# teams      -- overarching dictionary of data for every team
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
        # Calculate their divisional win %
        tied_teams[team] = float(wins) / games
    # Sort the tied_teams by divisional win %
    tied_teams = sorted(tied_teams.items(), key=operator.itemgetter(1), reverse=True)
    # Return the tied_teams sorted by their win % against the division
    return tied_teams


# Tie-breaker to determine if the eigth_seed has clinched a better division record than the non_eigth_seed
# teams                    -- overarching dictionary of data for every team
# elig_playoff_teams       -- current 1-8 seeds in the two teams' conference
# conf                     -- teams in the two teams' conference
# eigth_seed               -- the current eigth seed in conf
# non_eigth_seed           -- the potential eliminated team
# scores                   -- every score from the season (to check for schedule purposes)
# date                     -- the current date
# other_elig_playoff_teams -- current 1-8 seeds in the other conference
def determine_div_tiebreak(teams, elig_playoff_teams, conf, eigth_seed, non_eigth_seed, scores, date, other_elig_playoff_teams):
    division = teams[eigth_seed]["Division"] # This will be the same as the non_eigth_seed's division
    # Keep track of the two teams' worst and best possible divisional record
    tied_teams = {
        eigth_seed: {
            "Wins": 0,
            "Games": 0
        },
        non_eigth_seed: {
            "Wins": 0,
            "Games": 0
        }
    }
    # Go through both teams
    for tied_team in tied_teams:
        # Find their divisional opponents and record their current divisional record
        division_teams = (other_team for other_team in conf if teams[other_team]["Division"] == division and other_team != tied_team)
        for div_team in division_teams:
            tied_teams[tied_team]["Wins"] += teams[tied_team]["Schedule"][div_team]["Wins"]
            tied_teams[tied_team]["Games"] += teams[tied_team]["Schedule"][div_team]["Games"]
    # Determine the non_eigth_seed's best possible divisional record
    div_games_left = TOTAL_DIVISIONAL_GAMES - tied_teams[non_eigth_seed]["Games"]
    possible_div_wins = div_games_left + tied_teams[non_eigth_seed]["Wins"]
    # If the eigth_seed already has a better divisional record than the non_eigth_seed's could potentially have, then the non_eigth_seed is eliminated
    if tied_teams[eigth_seed]["Wins"] > possible_div_wins:
        teams[non_eigth_seed]["Eliminated"] = True
        teams[non_eigth_seed]["Elimination Date"] = date
    # If they could potentially have the same divisional record, then go to the next tiebreak, conference record
    elif tied_teams[eigth_seed]["Wins"] == possible_div_wins:
        teams = determine_conf_tiebreak(teams, elig_playoff_teams, conf, eigth_seed, non_eigth_seed, scores, date, other_elig_playoff_teams)

    return teams


# Tie-breaker to determine if the eigth_seed has clinched a better conference record than the non_eigth_seed
# teams                    -- overarching dictionary of data for every team
# elig_playoff_teams       -- current 1-8 seeds in the two teams' conference
# conf                     -- teams in the two teams' conference
# eigth_seed               -- the current eigth seed in conf
# non_eigth_seed           -- the potential eliminated team
# scores                   -- every score from the season (to check for schedule purposes)
# date                     -- the current date
# other_elig_playoff_teams -- current 1-8 seeds in the other conference
def determine_conf_tiebreak(teams, elig_playoff_teams, conf, eigth_seed, non_eigth_seed, scores, date, other_elig_playoff_teams):
    # Keep track of the two teams' worst and best possible divisional record
    tied_teams = {
        eigth_seed: {
            "Wins": 0,
            "Games": 0
        },
        non_eigth_seed: {
            "Wins": 0,
            "Games": 0
        }
    }
    # Go through both teams
    for tied_team in tied_teams:
        # Find their conference opponents and record their current conference record
        conf_teams = (other_team for other_team in conf if other_team != tied_team)
        for conf_team in conf_teams:
            tied_teams[tied_team]["Wins"] += teams[tied_team]["Schedule"][conf_team]["Wins"]
            tied_teams[tied_team]["Games"] += teams[tied_team]["Schedule"][conf_team]["Games"]
    # Determine the non_eigth_seed's best possible conference record
    conf_games_left = TOTAL_CONFERENCE_GAMES - tied_teams[non_eigth_seed]["Games"]
    possible_conf_wins = conf_games_left + tied_teams[non_eigth_seed]["Wins"]
    # If the eigth seed already has more conference wins than the non_eigth_seed could potentially have, then the non_eigth_seed is eliminated
    if tied_teams[eigth_seed]["Wins"] > possible_conf_wins:
        teams[non_eigth_seed]["Eliminated"] = True
        teams[non_eigth_seed]["Elimination Date"] = date
    # If they could potentially have the same conference record, then go to the next tiebreak, win % against same conference playoff eligible teams
    elif tied_teams[eigth_seed]["Wins"] == possible_conf_wins:
        # Determine who would own the better win % against same-conf playoff eligible opponents; non_eigth_seed will be eliminated in the function
        teams, next_tiebreak = determine_playoff_record(teams, elig_playoff_teams, tied_teams, eigth_seed, non_eigth_seed, scores, date)
        # If they could potentially have the same win % against same conf playoff eligible opponents (unlikely), then determine their win % against the other conf playoff eligible teams
        if next_tiebreak:
            # Determine who would own the better win % against other-conf playoff eligible opponents; non_eigth_seed will be eliminated in the function
            teams, final_tiebreak = determine_playoff_record(teams, other_elig_playoff_teams, tied_teams, eigth_seed, non_eigth_seed, scores, date)

    return teams


# Tie-breaker to determine who between the eigth_seed and non_eigth_seed owns a better record against the specified set of playoff teams
# teams              -- overarching dictionary of data for every team
# elig_playoff_teams -- current 1-8 seeds in the two teams' conference
# tied_teams         -- eigth_seed and non_eigth_seed, mapped to their record against the elig_playoff_teams
# eigth_seed         -- the current eigth seed in conf
# non_eigth_seed     -- the potential eliminated team
# scores             -- every score from the season (to check for schedule purposes)
# date               -- the current date
def determine_playoff_record(teams, elig_playoff_teams, tied_teams, eigth_seed, non_eigth_seed, scores, date):
    # Go through both teams
    for team in tied_teams:
        # Reset their tied_team data
        tied_teams[team]["Wins"] = 0
        tied_teams[team]["Games"] = 0
        tied_teams[team]["Games Scheduled"] = 0
        # Record their record against each elig_playoff_team
        for playoff_team in elig_playoff_teams:
            # Edge case for the eigth_seed
            if playoff_team != team:
                tied_teams[team]["Wins"] += teams[team]["Schedule"][playoff_team]["Wins"]
                tied_teams[team]["Games"] += teams[team]["Schedule"][playoff_team]["Games"]
        # Determine how many games they have left against the playoff teams (varies for same conference, 16 for opposite conference)
        for score in scores:
            # Edge case for eigth seed, make sure it's a game between a playoff team and the eigth_seed/non_eigth_seed
            if ((score[1] in elig_playoff_teams and score[1] != team) or (score[2] in elig_playoff_teams and score[2] != team)) \
                 and (score[1] == team or score[2] == team):
                tied_teams[team]["Games Scheduled"] += 1
    # Determine the non_eigth_seed's best potential win %
    possible_wins = tied_teams[non_eigth_seed]["Wins"] + (tied_teams[non_eigth_seed]["Games Scheduled"] - tied_teams[non_eigth_seed]["Games"])
    possible_win_perc = float(possible_wins) / tied_teams[non_eigth_seed]["Games Scheduled"]
    # Flag to determine whether to investigate the eigth_seed/non_eigth_seed's record against opposite conference playoff opponents
    next_tiebreak = False
    # If the non_eigth_seed could potentially have the same win % as the eigth_seed against elig_playoff_teams
    if possible_win_perc == (float(tied_teams[eigth_seed]["Wins"]) / tied_teams[eigth_seed]["Games Scheduled"]):
        next_tiebreak = True
    # If the eigth_seed clinched a better record against elig_playoff_teams than the non_eigth_seed is eliminated
    elif (float(tied_teams[eigth_seed]["Wins"]) / tied_teams[eigth_seed]["Games Scheduled"]) > possible_win_perc:
        teams[non_eigth_seed]["Eliminated"] = True
        teams[non_eigth_seed]["Elimination Date"] = date

    return teams, next_tiebreak


# Helper function to determine the 8th place team. This team will then be cross-examined against the non-eliminated, non-playoff teams
# conf  -- teams in the two teams' conference
# teams -- overarching dictionary of data for every team
def determine_8th_place(conf, teams):
    # List of the current non eliminated teams in the specified conference
    non_eliminated_teams = []
    # Go through each team in the conference
    for team in conf:
        if not teams[team]["Eliminated"]:
            win_perc = teams[team]["Wins"] / float(teams[team]["Games"])
            non_eliminated_teams.append((team,win_perc))
    # Sort the non eliminated teams by descending record
    playoff_teams = sorted(non_eliminated_teams, key=lambda x: x[1], reverse=True)
    
    # Playing around with the results, I found that there never existed a 3-way tie for 6th or 8th and a 4-way tie for 6th.
    # However, there did exist cases with a 2-way tie for 7th, 2-way tie for 8th, and a 3-way tie for 7th.
    # In order to handle these scenarios, I will consider 9 'playoff' teams in case one of these ties exist.
    # In all of these scenarios, the 8th seed must be determined.
    playoff_teams = playoff_teams[:9]

    # There exists a 3-way tie for 7th.
    if len(playoff_teams) == 9 and playoff_teams[6][1] == playoff_teams[7][1] == playoff_teams[8][1]:
        # Printing the results here, there was never a division leader involved in a 3-way tie for 7th. 
        # Onto the next tiebreak, best winning percentage in all games among the tied teams
        win_perc = playoff_teams[6][1]
        tied_teams = {
            playoff_teams[6][0]: 1.0, 
            playoff_teams[7][0]: 1.0, 
            playoff_teams[8][0]: 1.0
        }
        # Determine the best winning percentage in all games among the tied teams
        tied_teams = break_3way_tie(tied_teams, teams)
        # tied_teams[0] is the 7th seed, tied_teams[1] is the 8th seed based on the win % among tied teams
        playoff_teams[6] = (tied_teams[0][0], win_perc)
        playoff_teams[7] = (tied_teams[1][0], win_perc)

    # There exists a 2-way tie for 8th.
    elif len(playoff_teams) == 9 and playoff_teams[7][1] == playoff_teams[8][1]:
        # There was never a division leader involved in a tie for 8th.
        # Onto the next tiebreak, best winning percentage in all games among the tied teams
        # win_perc = eigth seed's record against the ninth seed
        win_perc = float(teams[playoff_teams[7][0]]["Schedule"][playoff_teams[8][0]]["Wins"]) / \
                    teams[playoff_teams[7][0]]["Schedule"][playoff_teams[8][0]]["Games"]
        if win_perc > 0.5:
            # eigth seed is already in place
            pass
        elif win_perc < 0.5:
            # Update the eigth seed
            playoff_teams[7] = (playoff_teams[8][0], playoff_teams[8][1])
        else:
            # Printing the results, this case was reached once between the Bulls and Hornets
            # Neither of them were division leaders or in the same division. Onto the next tiebreak: conference record
            tied_teams = {
                playoff_teams[7][0]: 1.0,
                playoff_teams[8][0]: 1.0
            }
            # Determine the best winning percentage in the conference
            tied_teams = rank_conf_record(tied_teams, conf, teams)
            # tied_teams[0] is the 8th seed based on the conference record tiebreak
            playoff_teams[7] = (tied_teams[0][0], playoff_teams[7][1])

    # There exists a 2-way tie for 7th.
    elif playoff_teams[6][1] == playoff_teams[7][1]:
        # win_perc = seventh seed's record against the eigth seed
        win_perc = float(teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Wins"]) / \
                   teams[playoff_teams[6][0]]["Schedule"][playoff_teams[7][0]]["Games"]
        if win_perc > 0.5:
            # seventh and eigth seed's are properly ranked
            pass
        elif win_perc < 0.5:
            # Flip the seventh and eigth seeds
            temp = playoff_teams[6][0]
            playoff_teams[6] = (playoff_teams[7][0], playoff_teams[7][1])
            playoff_teams[7] = (temp, playoff_teams[6][1])
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
            playoff_teams[6] = (tied_teams[0][0], playoff_teams[7][1])
            playoff_teams[7] = (tied_teams[1][0], playoff_teams[7][1])

    # Return the playoff team names and ignore their win %
    playoff_teams = [team_name[0] for team_name in playoff_teams]
    return playoff_teams[0:8]


# Wrapper function to compare elimination eligible teams to the current eigth seed and potentially eliminate them
# conf                     -- teams in the two teams' conference
# elig_playoff_teams       -- current 1-8 seeds in the two teams' conference
# teams                    -- overarching dictionary of data for every team
# eigth_seed               -- the current eigth seed in conf
# scores                   -- every score from the season (to check for schedule purposes)
# date                     -- the current date
# other_elig_playoff_teams -- current 1-8 seeds in the other conference
def eliminate(conf, elig_playoff_teams, teams, eigth_seed, scores, date, other_elig_playoff_teams):
    # Eigth seed's win percentage
    win_perc = float(teams[eigth_seed]["Wins"]) / teams[eigth_seed]["Games"]
    # Teams that aren't currently seed 1-8 but also not eliminated
    losing_non_elim_teams = (team for team in conf if team not in elig_playoff_teams and not teams[team]["Eliminated"])
    # For each losing non-eliminated team determine their best possible record if they won out
    for team in losing_non_elim_teams:
        games_left = TOTAL_GAMES - teams[team]["Games"]
        possible_wins = games_left + teams[team]["Wins"]
        # Extreme scenario: current eigth seed could lose out and current non playoff team could win out
        if teams[eigth_seed]["Wins"] == possible_wins:
            # Flag indicating if they're in the same division
            division_flag = ( teams[eigth_seed]["Division"] == teams[team]["Division"] )
            # If team is in the same division as the eigth seed, then they play 4 times.
            if division_flag and teams[eigth_seed]["Schedule"][team]["Wins"] >= 3:
                # Eigth seed has won majority of games, other team is eliminated
                teams[team]["Eliminated"] = True
                teams[team]["Elimination Date"] = date
            
            # Scenario when eigth seed and losing team are in same division and split 2-2 against each other
            elif division_flag and teams[eigth_seed]["Schedule"][team]["Wins"] == 2:
                teams = determine_div_tiebreak(teams, elig_playoff_teams, conf, eigth_seed, team, scores, date, other_elig_playoff_teams)

            # Teams aren't in the same division
            elif not division_flag:
                # Loop through the scores and determine how many times these teams play each other (Some conference opponents are played 4 times a season while others only 3).
                games_against_eachother = 0
                for score in scores:
                    if (score[1] == eigth_seed or score[1] == team) and (score[2] == eigth_seed or score[2] == team):
                        games_against_eachother += 1
                # Eigth seed has clinched a better record than the current non playoff team
                if games_against_eachother == 4 and teams[eigth_seed]["Schedule"][team]["Wins"] >= 3:
                    teams[team]["Eliminated"] = True
                    teams[team]["Elimination Date"] = date
                # Eigth seed has clinched a better record than the current non playoff team
                elif games_against_eachother == 3 and teams[eigth_seed]["Schedule"][team]["Wins"] >= 2:
                    teams[team]["Eliminated"] = True
                    teams[team]["Elimination Date"] = date
                # Eigth seed and current non playoff team could potentially split the remainder of their games, onto the conference record tiebreak
                elif games_against_eachother == 4 and teams[eigth_seed]["Schedule"][team]["Wins"] == 2:
                    teams = determine_conf_tiebreak(teams, elig_playoff_teams, conf, eigth_seed, team, scores, date, other_elig_playoff_teams)
        # If the eigth seed has already clinched a better record than the current non-playoff team, then the non-playoff team is eliminated
        elif teams[eigth_seed]["Wins"] > possible_wins:
            teams[team]["Eliminated"] = True
            teams[team]["Elimination Date"] = date

    return teams


# Wrapper function to perform elimination functionality
# teams  -- overarching dictionary of data for every team
# scores -- every score from the season (to check for schedule purposes)
# date   -- the current date
def elimination_check(teams, scores, date):

    # Break down the teams by conference
    east = []
    west = []
    for team in teams:
        if teams[team]["Conference"] == "East":
            east.append(team)
        else:
            west.append(team)

    # Eliminate teams from each conference in parallel
    # Grab the eight playoff teams from each conference
    e_elig_playoff_teams = determine_8th_place(east, teams)
    w_elig_playoff_teams = determine_8th_place(west, teams)
    # Grab the eigth seed from each conference
    e_eigth_seed = e_elig_playoff_teams[7]
    w_eigth_seed = w_elig_playoff_teams[7]
    # Potentially eliminate eligible teams from each conference
    teams = eliminate(east, e_elig_playoff_teams, teams, e_eigth_seed, scores, date, w_elig_playoff_teams)
    teams = eliminate(west, w_elig_playoff_teams, teams, w_eigth_seed, scores, date, e_elig_playoff_teams)

    return teams

