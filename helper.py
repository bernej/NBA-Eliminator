import datetime, xlrd

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

    for team in team_data:
        teams[team[0]] = {
            "Division": team[1],
            "Conference": team[2],
            "Games": 0,
            "Wins": 0,
            "Losses": 0,
            "Eliminated": False
        }

    return teams

def loser_elimination_check(losing_team, teams, date):
    playoff_teams = 0
    tiebreak_teams = []

    # Break it down by conference
    conf = (team for team in teams if teams[team]["Conference"] == teams[losing_team]["Conference"])

    for team in conf:
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

def determine_last_place(conf, teams):
    # Find the lowest winning non-eliminated team
    worst_non_elim_team = ""
    non_elim_win_perc = 1.0

    playoff_teams = []

    for team in conf:
        if not teams[team]["Eliminated"]:
            win_perc = teams[team]["Wins"] / float(teams[team]["Games"])
            playoff_teams.append((team,win_perc))

    playoff_teams = sorted(playoff_teams, key=lambda x: x[1], reverse=True)
    # Playing around
    playoff_teams = playoff_teams[:9]

    # Determined that there were no 3-way ties for the 6th seed by checking:
    # playoff_teams[5][1] == playoff_teams[6][1] and playoff_teams[6][1] == playoff_teams[7][1]
    # Therefore, we must break the two-way tie for 7th place, then see if the 8th place
    # team can break the tie with elimination-eligible non-playoff teams
    if playoff_teams[7][1] == playoff_teams[8][1]:
        print playoff_teams[7][0] + " and " + playoff_teams[8][0]

    return worst_non_elim_team, non_elim_win_perc


def winner_elimination_check(winning_team, teams, date):

   # Break it down by conference
    conf = (team for team in teams if teams[team]["Conference"] == teams[winning_team]["Conference"])

    worst_team_alive, win_perc = determine_last_place(conf, teams)

    # todo: 2 teams could get eliminated from 1 win. possibly find lowest win totals

    return teams





