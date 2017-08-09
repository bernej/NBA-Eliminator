import sys
from helper import *

def main(args):

    path = 'Analytics_Attachment.xlsx'

    teams = read_sheet(path, 0)

    scores = read_sheet(path, 1)

    for team in teams:
        print team

    for score in scores:
        print score

    # todo: figure out how to write to xlsx file

if __name__ == "__main__":
    main(sys.argv[0:])