from excel_class import excel_hours
import os.path
import sys
import argparse


def check_and_execute(Path_to_file: str):
    """
        Function to check: 
        if excel exists (create new file and insert data from user) or not (open exisiting and insert data from user)
    Args:
        Path_to_file (str): Path to excel file
    """
    def user_choice():

        while True:
            choice = input('''
Press 1 to add hours to cell
Press 2 to detele hours from cell
Press 3 to exit program

            ''')
            if choice == "1" or choice == "2":
                new_excel.user_action(choice)
            else:
                sys.exit(0)

    new_excel = excel_hours(Path_to_file)
    if os.path.isfile(Path_to_file):
        user_choice()
    else:
        new_excel.create_excel()
        user_choice()


def parse_output():
    parser = argparse.ArgumentParser()
    parser.add_argument('-of', '--outputfile')  # keep output
    args = parser.parse_args()
    return vars(args)


if __name__ == "__main__":
    # main invoke
    Path_to_file = parse_output()['outputfile']+".xlsx"
    check_and_execute(Path_to_file)
