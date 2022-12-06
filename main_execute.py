from excel_class import excel_hours
import os.path


def check_and_execute(File_name: str, Path_to_file: str):
    """
        Function to check: 
        if excel exists (create new file and insert data from user) or not (open exisiting and insert data from user)
    Args:
        File_name (str): Excel file name
        Path_to_file (str): Path to excel file
    """
    new_excel = excel_hours(File_name, Path_to_file)
    if os.path.isfile(Path_to_file + File_name):
        new_excel.insert_hours_by_user_choice()
    else:
        new_excel.create_excel()
        new_excel.insert_hours_by_user_choice()


if __name__ == "__main__":
    # main invoke
    File_name = "Work_hours.xlsx"
    Path_to_file = "C:/Users/Admin/Desktop/"
    check_and_execute(File_name, Path_to_file)
