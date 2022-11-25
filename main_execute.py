from excel_class import excel_hours
import os.path


def check_and_execute(File_name:str,Path_to_file:str):
    if os.path.isfile(Path_to_file + File_name):
        print("false")
    else:
        new_excel = excel_hours(File_name,Path_to_file)
        new_excel.create_excel()
        
    


if __name__ == "__main__":
    #invoke
    File_name = "Work_hours.xlsx"
    Path_to_file = "C:/Users/onvtsh/Desktop/" 
    check_and_execute(File_name,Path_to_file)