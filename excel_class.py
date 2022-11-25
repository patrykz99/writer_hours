import openpyxl
import xlsxwriter



class excel_hours():
    def __init__(self,fileName:str,Path_to_file:str) -> None:
        self.fileName = fileName
        self.path = Path_to_file
        
    def create_excel(self):
        new_excel_file = openpyxl.Workbook()
        sheet1 = new_excel_file.active
        
        
        
        
        
        new_excel_file.save(self.path+self.fileName)
        
    def open_and_fill(self):
        pass
        
        
        