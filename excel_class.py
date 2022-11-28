import openpyxl
import xlsxwriter
import os
from time import sleep
from datetime import date


class excel_hours():
    def __init__(self, fileName: str, Path_to_file: str) -> None:
        self.fileName = fileName
        self.months = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December']
        self.path = Path_to_file

    def create_excel(self):
        new_excel_file = xlsxwriter.Workbook(
            self.path+self.fileName)
        sheet1 = new_excel_file.add_worksheet()

        date_today = date.today()
        cell_border = new_excel_file.add_format(
            {'border': 1, 'align': 'center'})
        fst_cell = new_excel_file.add_format(
            {'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
        # column 2 - 9
        # rows 2 - 34
        # create table
        '''make border'''
        for c in range(2, 10):
            for r in range(3, 36):
                sheet1.write(r, c, None, cell_border)

        ''' stylize and fill table '''
        for r in range(3, 36):
            if r == 3:
                sheet1.merge_range(r, 2, r, 3, 'Month/Date', cell_border)
            elif r == 4:
                sheet1.merge_range(r, 2, r, 3, self.months[0], cell_border)
            else:
                first_day_in_the_year = date_today.replace(
                    month=1, day=r-4).strftime("%d/%m/%Y")
                sheet1.merge_range(
                    r, 2, r, 3, first_day_in_the_year, cell_border)

        sheet1.merge_range(2, 2, 2, 9, 'Table with work hours', fst_cell)
        sheet1.merge_range(3, 4, 4, 5, 'From hour to hour', cell_border)
        sheet1.merge_range(3, 6, 4, 7, 'Amount of hours', cell_border)
        sheet1.merge_range(3, 8, 4, 9, 'Total hours per week', cell_border)
        new_excel_file.close()
        # dodac zeby np. po 5 min usuwalo ten plik to tylko do modyfikowania

        sleep(180)
        os.remove(self.path+self.fileName)

    def open_and_fill(self):
        pass
