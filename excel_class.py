import openpyxl
import xlsxwriter
import os
from time import sleep
import datetime


class excel_hours():
    def __init__(self, fileName: str, Path_to_file: str) -> None:
        self.fileName = fileName
        self.months = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July',
                       8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}
        self.path = Path_to_file

    def new_sheet_and_fill(self, excel_file, worksheet_name, current_month):
        """
        Method to create new excel worksheet, create table inside and customize it 

        Args:
            excel_file (str): Variable for name of file
            worksheet_name (str): Variable for worksheet name
            current_month (int): key in the dictionary so that extract value(month name) from dictionary and insert into table 
        """
        def make_dates_for_respective_months(row_number):
            '''fill dates and mark weekends'''
            if (current_month == 2) and (row_number - 4 <= 28):
                date = datetime.date(
                    year=2022, month=current_month, day=row_number-4)
                if date.weekday() in (5, 6):
                    for col in range(4, 10):
                        worksheet_name.write(row_number, col, None, excel_file.add_format(
                            {'bg_color': 'red'}))
                else:
                    for col in range(4, 9, 2):
                        if col == 8:
                            if row_number == 32:
                                worksheet_name.merge_range(
                                    row_number, col, row_number, col+1, None, excel_file.add_format({'bottom': 1, 'right': 1}))
                                break
                            worksheet_name.merge_range(
                                row_number, col, row_number, col+1, None, excel_file.add_format({'right': 1}))
                            break
                        worksheet_name.merge_range(
                            row_number, col, row_number, col+1, None, cell_border)
                '''Write dates'''

                worksheet_name.merge_range(
                    row_number, 2, row_number, 3, date.strftime('%d/%m/%Y'), cell_border)
            elif (current_month % 2 == 0 and current_month != 2 and current_month < 8) and (row_number - 4 <= 30)\
                    or (current_month % 2 != 0 and current_month != 2 and current_month >= 8) and (row_number - 4 <= 30):
                date = datetime.date(
                    year=2022, month=current_month, day=row_number-4)
                if date.weekday() in (5, 6):
                    for col in range(4, 10):
                        worksheet_name.write(row_number, col, None, excel_file.add_format(
                            {'bg_color': 'red'}))
                else:
                    for col in range(4, 9, 2):
                        if col == 8:
                            if row_number == 34:
                                worksheet_name.merge_range(
                                    row_number, col, row_number, col+1, None, excel_file.add_format({'bottom': 1, 'right': 1}))
                                break
                            worksheet_name.merge_range(
                                row_number, col, row_number, col+1, None, excel_file.add_format({'right': 1}))
                            break
                        worksheet_name.merge_range(
                            row_number, col, row_number, col+1, None, cell_border)

                worksheet_name.merge_range(
                    row_number, 2, row_number, 3, date.strftime('%d/%m/%Y'), cell_border)
            else:
                if (current_month % 2 != 0 and current_month < 8) or (current_month % 2 == 0 and current_month >= 8):
                    date = datetime.date(
                        year=2022, month=current_month, day=row_number-4)
                    if date.weekday() in (5, 6):
                        for col in range(4, 10):
                            worksheet_name.write(row_number, col, None, excel_file.add_format(
                                {'bg_color': 'red'}))
                    else:
                        for col in range(4, 9, 2):
                            if col == 8:
                                if row_number == 35:
                                    worksheet_name.merge_range(
                                        row_number, col, row_number, col+1, None, excel_file.add_format({'bottom': 1, 'right': 1}))
                                    break
                                worksheet_name.merge_range(
                                    row_number, col, row_number, col+1, None, excel_file.add_format({'right': 1}))
                                break
                            worksheet_name.merge_range(
                                row_number, col, row_number, col+1, None, cell_border)

                    worksheet_name.merge_range(
                        row_number, 2, row_number, 3, date.strftime('%d/%m/%Y'), cell_border)

        def make_border_for_cells(len_month=36):
            for c in range(2, 10):
                for r in range(3, len_month):
                    worksheet_name.write(r, c, None, cell_border)

        '''main triggering for fill and stylize table'''

        cell_border = excel_file.add_format(
            {'border': 1, 'align': 'center'})
        fst_cell = excel_file.add_format(
            {'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})

        if current_month == 2:
            make_border_for_cells(33)
        elif current_month % 2 == 0:
            if current_month > 7:
                make_border_for_cells()
            else:
                make_border_for_cells(35)
        else:
            if current_month > 7:
                make_border_for_cells(35)
            else:
                make_border_for_cells()

        worksheet_name.merge_range(
            2, 2, 2, 9, 'Table with work hours', fst_cell)
        worksheet_name.merge_range(
            3, 4, 4, 5, 'From hour to hour', cell_border)
        worksheet_name.merge_range(3, 6, 4, 7, 'Amount of hours', cell_border)
        worksheet_name.merge_range(
            3, 8, 4, 9, 'Total hours per week', cell_border)

        # stylize and fill table + mark when is weekend
        for r in range(3, 36):
            if excel_file.get_worksheet_by_name(self.months[current_month][:3]):
                if r == 3:
                    worksheet_name.merge_range(
                        r, 2, r, 3, 'Month/Date', cell_border)
                elif r == 4:
                    worksheet_name.merge_range(
                        r, 2, r, 3, self.months[current_month], cell_border)
                else:
                    make_dates_for_respective_months(r)

    def insert_datas_by_user(self):
        # respond_date = input(
        #     'Please enter what date you want fill: Custom date or today\'s date (custom/today) ')
        respond_hours = input(
            'Please enter From what hour to what hour u worked, amount of hours you worked (separate with a space) ')
        # 15-16.30 1.30 15-17 2
        to_write = respond_hours.split(' ')
        date_today = datetime.date.today().strftime('%d/%m/%Y')
        wb = openpyxl.load_workbook(filename=self.path+self.fileName)
        if self.months[int(date_today[3:5])][0:3] in wb.sheetnames:
            ws_current = wb[self.months[int(date_today[3:5])][0:3]]
            for row in ws_current.iter_rows(min_row=6, min_col=3, max_col=4, max_row=36):
                for cell in row:
                    if cell.value == date_today:
                        # skonczylem tutaj tu bedzie trzeba dodac zeby wpisywalo te godziny w cell'u obok
                        print('jest git')

        # print(date_today.strftime('%d/%m/%Y'))
        print(to_write, date_today[3:5], type(self.months.keys()))

    def create_excel(self):
        new_excel_file = xlsxwriter.Workbook(
            self.path+self.fileName)

        for k, v in self.months.items():
            Work_sheet = new_excel_file.add_worksheet(f'{v[:3]}')
            self.new_sheet_and_fill(new_excel_file, Work_sheet, k)

        new_excel_file.close()
        # dodac zeby np. po 5 min usuwalo ten plik to tylko do modyfikowania

        # sleep(180)
        # os.remove(self.path+self.fileName)
