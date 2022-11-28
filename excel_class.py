import openpyxl
import xlsxwriter


class excel_hours():
    def __init__(self, fileName: str, Path_to_file: str) -> None:
        self.fileName = fileName
        self.path = Path_to_file

    def create_excel(self):
        new_excel_file = xlsxwriter.Workbook(
            self.path+self.fileName)
        sheet1 = new_excel_file.add_worksheet()
        cell_border = new_excel_file.add_format({'border': 1})
        for c in range(3, 11):
            for r in range(6, 42):
                sheet1.write(r, c, None, cell_border)

        new_excel_file.close()
        # dodac zeby np. po 5 min usuwalo ten plik to tylko do modyfikowania

    def open_and_fill(self):
        pass
