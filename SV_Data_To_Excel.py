import win32com.client as win32


class ExcelOperate():
    def __init__(self):
        self.xl = win32.gencache.EnsureDispatch('Excel.Application')
        self.ss = self.xl.Workbooks.Add()
        self.sh = self.ss.ActiveSheet
        self.sh.Cells(1, 1).Value = 'Chinese Name'
        self.sh.Cells(1, 2).Value = 'English Name'
        self.sh.Cells(1, 3).Value = 'Spell Type'
        self.sh.Cells(1, 4).Value = 'Spell Description'
        self.sh.Rows(1).Font.Bold = True    # titles font are bold
        self.sh.Cells(1, 1).Interior.ColorIndex = 4     # background color of row 1 is green
        self.sh.Cells(1, 2).Interior.ColorIndex = 4
        self.sh.Cells(1, 3).Interior.ColorIndex = 4
        self.sh.Cells(1, 4).Interior.ColorIndex = 4
        self.xl.Visible = True      # the excel will show on the screen

    def SaveToExcel(self, Row, line, Data):
        """
        Write Data to a Excel file in the specified cell(row, line)
        :param Row: the row number of the cell, no less than 1
        :param line: the line number of the cell, no less than 1
        :param Data: the data to be written in the specified cell
        :return: NULL
        """
        self.sh.Cells(Row, line).Value = Data

    def CloseExcelFile(self):
        """
        Close the excel file
        :return: NULL
        """
        self.ss.close(False)    # SaveChanges = False
        self.xl.Application.Quit()

    def AdjustFontSize(self):
        """
        adjust some parameters of the excel file so better for reading
        :return: NULL
        """
        self.sh.Columns("A").Delete()


