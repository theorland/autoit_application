import gspread
import ICS_Config
import ICS_Email_Tester
import datetime

import pyspeedtest

class ICS_Monitor :
    def __init__(self,client):

        AppConfig = ICS_Config.ICS_Config()

        self.cfg_sheet = AppConfig.sheet_load()

        self.SS = client.open_by_key(self.cfg_sheet.id)
        self.cfg_email = AppConfig.email_load()

    def update_result(self,curr_sheet,result):
        curr_sheet.insert_row([""], 4)
        if (curr_sheet.row_count >= self.cfg_sheet.max_row):
            curr_sheet.delete_row(self.cfg_sheet.max_row)
        cell_list = curr_sheet.range("A4:D4")
        cell_list[0].value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cell_list[1].value = round(result.delay)
        cell_list[2].value = result.stat
        cell_list[3].value = result.cond

        curr_sheet.update_cells(cell_list)
        curr_sheet.update_cell(2, 1, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def test_email(self):

        EmailTester = ICS_Email_Tester.ICS_Email_Tester()
        curr_sheet = self.SS.worksheet(self.cfg_sheet.pop)
        result = EmailTester.pop_tester(self.cfg_email)
        result = EmailTester.add_condition(result)

        self.update_result(curr_sheet,result)

        curr_sheet = self.SS.worksheet(self.cfg_sheet.pop_ssl)
        result = EmailTester.pop_tester(self.cfg_email)
        result = EmailTester.add_condition(result)

        self.update_result(curr_sheet, result)


