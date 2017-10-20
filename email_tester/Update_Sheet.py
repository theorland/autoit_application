import datetime
from ICS_RDP_Tester import ICS_RDP_Tester
from ICS_Email_Tester import ICS_Email_Tester
from ICS_Shared_Config import ICS_Shared_Config
from threading import Thread


class ICS_Monitor:
    def __init__(self):

        AppConfig = ICS_Shared_Config.Config()

        self.cfg_sheet = AppConfig.sheet_load()
        client = ICS_Shared_Config.Gspread_Client()

        self.SS = client.open_by_key(self.cfg_sheet.id)
        self.cfg_email = AppConfig.email_load()
        self.cfg_rdp = AppConfig.remote_load()

    def update_result(self,curr_sheet,result):
        curr_sheet.insert_row([""], 4)

        if (curr_sheet.row_count >= int(self.cfg_sheet.max_row)):
            curr_sheet.resize(self.cfg_sheet.max_row)

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
        result = EmailTester.pop_ssl_tester(self.cfg_email)
        result = EmailTester.add_condition(result)

        self.update_result(curr_sheet, result)

    def test_remote(self):
        tester = ICS_RDP_Tester(self.cfg_rdp)
        curr_sheet = self.SS.worksheet(self.cfg_sheet.rdp)
        result = tester.run_test()

        self.update_result(curr_sheet,result)

class ICS_Monitor_POP(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 1
        self.name = "Monitor_POP"
        self.counter = 1

    def run(self):
        ICS_Shared_Config.log("POP Test : Start ")
        try:
            EmailTester = ICS_Email_Tester()
            curr_sheet = self.SS.worksheet(self.cfg_sheet.pop)

            result = EmailTester.pop_tester(self.cfg_email)
            result = EmailTester.add_condition(result)

            self.update_result(curr_sheet, result)
            ICS_Shared_Config.log("POP Test : Complete ")
        except:
            ICS_Shared_Config.log("POP Test : Failed ")

class ICS_Monitor_POP_SSL(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 2
        self.name = "Monitor_POP_SSL"
        self.counter = 2

    def run(self):

        ICS_Shared_Config.log("POP_SSL Test : Start ")
        try:
            EmailTester = ICS_Email_Tester()

            curr_sheet = self.SS.worksheet(self.cfg_sheet.pop_ssl)
            result = EmailTester.pop_ssl_tester(self.cfg_email)
            result = EmailTester.add_condition(result)

            self.update_result(curr_sheet, result)
            ICS_Shared_Config.log("POP_SSL Test : Complete ")
        except:
            ICS_Shared_Config.log("POP_SSL Test : Failed ")


class ICS_Monitor_Remote(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 3
        self.name = "Monitor_RDP"
        self.counter = 3


    def run(self):
        ICS_Shared_Config.log("RDP Test : Start ")
        try:
            tester = ICS_RDP_Tester(self.cfg_rdp)
            curr_sheet = self.SS.worksheet(self.cfg_sheet.rdp)
            result = tester.run_test()

            self.update_result(curr_sheet, result)

            ICS_Shared_Config.log("RDP Test : Complete ")
        except:
            ICS_Shared_Config.log("RDP Test : Failed ")


