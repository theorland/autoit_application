from datetime import datetime
from ICS_RDP_Tester import ICS_RDP_Tester
from ICS_Email_Tester import ICS_Email_Tester
from ICS_Shared_Config import ICS_Shared_Config
from ICS_Http_Tester import ICS_Http_Tester
from threading import Thread
from collections import namedtuple
import ICS_Speed_Test, os

class ICS_Monitor:
    SheetCriteria = namedtuple('SheetCriteria', \
        ['great','good', 'average'])

    def __init__(self):

        AppConfig = ICS_Shared_Config.Config()
        self.cfg_sheet = AppConfig.sheet_load()

        self.SS = ICS_Shared_Config.Gspread_Open(self.cfg_sheet.id)

        self.cfg_email = AppConfig.email_load()
        self.cfg_rdp = AppConfig.remote_load()
        self.cfg_speedtest = AppConfig.speedtest_load()

    def update_criteria(self,curr_sheet, criteria):
        cell_list = curr_sheet.range("B1:D1")# type SpreadSheet

        cell_list[0].value = criteria.great
        cell_list[1].value = criteria.good
        cell_list[2].value = criteria.average

        curr_sheet.update_cells(cell_list)


    def update_result(self,curr_sheet,result):
        curr_sheet.insert_row([""], 4)

        if (curr_sheet.row_count >= int(self.cfg_sheet.max_row)):
            curr_sheet.resize(self.cfg_sheet.max_row)

        cell_list = curr_sheet.range("A4:D4")
        cell_list[0].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cell_list[1].value = round(result.delay)
        cell_list[2].value = result.stat
        cell_list[3].value = result.cond

        curr_sheet.update_cells(cell_list)
        curr_sheet.update_cell(2, 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def generate_criteria(self):
        return 0

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

        self.SS.worksheet(self.cfg_sheet.pop)

    def generate_criteria(self):
        sheet = self.SS.worksheet(self.cfg_sheet.pop)
        tester = ICS_Email_Tester(self.cfg_email)
        criteria = ICS_Monitor.SheetCriteria(
            great=tester.DELAY_GREAT,
            good=tester.DELAY_GOOD,
            average=tester.DELAY_AVERAGE
        )
        self.update_criteria(sheet, criteria)

    def run(self):
        ICS_Shared_Config.log("POP Test : Start ")

        try:
            EmailTester = ICS_Email_Tester(self.cfg_email)

            curr_sheet = self.SS.worksheet(self.cfg_sheet.pop)

            result = EmailTester.pop_tester()
            result = EmailTester.add_condition(result)

            self.update_result(curr_sheet, result)
            ICS_Shared_Config.log("POP Test : Complete ")
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))

            ICS_Shared_Config.log("POP Test : Failed ")

class ICS_Monitor_POP_SSL(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 2
        self.name = "Monitor_POP_SSL"
        self.counter = 2

    def generate_criteria(self):
        sheet = self.SS.worksheet(self.cfg_sheet.pop_ssl)
        tester = ICS_Email_Tester(self.cfg_email)
        criteria = ICS_Monitor.SheetCriteria(
            great=tester.DELAY_GREAT,
            good=tester.DELAY_GOOD,
            average=tester.DELAY_AVERAGE
        )
        self.update_criteria(sheet, criteria)

    def run(self):

        ICS_Shared_Config.log("POP_SSL Test : Start ")

        try:
            EmailTester = ICS_Email_Tester(self.cfg_email)

            curr_sheet = self.SS.worksheet(self.cfg_sheet.pop_ssl)
            result = EmailTester.pop_ssl_tester()
            result = EmailTester.add_condition(result)

            self.update_result(curr_sheet, result)
            ICS_Shared_Config.log("POP_SSL Test : Complete ")
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))

            ICS_Shared_Config.log("POP_SSL Test : Failed ")


class ICS_Monitor_Remote(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 3
        self.name = "Monitor_RDP"
        self.counter = 3

    def generate_criteria(self):
        sheet = self.SS.worksheet(self.cfg_sheet.rdp)
        tester = ICS_RDP_Tester(self.cfg_rdp)
        criteria = ICS_Monitor.SheetCriteria(
            great=tester.DELAY_GREAT,
            good = tester.DELAY_GOOD,
            average= tester.DELAY_AVERAGE
        )
        self.update_criteria(sheet,criteria)

    def run(self):
        ICS_Shared_Config.log("RDP Test : Start ")

        try:
            tester = ICS_RDP_Tester(self.cfg_rdp)
            curr_sheet = self.SS.worksheet(self.cfg_sheet.rdp)
            result = tester.run_test()

            self.update_result(curr_sheet, result)

            ICS_Shared_Config.log("RDP Test : Complete ")
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))

            ICS_Shared_Config.log("RDP Test : Failed ")



class ICS_Monitor_Http(ICS_Monitor,Thread):

    def __init__(self,config,counter : int ):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 10 + counter
        self.name = "Monitor_Http_Test"
        self.counter = 10 + counter

        self.tester = ICS_Http_Tester(config)
        self.SS = ICS_Shared_Config.Gspread_Open(self.cfg_sheet.id)

    def generate_criteria(self):
        sheet = self.SS.worksheet(self.tester.config.sheet)
        criteria = ICS_Monitor.SheetCriteria(
            great=self.tester.DELAY_GREAT,
            good = self.tester.DELAY_GOOD,
            average= self.tester.DELAY_AVERAGE
        )
        self.update_criteria(sheet,criteria)

    def run(self):

        try:
            ICS_Shared_Config.log("HTTP Test : " + self.tester.config.sheet)
            curr_sheet = self.SS.worksheet(self.tester.config.sheet)
            result = self.tester.test()
            self.update_result(curr_sheet,result)
            ICS_Shared_Config.log("HTTP Test : Done for %s " % self.tester.config.url)
        except Exception as e:

            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))
            ICS_Shared_Config.log("HTTP Test: Failed ")

    @staticmethod
    def PreparingTester() -> "list[ICS_Monitor_Http]" :
        app_config = ICS_Shared_Config.Config()
        all_http_config = app_config.http_load()

        result = []

        for config in all_http_config:
            new_test = ICS_Monitor_Http(config,len(result))
            result.append(new_test)

        return result



class ICS_Monitor_Speedtest(ICS_Monitor,Thread):
    def __init__(self):
        Thread.__init__(self)
        ICS_Monitor.__init__(self)
        self.threadID = 4
        self.name = "Monitor_Speedtest"
        self.counter = 4

    def run(self):
        ICS_Shared_Config.log("Speedtest : ")

        try:
            ICS_Speed_Test.Host = self.cfg_speedtest

            curr_sheet = self.SS.worksheet(self.cfg_sheet.speedtest)
            result = ICS_Speed_Test.do_speed_test()

            self.update_result(curr_sheet, result)

            ICS_Shared_Config.log("Speedtest : Done for %s " %  ICS_Speed_Test.Host )
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))

            ICS_Shared_Config.log("Speedtest: Failed ")

