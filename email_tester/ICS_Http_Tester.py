from splinter import Browser
from splinter.driver.webdriver.firefox import WebDriver
from splinter.request_handler.status_code import StatusCode
from datetime import datetime
from ICS_Config import ICS_Http_Tester_Type
from collections import namedtuple
from ICS_Shared_Config import ICS_Shared_Config
import json, os

class ICS_Http_Tester:
    Result_Type = namedtuple('Email_Test_Result', \
        ['start', 'end', 'delay', 'num', 'cond', 'stat'])

    def __init__(self,config):
        self.config = config # type: ICS_Http_Tester_Type.Config_Type
        ''' INITIALLIZE DEFAULT VALUES
        '''

        self.DELAY_AVERAGE = 500
        self.DELAY_GOOD = 1000
        self.DELAY_GREAT  = 500
        self.parseConfig(self.config.cond)

    def add_condition(self,Result : Result_Type):
        global DELAY_GREAT, DELAY_GOOD, DELAY_AVERAGE
        status = "BAD"
        if Result.stat !="None":
            if Result.delay<=self.DELAY_GREAT:
                status = "GREAT"
            elif Result.delay<=self.DELAY_GOOD:
                status = "GOOD"
            elif Result.delay<=self.DELAY_AVERAGE:
                status = "AVERAGE"
            Result = Result._replace(cond=status)
        return Result

    def parseConfig(self,config_delay_str : str):

        config_delay = json.loads(config_delay_str)
        self.DELAY_AVERAGE = config_delay["average"]
        self.DELAY_GOOD = config_delay["good"]
        self.DELAY_GREAT = config_delay["great"]

    def test(self) -> "ICS_Http_Tester.Result_Type" :



        code_result = ""
        start_time = datetime.now()
        result = self.Result_Type(
            datetime.now(), datetime.now(), \
            0, 0, "BAD", "None")

        code_result = "None"

        try:
            with Browser(self.config.browser) as firefox:  # type: WebDriver
                start_time = datetime.now()
                firefox.visit(self.config.url)
                status_code = firefox.status_code  # type: StatusCode
                code_result = str(status_code.code) + " : " + status_code.reason
                firefox.quit()
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))
            ICS_Shared_Config.log("HTTP Error " + str(self.config.url) + ":" + str(self.config.browser))

        end_time = datetime.now()
        diff = (end_time - start_time).total_seconds()* 1000

        result = result._replace( \
            start=start_time, end=end_time, delay=diff, stat=code_result)

        result = self.add_condition(result)
        return result









