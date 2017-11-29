from splinter import Browser
from splinter.driver.webdriver.firefox import WebDriver
from splinter.request_handler.status_code import StatusCode
from datetime import datetime
from ICS_Shared_Config import ICS_Shared_Config

class : 
def __init__(self):
    config = ICS_Shared_Config.Config()


def test(self):
    start_time = datetime.now()
    code_result = ""
    with Browser("firefox") as firefox : # type: WebDriver
        firefox.visit("http://is-intl.com")
        status_code = firefox.status_code # type: StatusCode
        code_result = str(status_code.code) + " : "  + status_code.reason
        firefox.quit()
    end_time = datetime.now()

    diff = (end_time-start_time).total_seconds()





