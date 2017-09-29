import pyping
from collections import namedtuple
from ICS_Email_Tester import ICS_Email_Tester
from datetime import datetime
import json


class ICS_RDP_Tester:
    Ping_Config_Type = namedtuple("Ping_Config", ["host", "timeout", "count", "packet_size"])
    def __init__(self,config):
        self.config = config

    def run_test(self):
        result = pyping.ping(hostname = self.config.host, \
                             timeout = int(self.config.timeout),  \
                             count = int(self.config.count), \
                             packet_size = int(self.config.packet_size))
        print(str(datetime.now()) + ": Pinging " + self.config.host + " Complete")

        conclusion = self.conclusion(result)
        return conclusion
    def conclusion(self,result):

        conclusion = "GREAT"
        result.max_rtt = float(result.max_rtt)
        result_output = ICS_Email_Tester.Result_Type( \
            start =0, \
            end = 0, \
            delay = result.max_rtt, \
            num = self.config.count, \
            cond = conclusion, \
            stat = -1 * result.packet_lost)
        print result.__dict__
        if result.packet_lost>2:
            conclusion= "BAD"
        elif result.max_rtt>=500:
            conclusion = "AVERAGE"
        elif result.max_rtt>=200:
            conclusion = "GOOD"
        result_output = result_output._replace( cond =conclusion)

        return result_output

