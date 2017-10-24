import pyping3 as pyping

from collections import namedtuple
from ICS_Shared_Config import ICS_Shared_Config

MAX_PING_LOST = 2
DELAY_GREAT = 200
DELAY_GOOD = 500

class ICS_RDP_Tester:
    Result_Type = namedtuple('Email_Test_Result', \
                    ['start', 'end', 'delay', 'num', 'cond', 'stat'])

    def __init__(self,config):
        self.config = config

    def run_test(self):
        result = pyping.ping(hostname = self.config.host, \
                             timeout = int(self.config.timeout),  \
                             count = int(self.config.count), \
                             packet_size = int(self.config.packet_size))
        ICS_Shared_Config.log("Pinging " + self.config.host + " Complete")

        conclusion = self.conclusion(result)
        return conclusion
    def conclusion(self,result):

        conclusion = "GREAT"
        result.max_rtt = float(result.max_rtt)
        result_output = self.Result_Type( \
            start =0, \
            end = 0, \
            delay = result.max_rtt, \
            num = self.config.count, \
            cond = conclusion, \
            stat = -1 * result.packet_lost)
        ICS_Shared_Config.log(result_output)
        if result.packet_lost>MAX_PING_LOST:
            conclusion= "BAD"
        elif result.max_rtt>=DELAY_GOOD:
            conclusion = "AVERAGE"
        elif result.max_rtt>=DELAY_GREAT:
            conclusion = "GOOD"
        result_output = result_output._replace( cond =conclusion)

        return result_output

