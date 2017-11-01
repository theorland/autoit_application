import pyping3 as pyping
from collections import namedtuple
from ICS_Shared_Config import ICS_Shared_Config
from ICS_Config import ICS_RDP_Tester_Type
import json

MAX_PING_LOST = 2
DELAY_GREAT = 200
DELAY_GOOD = 500
DELAY_AVERAGE = 1000

class ICS_RDP_Tester:
    Result_Type = namedtuple('Email_Test_Result', \
                    ['start', 'end', 'delay', 'num', 'cond', 'stat'])

    def __init__(self,config:ICS_RDP_Tester_Type.Ping_Config_Type):
        self.config = config

        self.parse_cond(config.cond)


    def run_test(self):
        result = pyping.ping(hostname = self.config.host, \
                             timeout = int(self.config.timeout),  \
                             count = int(self.config.count), \
                             packet_size = int(self.config.packet_size))
        ICS_Shared_Config.log("Pinging " + self.config.host + " Complete")

        conclusion = self.conclusion(result)
        return conclusion
    def conclusion(self,result : pyping.Response):
        global DELAY_GREAT, DELAY_GOOD, DELAY_AVERAGE, MAX_PING_LOST
        conclusion = "GREAT"
        result.max_rtt = float(result.max_rtt)

        result_output = self.Result_Type( \
            start =0, \
            end = 0, \
            delay = result.max_rtt, \
            num = self.config.count, \
            cond = conclusion, \
            stat = -1 * result.packet_lost)


        if result.packet_lost>MAX_PING_LOST:

            conclusion= "BAD"
        elif result.max_rtt>=DELAY_AVERAGE:

            conclusion = "BAD"
        elif result.max_rtt>=DELAY_GOOD:

            conclusion = "AVERAGE"
        elif result.max_rtt>=DELAY_GREAT:

            conclusion = "GOOD"

        result_output = result_output._replace( cond =conclusion)

        ICS_Shared_Config.log(result_output)

        return result_output

    def parse_cond(self,_config : str):
        global DELAY_GREAT, DELAY_GOOD, DELAY_AVERAGE, MAX_PING_LOST
        if (len(_config)<=0):
            return None
        _config = json.loads(_config)

        DELAY_GREAT = _config["great"]
        DELAY_GOOD = _config["good"]
        DELAY_AVERAGE = _config["average"]
        MAX_PING_LOST = _config["max_lost"]
