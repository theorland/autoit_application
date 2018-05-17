import poplib, collections
import datetime
import json, os
from ICS_Shared_Config import ICS_Shared_Config
from ICS_Config import ICS_Email_Tester_Type



class ICS_Email_Tester:

    Result_Type = collections.namedtuple('Email_Test_Result', \
        ['start', 'end', 'delay', 'num', 'cond', 'stat'])

    def __init__(self,_config:ICS_Email_Tester_Type.Config_Type):
        self.config = _config
        self.parse_cond(_config.cond)

        self.DELAY_GREAT = 1000
        self.DELAY_GOOD = 2000
        self.DELAY_AVERAGE = 4000

    def parse_cond(self,_config : str):

        if (len(_config)<=0):
            return None
        _config = json.loads(_config)

        self.DELAY_GREAT = _config["great"]
        self.DELAY_GOOD = _config["good"]
        self.DELAY_AVERAGE = _config["average"]

    def add_condition(self,Result : Result_Type):
        status = "BAD"
        if Result.stat>0:
            if Result.delay<=self.DELAY_GREAT:
                    status = "GREAT"
            elif Result.delay<=self.DELAY_GOOD:
                    status = "GOOD"
            elif Result.delay<=self.DELAY_AVERAGE:
                    status = "AVERAGE"
        Result = Result._replace(cond=status)
        return Result

    def pop_tester(self):
        Config = self.config

        result = self.Result_Type( \
                datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
        stat = -1

        try:
            stat = self.pop_raw_run(Config)
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))
            ICS_Shared_Config.log("POP Error " + str(Config.host) + ":" + str(Config.port) )
            stat= -1

        tmp_end = datetime.datetime.now()
        tmp_delay = (tmp_end - result.start).total_seconds() * 1000

        Result = result._replace(end=tmp_end, delay=tmp_delay, stat= stat)

        return Result

    def pop_ssl_tester(self):

        Config = self.config

        result_ssl = self.Result_Type(\
                datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
        stat = -1

        try:

            stat  = self.pop_ssl_raw_run(Config)
        except Exception as e:
            ICS_Shared_Config.log(os.path.basename(__file__) + " : " + str(e))

            ICS_Shared_Config.log("POP SSL Error " + str(Config.host) + ":" + str(Config.port_ssl))
            stat = -1

        tmp_end = datetime.datetime.now()
        tmp_delay = (tmp_end - result_ssl.start).total_seconds() * 1000

        result_ssl = result_ssl._replace(end=tmp_end, delay=tmp_delay, stat= stat)
        return result_ssl

    def pop_ssl_raw_run(self,Config:ICS_Email_Tester_Type):

        M = poplib.POP3_SSL(Config.host, Config.port_ssl)
        M.user(Config.username)
        M.pass_(Config.password)

        ICS_Shared_Config.log("POP SSL Success " + str(Config.host) + ":" + str(Config.port_ssl) \
                              + " --> " + str(M.getwelcome()))
        stat = M.stat()[0]
        M.quit()

        return stat

    def pop_raw_run(self,Config:ICS_Email_Tester_Type):

        M = poplib.POP3(Config.host, Config.port)
        M.user(Config.username)
        M.pass_(Config.password)
        ICS_Shared_Config.log("POP Success " + str(Config.host) + ":" + str(Config.port) \
                              + " --> " + str(M.getwelcome()))
        stat = M.stat()[0]
        M.quit()
        return stat
