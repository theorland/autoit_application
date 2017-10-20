import poplib
import collections
import datetime
from ICS_Shared_Config import ICS_Shared_Config

class ICS_Email_Tester:
    Result_Type = collections.namedtuple('Email_Test_Result', \
        ['start', 'end', 'delay', 'num', 'cond', 'stat'])


    def add_condition(self,Result):
        status = "BAD"
        if Result.stat>0:
            if Result.delay<=1000:
                    status = "GREAT"
            elif Result.delay<=2000:
                    status = "GOOD"
            elif Result.delay<=4000:
                    status = "AVERAGE"
        Result = Result._replace(cond=status)
        return Result

    def pop_tester(self,Config):
        result = self.Result_Type( \
                datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
        stat = -1
        try:
                M = poplib.POP3(Config.host, Config.port)
                M.user(Config.username)
                M.pass_(Config.password)
                ICS_Shared_Config.log("POP Success " + Config.host + ":" + str(Config.port) \
                        + " --> " + M.getwelcome())
                stat = M.stat()[0]
                M.quit()
        except :
            ICS_Shared_Config.log("POP Error" + str(Config.host) + ":" + str(Config.port) )
            stat= -1

        tmp_end = datetime.datetime.now()
        tmp_delay = (tmp_end - result.start).total_seconds() * 1000

        Result = result._replace(end=tmp_end, delay=tmp_delay, stat= stat)

        return Result

    def pop_ssl_tester(self,Config):
        result_ssl = self.Result_Type(\
                datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
        stat = -1
        try:

            M = poplib.POP3_SSL(Config.host, Config.port_ssl)
            M.user(Config.username)
            M.pass_(Config.password)

            ICS_Shared_Config.log("POP SSL Success " + Config.host + ":" + str(Config.port_ssl) \
                  + " --> " + M.getwelcome())
            stat = M.stat()[0]
            M.quit()
        except:
            ICS_Shared_Config.log("POP SSL Error " + Config.host + ":" + str(Config.port_ssl))
            stat = -1

        tmp_end = datetime.datetime.now()
        tmp_delay = (tmp_end - result_ssl.start).total_seconds() * 1000

        result_ssl = result_ssl._replace(end=tmp_end, delay=tmp_delay, stat= stat)
        return result_ssl

