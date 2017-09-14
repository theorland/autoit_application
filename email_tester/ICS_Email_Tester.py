import poplib
import collections
import datetime

class ICS_Email_Tester:
        Config_Type = collections.namedtuple('Email_Config', \
        ['host','username', 'password','port','port_ssl'])
        Result_Type  = collections.namedtuple('Email_Test_Result', \
        ['start','end','delay', 'num','cond','stat'])
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
                Result = self.Result_Type( \
                        datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
                stat = -1;
                try:
                        M = poplib.POP3(Config.host, Config.port)
                        M.user(Config.username)
                        M.pass_(Config.password)
                        print(str(datetime.datetime.now())+": POP " + Config.host + ":" + str(Config.port) \
                                + " --> " + M.getwelcome())
                        stat = M.stat()[0]
                        M.quit()
                except :
                        print(str(datetime.datetime.now())+": POP Error" + str(Config.host) + ":" + str(Config.port) )
                        stat= -1

                tmp_end = datetime.datetime.now()
                tmp_delay = (tmp_end - Result.start).total_seconds() * 1000

                Result = Result._replace(end=tmp_end, delay=tmp_delay, stat= stat)

                return Result

        def pop_ssl_tester(self,Config):
                Result_SSL = self.Result_Type(\
                        datetime.datetime.now(), datetime.datetime.now(), 0, 0, 0, 0)
                stat = -1;
                try:
                        M = poplib.POP3_SSL(Config.host, Config.port_ssl)
                        M.user(Config.username)
                        M.pass_(Config.password)

                        print(str(datetime.datetime.now())+": POP SSL " + Config.host + ":" + str(Config.port) \
                              + " --> " + M.getwelcome())
                        stat = M.stat()[0]
                        M.quit()
                except:
                        print(str(datetime.datetime.now())+": POP SSL Error" + Config.host + ":" + str(Config.port))
                        stat = -1

                tmp_end = datetime.datetime.now()
                tmp_delay = (tmp_end - Result_SSL.start).total_seconds() * 1000

                Result_SSL = Result_SSL._replace(end=tmp_end, delay=tmp_delay, stat= stat)
                return Result_SSL
