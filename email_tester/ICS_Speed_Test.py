import pyspeedtest
from collections import namedtuple
from ICS_Shared_Config import ICS_Shared_Config
from datetime import datetime
import json



Result_Type = namedtuple('Email_Test_Result', \
    ['start', 'end', 'delay', 'num', 'cond', 'stat'])
Result_Test = namedtuple('Speedtest_Result', \
    ['host','download','upload','ping'])
Host = "speedtest-xg.glbb.ne.jp:8080"

DELAY_GREAT  = 200
DELAY_GOOD  = 500
DELAY_AVERAGE = 1000


def calc_cond(speed_test_result = None ) -> "str":

    ping_speed = speed_test_result.ping
    cond = "BAD"
    if speed_test_result == None:
        return cond
    if ping_speed<=DELAY_GREAT:
        cond="GREAT"
    elif ping_speed<=DELAY_GOOD:
        cond="GOOD"
    elif ping_speed<=DELAY_AVERAGE:
        cond="AVERAGE"
    return cond

def do_speed_test():

    speedtest = pyspeedtest.SpeedTest(Host, 0, 2)

    ICS_Shared_Config.log("Speedtest Start on %s " % Host)

    result = Result_Type(start=datetime.now(), end=datetime.now(), \
                        delay = -1, num = -1, cond = "BAD", stat= -1)

    speed_test_result = Result_Test( \
        host=speedtest.host, \
        ping=speedtest.ping(), \
        download=speedtest.download(), \
        upload=speedtest.upload())

    ICS_Shared_Config.log('Ping: %d ms, DL: %s, UL: %s' % \
                          (speed_test_result.ping, \
                           pyspeedtest.pretty_speed(speed_test_result.download), \
                           pyspeedtest.pretty_speed(speed_test_result.upload)))

    ICS_Shared_Config.log("Speedtest Complete")
    result = result._replace(end=datetime.now(), \
                             delay=speed_test_result.ping,
                             stat="DL: %s, UL: %s' " % ( \
                                 pyspeedtest.pretty_speed(speed_test_result.download), \
                                 pyspeedtest.pretty_speed(speed_test_result.upload)),
                             cond=calc_cond(speed_test_result))

    return result

'''
try:
    speedtest = pyspeedtest.SpeedTest("speedtest-xg.glbb.ne.jp:8080", 0, 2)
    print('Server: %s ' % speedtest.host)
    print('Ping: %d ms' % speedtest.ping())
    print('Download speed: %s' % pyspeedtest.pretty_speed(speedtest.download()))
    print('Upload speed: %s' % pyspeedtest(speedtest.upload()))
except Exception as e:
    pyspeedtest.LOG.error(e)
sys.exit(1)
'''