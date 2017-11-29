import ICS_Email_Tester, ICS_RDP_Tester, ICS_Config, ICS_Http_Tester
import unittest
import os, sys
import pyping3 as pyping

class ICS_Condition_Tester(unittest.TestCase):
    def test_cond_email(self):
        email = ICS_Email_Tester.ICS_Email_Tester(ICS_Config.ICS_Email_Tester_Type.Config_Type("","","","","",""))
        email.parse_cond('{"great":300,"good":700,"average":1200}')
        self.assertEqual(ICS_Email_Tester.DELAY_GREAT,300)
        self.assertEqual(ICS_Email_Tester.DELAY_GOOD, 700)
        self.assertEqual(ICS_Email_Tester.DELAY_AVERAGE, 1200)

    def test_cond_rdp(self):
        rdp = ICS_RDP_Tester.ICS_RDP_Tester(ICS_Config.ICS_RDP_Tester_Type.Ping_Config_Type("","","","",""))
        rdp.parse_cond('{"great":1000,"good":1500,"average":1800,"max_lost":5}')
        self.assertEqual(ICS_RDP_Tester.DELAY_GREAT,1000)
        self.assertEqual(ICS_RDP_Tester.DELAY_GOOD, 1500)
        self.assertEqual(ICS_RDP_Tester.DELAY_AVERAGE, 1800)
        self.assertEqual(ICS_RDP_Tester.MAX_PING_LOST, 5)

        response = pyping.Response()
        response.packet_lost = 1
        response.max_rtt = 500
        self.assertEqual(rdp.conclusion(response).cond,"GREAT")

        response.packet_lost = 1
        response.max_rtt = 1200
        self.assertEqual(rdp.conclusion(response).cond, "GOOD")

        response.packet_lost = 1
        response.max_rtt = 1600
        self.assertEqual(rdp.conclusion(response).cond, "AVERAGE")

        response.packet_lost = 6
        response.max_rtt = 1600
        self.assertEqual(rdp.conclusion(response).cond, "BAD")

        response.packet_lost = 5
        response.max_rtt = 2000
        self.assertEqual(rdp.conclusion(response).cond, "BAD")

    def test_timeout(self):
        application_path = "."
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        ICS_Config.ICS_Config.Current_Setting_Path = os.path.join(application_path,"config.ini")
        config = ICS_Config.ICS_Config()
        self.assertEqual(config.timeout(),334)



    def test_all_config(self):
        application_path = "."
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        ICS_Config.ICS_Config.Current_Setting_Path = os.path.join(application_path,"config.ini")
        config = ICS_Config.ICS_Config()
        config.email_load()
        config.remote_load()
        config.sheet_load()
        config.run_file()
        config.sleep()
        config.timeout()

    def test_http(self):
        ICS_Http_Tester

