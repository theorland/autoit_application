import ConfigParser
from collections import namedtuple

class ICS_Config:

    File_Path = ""
    Current_Setting_Path = ""
    Sheet_Config = namedtuple("Sheet_Config",["id","pop","pop_ssl","rdp","max_row"])

    def __init__(self):
        self.values = {"run": None, \
                        "sleep": None, \
                        "remote": None, \
                        "email": None, \
                        "sheet": None}
        self.ini_file = ConfigParser.ConfigParser()
        self.ini_file.read(ICS_Config.Current_Setting_Path)

    def run_file(self):
        if  self.values["run"] is None :
            read_value = self.ini_file.get("run","file")

            self.values["run"] = read_value
        else:
            read_value = self.values["run"]
        return read_value

    def sleep(self):
        if self.values["sleep"] is None:
            read_value = self.ini_file.get("run","sleep")

            self.values["sleep"] = read_value
        else:
            read_value = self.values["sleep"]
        return read_value

    def remote_load(self):
        if self.values["remote"] is None:

            config_value = ICS_RDP_Tester_Type.Ping_Config_Type(\
                host = self.ini_file.get("remote","host","127.0.0.1"), \
                timeout =self.ini_file.get("remote","timeout","1000"), \
                count = self.ini_file.get("remote","count","5"),  \
                packet_size =  self.ini_file.get("remote","packet_size","55") )

            self.values["remote"] = config_value
        else:
            config_value = self.values["remote"]
        return config_value

    def email_load(self):
        if self.values["email"] is None:
            config_value = ICS_Email_Tester_Type.Config_Type( \
                host = self.ini_file.get("email","host","127.0.0.1"), \
                username= self.ini_file.get("email", "username", "who@is-indonesia.com"), \
                password = self.ini_file.get("email", "password", "test"), \
                port = self.ini_file.get("email","pop_port",110), \
                port_ssl = self.ini_file.get("email","pop_ssl_port",995) )

            self.values["email"]
        else:
            config_value = self.values["email"]
        return config_value

    def sheet_load(self):
        if self.values["sheet"] is None:
            config_value = self.Sheet_Config(\
                id=self.ini_file.get("sheet","id"), \
                pop=self.ini_file.get("sheet", "pop"),
                pop_ssl = self.ini_file.get("sheet","pop_ssl"),\
                rdp = self.ini_file.get("sheet","rdp"), \
                max_row = self.ini_file.get("sheet","max_row"))
            self.values["sheet"] = config_value
        else:
            config_value = self.values["sheet"]

        return config_value

class ICS_Email_Tester_Type:

    Config_Type = namedtuple('Email_Config', \
        ['host','username', 'password','port','port_ssl'])
    Result_Type  = namedtuple('Email_Test_Result', \
        ['start','end','delay', 'num','cond','stat'])

class ICS_RDP_Tester_Type:
    Ping_Config_Type = namedtuple("Ping_Config", ["host", "timeout", "count", "packet_size"])