import configparser
from collections import namedtuple

class ICS_Config:

    File_Path = ""
    Current_Setting_Path = ""
    Sheet_Config = namedtuple("Sheet_Config",["id","pop","pop_ssl","rdp","speedtest","max_row"])

    def __init__(self):
        self.values = {"run": None, \
                        "sleep": None, \
                        "remote": None, \
                        "email": None, \
                        "sheet": None, \
                        "speedtest" : None, \
                       "http" : None, \
                       "timeout" : None}
        self.ini_file = configparser.ConfigParser()

        self.ini_file.read(ICS_Config.Current_Setting_Path)


    def run_file(self):
        if  self.values["run"] is None :
            read_value = self.ini_file.get("run","file")

            self.values["run"] = read_value
        else:
            read_value = self.values["run"]
        return read_value

    def sleep(self):
        read_value = None
        if self.values["sleep"] is None:
            read_value = self.ini_file.getint("run","sleep")

            self.values["sleep"] = read_value
        else:
            read_value = self.values["sleep"]
        return read_value

    def timeout(self):
        read_value = None
        if self.values["timeout"] is None:
            read_value = self.ini_file.getint("run", "timeout")

            self.values["timeout"] = read_value
        else:

            read_value = self.values["timeout"]

        return read_value

    '''
        SPEEDTEST
    '''

    def speedtest_load(self):
        read_value = None
        if self.values["speedtest"] is None:
            read_value = self.ini_file.get("speedtest","url")
            self.values["speedtest"] = read_value
        else:
            read_value = self.values["speedtest"]
        return read_value


    def remote_load(self):

        if self.values["remote"] is None:

            config_value = ICS_RDP_Tester_Type.Ping_Config_Type(\
                host = self.ini_file.get("remote","host"), \
                timeout =self.ini_file.getint("remote","timeout"), \
                count = self.ini_file.get("remote","count"),  \
                packet_size =  self.ini_file.get("remote","packet_size"), \
                cond =self.ini_file.get("remote", "condition"), \
                )

            self.values["remote"] = config_value
        else:
            config_value = self.values["remote"]
        return config_value


    def email_load(self):
        if self.values["email"] is None:
            config_value = ICS_Email_Tester_Type.Config_Type( \
                host = self.ini_file.get("email","host"), \
                username= self.ini_file.get("email", "username"), \
                password = self.ini_file.get("email", "password"), \
                port = self.ini_file.get("email","pop_port"), \
                port_ssl = self.ini_file.get("email","pop_ssl_port"),\
                cond = self.ini_file.get("email","condition"))

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
                speedtest = self.ini_file.get("sheet","speedtest"), \
                max_row = self.ini_file.get("sheet","max_row"))

            self.values["sheet"] = config_value
        else:
            config_value = self.values["sheet"]

        return config_value
		
    def http_load(self):
        config_value = None
        if self.values["http"] is None:
            config_value = []

            counter = 0
            section_name = "http" + str(counter)

            while self.ini_file.has_section(section_name):
                new_config = ICS_Http_Tester_Type.Config_Type( \
                    url = self.ini_file.get(section_name, "url"), \
                    browser =  self.ini_file.get(section_name, "browser"), \
                    sheet = self.ini_file.get(section_name, "sheet"), \
                    cond= self.ini_file.get(section_name, "condition"), \
                )
                config_value.append(new_config)
                counter+=1
                section_name = "http" + str(counter)
        else :
            config_value = self.values["http"]


        return config_value


class ICS_Email_Tester_Type:
    Config_Type = namedtuple('Email_Config', \
        ['host','username', 'password','port','port_ssl','cond'])
    Result_Type  = namedtuple('Email_Test_Result', \
        ['start','end','delay', 'num','stat'])

class ICS_RDP_Tester_Type:
    Ping_Config_Type = namedtuple( \
        "Ping_Config", \
        ["host", "timeout", "count", "packet_size", "cond"])

class ICS_Http_Tester_Type:
    Config_Type = namedtuple(\
        "HTTP_config", \
        ["browser", "url", "cond","sheet"])