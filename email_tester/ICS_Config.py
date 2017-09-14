import ConfigParser
import ICS_Email_Tester
from collections import namedtuple

class ICS_Config:
    File_Path = ""
    Current_Setting_Path = ""
    Sheet_Config = namedtuple("Sheet_Config",["id","pop","pop_ssl","max_row"])
    def run_file(self):
        ini_config = ConfigParser.ConfigParser()
        ini_config.read(self.Current_Setting_Path)
        read_value = ini_config.get("run","file")
        return read_value
    def sleep(self):
        ini_config = ConfigParser.ConfigParser()
        ini_config.read(self.Current_Setting_Path)
        read_value = ini_config.get("run","sleep")
        return read_value
    def email_load(self):
        ini_config = ConfigParser.ConfigParser()

        ini_config.read(self.Current_Setting_Path)
        config_value = ICS_Email_Tester.ICS_Email_Tester.Config_Type( \
            host = ini_config.get("email","host","203.146.26.11"), \
            username= ini_config.get("email", "username", "who@is-indonesia.com"), \
            password = ini_config.get("email", "password", "test"), \
            port = ini_config.get("email","pop_port",110), \
            port_ssl = ini_config.get("email","pop_ssl_port",995) )

        return config_value
    def sheet_load(self):
        ini_config = ConfigParser.ConfigParser()
        ini_config.read(self.Current_Setting_Path)
        config_value = self.Sheet_Config(\
            id=ini_config.get("sheet","id"), \
            pop=ini_config.get("sheet", "pop"),\
            pop_ssl = ini_config.get("sheet","pop_ssl"),\
            max_row = ini_config.get("sheet","max_row"))
        return config_value

