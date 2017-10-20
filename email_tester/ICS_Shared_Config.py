from ICS_Config import ICS_Config
import logging
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from datetime import datetime
import time
import os
import sys


class ICS_Shared_Config:
    Var_Config = None
    @staticmethod
    def Config():
        if ICS_Shared_Config.Var_Config is None:
            ICS_Shared_Config.Var_Config = ICS_Config()
        return ICS_Shared_Config.Var_Config

    Client_Scope = ['https://spreadsheets.google.com/feeds', \
             'https://www.googleapis.com/auth/drive']
    Client_Secret_Json_Path = None
    Client_Secret = None
    Client_Timeout = 600
    @staticmethod
    def Gspread_Client(renew=False):
        if ICS_Shared_Config.Client_Secret is None or renew==True:
            success = False
            while success == False:
                try:
                    creds = ServiceAccountCredentials.from_json_keyfile_name( \
                        ICS_Shared_Config.Client_Secret_Json_Path, \
                        ICS_Shared_Config.Client_Scope)
                    ICS_Shared_Config.Client_Secret = gspread.authorize(creds)

                    ICS_Shared_Config.log("Refresh the token success")
                    success = True
                except:
                    ICS_Shared_Config.log("Refresh the token failed will redo in next 10 minutes" )
                    time.sleep(ICS_Shared_Config.Client_Timeout)
        return ICS_Shared_Config.Client_Secret

    @staticmethod
    def Initialization():
        log_local_name = datetime.now().strftime("%Y-%m-%d.log")
        logging.basicConfig(filename=log_local_name , level=logging.INFO)

        application_path = "."
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)

        ICS_Shared_Config.log("Testing Email and RDP: application path " + application_path)

        ICS_Shared_Config.Client_Secret_Json_Path= os.path.join(application_path ,'client_secret.json')
        ICS_Config.Current_Setting_Path = os.path.join(application_path ,'email_tester.ini')

        ICS_Shared_Config.log("Testing Email: setting file in " + ICS_Shared_Config.Client_Secret_Json_Path)
        ICS_Shared_Config.log("Testing Email: setting file in " + ICS_Config.Current_Setting_Path)

    @staticmethod
    def log(text_to_log):
        text_to_log = str(text_to_log)
        text_to_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " : " +  text_to_log
        print(text_to_log)
        logging.info(text_to_log)

