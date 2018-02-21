#cs
   Created by : theo (lourenzoisthebest@gmail.com)

   Last Modified : 2018-02-20
   Fix Splash Function
   add Obfuscate
   Must Declare vars

   Last Modified : 2017-06-30
   Fixing printer function, disable and enable network for make sure network share fixed
#ce
#RequireAdmin
#include <AutoItConstants.au3>
#include <FontConstants.au3>

#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/SO

Opt('MustDeclareVars', 1)

Global Const $FILE_INI_PATH = @ScriptDir &"\config\fix_printer.ini"
Global $CONFIG_NETWORK
Global $LOG_TEXT = ""
Global $DELAY = 1000
Global $MY_COMPUTER = "ICS-WIN10"

$LOG_TEXT = "Loading Ini File" & @CRLF  & $LOG_TEXT
cust_SplashTextOn($LOG_TEXT);

load_ini_file()

SplashTextOn("Fixing Progress","Fixing printer in progress", _
-1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
restart_printer()
down_up_network($CONFIG_NETWORK)

$LOG_TEXT = "WAITING FOR YOU "& @CRLF & @CRLF  & $LOG_TEXT
cust_SplashTextOn($LOG_TEXT);
SplashOff()

Func cust_SplashTextOn($text)
   SplashTextOn("Fixing Printer and Network", _
	  $text, _
	  800 ,600,  -1, -1, $DLG_NOTITLE  OR  $DLG_TEXTVCENTER  OR $DLG_CENTERONTOP, "" , 20, $FW_HEAVY  )
EndFunc

Func clear_printer_queue()
   $LOG_TEXT = "Clear Printer Queue of " &  $MY_COMPUTER & @CRLF  & $LOG_TEXT
   cust_SplashTextOn($LOG_TEXT);
   Local $path = "winmgmts:\\" & $MY_COMPUTER & "\root\cimv2"
   Local $objWMIService = ObjGet($path )
   Local $colPrintJobs =  $objWMIService.ExecQuery ("Select * from Win32_PrintJob Where Size > 1")
   For $objPrintJob in $colPrintJobs
	   $objPrintJob.Delete_
   Next
EndFunc


Func down_up_network($name)
   $LOG_TEXT = "Restart Network " &  $name & @CRLF  & $LOG_TEXT
   cust_SplashTextOn($LOG_TEXT);
   ShellExecuteWait("netsh.exe", 'int set interface "' & $name & '" disable', "", "", @SW_HIDE)
   Sleep($DELAY)

   ShellExecuteWait("netsh.exe", 'int set interface "' & $name &  '" enable', "", "", @SW_HIDE)
EndFunc

Func restart_printer()
   $LOG_TEXT = "Stop Printer Service " & @CRLF  & $LOG_TEXT
   cust_SplashTextOn($LOG_TEXT);
   ShellExecuteWait("net.exe", 'stop spooler', "", "", @SW_HIDE)
   Sleep($DELAY)
   clear_printer_queue()
   Sleep($DELAY)
   $LOG_TEXT = "Start Printer Service " & @CRLF  & $LOG_TEXT
   cust_SplashTextOn($LOG_TEXT);
   ShellExecuteWait("net.exe", 'start spooler', "", "", @SW_HIDE)
EndFunc

Func load_ini_file()

 $LOG_TEXT = "Load ini file from " & $FILE_INI_PATH & @CRLF  & $LOG_TEXT
 cust_SplashTextOn($LOG_TEXT);

 $CONFIG_NETWORK = IniRead($FILE_INI_PATH,"IF","NET","")

 $MY_COMPUTER = IniRead($FILE_INI_PATH,"GLOBAL","MY","")
 $DELAY = IniRead($FILE_INI_PATH,"GLOBAL","DELAY","")
EndFunc
