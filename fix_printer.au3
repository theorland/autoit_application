#cs ----------------------------------------------------
   Created by : theo (lourenzoisthebest@gmail.com)
   Last Modified : 2017-06-30
   Fixing printer function, disable and enable network for make sure network share fixed
#ce ----------------------------------------------------
#RequireAdmin
#include <AutoItConstants.au3>
#include <FontConstants.au3>
Global Const $FILE_INI_PATH = @ScriptDir &"\config\fix_printer.ini"
Global $CONFIG_NETWORK
load_ini_file()

SplashTextOn("Fixing Progress","Fixing printer in progress", _
-1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
restart_printer()
down_up_network($CONFIG_NETWORK)
SplashOff()

Func down_up_network($name)
   ShellExecuteWait("netsh.exe", 'int set interface "' & $name & '" disable', "", "", @SW_SHOW)
   Sleep(1000)
   ShellExecuteWait("netsh.exe", 'int set interface "' & $name &  '" enable', "", "", @SW_SHOW)
EndFunc

Func restart_printer()
   ShellExecuteWait("net.exe", 'stop spooler', "", "", @SW_SHOW)
   ShellExecuteWait("net.exe", 'start spooler', "", "", @SW_SHOW)
EndFunc

Func load_ini_file()
 $CONFIG_NETWORK = IniRead($FILE_INI_PATH,"IF","NET","")
EndFunc
