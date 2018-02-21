#cs ----------------------------------------------------
   Created by : theo (lourenzoisthebest@gmail.com)
   Created Time : 2017-07-04
   Fixing printer function, disable and enable network for make sure network share fixed
--------------------------------------------------------

Modified by : theo (lourenzoisthebest@gmail.com)
Modified Time : 2018-02-20
Add Obfuscator
Change ini file to exe file
Add delay to ini config
Fix Splash parameter

Modified by : theo (lourenzoisthebest@gmail.com)
Modified Time : 2017-07-05
Add More Logging Information, and delay as variable


#ce ----------------------------------------------------

;#RequireAdmin

#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/SO

#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <Array.au3>
#include <File.au3>

Global Const $FILE_INI_PATH = get_ini_path()
Global $LIST_APP[0]
Global $LOG_TEXT = ""
Global $DELAY_FORCE = 3000

$LOG_TEXT = @CRLF & "Starting Closing Apps Process" & $LOG_TEXT
cust_SplashTextOn($LOG_TEXT)

load_ini_file()

close_all_app()

SplashOff()

Func close_all_app()
   Local $process_name = "", $status_kill = "";
   For $process_name In $LIST_APP


	  cust_SplashTextOn("Now Closing '"  & $process_name & "'"  & @CRLF & $LOG_TEXT);
	  $status_kill = ""
	  ProcessClose ( $process_name )
	  Sleep($DELAY_FORCE)
	  If ProcessExists($process_name) Then
		 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
		 $status_kill = " but KILLED "
	  EndIf
	  $LOG_TEXT = @CRLF &  'Closing "' & $process_name & '" Completed' &  $status_kill & $LOG_TEXT
   Next

   cust_SplashTextOn( "COMPLETE WAITING FOR YOU" & @CRLF  & @CRLF & $LOG_TEXT);

   Sleep($DELAY_FORCE)
EndFunc

Func cust_SplashTextOn($text)
   SplashTextOn("Closing All Hanging Application", _
	  $text, _
	  800 ,600,  -1, -1, $DLG_NOTITLE  OR  $DLG_TEXTVCENTER  OR $DLG_CENTERONTOP, "" , 20, $FW_HEAVY  )
EndFunc


Func load_ini_file()


   $LOG_TEXT = @CRLF & "Loading configuration file : "  & $FILE_INI_PATH & $LOG_TEXT
   cust_SplashTextOn($LOG_TEXT)

   Local $curr_i = 0, $curr_app, $curr_name
   Local $curr_name = "APP" & $curr_i
   Local $curr_app = IniRead($FILE_INI_PATH,"APP_LIST",$curr_name,"N/A")



   While $curr_app <> "N/A"
	  _ArrayAdd( $LIST_APP, $curr_app )
	  $curr_i = $curr_i + 1
	  $curr_name = "APP" & $curr_i
	  $curr_app = IniRead($FILE_INI_PATH,"APP_LIST",$curr_name,"N/A")
   WEnd
   _ArrayDelete($LIST_APP,0)

   $DELAY_FORCE = Int(IniRead($FILE_INI_PATH,"GLOBAL","DELAY",3000))

EndFunc


Func get_ini_path()
   Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
   Local $aPathSplit = _PathSplit(@ScriptFullPath, $sDrive, $sDir, $sFileName, $sExtension)
   return @ScriptDir & "\config\" & $sFileName & ".ini"
EndFunc
