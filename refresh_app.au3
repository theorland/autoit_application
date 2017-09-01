#cs ----------------------------------------------------
   Created by : theo (lourenzoisthebest@gmail.com)
   Created Time : 2017-07-04
   Fixing printer function, disable and enable network for make sure network share fixed
--------------------------------------------------------

Modified by : theo (lourenzoisthebest@gmail.com)
Modified Time : 2017-07-05
Add More Logging Information, and delay as variable

#ce ----------------------------------------------------

;#RequireAdmin
#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <Array.au3>

Global Const $FILE_INI_PATH = @ScriptDir & "\config\refresh_app.ini"
Global $LIST_APP[0]
Global $LOG_TEXT = ""
Global $DELAY_FORCE = 3000

SplashTextOn("Fixing Progress","Starting Closing Apps Process", _
-1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )

load_ini_file()

close_all_app()

SplashOff()

Func close_all_app()
   Local $process_name = ""
   For $process_name In $LIST_APP

	  SplashTextOn("Fixing Progress", _
	  "Now Closing '"  & $process_name & "'" & $LOG_TEXT, _
		 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
	  ProcessClose ( $process_name )
	  Sleep($DELAY_FORCE)
	  If ProcessExists($process_name) Then
		 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
	  EndIf
	  $LOG_TEXT =@CRLF &  'Closing "' & $process_name & '" Completed' & $LOG_TEXT
   Next
EndFunc

Func load_ini_file()

   Local $curr_i = 1, $curr_app, $curr_name
   Local $curr_name = "APP" & $curr_i
   Local $curr_app = IniRead($FILE_INI_PATH,"APP_LIST",$curr_name,"N/A")
   $LOG_TEXT = @CRLF & "Loading configuration file" & $LOG_TEXT

   While $curr_app <> "N/A"
	  _ArrayAdd( $LIST_APP, $curr_app )
	  $curr_i = $curr_i + 1
	  $curr_name = "APP" & $curr_i
	  $curr_app = IniRead($FILE_INI_PATH,"APP_LIST",$curr_name,"N/A")
   WEnd
   _ArrayDelete($LIST_APP,0)

EndFunc