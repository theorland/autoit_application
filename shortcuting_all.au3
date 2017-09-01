#cs ----------------------------------------------------------------------------
   Created by : theo (lourenzoisthebest@gmail.com)
   Last Modified : 2017-06-30
   Create shortcut for all exe file here
#ce ----------------------------------------------------------------------------
#RequireAdmin
#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

Global $FILE_INI_PATH = @ScriptDir &"\config\shortcuting_all.ini"
Global $LIST_APP[1][2]
;Global $OUTPUT_DIR = @ScriptDir  & "\Programs"

Global $OUTPUT_DIR = @StartMenuCommonDir  & "\Programs"
Global $LOG_TEXT = "";


; MAIN PROGRAM
load_ini_file()

create_shortcut_all()

MsgBox ( $MB_SYSTEMMODAL + $MB_ICONINFORMATION +  $MB_OK,  _
    "Shortcuting Progress", $LOG_TEXT)

SplashOff()


Func create_shortcut_all()
   Local $output_link, $exe_path, $stats , $i, $entry[2]

   For $i = 0 To UBound($LIST_APP)-1

	  $entry[0] = $LIST_APP[$i][0]
	  $entry[1] = $LIST_APP[$i][1]

	  $exe_path = @ScriptDir & "\" & $entry[0]
	  $output_link = $OUTPUT_DIR & "\" & $entry[1] & ".lnk"

	  SplashTextOn("Path2", $LIST_APP[$i][0] & " = " & $LIST_APP[$i][1])

	  If Not FileExists($exe_path) Then
		 $LOG_TEXT = $exe_path & " Not Exists" & @CRLF & $LOG_TEXT
		 ContinueLoop
	  EndIf

	  $stats = FileCreateShortcut(@scriptDir & "\" & $entry[0] _
		 ,$output_link)
#cs
	  If $stats = 1 Then
		 Local $aDetails = FileGetShortcut($output_link)
		 If Not @error Then
			   MsgBox($MB_SYSTEMMODAL, "", "Path: " & $aDetails[0] & @CRLF & _
					  "Working directory: " & $aDetails[1] & @CRLF & _
					  "Arguments: " & $aDetails[2] & @CRLF & _
					  "Description: " & $aDetails[3] & @CRLF & _
					  "Icon filename: " & $aDetails[4] & @CRLF & _
					  "Icon index: " & $aDetails[5] & @CRLF & _
					  "Shortcut state: " & $aDetails[6] & @CRLF)
		 EndIf

	  EndIf
#ce
	  If $stats = 1 Then
		 $LOG_TEXT = $output_link & " Created" & @CRLF & $LOG_TEXT
	  Else
		 $LOG_TEXT = $output_link & " Can't Created" & @CRLF & $LOG_TEXT
	  EndIf
   Next
EndFunc


Func load_ini_file()

   Local $curr_i = 1, $new_app, $curr_name
   Local $curr_name = "APP" & $curr_i, _
		 $curr_sc = "SC" & $curr_i

   $OUTPUT_DIR = IniRead($FILE_INI_PATH, _
		 "OUTPUT","DIR",  $OUTPUT_DIR)

   Local $new_app = IniRead($FILE_INI_PATH, _
			"APP_LIST",$curr_name,"N/A"), _
		 $new_sc = IniRead($FILE_INI_PATH, _
			"APP_LIST",$curr_sc,"N/A"), _
		 $new_entry
   $LOG_TEXT = @CRLF & "Loading configuration file" & $LOG_TEXT

   While $new_app <> "N/A"
	  ; ReDim $LIST_APP[UBound($LIST_APP)+1][2]

	  Dim $new_entry[1][2] = [[$new_app , $new_sc ]]
	  _ArrayConcatenate( $LIST_APP, $new_entry )
	  $curr_i = $curr_i + 1

	  $curr_name = "APP" & $curr_i
	  $curr_sc   = "SC" & $curr_i

	  $new_app = IniRead($FILE_INI_PATH, _
			"APP_LIST",$curr_name,"N/A")
	  $new_sc  = IniRead($FILE_INI_PATH, _
			"APP_LIST",$curr_sc,"N/A")

	  IF $new_sc = "N/A" Then
		 $new_sc =$new_app
	  EndIf

   WEnd
   _ArrayDelete($LIST_APP, 0 )
   _ArrayDisplay($LIST_APP, 'Ini File of "shortcuting_all.ini"')

EndFunc