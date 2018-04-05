#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; #PROGRAM# =====================================================================================================================
; Name ..........: _OL_Warnings
; Description ...: Check if Outlook warning windows pops up, if so bypass them, by clicking "allow".
; Syntax.........: Run(_OL_Warnings.exe $iOL_ProcessID [$iOL_WinCheckTime=1000[ $iOL_ControlCheckTime=1000[ $sOL_Version=""[ $iOL_Language=1033]]]])
; Parameters ....: $iOL_ProcessID     - The ProcessID of the calling script we should monitor and exit when finished
;                  $iOL_WinCheckTime  - Optional: How long, in milliseconds, we will wait before we check for the warning windows (default = 1000)
;                  $iOL_CtrlCheckTime - Optional: How long, in milliseconds, we will wait before we check that the controls we click are enabled (default = 1000)
;                  $sOL_Version       - Optional: Outlook version number e.g. 14.0.0.4760 (default = "")
;                  $iOL_Language      - Optional: Installed Outlook language. Please see http://msdn.microsoft.com/en-us/library/aa170976.aspx (default = 1033 = English US)
; Return values .: Failure - Sets returns value:
;                  |1 - ProcessID is missing as first parameter
;                  |2 - Specified ProcessID does not exist at startup
; Author ........: Wooltown
; Modified ......: water
; Remarks .......: If Outlook has security settings enabled, then warning windows will pop up, requiring manual key pressing.
;                  This function makes the necessary key presses. It is called by _OL_Open if you set $fOL_WarningClick to True or anytime by your script.
;                  The exe runs until the calling script is terminated.
;+
;                  Compile this script into an exe. Default location is the directory where the calling script is located.
;                  If you run Outlook in a different language then please change window title ($sWindowTitle) and text ($sWindowText).
;                  The script uses $sWindowTitle="Microsoft Outlook" for >= Outlook 2007 and "Microsoft Office Outlook" for < Outlook 2007.
;+
;                  To test the exe in case of problems please run the following DOS bat file:
;                    _OL_Warnings.exe ProcessID [$iOL_WinCheckTime[ $iOL_ControlCheckTime]]
;                    echo %ErrorLevel%
;                  If the output is not 0 then one of the errors described above has happened (section: Return values).
;                  These return codes can't be checked by _OL_Open because _OL_Warnings is called using "Run" which can not return any data to the calling process.
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Opt("TrayIconHide", 1)          					; 0=show, 1=hide tray icon
Opt("WinSearchChildren", 1)     					; 0=no, 1=search children also
Global $sWindowTitle = "Microsoft Outlook"			; Window title for >= Outlook 2007. Will be set for other Outlook versions below
Global $sWindowText  = "A program is trying to"
Global $iOL_ProcessID, $iOL_WinCheckTime = 1000, $iOL_CtrlCheckTime = 1000, $sOL_Version = "", $iOL_Language = 1033
; Process command line parameters
If $CmdLine[0] = 0 Then Exit 1						; No ProcessID was specified
$iOL_ProcessID = $CmdLine[1]
If Not ProcessExists($iOL_ProcessID) Then Exit 2
If $CmdLine[0] > 1 Then $iOL_WinCheckTime  = $CmdLine[2]
If $CmdLine[0] > 2 Then $iOL_CtrlCheckTime = $CmdLine[3]
If $CmdLine[0] > 3 Then $sOL_Version = $CmdLine[4]
If $CmdLine[0] > 4 Then $iOL_Language = $CmdLine[5]
; Set window title for < Outlook 2007
If $sOL_Version <> "" Then
	Global $aVersion = StringSplit($sOL_Version, '.')
	If IsArray($aVersion) And $aVersion[1] < 12 Then $sWindowTitle = "Microsoft Office Outlook"
EndIf
While 1
	If WinExists($sWindowTitle, $sWindowText) Then
		Local $aOL_WinSize = WinGetPos($sWindowTitle, $sWindowText)
		ToolTip(@CRLF & "OL_Warning will automatically click these buttons" & @CRLF, $aOL_WinSize[0] + 50, $aOL_WinSize[1] + 70, "Don't touch")
		While 1
			WinActivate($sWindowTitle, $sWindowText)
			If ControlCommand($sWindowTitle, $sWindowText, "Button3", "IsEnabled") Then
				ControlFocus($sWindowTitle, $sWindowText, "[CLASS:Button; INSTANCE:3]")
				ControlClick($sWindowTitle, $sWindowText, "Button3")
			EndIf
			If ControlCommand($sWindowTitle, $sWindowText, "Button4", "IsEnabled") Then
				ControlFocus($sWindowTitle, $sWindowText, "[CLASS:Button; INSTANCE:4]")
				Send("{SPACE}")
				ToolTip("")
				ExitLoop
			EndIf
			Sleep($iOL_CtrlCheckTime)
			If Not WinExists($sWindowTitle, $sWindowText) Then ExitLoop
		WEnd
	EndIf
	Sleep($iOL_WinCheckTime)
	If Not ProcessExists($iOL_ProcessID) Then Exit
WEnd