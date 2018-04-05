#include-once
#include <OutlookEX.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Global $sTitle = "OutlookEX UDF - Manage Test Environment"
Global $iDontAsk = IniRead("_OL_TestEnvironment.ini", "Configuration", "DontAsk", "4") ; checked = 1, unchecked = 4
Global $iDontDelete = IniRead("_OL_TestEnvironment.ini", "Configuration", "DontDelete", "4") ; checked = 1, unchecked = 4
Global $oOutlook = _OL_Open()

Global $nMsg
#Region ### START Koda GUI section ### Form=h:\tools\autoit3\outlook\_ol_testenvironment.kxf
GUICreate($sTitle, 546, 300, 192, 114)
GUICtrlCreateGroup("Configuration", 8, 8, 529, 167)
Global $Checkbox1 = GUICtrlCreateCheckbox("Don't ask", 19, 24, 73, 25)
Global $Checkbox2 = GUICtrlCreateCheckbox("Don't delete", 19, 102, 80, 25)
Global $BtnSave = GUICtrlCreateButton("Save", 19, 142, 65, 25, $WS_GROUP)
GUICtrlCreateLabel("To ensure that every example script has the same test environment the first step of every " & _
		"example script is to delete and recreate the folder 'Outlook-UDF-Test', its subfolders and the test items." & @CRLF & _
		"If you mark the checkbox the example scripts don't ask and delete/recreate the test environment automatically." & @CRLF & @CRLF & _
		"If you want to make manual changes to the test environment you can suppress the deletion/recreation of the test environment " & _
		"by every example script." & @CRLF & @CRLF & _
		"Press the 'Save' button to save this setting.", 108, 29, 416, 135)
GUICtrlSetState($Checkbox1, $iDontAsk)
GUICtrlSetState($Checkbox2, $iDontDelete)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("Test Environment", 8, 198, 529, 99)
Global $BtnCreate = GUICtrlCreateButton("Create", 19, 264, 65, 25, $WS_GROUP)
Global $BtnDelete = GUICtrlCreateButton("Delete", 104, 264, 65, 25, $WS_GROUP)
GUICtrlCreateLabel("Press the 'Create' button to delete/recreate the test environment now." & @CRLF & "Press the 'Delete' button to delete the test environment now.", 19, 222, 428, 30)
Global $BtnExit = GUICtrlCreateButton("Exit", 461, 264, 65, 25, $WS_GROUP)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BtnExit
			Exit
		Case $BtnSave
			IniWrite("_OL_TestEnvironment.ini", "Configuration", "DontAsk", GUICtrlRead($Checkbox1))
			IniWrite("_OL_TestEnvironment.ini", "Configuration", "DontDelete", GUICtrlRead($Checkbox2))
		Case $BtnCreate
			$iDontAsk = GUICtrlRead($Checkbox1)
			$iDontDelete = GUICtrlRead($Checkbox2)
			_Create()
		Case $BtnDelete
			$iDontAsk = GUICtrlRead($Checkbox1)
			$iDontDelete = GUICtrlRead($Checkbox2)
			_Delete()
	EndSwitch
WEnd

; #FUNCTION# ====================================================================================================================
; Name...........: _Create
; Description ...: Delete and recreate the OutlookEX UDF test environment.
; Author ........: water
; ===============================================================================================================================
Func _Create()

	Local $iFolderExists = _OL_FolderExists($oOutlook, "*\Outlook-UDF-Test")
	If $iFolderExists = 1 Then
		If $iDontDelete = $GUI_UNCHECKED Then
			If $iDontAsk = $GUI_UNCHECKED Then
				If MsgBox(36, $sTitle, "Testenvironment already exists. Should it be deleted and recreated?") = 7 Then Return SetError(1, 0, 0)
			EndIf
			_OL_FolderDelete($oOutlook, "*\Outlook-UDF-Test")
			If @error <> 0 Then Return MsgBox($MB_ICONERROR, $sTitle, "Error Deleting 'Outlook-UDF-Test'" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		Else
			Return 1
		EndIf
	EndIf
	_OL_TestEnvironmentCreate($oOutlook, $iDontAsk, $iDontDelete)
	Select
		Case @error = 0
		Case @error >= 300 And @error <= 399
			Return MsgBox($MB_ICONERROR, $sTitle, "Error Creating Test Folder Structure 'Outlook-UDF-Test\SourceFolder'" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		Case @error >= 400 And @error <= 499
			Return MsgBox($MB_ICONERROR, $sTitle, "Error Creating Test Folder Structure 'Outlook-UDF-Test\TargetFolder'" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		Case @error >= 500 And @error <= 599
			Return MsgBox($MB_ICONERROR, $sTitle, "Error Creating Test Items in Source Folder" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		Case @error >= 600 And @error <= 699
			Return MsgBox($MB_ICONERROR, $sTitle, "Error Creating Test Items in Target Folder" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		Case Else
			Return MsgBox($MB_ICONERROR, $sTitle, "Error Creating Test Items in Target Folder" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
	EndSelect
	If $iFolderExists Then
		If $iDontDelete = $GUI_UNCHECKED Then
			MsgBox(64, $sTitle, "Test Folder Structure and Test items successfully deleted/created!")
		Else
			MsgBox(64, $sTitle, "Test Folder Structure and Test items have not been deleted/recreated due to the ""Don't delete"" flag!")
		EndIf
	Else
		MsgBox(64, $sTitle, "Test Folder Structure and Test items successfully created!")
	EndIf

EndFunc   ;==>_Create

; #FUNCTION# ====================================================================================================================
; Name...........: _Delete
; Description ...: Delete the OutlookEX UDF test environment.
; Author ........: water
; ===============================================================================================================================
Func _Delete()

	If Not _OL_FolderExists($oOutlook, "*\Outlook-UDF-Test") Then Return MsgBox(65, $sTitle, "Folder Structure does not exist!")
	If $iDontDelete = $GUI_UNCHECKED Then
		If $iDontAsk = $GUI_UNCHECKED Then
			If MsgBox(36, $sTitle, "Testenvironment already exists. Should it be deleted?") = 7 Then Return SetError(1, 0, 0)
		EndIf
		_OL_FolderDelete($oOutlook, "*\Outlook-UDF-Test")
		If @error <> 0 Then Return MsgBox($MB_ICONERROR, $sTitle, "Error Deleting 'Outlook-UDF-Test'" & @CRLF & "@error: " & @error & ", @extended: " & @extended)
		MsgBox(64, $sTitle, "Test Folder Structure and Test Items successfully deleted!")
	Else
		Return MsgBox(64, $sTitle, "Test Folder Structure and Test items have not been deleted due to the ""Don't delete"" flag!")
	EndIf

EndFunc   ;==>_Delete
