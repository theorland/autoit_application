#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $iResult = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error)

; *****************************************************************************
; Example 1
; Check an existing folder
; *****************************************************************************
If _OL_FolderExists($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Contacts") = 0 Then
	If @extended = 4 Then
		MsgBox(16, "OutlookEX UDF: _OL_FolderExist Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Contacts' does not exist.")
	Else
		MsgBox(16, "OutlookEX UDF: _OL_FolderExist Example Script", "Error checking folder 'Outlook-UDF-Test\TargetFolder\Contacts'. @error = " & @error)
	EndIf
Else
	MsgBox(64, "OutlookEX UDF: _OL_FolderExist Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Contacts' exists.")
EndIf

; *****************************************************************************
; Example 2
; Check an non existing folder
; *****************************************************************************
If _OL_FolderExists($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Contacts-XY") = 0 Then
	If @extended = 4 Then
		MsgBox(16, "OutlookEX UDF: _OL_FolderExist Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Contacts-XY' does not exist.")
	Else
		MsgBox(16, "OutlookEX UDF: _OL_FolderExist Example Script", "Error checking folder 'Outlook-UDF-Test\TargetFolder\Contacts-XY'. @error = " & @error)
	EndIf
Else
	MsgBox(64, "OutlookEX UDF: _OL_FolderExist Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Contacts-XY' exists.")
EndIf

_OL_Close($oOutlook)