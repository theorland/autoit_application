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
; Find folders by name
; *****************************************************************************
Global $aFolder = _OL_FolderFind($oOutlook, "*\Outlook-UDF-Test", 99, "Notes")
If @error = 4 Then
	MsgBox(16, "OutlookEX UDF: _OL_FolderFind Example Script", "Error searching for folder. @error = " & @error & ", @extended = " & @extended)
Else
	_ArrayDisplay($aFolder, "Folders searched by name")
EndIf

; *****************************************************************************
; Example 2
; Find folders by default item type
; *****************************************************************************
$aFolder = _OL_FolderFind($oOutlook, "*\Outlook-UDF-Test", 99, "", 1, $olMailItem)
If @error = 4 Then
	MsgBox(16, "OutlookEX UDF: _OL_FolderFind Example Script", "Error searching for folder. @error = " & @error & ", @extended = " & @extended)
Else
	_ArrayDisplay($aFolder, "Folders searched by default item type")
EndIf

_OL_Close($oOutlook)