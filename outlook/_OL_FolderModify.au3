#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Changes the description of a folder
; *****************************************************************************
; Display the properties of the folder
Global $aFolder = _OL_FolderGet($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Mail")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderModify Example Script", "Error accessing folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aFolder, "Description (Element 12) is empty!")
; Change the description of the folder
Global $oFolder = _OL_FolderModify($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Mail", "", Default, "TestDescription")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderModify Example Script", "Error modifying folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Display the properties of the folder
$aFolder = _OL_FolderGet($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Mail")
_ArrayDisplay($aFolder, "Description (Element 12) has changed!")
MsgBox(64, "OutlookEX UDF: _OL_FolderModify Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Mail' successfully modified.")

_OL_Close($oOutlook)