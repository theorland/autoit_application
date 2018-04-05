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
; Copy a folder, subfolders and items to another folder
; *****************************************************************************
Global $oFolder = _OL_FolderCopy($oOutlook, "*\Outlook-UDF-Test\SourceFolder", "*\Outlook-UDF-Test\TargetFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderCopy Example Script", "Error copying folder 'Outlook-UDF-Test\SourceFolder' to folder 'Outlook-UDF-Test\TargetFolder'. @error = " & @error)
; Show folder
$oFolder.Display
MsgBox(64, "OutlookEX UDF: _OL_FolderCreate Example Script", "Folder 'Outlook-UDF-Test\SourceFolder' successfully copied to folder 'Outlook-UDF-Test\TargetFolder'.")

_OL_Close($oOutlook)