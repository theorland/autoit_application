#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error)

; *****************************************************************************
; Example 1
; Move a folder, subfolders and items to another folder
; *****************************************************************************
Global $oFolder = _OL_FolderMove($oOutlook, "*\Outlook-UDF-Test\SourceFolder", "*\Outlook-UDF-Test\TargetFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderMove Example Script", "Error moving folder 'Outlook-UDF-Test\SourceFolder' to folder 'Outlook-UDF-Test\TargetFolder'. @error = " & @error)
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder")
$aResult[1].Display
MsgBox(64, "OutlookEX UDF: _OL_FolderCreate Example Script", "Folder 'Outlook-UDF-Test\SourceFolder' successfully moved to folder 'Outlook-UDF-Test\TargetFolder'.")

_OL_Close($oOutlook)