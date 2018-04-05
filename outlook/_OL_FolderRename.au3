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
; Renames a folder
; *****************************************************************************
Global $oFolder = _OL_FolderRename($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Mail", "Mail-Renamed")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderRename Example Script", "Error renaming folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error)
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Mail-Renamed")
$aResult[1].Display
MsgBox(64, "OutlookEX UDF: _OL_FolderRename Example Script", "Folder 'Outlook-UDF-Test\TargetFolder\Mail' successfully renamed.")

_OL_Close($oOutlook)