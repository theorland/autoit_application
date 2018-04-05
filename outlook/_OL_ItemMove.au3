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
; Find a task and move it to another folder
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Tasks", $olTask, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemMove Example Script", "Could not find a task item in folder 'Outlook-UDF-Test\SourceFolder\Tasks'. @error = " & @error)
_OL_ItemMove($oOutlook, $aOL_Item[1][0], Default, "*\Outlook-UDF-Test\TargetFolder\Tasks")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemMove Example Script", "Error moving specified task. @error = " & @error)

; Show folder
Global $oFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Tasks")
$oFolder[1].Display

MsgBox(64, "OutlookEX UDF: _OL_ItemMove Example Script", "Task successfully moved to 'Outlook-UDF-Test\TargetFolder\Tasks'!")

_OL_Close($oOutlook)