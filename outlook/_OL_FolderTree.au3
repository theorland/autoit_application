#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
;Global $iResult = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Display a tree of all folders in your mailbox
; *****************************************************************************
Global $aResult = _OL_FolderTree($oOutlook, "*")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderTree Example Script", "Error accessing root folder. @error = " & @error)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_FolderTree Example Script - All folders")

; *****************************************************************************
; Example 2
; Display a tree of folders in the test environment
; *****************************************************************************
$aResult = _OL_FolderTree($oOutlook, "*\Outlook-UDF-Test")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderTree Example Script", "Error accessing folder 'Outlook-UDF-Test'. @error = " & @error)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_FolderTree Example Script - Tree starting with '*\Outlook-UDF-Test'")

_OL_Close($oOutlook)