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
; Set folder "*\Outlook-UDF-Test\SourceFolder\Mail" as the current folder
; *****************************************************************************
_OL_FolderSet($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderSet Example Script", "Error setting a new current folderving mail item to C:\temp\. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_FolderSet Example Script", "Folder '*\Outlook-UDF-Test\SourceFolder\Mail' is set as the new current folder!")

_OL_Close($oOutlook)