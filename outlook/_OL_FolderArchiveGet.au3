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
; Get the Auto-Archiving properties of a folder
; *****************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail")
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Error accessing folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
Global $aProperties = _OL_FolderArchiveGet($aFolder[1])
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveGet Example Script", "Error getting Auto-Archiving properties for folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aProperties)

_OL_Close($oOutlook)