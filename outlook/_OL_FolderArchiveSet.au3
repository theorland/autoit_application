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
; Disable Auto-Archiving for a single folder
; *****************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail")
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Error accessing folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_OL_FolderArchiveSet($aFolder[1], False, False)
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Error setting Auto-Archiving for folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Auto-Archiving for folder '*\Outlook-UDF-Test\SourceFolder\Mail' successfully disabled.")

; *****************************************************************************
; Example 2
; Set Auto-Archiving for a folder and all subfolders
; *****************************************************************************
$aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder")
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Error accessing folder '*\Outlook-UDF-Test\SourceFolder'. @error = " & @error & ", @extended = " & @extended)
_OL_FolderArchiveSet($aFolder[1], True, True , True, Default, 0, 999, 1)
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Error setting Auto-Archiving for folder '*\Outlook-UDF-Test\SourceFolder'. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_FolderArchiveSet Example Script", "Auto-Archiving for folder '*\Outlook-UDF-Test\SourceFolder' and all subfolders successfully set.")

_OL_Close($oOutlook)