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
; Find a mail and save the item without attachements to C:\temp\Outlook-UDF-Test\Dir1
; Rename the item if it already exists and return the full path of the saved item
; *****************************************************************************
Global $sSaveDir = "C:\temp\Outlook-UDF-Test\Dir1\"
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSave Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
Global $sPath = _OL_ItemSave($oOutlook, $aOL_Item[1][0], Default, $sSaveDir, $olHTML, 1 + 16 + 32)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSave Example Script", "Error saving mail item to " & $sSaveDir & ". @error = " & @error & ", @extended = " & @extended)
ShellExecute($sPath)
MsgBox(64, "OutlookEX UDF: _OL_ItemSave Example Script", "Example 1: Item successfully saved as " & $sPath & "!")

; *****************************************************************************
; Example 2
; Find a mail and save the item plus attachements to C:\temp\Outlook-UDF-Test\Dir2
; Rename the item & attachments if they already exist
; *****************************************************************************
$sSaveDir = "C:\temp\Outlook-UDF-Test\Dir2\"
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSave Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_OL_ItemSave($oOutlook, $aOL_Item[1][0], Default, $sSaveDir, $olHTML, 3 + 16)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSave Example Script", "Error saving mail item to " & $sSaveDir & ". @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_ItemSave Example Script", "Example 2: Item plus attachments successfully saved to " & $sSaveDir & "!")

_OL_Close($oOutlook)
