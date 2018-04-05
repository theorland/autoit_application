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
; Attach a file to a mail (link to the file)
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
Global $oItem = _OL_ItemAttachmentAdd($oOutlook, $aOL_Item[1][0], Default, @ScriptDir & "\The_Outlook.jpg," & $olByReference)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "Error adding attachment to mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Display item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "A link to the file was successfully added to the mail item!")

; *****************************************************************************
; Example 2
; Attach a file to a task (copy file)
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Tasks", $olTask, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "Could not find a task item in folder 'Outlook-UDF-Test\SourceFolder\Tasks'. @error = " & @error)
$oItem = _OL_ItemAttachmentAdd($oOutlook, $aOL_Item[1][0], Default, @ScriptDir & "\The_Outlook.jpg")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "Error adding attachment to task item in folder 'Outlook-UDF-Test\SourceFolder\Tasks'. @error = " & @error & ", @extended = " & @extended)
; Display item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemAttachmentAdd Example Script", "A copy of the file was successfully added to the task item!")

_OL_Close($oOutlook)