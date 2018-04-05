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
; Find a mail and save the first attachements to C:\temp\Outlook-UDF-Test\Dir2
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentSave Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_OL_ItemAttachmentSave($oOutlook, $aOL_Item[1][0], Default, 1, "C:\temp\Outlook-UDF-Test\Dir2\Attachment2.jpg")

If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentSave Example Script", "Error saving mail item to C:\temp\Outlook-UDF-Test\Dir2\. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_ItemAttachmentSave Example Script", "Attachment successfully saved!")

_OL_Close($oOutlook)