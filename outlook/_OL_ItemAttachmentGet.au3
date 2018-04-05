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
; Get a list of all attachments of a mail item
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentGet Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
$Result = _OL_ItemAttachmentGet($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentGet Example Script", "Error getting list of attachments for mail in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($Result, "OutlookEX UDF: _OL_ItemAttachmentGet Example Script")

_OL_Close($oOutlook)