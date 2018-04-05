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
; Reply to a mail
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
Global $oResult = _OL_ItemReply($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Error replying to a mail in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Display created reply item
$oResult.Display

MsgBox(64, "OutlookEX UDF: _OL_ItemReply Example Script", "Reply to mail item successfully created!")

; *****************************************************************************
; Example 2
; Reply to all recipients of a mail
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contact'. @error = " & @error)
$oResult = _OL_ItemReply($oOutlook, $aOL_Item[1][0], Default, True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Error replying to a  mail in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Display created reply item
$oResult.Display

MsgBox(64, "OutlookEX UDF: _OL_ItemReply Example Script", "ReplyAll to mail item successfully created!")

; *****************************************************************************
; Example 3
; Reply to a meeting request and accept
; *****************************************************************************
#cs
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contact'. @error = " & @error)
$oResult = _OL_ItemReply($oOutlook, $aOL_Item[1][0], Default, True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemReply Example Script", "Error replying to a  mail in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Display created reply item
$oResult.Display

MsgBox(64, "OutlookEX UDF: _OL_ItemReply Example Script", "ReplyAll to mail item successfully created!")
#ce

_OL_Close($oOutlook)