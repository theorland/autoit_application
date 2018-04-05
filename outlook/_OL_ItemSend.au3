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
; Send a mail item
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSend Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
$Result = _OL_ItemSend($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSend Example Script", "Error sending mail from folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)

MsgBox(64, "OutlookEX UDF: _OL_ItemSend Example Script", "Mail successfully sent!")

_OL_Close($oOutlook)