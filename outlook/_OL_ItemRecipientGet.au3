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
; Get a list of all recipients of a meeting
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, '[Subject]="TestAppointment"', "", "", "EntryID,Subject")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientGet Example Script", "Could not find a meting item in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
$Result = _OL_ItemRecipientGet($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientGet Example Script", "Error getting member list of distribution list in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($Result, "OutlookEX UDF: _OL_ItemRecipientGet Example Script", "", 0, "|", "Recipient object|Name|EntryID")

_OL_Close($oOutlook)