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
; Delete first recipient from an appointment
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientDelete Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
Global $oItem = _OL_ItemRecipientDelete($oOutlook, $aOL_Item[1][0], Default, 1)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientDelete Example Script", "Error deleting a recipient from appointment item in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemRecipientDelete Example Script", "Recipient successfully deleted from appointment item!")

_OL_Close($oOutlook)