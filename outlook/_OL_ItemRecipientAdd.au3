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
; Add an optional recipient (the current user) to a meeting
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Could not find an appointment item in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
Global $oItem = _OL_ItemRecipientAdd($oOutlook, $aOL_Item[1][0], Default, $olOptional, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error adding recipient to appointment in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)
; Display item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Recipient successfully added to the appointment!")

_OL_Close($oOutlook)