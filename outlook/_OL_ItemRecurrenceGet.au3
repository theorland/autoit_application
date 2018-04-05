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
; Search recurring appointments and show the recurrence information
; *****************************************************************************
Global $aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "[IsRecurring]=True", "", "", "EntryID,Subject", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemRecurrenceGet Example Script - Found recurring appointments")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceGet Example Script", "Error finding a recurring appointment. @error = " & @error & ", @extended: " & @extended)
EndIf

$aItems = _OL_ItemRecurrenceGet($oOutlook, $aItems[1][0], Default)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemRecurrenceGet Example Script - Properties for first recurring appointment")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceGet Example Script", "Error getting recurrence information. @error = " & @error & ", @extended: " & @extended)
EndIf

_OL_Close($oOutlook)