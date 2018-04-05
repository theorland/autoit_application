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
; Add recurrence: Daily with defined start and end date/time
; *****************************************************************************
Global $aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, '[Subject]="TestAppointment" AND [IsRecurring]=True', "", "", "EntryID,Subject,Start,End", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemRecurrenceExceptionSet Example Script - Found recurring appointments")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceExceptionSet Example Script", "Error finding a recurring appointment. @error = " & @error & ", @extended: " & @extended)
EndIf
; Define exception
Global $oItem = _OL_ItemRecurrenceExceptionSet($oOutlook, $aItems[1][0], Default, _NowCalcDate() & " 08:00:00", _NowCalcDate() & " 09:00:00", _NowCalcDate() & " 14:00:00", "Exception", "ExceptionBody")
If @error = 0 Then
	; Display item
	$oItem.Display
	MsgBox(64, "OutlookEX UDF: _OL_ItemRecurrenceExceptionSett Example Script", "Exception successfully defined.")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceExceptionSet Example Script", "Error setting exception. @error = " & @error & ", @extended: " & @extended)
EndIf

_OL_Close($oOutlook)