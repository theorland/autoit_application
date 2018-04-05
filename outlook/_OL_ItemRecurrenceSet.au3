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
; Add recurrence: Daily with defined start and end date/time
; *****************************************************************************
Global $sItem1 = _OL_ItemCreate($oOutlook, $olAppointmentItem, "*\Outlook-UDF-Test\SourceFolder\Calendar", "", "Subject=Recurrence Test 1", "Start=" & _NowCalcDate() & " 08:00:00", "End=" & _NowCalcDate() & " 11:30:00", _
		"Location=Building A, Room 10", "RequiredAttendees=" & $oOutlook.GetNameSpace("MAPI" ).CurrentUser.Name)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Error creating an appointment in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
; Add recurrence
Global $oItem = _OL_ItemRecurrenceSet($oOutlook, $sItem1, Default, _NowCalcDate(), "08:00:00", _DateAdd("D", 14, _NowCalcDate()), "11:30", $olRecursDaily, "", "", "")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Error adding recurrence information'. @error = " & @error)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Recurrence Test 1: Daily with defined start and end date/time - Success!")

; *****************************************************************************
; Example 2
; Add recurrence: Every 3rd month on the 2nd monday in 2011
; *****************************************************************************
Global $sItem2 = _OL_ItemCreate($oOutlook, $olAppointmentItem, "*\Outlook-UDF-Test\SourceFolder\Calendar", "", "Subject=Recurrence Test 2", "Start=2011/02/08 08:00:00", "End=2011/02/08 11:30:00", _
		"Location=Building A, Room 10", "RequiredAttendees=" & $oOutlook.GetNameSpace("MAPI" ).CurrentUser.Name)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Error creating an appointment in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
; Add recurrence
$oItem = _OL_ItemRecurrenceSet($oOutlook, $sItem2, Default, "2011/01/01", "08:00:00", "2011/12/31", "11:30", $olRecursMonthNTh, $olMonday, 3, 2)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Error adding recurrence information'. @error = " & @error)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemRecurrenceSet Example Script", "Recurrence Test 2: Every 3rd month on the 2nd monday in 2011 - Success!")

_OL_Close($oOutlook)