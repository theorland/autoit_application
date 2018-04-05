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
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar")
Global $aItems = _OL_ItemFind($oOutlook, $aFolder[1], $olAppointment, '[Subject]="TestAppointment" AND [IsRecurring]=True', "", "", "EntryID,Subject,Start,End", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemRecurrenceExceptionGet Example Script - Found recurring appointments")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceExceptionGet Example Script", "Error finding a recurring appointment. @error = " & @error & ", @extended: " & @extended)
EndIf
; Get exceptions
Global $aExceptions = _OL_ItemRecurrenceExceptionGet($oOutlook, $aItems[1][0], Default)
If @error <> 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_ItemRecurrenceExceptionGet Example Script", "Error getting exceptions. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aExceptions, "OutlookEX UDF: _OL_ItemRecurrenceExceptionGet Example Script - Found exceptions for first recurring item")
; Display folder
$aFolder[1].Display

MsgBox(64, "OutlookEX UDF: _OL_ItemRecurrenceExceptionGet Example Script", "Subject of first exception: '" & $aExceptions[1][0].Subject & "'")

_OL_Close($oOutlook)