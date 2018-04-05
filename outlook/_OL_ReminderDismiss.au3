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
; Dismiss the Test-Reminder
; *****************************************************************************
Global $aReminders = _OL_ReminderGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ReminderDismiss Example Script", "Error accessing reminders. @error = " & @error & ", @extended: " & @extended)
; Find the Test-Reminder created by function _OL_TestEnvironmentCreate
Global $bFound = False
For $iIndex = 1 To $aReminders[0][0]
	If StringLeft($aReminders[$iIndex][0], 15) = "TestAppointment" Then
		$bFound = True
		ExitLoop
	EndIf
Next
; If found dismiss the reminder
If $bFound = True Then
	If MsgBox(36, "OutlookEX UDF: _OL_ReminderDismiss Example Script", "The reminder for appointment" & @CRLF & "  " & $aReminders[$iIndex][0] & @CRLF & _
		"  Original reminder date/time: " & $aReminders[$iIndex][6] & @CRLF & _
		"will be dismissed. OK?") = 6 Then
		_OL_ReminderDismiss($oOutlook, $iIndex)
		If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ReminderDismiss Example Script", "Error dismissing the reminder. @error = " & @error & ", @extended: " & @extended)
	EndIf
Else
	MsgBox(16, "OutlookEX UDF: _OL_ReminderDismiss Example Script", "Could not find reminder for appointment 'TestAppointment*'.")
EndIf

_OL_Close($oOutlook)