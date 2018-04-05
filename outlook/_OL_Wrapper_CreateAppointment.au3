#include <OutlookEX.au3>
#include <Date.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Create an appointment and set recurrence properties.
; *****************************************************************************
Global $sCurrentUser = $oOutlook.GetNameSpace("MAPI").CurrentUser.Name
Global $sStart = StringLeft(_Nowcalc(),16)
Global $sEnd   = StringLeft(_DateAdd("h", 3, _NowCalc()), 16)
_OL_Wrapper_CreateAppointment($oOutlook, "TestMeeting", $sStart, $sEnd, "My office", False, "Testbody", _
	15, $olBusy, $olImportanceHigh, $olPrivate, $olRecursWeekly, $sStart, _DateAdd("w", 3, $sEnd), 1)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OutlookCreateAppointment Wrapper Script", "Error creating appointment. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OutlookCreateAppointment Wrapper Script", "Appointment successfully created '" & $sCurrentUser & "'!")

_OL_Close($oOutlook)