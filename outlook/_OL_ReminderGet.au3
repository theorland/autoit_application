#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get all visible reminders
; *****************************************************************************
Global $aReminders = _OL_ReminderGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ReminderGet Example Script", "Error accessing reminders. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aReminders, "OutlookEX UDF: All visible reminders for the current user", "", 0, "|", "Title|OlObjectClass|visible?|Reminder object|Item object|Next occurrence|Original date")

_OL_Close($oOutlook)