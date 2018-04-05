#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Deliver all undelivered messages submitted in the current session and
; receive mail for all accounts in the current profile.
; Show progress dialog.
; *****************************************************************************
_OL_ItemSendReceive($oOutlook, True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemSendReceive Example Script", "Error sending/receiving mail. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_ItemSendReceive Example Script", "Mail successfully sent/received!")

_OL_Close($oOutlook)