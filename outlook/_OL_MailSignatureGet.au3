#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all Email signatures for the current mail account
; *****************************************************************************
Global $aSignatures = _OL_MailSignatureGet()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_MailSignatureGet Example Script", "Error accessing mail signatures. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aSignatures, "OutlookEX UDF: All email signatures for the current mail account", "", 0, "|", "Name|Used for new messages?|Used for reply messages?")

_OL_Close($oOutlook)