#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Set Out Of Office message for internal mails.
; Does not always work! Unfortunately I don't know why!
; *****************************************************************************
Global $iOOF = _OL_OOFSet($oOutlook, "*", True, "This is a new Out-Of-Office message!")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_OOFSet Example Script", "Error setting the OOF message. @error = " & @error & ", @extended: " & @extended)

_OL_Close($oOutlook)