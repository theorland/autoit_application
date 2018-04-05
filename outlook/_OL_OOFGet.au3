#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Display info about Out of Office message settings.
; *****************************************************************************
Global $aOOF = _OL_OOFGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_OOFGet Example Script", "Error accessing the OOF message settings. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aOOF, "OutlookEX UDF: Out of Office message settings")

_OL_Close($oOutlook)