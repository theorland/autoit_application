#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all groups in the Outlookbar
; *****************************************************************************
Global $aResult = _OL_BarGroupGet($oOutlook)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_BarGroupGet Example Script", "", 0, "|", "Name|View type")

_OL_Close($oOutlook)