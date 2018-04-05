#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get all rules
; *****************************************************************************
Global $aRules = _OL_RuleGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleGet Example Script", "Error accessing rules. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aRules, "OutlookEX UDF: All rules for the currently logged on users default store", "", 0, "|", "Object|Enabled?|Execution order|Client side rule?|Name|Send or Receive rule")

_OL_Close($oOutlook)