#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get all rules and display the enabled actions for the first rule
; *****************************************************************************
Global $aRules = _OL_RuleGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleActionGet Example Script", "Error accessing rules. @error = " & @error & ", @extended: " & @extended)

Global $aActions = _OL_RuleActionGet($aRules[1][0])
_ArrayDisplay($aActions, "OutlookEX UDF: All actions for rule '" & $aRules[1][4] & "'", "", 0, "|", "OlRuleActionType|OlObjectClass|Enabled?|Depending on the OlRuleActionType| | | | | | | | | | | ")

_OL_Close($oOutlook)