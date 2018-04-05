#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get all rules and display the active conditions for the first rule
; *****************************************************************************
Global $aRules = _OL_RuleGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleConditionGet Example Script", "Error accessing rules. @error = " & @error & ", @extended: " & @extended)

Global $aConditions = _OL_RuleConditionGet($aRules[1][0])
_ArrayDisplay($aConditions, "OutlookEX UDF: All conditions for rule '" & $aRules[1][4] & "'", "", 0, "|", "OlRuleConditionType|OlObjectClass|Enabled?|Depending on the OlRuleConditionType| | | | | | | | | | | ")

; *****************************************************************************
; Example 2
; Display the active exceptions to the conditions for the first rule
; *****************************************************************************
$aConditions = _OL_RuleConditionGet($aRules[1][0], True, True)
_ArrayDisplay($aConditions, "OutlookEX UDF: All exceptions for rule '" & $aRules[1][4] & "'")

_OL_Close($oOutlook)