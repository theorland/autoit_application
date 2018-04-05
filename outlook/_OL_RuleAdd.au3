#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_RuleAdd Example Script", "This script creates three rules 'Outlook-UDF-Test - *' for incoming/outgoing messages." & @CRLF & _
		"To remove this rules please run _OL_RuleDelete.au3." & @CRLF & @CRLF & _
		"Are you sure you want to create this rules?")
If $iReply <> 6 Then Exit

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Create a rule to be used on incomming messages named "Outlook-UDF-Test - AssignToCategory"
; The new message will be assigned a category of "Outlook-UDF-Test" if the body contains text "AssignToCategory"
; Exception is that all messages with subject "test" or "Outlook-UDF-Test" will be ignored
; *****************************************************************************
Global $oResult
; Create the rule
$oResult = _OL_RuleAdd($oOL, "*", "Outlook-UDF-Test - AssignToCategory")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error creating rule 1. @error = " & @error & ", @extended: " & @extended)
; Add the rule action
$oResult = _OL_RuleActionSet($oOL, "*", "Outlook-UDF-Test - AssignToCategory", $olRuleActionAssignToCategory, True, "Outlook-UDF-Test")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule action for rule 1. @error = " & @error & ", @extended: " & @extended)
; Add the rule condition
$oResult = _OL_RuleConditionSet($oOL, "*", "Outlook-UDF-Test - AssignToCategory", $olConditionBody, True, False, "AssignToCategory")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule condition for rule 1. @error = " & @error & ", @extended: " & @extended)
; Add the exceptions to the rule condition
$oResult = _OL_RuleConditionSet($oOL, "*", "Outlook-UDF-Test - AssignToCategory", $olConditionSubject, True, True, "test|Outlook-UDF-Test")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule condition for rule 1. @error = " & @error & ", @extended: " & @extended)

MsgBox(64, "OutlookEX UDF: _OL_RuleAdd Example Script", "Rule 1 'Outlook-UDF-Test - AssignToCategory' + Action + Condition + Condition Exceptions successfully created!")

; *****************************************************************************
; Example 2
; Create a rule to be used on outgoing messages named "Outlook-UDF-Test - CcMessage"
; The new message will be sent to as CC to the current user if the body contains text "Archive"
; The rule is executed as number 2 in list of active rules
; *****************************************************************************
; Create the rule
$oResult = _OL_RuleAdd($oOL, "*", "Outlook-UDF-Test - CcMessage", True, $olRuleSend, 2)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error creating rule 2. @error = " & @error & ", @extended: " & @extended)
; Add the rule action
$oResult = _OL_RuleActionSet($oOL, "*", "Outlook-UDF-Test - CcMessage", $olRuleActionCcMessage, True, "Thomas Rupp")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule action for rule 2. @error = " & @error & ", @extended: " & @extended)
; Add the rule condition
$oResult = _OL_RuleConditionSet($oOL, "*", "Outlook-UDF-Test - CcMessage", $olConditionBody, True, False, "Archive")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule condition for rule 2. @error = " & @error & ", @extended: " & @extended)

MsgBox(64, "OutlookEX UDF: _OL_RuleAdd Example Script", "Rule 2 'Outlook-UDF-Test - CcMessage' + Action + Condition successfully created!")

; *****************************************************************************
; Example 3
; Create a rule to be used on incoming messages named "Outlook-UDF-Test - MoveToFolder"
; The new message will be moved to folder "\\*\Outlook-UDF-Test\TargetFolder\Mail"
; if the subject contains "TestMail"
; *****************************************************************************
Global $aFolder = _OL_FolderAccess($oOL, "*\Outlook-UDF-Test\TargetFolder\Mail")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error creating rule 3. Can't access target folder. @error = " & @error & ", @extended: " & @extended)
; Create the rule
$oResult = _OL_RuleAdd($oOL, "*", "Outlook-UDF-Test - MoveToFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error creating rule 3. @error = " & @error & ", @extended: " & @extended)
; Add the rule action
$oResult = _OL_RuleActionSet($oOL, "*", "Outlook-UDF-Test - MoveToFolder", $olRuleActionMoveToFolder, True, $aFolder[1])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule action for rule 3. @error = " & @error & ", @extended: " & @extended)
; Add the rule condition
$oResult = _OL_RuleConditionSet($oOL, "*", "Outlook-UDF-Test - MoveToFolder", $olConditionSubject, True, False, "TestMail")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleAdd Example Script", "Error adding rule condition for rule 3. @error = " & @error & ", @extended: " & @extended)

MsgBox(64, "OutlookEX UDF: _OL_RuleAdd Example Script", "Rule 3 'Outlook-UDF-Test - MoveToFolder' + Action + Condition successfully created!")
_OL_Close($oOL)