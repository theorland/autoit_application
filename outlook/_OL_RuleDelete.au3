#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_RuleDelete Example Script", "This script deletes all rules where the name starts with 'Outlook-UDF-Test'." & @CRLF & @CRLF & _
		"Are you sure you want to delete this rules?")
If $iReply <> 6 Then Exit

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get a list of rules (active and inactive) where the name starts with "Outlook-UDF-Test"
; Display the array of rules and ask the user if they should be deleted.
; *****************************************************************************
Global $aResult = _OL_RuleGet($oOL, "*", False)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleDelete Example Script", "Error retrieving list of rules. @error = " & @error & ", @extended: " & @extended)
; Remove all rules where the name doesn't start with "Outlook-UDF-Test" from the array
For $iIndex = $aResult[0][0] To 1 Step -1
	If Stringleft($aResult[$iIndex][4], 16) <> "Outlook-UDF-Test" Then _ArrayDelete($aResult, $iIndex)
Next
$aResult[0][0] = UBound($aResult, 1) - 1
; Display the rules to be deleted and ask the user for OK
Global $sRules = "Should the following rules be deleted?" & @CRLF & @CRLF
For $iIndex = 1 To $aResult[0][0]
	$sRules = $sRules & $aResult[$iIndex][4] & @CRLF
Next
$iReply = MsgBox(308, "OutlookEX UDF: _OL_RuleDelete Example Script", $sRules)
If $iReply <> 6 Then Exit
; Delete the selected rules
For $iIndex = 1 To $aResult[0][0]
	_OL_RuleDelete($oOL, "*", $aResult[$iIndex][4])
	If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleDelete Example Script", "Error deleting rule '" & $aResult[$iIndex][4] & "'. @error = " & @error & ", @extended: " & @extended)
Next

MsgBox(64, "OutlookEX UDF: _OL_RuleDelete Example Script", "All specified rules successfully deleted!")

_OL_Close($oOL)