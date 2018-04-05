#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_RuleExecute Example Script", "This script executes rule 'Outlook-UDF-Test - MoveToFolder' against folder '*\Outlook-UDF-Test\SourceFolder\Mail'." & @CRLF & @CRLF & _
		"Are you sure you want to run this example script?")
If $iReply <> 6 Then Exit

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Execute rule "Outlook-UDF-Test - MoveToFolder" against folder
; "*\Outlook-UDF-Test\SourceFolder\Mail"
; *****************************************************************************
; Access folder to execute rule against
Global $aFolder = _OL_FolderAccess($oOL, "*\Outlook-UDF-Test\SourceFolder\Mail")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleExecute Example Script", "Error accessing source folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended: " & @extended)
; Execute rule
_OL_RuleExecute($oOl, "*", "Outlook-UDF-Test - MoveToFolder", $aFolder[1])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_RuleExecute Example Script", "Error executing rule 'Outlook-UDF-Test - MoveToFolder'. @error = " & @error & ", @extended: " & @extended)

MsgBox(64, "OutlookEX UDF: _OL_RuleExecute Example Script", "Rule 'Outlook-UDF-Test - MoveToFolder' successfully executed!")

_OL_Close($oOL)