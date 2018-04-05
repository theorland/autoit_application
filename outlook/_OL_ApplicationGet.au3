#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get information about the Outlook application
; *****************************************************************************
Global $aResult = _OL_ApplicationGet($oOutlook)
ConsoleWrite($aresult[0] & @CRLF)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ApplicationGet Example Script", "Error getting info about Outlook. @error = " & @error & ", @extended = " & @extended)
Global $aDisplay[$aResult[0]+1][2] = [[$aResult[0], 2], ["", "default profile"],["", "Execution mode language"],["", "Help language"],["", "Install language"], _
	["", "User interface language"], ["", "application name"],["", "Product code (GUID)"],["", "Product version"]]
For $i = 1 To $aResult[0]
	$aDisplay[$i][0] = $aResult[$i]
Next
_ArrayDisplay($aDisplay, "OutlookEX UDF: _OL_AccountGet Example Script")

_OL_Close($oOutlook)