#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Delete the signature "Outlook-UDF-Test".
; *****************************************************************************
Global $iResult = _OL_MaiLSignatureDelete("Outlook-UDF-Test")
If @error = 3 Then
	MsgBox(16, "OutlookEX UDF: _OL_MailSignatureDelete Example Script", "Signature 'Outlook-UDF-Test' could not be deleted." & @CRLF &  "This signature does not exist." & @CRLF & "Please use '_OL_MailSignatureCreate' to create the signature.")
ElseIf @error <> 0 Then
	MsgBox(16, "OutlookEX UDF: _OL_MailSignatureDelete Example Script", "Signature 'Outlook-UDF-Test' could not be deleted. @error = " & @error & ", @extended: " & @extended)
Else
	MsgBox(64, "OutlookEX UDF: _OL_MailSignatureDelete Example Script", "Signature 'Outlook-UDF-Test' successfully deleted.")
EndIf
_OL_Close($oOutlook)