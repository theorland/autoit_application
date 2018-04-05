#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_MailSignatureSet Example Script", "This script sets signature 'Outlook-UDF-Test' as the default signature for new messages." & @CRLF & _
		"To remove this setting please call the Outlook signature/wallpaper wizard." & @CRLF & @CRLF & _
		"Are you sure you want to set the default signature?")
If $iReply <> 6 Then Exit

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Sets the signature "Outlook-UDF-Test" as the default signature for
; new messages.
; The default signature for replies is unchanged.
; *****************************************************************************
_OL_MailSignatureSet("Outlook-UDF-Test", Default)
If @error = 3 Then
	MsgBox(16, "OutlookEX UDF: _OL_MailSignatureSet Example Script", "Signature 'Outlook-UDF-Test' could not be set." & @CRLF &  "This signature does not exist." & @CRLF & "Please use '_OL_MailSignatureCreate' to create the signature.")
ElseIf @error <> 0 Then
	MsgBox(16, "OutlookEX UDF: _OL_MailSignatureSet Example Script", "Signature 'Outlook-UDF-Test' could not be set. @error = " & @error & ", @extended: " & @extended)
Else
	MsgBox(64, "OutlookEX UDF: _OL_MailSignatureSet Example Script", "Signature 'Outlook-UDF-Test' set as the default signature for new messages.")
EndIf

_OL_Close($oOutlook)