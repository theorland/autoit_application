#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Send a html mail to the current user.
; Add an attachment and set importance to high.
; *****************************************************************************
Global $sCurrentUser = $oOutlook.GetNameSpace("MAPI").CurrentUser.Name
_OL_Wrapper_SendMail($oOutlook, $sCurrentUser, "", "", "TestSubject", "Body<br><b>fett</b> normal.", @ScriptDir & "\_OL_Wrapper_SendMail.au3", $olFormatHTML, $olImportanceHigh)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OutlookSendMail Wrapper Script", "Error sending mail. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OutlookSendMail Wrapper Script", "Mail successfully sent to user '" & $sCurrentUser & "'!")

_OL_Close($oOutlook)