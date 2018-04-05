#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error Then Exit MsgBox(16, "OutlookEX UDF - _OL_RecipientSelect Example Script", "Error running _OL_Open. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; *****************************************************************************
Global $avResult = _OL_ItemRecipientSelect($oOutlook, "Jon Doe", Default, Default, "suggested contacts", Default, Default, "Test")

If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientSelect Example Script", "Error frunning _OL_ItemRecipientSelect. @error = " & @error & ", @extended = " & @extended)
If IsBool($avResult) Then
	MsgBox(64, "OutlookEX UDF: _OL_ItemRecipientSelect Example Script", "The user cancelled the selection!")
Else
	For $i = 0 To UBound($avResult) - 1
		ConsoleWrite($avResult[$i].Name & @CRLF)
	Next
EndIf

_OL_Close($oOutlook)