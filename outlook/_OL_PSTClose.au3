#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Close the PST archive created by _OL_PSTCreate by displayname
; *****************************************************************************
Global $oPST = _OL_PSTClose($oOutlook, "Outlook-UDF-PST")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_PSTClose Example Script", "Error closing 'Outlook-UDF-PST' archive. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_PSTClose Example Script", "Archive 'Outlook-UDF-PST' successfully closed!")

_OL_Close($oOutlook)