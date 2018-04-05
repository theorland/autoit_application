#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get information about the currently accessed PST archives
; *****************************************************************************
Global $aPST = _OL_PSTGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_PSTGet Example Script", "Error accessing PST archives. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aPST, "OutlookEX UDF: All accessed PST archives", _
	"", 0, "|", "Displayname|Folder object|Path to the PST")

_OL_Close($oOutlook)