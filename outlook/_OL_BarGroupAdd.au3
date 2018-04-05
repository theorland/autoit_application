#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_BarGroupAdd Example Script", "This script adds group 'Outlook-UDF-Test' to the Outlookbar." & @CRLF & _
		"To delete the group please run '_OL_BargroupDelete'." & @CRLF & @CRLF & _
		"Are you sure you want to create the group?")
If $iReply <> 6 Then Exit

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Adds a group in the Outlookbar on position 2
; *****************************************************************************
_OL_BarGroupAdd($oOutlook, "Outlook-UDF-Test", 2)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarGroupAdd Example Script", "Error adding group to the OutlookBar. @error = " & @error & ", @extended: " & @extended)
; Activate the OutlookBar
$oOutlook.ActiveExplorer.ShowPane($olOutlookBar, True)
MsgBox(64, "OutlookEX UDF: _OL_BarGroupAdd Example Script", "Group 'Outlook-UDF-Test' successfully added to the OutlookBar!")

_OL_Close($oOutlook)