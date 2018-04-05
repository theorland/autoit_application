#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Delete the "Outlook-UDF-Test" group from the Outlookbar
; (created by example script _OL_BarShortcutAdd)
; *****************************************************************************
; Get list of groups in the OutlookBar
Global $aResult = _OL_BarGroupGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarGroupDelete Example Script", "Error getting list of groups in OutlookBar. @error = " & @error & ", @extended = " & @extended)
; Search for group "Outlook-UDF-Test"
Global $iFound = 0
For $iIndex = 1 To $aResult[0][0]
	If $aResult[$iIndex][0] = "Outlook-UDF-Test" Then
		$iFound = $iIndex
		ExitLoop
	EndIf
Next
; Group not found - exit
If $iFound = 0 Then Exit MsgBox(64, "OutlookEX UDF: _OL_BarGroupDelete Example Script", "Group 'Outlook-UDF-Test' could not be found. Please use example '_OL_BarGroupAdd' to create.")
; Delete group "Outlook-UDF-Test". The group is accessed by its index value
_OL_BarGroupDelete($oOutlook, $iFound)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarGroupDelete Example Script", "Error deleting group 'Outlook-UDF-Test'. @error = " & @error & ", @extended = " & @extended)
; Activate the OutlookBar
$oOutlook.ActiveExplorer.ShowPane($olOutlookBar, True)
MsgBox(64, "OutlookEX UDF: _OL_BarGroupDelete Example Script", "Group 'Outlook-UDF-Test' successfully deleted!")

_OL_Close($oOutlook)