#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Add a shortcut to the "Outlook-UDF-Test" group in the Outlookbar
; *****************************************************************************
; Get list of groups in the OutlookBar
Global $aResult = _OL_BarGroupGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutAdd Example Script", "Error getting list of groups in OutlookBar. @error = " & @error & ", @extended = " & @extended)
; Search for group "Outlook-UDF-Test"
Global $iFound = 0
For $iIndex = 1 To $aResult[0][0]
	If $aResult[$iIndex][0] = "Outlook-UDF-Test" Then
		$iFound = $iIndex
		ExitLoop
	EndIf
Next
; Group not found - create it
If $iFound = 0 Then
	_OL_BarGroupAdd($oOutlook, "Outlook-UDF-Test", 2)
	If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutAdd Example Script", "Error adding group 'Outlook-UDF-Test' to the OutlookBar. @error = " & @error & ", @extended: " & @extended)
	MsgBox(64, "OutlookEX UDF: _OL_BarShortcutAdd Example Script", "Group 'Outlook-UDF-Test' has been added to the OutlookBar")
	$iFound = 2
EndIf
; Add shortcut to the group "Outlook-UDF-Test". The group is accessed by its index value
_OL_BarShortcutAdd($oOutlook, $iFound, "Outlook-UDF-Test-Shortcut", "http://www.autoitscript.com")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutAdd Example Script", "Error adding shortcut to group 'Outlook-UDF-Test'. @error = " & @error & ", @extended = " & @extended)
; Activate the OutlookBar and show the created shortcut
$oOutlook.ActiveExplorer.ShowPane($olOutlookBar, True)
MsgBox(64, "OutlookEX UDF: _OL_BarShortcutAdd Example Script", "Shortcut successfully added to group 'Outlook-UDF-Test'!")

_OL_Close($oOutlook)