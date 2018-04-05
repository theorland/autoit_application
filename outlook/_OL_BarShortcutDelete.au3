#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Delete the "Outlook-UDF-Test-Shortcut" from the "Outlook-UDF-Test" group
; in the Outlookbar (created by example script _OL_BarShortcutAdd)
; *****************************************************************************
; Get list of groups in the OutlookBar
Global $aResult = _OL_BarGroupGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Error getting list of groups in OutlookBar. @error = " & @error & ", @extended = " & @extended)
; Search for group "Outlook-UDF-Test"
Global $iFound1 = 0
For $iIndex = 1 To $aResult[0][0]
	If $aResult[$iIndex][0] = "Outlook-UDF-Test" Then
		$iFound1 = $iIndex
		ExitLoop
	EndIf
Next
; Group not found - exit
If $iFound1 = 0 Then Exit MsgBox(64, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Group 'Outlook-UDF-Test' could not be found. Please use example '_OL_BarShortcutAdd' to create.")
; Get list of shortcuts in group Outlook-UDF-Test
$aResult = _OL_BarShortcutGet($oOutlook, "Outlook-UDF-Test")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Error accessing group 'Outlook-UDF-Test'. @error = " & @error & ", @extended = " & @extended)
; Search for shortcut "Outlook-UDF-Test-Shortcut"
Global $iFound2 = 0
For $iIndex = 1 To $aResult[0][0]
	If $aResult[$iIndex][0] = "Outlook-UDF-Test-Shortcut" Then
		$iFound2 = $iIndex
		ExitLoop
	EndIf
Next
; Shortcut not found - exit
If $iFound2 = 0 Then Exit MsgBox(64, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Shortcut 'Outlook-UDF-Test-Shortcut' not found in group 'Outlook-UDF-Test'. Please use example '_OL_BarShortcutAdd' to create.")
; Delete shortcut from group "Outlook-UDF-Test". The group is accessed by its index value
_OL_BarShortcutDelete($oOutlook, $iFound1, $iFound2)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Error deleting shortcut 'Outlook-UDF-Test-Shortcut' from group 'Outlook-UDF-Test'. @error = " & @error & ", @extended = " & @extended)
; Activate the OutlookBar
$oOutlook.ActiveExplorer.ShowPane($olOutlookBar, True)
MsgBox(64, "OutlookEX UDF: _OL_BarShortcutDelete Example Script", "Shortcut 'Outlook-UDF-Test-Shortcut' successfully deleted from group 'Outlook-UDF-Test'!")

_OL_Close($oOutlook)