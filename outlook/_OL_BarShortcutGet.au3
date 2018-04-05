#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all shortcuts in the first group of the Outlookbar
; *****************************************************************************
Global $aResult = _OL_BarGroupGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutGet Example Script", "Error getting list of groups in OutlookBar. @error = " & @error & ", @extended = " & @extended)
Global $sGroup = $aResult[1][0]
$aResult = _OL_BarShortcutGet($oOutlook, $sGroup)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_BarShortcutGet Example Script", "Error getting list of shortcuts from the first group in OutlookBar. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_BarShortcutGet Example Script - Group: " & $sGroup, "", 0, "|", "Name|Target")

_OL_Close($oOutlook)