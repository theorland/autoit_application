#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Returns all categories by which Outlook items can be grouped
; *****************************************************************************
Global $aOL_Result = _OL_CategoryGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_CategoryGet Example Script", "Error accessing categories. @error = " & @error)
_ArrayDisplay($aOL_Result, "OutlookEX UDF: All categories by which Outlook items can be grouped", "", 0, "|", "Border color|Gradient bottom color|Gradient top color|CategoryID|Color|Name|ShortcutKey")

_OL_Close($oOutlook)