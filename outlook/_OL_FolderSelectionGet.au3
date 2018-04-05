#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get information about selected items in the the current folder.
; *****************************************************************************
Global $aResult = _OL_FolderSelectionGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderSelectionGet Example Script", "Error accessing current folder. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: All items selected in the current folder", "", 0, "|", "Object|EntryID|OlObjectClass")

_OL_Close($oOutlook)