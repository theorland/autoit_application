#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get information about the current folder.
; *****************************************************************************
Global $aResult = _OL_FolderGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderGet Example Script", "Error accessing current folder. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: properties of the current folder")

_OL_Close($oOutlook)