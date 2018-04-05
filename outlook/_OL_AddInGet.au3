#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Display all COM Addins
; *****************************************************************************
Global $aResult = _OL_AddInGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_AddInGet Example Script", "Error accessing the COM Addins. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: Installed COM Addins", "", 0, "|", "|Object|Active|Description|GUID|ProgID")

_OL_Close($oOutlook)
Exit
