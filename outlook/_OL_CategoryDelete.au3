#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Deletes a category
; *****************************************************************************
_OL_CategoryDelete($oOutlook, "Outlook-UDF-Test")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_CategoryDelete Example Script", "Error deleting category 'Outlook-UDF-Test'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_CategoryDelete Example Script", "Category 'Outlook-UDF-Test' successfully deleted!")

_OL_Close($oOutlook)