#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Adds a category with color orange
; *****************************************************************************
_OL_CategoryAdd($oOutlook, "Outlook-UDF-Test", $olCategoryColorOrange)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_CategoryAdd Example Script", "Error adding category. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_CategoryAdd Example Script", "Category 'Outlook-UDF-Test' successfully added!")

_OL_Close($oOutlook)