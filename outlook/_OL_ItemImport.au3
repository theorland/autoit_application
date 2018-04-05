#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Import all contacts from c:\temp\_OL_ItemExport.csv to
; *****************************************************************************
Global $iResult = _OL_ItemImport($oOutlook, "C:\temp\_OL_ItemExport.csv", "", "", 1, "*\Outlook-UDF-Test\TargetFolder\Contacts", $olContactItem)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemImport Example Script", "Error creating contacts from file 'C:\temp\_OL_ItemExport.csv' in folder '*\Outlook-UDF-Test\TargetFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_ItemImport Example Script", $iResult & " contact(s) successfully imported to folder '*\Outlook-UDF-Test\TargetFolder\Contacts'.")

_OL_Close($oOutlook)