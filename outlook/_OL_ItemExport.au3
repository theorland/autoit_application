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
; Export all contacts from the test folder to c:\temp\_OL_ItemExport.csv
; *****************************************************************************
Global $aData = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", Default, "", "", "", "FirstName, LastName")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemExport Example Script", "Error getting contacts from folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
Global $iResult = _OL_ItemExport("C:\temp\_OL_ItemExport.csv", "", "", 1, "FirstName,LastName", $aData)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemExport Example Script", "Error exporting contacts to file 'C:\temp\_OL_ItemExport.csv'. @error = " & @error & ", @extended = " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_ItemExport Example Script", $iResult & " contact(s) successfully exported to file 'C:\temp\_OL_ItemExport.csv'.")

_OL_Close($oOutlook)