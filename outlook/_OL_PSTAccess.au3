#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Access the PST archive C:\temp\Outlook-UDF-Test.pst created by _OL_PSTCreate
; *****************************************************************************
If Not FileExists("C:\temp\Outlook-UDF-Test.pst") Then Exit MsgBox(16, "OutlookEX UDF: _OL_PSTAccess Example Script", "PST 'C:\temp\Outlook-UDF-Test.pst' does not exist. Run  _OL_PSTCreate to create it!")
Global $oPST = _OL_PSTAccess($oOutlook, "C:\temp\Outlook-UDF-Test.pst", "Outlook-UDF-PST")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_PSTAccess Example Script", "Error accessing 'C:\temp\Outlook-UDF-Test.pst' archive. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_PSTAccess Example Script", "Archive 'C:\temp\Outlook-UDF-Test.pst' successfully accessed!")

_OL_Close($oOutlook)