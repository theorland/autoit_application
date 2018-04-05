#include <OutlookEX.au3>

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Create an empty PST archive in C:\temp
; *****************************************************************************
Global $oPST = _OL_PSTCreate($oOutlook, "C:\temp\Outlook-UDF-Test.pst", "Outlook-UDF-PST")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_PSTCreate Example Script", "Error creating 'C:\temp\Outlook-UDF-Test.pst' archive. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_PSTCreate Example Script", "Archive 'C:\temp\Outlook-UDF-Test.pst' successfully created!")

_OL_Close($oOutlook)