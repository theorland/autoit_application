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
; Get a list of all conflicts for the specified item
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "[Subject]='TestAppointment-Conflict'", "", "", "EntryID,Subject")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemConflictGet Example Script", "Could not find an appointment item in folder '*\Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
$Result = _OL_ItemConflictGet($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemConflictGet Example Script", "Error getting list of conflicts for appointment in folder '*\Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($Result, "OutlookEX UDF: _OL_ItemConflictGet Example Script", "", 0, "|", "Object in conflict|OlObjectClass|Name")

_OL_Close($oOutlook)