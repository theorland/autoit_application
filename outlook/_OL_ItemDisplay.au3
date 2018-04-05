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
; Find and display a note with default values
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Notes", $olNote, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Could not find a note item in folder 'Outlook-UDF-Test\SourceFolder\Notes'. @error = " & @error)
Global $oInspector = _OL_ItemDisplay($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Error displaying specified note. @error = " & @error)
MsgBox(64, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Note successfully displayed")
$oInspector.Close(1)

; *****************************************************************************
; Example 2
; Find a contact and display with size settings for the window
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
$oInspector =_OL_ItemDisplay($oOutlook, $aOL_Item[1][0], Default, 500, 500, 100, 100)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Error displaying specified contact. @error = " & @error)
MsgBox(64, "OutlookEX UDF: _OL_ItemDisplay Example Script", "Contact successfully displayed")
$oInspector.Close(1)

_OL_Close($oOutlook)