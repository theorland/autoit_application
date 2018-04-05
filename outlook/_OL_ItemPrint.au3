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
; Find and print a note
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Notes", $olNote, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemPrint Example Script", "Could not find a note item in folder 'Outlook-UDF-Test\SourceFolder\Notes'. @error = " & @error)
_OL_ItemPrint($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemPrint Example Script", "Error printing specified note. @error = " & @error)
MsgBox(64, "OutlookEX UDF: _OL_ItemPrint Example Script", "Note successfully printed")

; *****************************************************************************
; Example 2
; Find and print a contact
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemPrint Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
_OL_ItemPrint($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemPrint Example Script", "Error printing specified contact. @error = " & @error)
MsgBox(64, "OutlookEX UDF: _OL_ItemPrint Example Script", "Contact successfully printed")

_OL_Close($oOutlook)