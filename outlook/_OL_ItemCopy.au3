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
; Find a note and create a copy in the same folder
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Notes", $olNote, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCopy Example Script", "Could not find a note item in folder 'Outlook-UDF-Test\SourceFolder\Notes'. @error = " & @error)
_OL_ItemCopy($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCopy Example Script", "Error copying specified note. @error = " & @error)
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Notes")
$aResult[1].Display
MsgBox(64, "OutlookEX UDF: _OL_ItemCopy Example Script", "Note successfully copied in 'Outlook-UDF-Test\SourceFolder\Notes'!")

; *****************************************************************************
; Example 2
; Find a contact and create a copy in the target folder
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCopy Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
_OL_ItemCopy($oOutlook, $aOL_Item[1][0], Default, "*\Outlook-UDF-Test\TargetFolder\Contacts")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCopy Example Script", "Error copying specified contact to another folder. @error = " & @error)
$aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Contacts")
$aResult[1].Display
MsgBox(64, "OutlookEX UDF: _OL_ItemCopy Example Script", "Contact successfully copied to 'Outlook-UDF-Test\TargetFolder\Contacts'!")

_OL_Close($oOutlook)