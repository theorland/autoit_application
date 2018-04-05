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
; Find a mail and delete it
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDelete Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
_OL_ItemDelete($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemDelete Example Script", "Error deleting mail. @error = " & @error & ", @extended = " & @extended)

; Display Target folder
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail")
$aResult[1].Display

MsgBox(64, "OutlookEX UDF: _OL_ItemDelete Example Script", "Item successfully deleted from 'Outlook-UDF-Test\SourceFolder\Mail'!")

_OL_Close($oOutlook)