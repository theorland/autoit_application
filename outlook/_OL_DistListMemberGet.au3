#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $iResult = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get a list of all members of a distribution list
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olDistributionList, "", "", "", "EntryID")
If @error Or $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberGet Example Script", "Could not find a distribution list item in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
$iResult = _OL_DistListMemberGet($oOutlook, $aOL_Item[1][0], Default)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberGet Example Script", "Error getting member list of distribution list in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($iResult, "OutlookEX UDF: _OL_DistListMemberGet Example Script", "", 0, "|", "Recipient object|Name|EntryID")

_OL_Close($oOutlook)