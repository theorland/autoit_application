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
; Delete a member from a distribution list by name
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olDistributionList, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Could not find a distribution list item in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
Global $oItem = _OL_DistListMemberDelete($oOutlook, $aOL_Item[1][0], Default, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Error deleting a member from distribution list in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Member successfully deleted from distribution list!")

; *****************************************************************************
; Example 2
; Delete a member to a distribution list as object (name already resolved)
; *****************************************************************************
Global $sRecipient = InputBox("OutlookEX UDF: _OL_DistListMemberDelete Example Script","Please enter name of recipient to be deleted from the distribution list")
Global $oOL_Recipient = $oOutlook.Session.CreateRecipient($sRecipient)
If @error <> 0 Or Not IsObj($oOL_Recipient) Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Error creating recipient. @error = " & @error & ", @extended = " & @extended)
$oOL_Recipient.Resolve
If @error <> 0 Or Not $oOL_Recipient.Resolved Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Error resolving recipient. @error = " & @error & ", @extended = " & @extended)
$oItem = _OL_DistListMemberDelete($oOutlook, $aOL_Item[1][0], Default, $oOL_Recipient)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Error deleting a member from distribution list in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_DistListMemberDelete Example Script", "Member successfully deleted from distribution list!")

_OL_Close($oOutlook)