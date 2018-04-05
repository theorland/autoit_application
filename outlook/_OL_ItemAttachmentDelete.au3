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
; Delete an attachment from a mail item
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentDelete Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
; Find mail with attachments
Global $oItem, $fFound = False
For $iIndex = 1 To $aOL_Item[0][0]
	$oItem = $oOutlook.Session.GetItemFromID($aOL_Item[$iIndex][0])
	If $oItem.Attachments.Count > 0 Then
		$fFound = True
		ExitLoop
	EndIf
Next
If Not $fFound Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentDelete Example Script", "Could not find a mail with an attachment in 'Outlook-UDF-Test\SourceFolder\Mail'.")
$oItem = _OL_ItemAttachmentDelete($oOutlook, $aOL_Item[$iIndex][0], Default, 1)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemAttachmentDelete Example Script", "Error deleting an attachment from mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Show item
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemAttachmentDelete Example Script", "Attachment successfully deleted from mail item!")

_OL_Close($oOutlook)