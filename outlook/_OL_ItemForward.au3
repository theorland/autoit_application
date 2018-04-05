#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error running _OL_Open. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

_Example1($oOutlook)
_Example2($oOutlook)

_OL_Close($oOutlook)
Exit

; *****************************************************************************
; Example 1
; Forward a mail item with subject "Test" from the users Inbox
; to the current user
; *****************************************************************************
Func _Example1($oOutlook)

	; Access the Inbox of the current user
	Local $aFolder = _OL_FolderAccess($oOutlook, "", $olFolderInbox)
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error accessing the users Inbox. @error = " & @error & ", @extended = " & @extended)
	; Search a mail item in the Inbox with subject "Test"
	Local $aOL_Item = _OL_ItemFind($oOutlook, $aFolder[1], $olMail, "[Subject]='Test'", "", "", "EntryID")
	If $aOL_Item[0][0] = 0 Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Could not find a mail item with the subject 'Test' in the Inbox. @error = " & @error & ", @extended = " & @extended)
	; Create a copy of the mail item which can be forwarded
	Local $oForwardItem = _OL_ItemForward($oOutlook, $aOL_Item[1][0], Default, 0)
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error forwarding mail item. @error = " & @error & ", @extended = " & @extended)
	; Set the current user as recipient of the forwarded copy
	_OL_ItemRecipientAdd($oOutlook, $oForwardItem, Default, $olTo, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name)
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error in function _OL_ItemRecipientAdd. @error = " & @error & ", @extended = " & @extended)
	; Get some properties of the mail
	Local $aProperties = _OL_ItemGet($oOutlook, $oForwardItem, Default, "Body,BodyFormat,HTMLBody")
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error in function _OL_ItemGet. @error = " & @error & ", @extended = " & @extended)
	; Prepend some text to the body of the mail
	If  Int($aProperties[2][1]) = $olFormatHTML Then
		$aProperties[3][1] = "Modified Test Text<p>" & $aProperties[3][1]
		_OL_ItemModify($oOutlook, $oForwardItem, Default, "HTMLBody=" & $aProperties[3][1])
	Else
		$aProperties[1][1] = "Modified Test Text" & @CRLF & $aProperties[1][1]
		_OL_ItemModify($oOutlook, $oForwardItem, Default, "Body=" & $aProperties[1][1])
	EndIf
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error in function _OL_ItemModify. @error = " & @error & ", @extended = " & @extended)
	; Send item
	_OL_ItemSend($oOutlook, $oForwardItem)
    If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error sending the mail item. @error = " & @error & ", @extended = " & @extended)
	; Display success message
	MsgBox(64, "OutlookEX UDF: _OL_ItemForward Example Script", "Mail item successfully forwarded!")

EndFunc   ;==>_Example1

; *****************************************************************************
; Example 2
; Forward a contact item in Vcal format
; *****************************************************************************
Func _Example2($oOutlook)

	; Search for contact items in the source contact folder of the test environment
	Local $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "EntryID")
	If $aOL_Item[0][0] = 0 Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Could not find a contact item in folder 'Outlook-UDF-Test\SourceFolder\Contact'. @error = " & @error)
	; Create a copy of the contact item which can be forwarded
	Local $oForwardItem = _OL_ItemForward($oOutlook, $aOL_Item[1][0], Default, 1)
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error forwarding contact item. @error = " & @error & ", @extended = " & @extended)
	; Set the current user as recipient of the forwarded copy
	_OL_ItemRecipientAdd($oOutlook, $oForwardItem, Default, $olTo, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name)
	If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error in function _OL_ItemRecipientAdd. @error = " & @error & ", @extended = " & @extended)
	; Send item
	_OL_ItemSend($oOutlook, $oForwardItem)
    If @error Then Return MsgBox(16, "OutlookEX UDF: _OL_ItemForward Example Script", "Error sending the contact item. @error = " & @error & ", @extended = " & @extended)
	; Display success message
	MsgBox(64, "OutlookEX UDF: _OL_ItemForward Example Script", "Contact item successfully forwarded in Vcal/Vcard format!")

EndFunc   ;==>_Example2
