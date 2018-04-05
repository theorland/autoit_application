#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)
Global $aItems

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 1 - Search for contacts with firstname = TestFirstName
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, '[FirstName] = "TestFirstName"', "", "", "", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Find contacts by firstname")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemFind Example Script", "Error finding a contact. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 2 - Search for appointments with "Room" as location (partial match)
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "", "Location", "Room", "EntryID,Subject,Location", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Find appointments by partial search")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemFind Example Script", "Error finding an appointment. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 3 - Get number of items (contacts without distribution lists) in the contacts folder
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "", "", 4)
If @error = 0 Then
	MsgBox(64, "OutlookEX UDF: _OL_ItemFind Example Script", "Number of items found: " & $aItems)
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemFind Example Script", "Could not find an item in the contacts folders. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 4 - Get unread mails from a folder and all subfolders
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test", $olMail, "[UnRead]=True", "", "", "Subject,Body", "", 1)
If IsArray($aItems) Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Unread mails")
Else
	MsgBox(48, "OutlookEX UDF: _OL_ItemFind Example Script", "Could not find an unread mail. @error = " & @error & ", @extended: " & @extended)
EndIf

_OL_Close($oOutlook)