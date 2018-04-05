#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *******************************************************************************
; Example 1
; Get a list of user properties for a folder
; *******************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyGet Example Script", "Error accessing folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
Global $aFolderProperties = _OL_UserpropertyGet($oOutlook, Default, $aFolder[1])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyGet Example Script", "Error getting user properties for folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aFolderProperties, "Userproperties for folder '*\Outlook-UDF-Test\SourceFolder\Contacts'", "", 0, "|", "|name|type")

; *******************************************************************************
; Example 2
; Get a list of user properties for an item
; *******************************************************************************
Global $aItems = _OL_ItemFind($oOutlook, $aFolder[1], $olContact, '[LastName]="TestLastName"', "", "", "EntryID,LastName")
If @error <> 0 Or $aItems[0][0] = 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_UserpropertyGet Example Script", "Error finding a contact in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
Global $aItemProperties = _OL_UserpropertyGet($oOutlook, Default, $aItems[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyGet Example Script", "Error getting user properties for item 'TestLastName' in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
_ArrayDisplay($aItemProperties, "Userproperties for item 'TestLastName'", "", 0, "|", "|name|type|content")

_OL_Close($oOutlook)