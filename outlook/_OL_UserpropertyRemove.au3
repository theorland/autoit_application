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
; Remove a Userproperty from a folder
; *******************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "Error accessing folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
_OL_UserpropertyRemove($oOutlook, Default, $aFolder[1], "TestUserPropertyFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "Error removing user property 'TestUserPropertyFolder' from folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "User property 'TestUserPropertyFolder' successfully removed from folder '*\Outlook-UDF-Test\SourceFolder\Contacts'")

; *******************************************************************************
; Example 2
; Remove a Userproperty from an item
; *******************************************************************************
Global $aItems = _OL_ItemFind($oOutlook, $aFolder[1], $olContact, '[LastName]="TestLastName"', "", "", "EntryID,LastName")
If @error <> 0 Or $aItems[0][0] = 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "Error finding a contact in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
_OL_UserpropertyRemove($oOutlook, Default, $aItems[1][0], "TestUserPropertyItem")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "Error removing user property 'TestUserPropertyItem' from contact 'TestLastName' in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_UserpropertyRemove Example Script", "User property 'TestUserPropertyItem' successfully removed from contact 'TestLastName' in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'")

_OL_Close($oOutlook)