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
; Add a Userproperty to a folder so the property can be used in a view
; *******************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "Error accessing folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
_OL_UserPropertyAdd($oOutlook, Default, $aFolder[1], "TestUserPropertyFolder", $olText)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "Error adding user property 'TestUserPropertyFolder' to folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "User property 'TestUserPropertyFolder' successfully added to folder '*\Outlook-UDF-Test\SourceFolder\Contacts'")

; *******************************************************************************
; Example 2
; Add a Userproperty to an item
; *******************************************************************************
Global $aItems = _OL_ItemFind($oOutlook, $aFolder[1], $olContact, '[LastName]="TestLastName"', "", "", "EntryID,LastName")
If @error <> 0 Or $aItems[0][0] = 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "Error finding a contact in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
_OL_UserPropertyAdd($oOutlook, Default, $aItems[1][0], "TestUserPropertyItem", $olText, Default, "TestUserPropertyContent", True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "Error adding user property 'TestUserPropertyItem' to contact 'TestLastName' in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_UserpropertyAdd Example Script", "User property 'TestUserPropertyItem' successfully added to contact 'TestLastName' in folder '*\Outlook-UDF-Test\SourceFolder\Contacts'")

_OL_Close($oOutlook)