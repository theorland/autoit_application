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
; Find a contact and display the properties
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", Default, '[FirstName] = "TestFirstName"', "", "", "EntryID", "", 0, "")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemGet Example Script", "Could not find a contact in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
Global $aOL_Properties = _OL_ItemGet($oOutlook, $aOL_Item[1][0])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemGet Example Script", "Error accessing properties. @error = " & @error)
_ArrayDisplay($aOL_Properties, "OutlookEX UDF: All properties of a contact item (name, value, datatype)", "", 0, "|", "Name|Value|Type")

; *****************************************************************************
; Example 2
; Only display a few properties of the same item
; *****************************************************************************
$aOL_Properties = _OL_ItemGet($oOutlook, $aOL_Item[1][0], Default, "ConversationTopic,Initials,hasPicture")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemGet Example Script", "Error accessing properties. @error = " & @error)
_ArrayDisplay($aOL_Properties, "OutlookEX UDF: Get a few properties of a contact item (name, value, datatype)", "", 0, "|", "Name|Value|Type")

_OL_Close($oOutlook)