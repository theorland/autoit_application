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
; Modify the subject of a mail item
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemModify Example Script", "Could not find a mail item in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
_OL_ItemModify($oOutlook, $aOL_Item[1][0], Default, "Subject=Modified Subject")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemModify Example Script", "Error modifying a mail in folder 'Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)

; Display modified item
$oOutlook.Session.GetItemFromID($aOL_Item[1][0]).Display
MsgBox(64, "OutlookEX UDF: _OL_ItemModify Example Script", "Subject of mail item successfully modified!")

; *****************************************************************************
; Example 2
; Modify a contact and passing the properties as an array
; *****************************************************************************
$aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", Default, '[FirstName] = "TestFirstName"', "", "", "EntryID", "", 0, "")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemGet Example Script", "Could not find a contact in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)
Global $aOL_Properties[10] = ["Companies=Mega Company", "AssistantName=Best secretary ever","BusinessAddressStreet=Mainstreet 1", "BusinessAddressCity=Metropolis", _
	"BusinessAddressCountry=Atlantis","BusinessAddressPostalCode=0815","Title=Professor"]
_OL_ItemModify($oOutlook, $aOL_Item[1][0], Default, $aOL_Properties)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemModify Example Script", "Error modifying a contact in folder 'Outlook-UDF-Test\SourceFolder\Contacts'. @error = " & @error)

; Display modified item
$oOutlook.Session.GetItemFromID($aOL_Item[1][0]).Display
MsgBox(64, "OutlookEX UDF: _OL_ItemModify Example Script", "Contact item (Properties: Companies, AssistantName, BusinessAddressStreet, BusinessAddressCity=Metropolis, BusinessAddressCountry, BusinessAddressPostalCode, Title) successfully modified!")

_OL_Close($oOutlook)