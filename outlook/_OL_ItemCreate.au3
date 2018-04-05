#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oItem
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Create an appointment with subject, start- and enddate.
; Attendee is the current user.
; Set the body using Microsoft Word as editor.
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olAppointmentItem, "*\Outlook-UDF-Test\TargetFolder\Calendar", "", "Subject=TestSubject", "Start=" & _NowCalc(), "End=" & _DateAdd("h", 3, _NowCalc()), _
		"Location=Building A, Room 10", "RequiredAttendees=" & $oOutlook.GetNameSpace("MAPI" ).CurrentUser)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating an appointment in folder 'Outlook-UDF-Test\TargetFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)

; Set the body of the appointment using Word as editor
Global Const $wdCollapseStart = 1 ; Collapse the range to the starting point
Global Const $wdWord = 2 ; A word
Global Const $wdCharacter = 1 ; A character

Global $oDoc = $oItem.GetInspector.WordEditor ; Get the Microsoft Word Document Object Model
Global $oRange = $oDoc.Range ; Get the range object
$oRange.InsertAfter("This is a test") ; Insert some text
$oRange.Collapse($wdCollapseStart) ; Move the range start/end to the start of the document
$oRange.MoveStart($wdWord, 1) ; Move the range start/end to word 1
$oRange.MoveEnd($wdWord, 2) ; Move the range end two words to the right
$oRange.MoveEnd($wdCharacter, -1) ; Move the range end one character to the left (so the space isn't included)
$oRange.Font.Underline = True ; Set the font.underline property for the range
$oItem.Display()
MsgBox(64, "OutlookEX UDF: _OL_ItemCreate Example Script", "The body of this appointment has been created using Microsoft Word as editor.")
$oItem.Close($olSave)

; *****************************************************************************
; Example 2
; Create a contact with first- and lastname
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olContactItem, "*\Outlook-UDF-Test\TargetFolder\Contacts", "", "FirstName=TestFirstName", "LastName=TestLastName")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a contact in folder 'Outlook-UDF-Test\TargetFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)
; Add a picture to the contact
$oItem.AddPicture(@ScriptDir & "\The_Outlook.jpg")
$oItem.Save()

; *****************************************************************************
; Example 3
; Create a distribution list with importance set to high
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olDistributionListItem, "*\Outlook-UDF-Test\TargetFolder\Contacts", "", "Subject=TestDistributionList", "Importance=" & $olImportanceHigh)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a distributionlist in folder 'Outlook-UDF-Test\TargetFolder\Contacts'. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 4
; Create a html mail plus two attachments (plus one inline picture = attachment)
; but don't send it
; Inline picture using Content-ID
; http://stackoverflow.com/questions/9158706/how-to-embed-an-image-on-an-outlook-2007-vsto
; *****************************************************************************
; Create the item without setting the body. We first need to add the picture before we can refer to in by the HTML body.
$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "*\Outlook-UDF-Test\TargetFolder\Mail", "", "Subject=TestMail", "BodyFormat=" & $olFormatHTML)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)
; Add all attachments
$oItem = _OL_ItemAttachmentAdd($oOutlook, $oItem, Default, @ScriptDir & "\The_Outlook.jpg", @ScriptDir & "\_OL_ItemCopy.au3, 4", @ScriptDir & "\_OL_Foldertree.au3")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error adding an attachment to a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = "  & @error & ", @extended = " & @extended)
; Modify the item to add HTML  that refers to the picture
$oItem = _OL_ItemModify($oOutlook, $oItem, Default, "HTMLBody=Bodytext in <b>bold</b><img src='cid:The_Outlook.jpg'>Embedded image.")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error modifying the item in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = "  & @error & ", @extended = " & @extended)
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_ItemCreate Example Script", "Mail with inline picture created.")

; *****************************************************************************
; Example 5
; Create a mail from a template
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "*\Outlook-UDF-Test\TargetFolder\Mail", @ScriptDir & "\_OL_ItemCreate.oft")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 6
; Create a note with a body and a specified display width
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olNoteItem, "*\Outlook-UDF-Test\TargetFolder\Notes", "", "Body=TestNote", "Width=350")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a note in folder 'Outlook-UDF-Test\TargetFolder\Notes'. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 7
; Create a task with a startdate of today
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olTaskItem, "*\Outlook-UDF-Test\TargetFolder\Tasks", "", "Subject=TestSubject", "StartDate=" & _NowDate())
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", "Error creating a task in folder 'Outlook-UDF-Test\TargetFolder\Tasks'. @error = " & @error & ", @extended = " & @extended)

; Display Target folder
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder")
$aResult[1].Display

MsgBox(64, "OutlookEX UDF: _OL_ItemCreate Example Script", "All items successfully created in 'Outlook-UDF-Test\TargetFolder' and its subfolders!")

_OL_Close($oOutlook)