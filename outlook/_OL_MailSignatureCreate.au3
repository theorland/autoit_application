#include <OutlookEX.au3>

Global $iReply = MsgBox(308, "OutlookEX UDF: _OL_MailSignatureCreate Example Script", "This script creates signature 'Outlook-UDF-Test'." & @CRLF & _
		"To delete the signature please run '_OL_SignatureDelete'." & @CRLF & @CRLF & _
		"Are you sure you want to create the signature?")
If $iReply <> 6 Then Exit

Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Use Word to create the signature content.
; Then call _OL_MailSignatureCreate and pass the content.
; *****************************************************************************
Global $sGiven = "Firstname"
Global $sSurname = "Lastname"
Global $sAddress1 = "Streetname and number"
Global $sAddress2 = "Town"
Global $sPostcode = "PostalCode"
Global $sFax = "Your Faxnumber"
Global $sMobile = "Your Mobile Phone number"
Global $sTitle = "Your Title - if you have one"
Global $sCompany = "Name of your company"
Global $sPhone = "Your Phone number"
Global $sEmail = "yourmailaddress@yourcompany.TLD"
Global $sWeb = "http://www.yourcompany.TLD"
Global $sExt = "Extension"
Global $sPicture = @ScriptDir & "\The_outlook.jpg"

Global Const $END_OF_STORY = 6
Global $oWord, $oDoc, $oSelection, $oRange, $oTable, $oCell, $oCellRange, $oLink
Global $sStyle = "No Spacing"

; Set up word template
$oWord = ObjCreate("Word.Application")
$oDoc = $oWord.Documents.Add()
$oSelection = $oWord.Selection

; Sets initial font typeface, colour etc., inserts name and title
If $oWord.LanguageSettings.LanguageID($msoLanguageIDUI) = 1031 Then $sStyle = "Kein Leerraum"
$oSelection.Style = $sStyle
$oSelection.Font.Name = "Verdana"
$oSelection.Font.Size = 10
$oSelection.Font.Bold = True
$oSelection.Font.Color = 0x002D9A
$oSelection.TypeParagraph()
$oSelection.TypeText($sGiven & " " & $sSurname)
$oSelection.Font.Size = 7
$oSelection.TypeText(Chr(11))
$oSelection.Font.Size = 8
$oSelection.Font.Bold = False
$oSelection.TypeText($sTitle)

; Inserts a 2 column table to contain the Address (left) and the contact information (right)
$oSelection.TypeText(Chr(11))
$oSelection.TypeParagraph()
$oRange = $oSelection.Range
$oDoc.Tables.Add($oRange, 5, 2)
$oTable = $oDoc.Tables(1)
$oTable.Cell(1, 1).Range.Text = $sCompany
$oTable.Cell(2, 1).Range.Text = $sAddress1
$oTable.Cell(3, 1).Range.Text = $sAddress2
$oTable.Cell(4, 1).Range.Text = $sPostcode
$oTable.Cell(1, 2).Range.Text = "Tel: " & $sPhone & "  |  Ext: " & $sExt
$oTable.Cell(2, 2).Range.Text = "Fax: " & $sFax
$oTable.Cell(3, 2).Range.Text = "Mobile: " & $sMobile

; Creates a clickable hyperlink
$oCell = $oTable.Cell(4, 2)
$oCellRange = $oCell.Range
$oCell.Select
$oSelection.TypeText("Web: ")
$oLink = $oSelection.Hyperlinks.Add($oSelection.Range, $sWeb, Default, Default, $sWeb)
$oLink.Range.Font.Name = "Verdana"
$oLink.Range.Font.Size = 8
$oLink.Range.Font.Bold = False

; Creates a clickable mailto: email address
$oCell = $oTable.Cell(5, 2)
$oCellRange = $oCell.Range
$oCell.Select
$oSelection.typeText("Email: ")
$oLink = $oSelection.Hyperlinks.Add($oSelection.Range, "mailto: " & $sEmail, Default, Default, $sEmail)
$oLink.Range.Font.Name = "Verdana"
$oLink.Range.Font.Size = 8
$oLink.Range.Font.Bold = False
$oTable.AutoFitBehavior(1)
$oSelection.EndKey($END_OF_STORY)

; Insert logo
$oSelection.TypeText(Chr(11))
$oSelection.InlineShapes.AddPicture($sPicture)

; Select the whole text
$oSelection = $oDoc.Range()

; Create the Signature
Global $iResult = _OL_MaiLSignatureCreate("Outlook-UDF-Test", $oWord, $oSelection)
If @error <> 0 Then
	MsgBox(16, "OutlookEX UDF: _OL_MailSignatureCreate Example Script", "Signature 'Outlook-UDF-Test' could not be created. @error = " & @error & ", @extended: " & @extended)
Else
	MsgBox(64, "OutlookEX UDF: _OL_MailSignatureCreate Example Script", "Signature 'Outlook-UDF-Test' successfully created.")
EndIf

; End Word
$oDoc.Saved = True
$oWord.Quit

_OL_Close($oOutlook)