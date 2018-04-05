#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error Then Exit MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_Open. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Search the subject of mail items containing $sString1 or
; starting with $sString2
; Returns: Subject, CreationTime, Recipient
; *****************************************************************************
Global $sString1 = "Test", $sString2 = "Mail"
Global $aSearchArray[3][4] = [[2, 4],[0x0037001E, 3, $sString1, "or"],["subject", 2, $sString2, ""]]
Global $aResult = _OL_ItemSearch($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $aSearchArray, "subject,CreationTime,To")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 1")
EndIf

; *****************************************************************************
; Example 2
; Same as example 1 but filter specified in DASL format
; Returns: EntryID, Subject, CreationTime, Recipient
; *****************************************************************************
Global $sFilter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001E"" ci_phrasematch '" & $sString1 & "' or ""http://schemas.microsoft.com/mapi/proptag/0x0037001E"" ci_startswith '" & $sString2 & "'"
$aResult = _OL_ItemSearch($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $sFilter, "EntryID,Subject,CreationTime,To")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 2")
EndIf

; *****************************************************************************
; Example 3
; Search the subject of mail items with exact matching $sString3
; Returns: EntryID, Subject
; *****************************************************************************
Global $sString3 = "TestMail"
Global $aSearchArray[2][4] = [[1, 4],[0x0037001E, 1, $sString3]]
$aResult = _OL_ItemSearch($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $aSearchArray, "EntryID,Subject")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 3")
EndIf

; *****************************************************************************
; Example 4
; Search the body of mail items for phrase $sString4
; Returns: EntryID, Subject, max 255 characters of the body
; *****************************************************************************
Global $sString4 = "Bodytext"
$sFilter = "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch '" & $sString4 & "'"
$aResult = _OL_ItemSearch($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $sFilter, "EntryID,Subject,urn:schemas:httpmail:textdescription")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 4")
EndIf

; *****************************************************************************
; Example 5
; Search the contacts for a specific name.
; Returns: EntryID, FullName and HomeAddressCountry
; *****************************************************************************
Global $sString5 = "%FirstName"
$sFilter = "@SQL=""http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8005001f"" LIKE '" & $sString5 & "'"
$aResult = _OL_ItemSearch($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $sFilter, "EntryID,FullName,HomeAddressCountry")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 5")
EndIf

; *****************************************************************************
; Example 6
; Search the inbox for mails with at least one attachment
; Returns: EntryID, subject
; *****************************************************************************
; Access the default mail folder
Global $aFolder = _OL_FolderAccess($oOutlook, "", $olFolderInbox)
If @error Then Exit MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_FolderAccess. @error = " & @error & ", @extended = " & @extended)
$sFilter = "@SQL=(""urn:schemas:httpmail:hasattachment"" = 1)"
$aResult = _OL_ItemSearch($oOutlook, $aFolder[1], $sFilter, "EntryID,subject")
If @error Then
    MsgBox(16, "OutlookEX UDF - _OL_ItemSearch Example Script", "Error running _OL_ItemSearch. @error = " & @error & ", @extended = " & @extended)
Else
    _Arraydisplay($aResult, "Example 6")
EndIf

_OL_Close($oOutlook)