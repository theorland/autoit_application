#include <OutlookEX.au3>

Global $oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $sScope = $oOL.GetNamespace("MAPI").GetDefaultFolder($olFolderInbox).FolderPath

; *****************************************************************************
; Example 1
; Create a searchfolder with all unread mails of the inbox
; *****************************************************************************
Global $oSF = _OL_SearchFolderCreate($oOL, "SearchFolder - Unread", $sScope, '"urn:schemas:httpmail:read" = 0', True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_SearchFolderCreate Example Script", "Error returned by example 1. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 2
; Create a searchfolder with all mails where the senders name contains "Thomas"
; *****************************************************************************
$oSF = _OL_SearchFolderCreate($oOL, "SearchFolder - Sender", $sScope, '"urn:schemas:httpmail:fromname" LIKE ''%Thomas%''', True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_SearchFolderCreate Example Script", "Error returned by example 2. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 3
; Create a searchfolder with all unread mails where the senders name
; contains "Thomas"
; *****************************************************************************
$oSF = _OL_SearchFolderCreate($oOL, "SearchFolder - Sender AND Unread", $sScope, '"urn:schemas:httpmail:fromname" LIKE ''%Thomas%'' AND "urn:schemas:httpmail:read" = 0', True)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_SearchFolderCreate Example Script", "Error returned by example 3. @error = " & @error & ", @extended = " & @extended)

MsgBox(64, "OutlookEX UDF: _OL_SearchFolderCreate Example Script", "All searchfolders created successfully!")

_OL_Close($oOL)