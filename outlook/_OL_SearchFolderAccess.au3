#include <OutlookEX.au3>
#include <MsgBoxConstants.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Retrieve a list of all searchfolders in all stores
; Access the first serachfolder found and
; display all items
; *****************************************************************************
Global $aSearchFolders = _OL_SearchFolderGet($oOutlook)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_SearchFolderAccess Example Script", "Error retrieving list of searchfolders. @error = " & @error & ", @extended = " & @extended)

Global $aSearchFolder = _OL_SearchFolderAccess($oOutlook, $aSearchFolders[1][2], $aSearchFolders[1][0])
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_SearchFolderAccess Example Script", "Error accessing searchfolder. @error = " & @error & ", @extended = " & @extended)

Global $iItems = _OL_ItemFind($oOutlook, $aSearchFolder[1], $olMail, "", "", "", "Subject,Body", "", 4)
If @error = 0 Then
	MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_SearchFolderAccess Example Script", "Found " & $iItems & " items in the searchfolder.")
Else
	MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_SearchFolderAccess Example Script", "Error searching for items in the ssearchfolder. @error = " & @error & ", @extended: " & @extended)
EndIf

_OL_Close($oOutlook)