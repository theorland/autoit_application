#include <OutlookEX.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Retrieve a list of all searchfolders in all stores
; *****************************************************************************
Global $aResult = _OL_SearchFolderGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_SearchFolderGet Example Script", "Error getting list of searchfolders. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_SearchFolderGet Example Script", "", 0, Default, _
"Store displayname|StoreID|Searchfolder displayname|EntryID|Default item type|Folderpath")

_OL_Close($oOutlook)