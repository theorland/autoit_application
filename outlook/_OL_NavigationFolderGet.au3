#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get a list of all groups in the mail navigation module
; *****************************************************************************
Global $aModules = _OL_NavigationFolderGet($oOutlook, $olModuleMail)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_NavigationFolderGet Example Script", "Error getting groups of the mail navigation module. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aModules, "OutlookEX UDF: _OL_NavigationFolderGet Example Script", "", 0, "|", "Navigation group|Folder name|Folder path|IsSelected?|IsRemovable?|IsSideBySide?|Position")

_OL_Close($oOutlook)