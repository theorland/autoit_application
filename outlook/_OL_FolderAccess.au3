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
; Access a folder of the test environment
; *****************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\TargetFolder\Contacts")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderAccess Example Script", "Error accessing folder 'Outlook-UDF-Test\TargetFolder\Contacts'. @error = " & @error)
Global $aFolderDisplay[6][2] = [[$aFolder[0]],["","Folder object"],["","Default item type for the folder"],["", "StoreID where the folder resides"],["", "EntryID of the folder"],["", "Folder path"]]
$aFolderDisplay[1][0] = $aFolder[1]
$aFolderDisplay[2][0] = $aFolder[2]
$aFolderDisplay[3][0] = $aFolder[3]
$aFolderDisplay[4][0] = $aFolder[4]
$aFolderDisplay[5][0] = $aFolder[5]
_ArrayDisplay($aFolderDisplay, "Folder 'Outlook-UDF-Test\TargetFolder\Contacts' successfully accessed.")

; ***************************************************************************************
; Example 2
; Access the default contacts folder of the current user and display in a separate window
; ***************************************************************************************
$aFolder = _OL_FolderAccess($oOutlook, "", $olFolderContacts)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderAccess Example Script", "Error accessing the default contacts folder. @error = " & @error)
$aFolder[1].Display
MsgBox(64, "OutlookEX UDF: _OL_FolderAccess Example Script", "Default contacts folder successfully accessed and displayed.")

_OL_Close($oOutlook)