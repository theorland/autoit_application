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
; Delete a folder, all subfolders and the contained items
; *****************************************************************************
Global $oFolder = _OL_FolderDelete($oOutlook, "*\Outlook-UDF-Test\SourceFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderDelete Example Script", "Error deleting folder 'Outlook-UDF-Test\SourceFolder'. @error = " & @error & ", @extended = " & @extended)
Global $aResult = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test")
$aResult[1].Display
MsgBox(64, "OutlookEX UDF: _OL_FolderDelete Example Script", "Folder 'Outlook-UDF-Test\SourceFolder', all subfolders plus items successfully deleted.")

; *****************************************************************************
; Example 2
; Empty the trash folder.
; The folder itself will not be deleted as it is a system folder
; *****************************************************************************
If MsgBox(36, "OutlookEX UDF: _OL_FolderDelete Example Script", "The trash folder of your mailbox will now be deleted!" & @CRLF & "Do you want the script to continue?") = 7 Then Exit
Global $aTrashFolder = _OL_FolderAccess($oOutlook, "", $olFolderDeletedItems)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderDelete Example Script", "Error accessing trash folder. @error = " & @error & ", @extended = " & @extended)
$oFolder = _OL_FolderDelete($oOutlook, $aTrashFolder[1], 5)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderDelete Example Script", "Error deleting trash folder. @error = " & @error & ", @extended = " & @extended)
$aTrashFolder[1].Display
MsgBox(64, "OutlookEX UDF: _OL_FolderDelete Example Script", "Trash folder successfully deleted.")

_OL_Close($oOutlook)