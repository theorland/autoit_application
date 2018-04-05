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
; Create a single task folder
; *****************************************************************************
Global $oFolder = _OL_FolderCreate($oOutlook, "Test-Folder", $olFolderTasks, "*\Outlook-UDF-Test\SourceFolder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderCreate Example Script", "Error creating folder 'Test-Folder' in folder 'Outlook-UDF-Test\SourceFolder'. @error = " & @error)
$oFolder.Display
MsgBox(64, "OutlookEX UDF: _OL_FolderCreate Example Script", "Folder 'Test-Folder' (Type: tasks) successfully created in folder 'Outlook-UDF-Test\SourceFolder'.")

; *****************************************************************************
; Example 2
; Create a notes folder plus subfolders
; *****************************************************************************
$oFolder = _OL_FolderCreate($oOutlook, "Test-Folder2\Test-Folder3", $olFolderNotes, "*\Outlook-UDF-Test\SourceFolder\Test-Folder")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_FolderCreate Example Script", "Error creating folder 'Test-Folder2\Test-Folder3 in folder 'Outlook-UDF-Test\SourceFolder\Test-Folder'. @error = " & @error)
$oFolder.Display
MsgBox(64, "OutlookEX UDF: _OL_FolderCreate Example Script", "Folder 'Test-Folder2\Test-Folder3' (Type: notes) successfully created in folder 'Outlook-UDF-Test\SourceFolder\Test-Folder'.")

_OL_Close($oOutlook)