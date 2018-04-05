#include <OutlookEX.au3>
#include <MsgBoxConstants.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Calculate size of your Inbox (omitting subfolders)
; *****************************************************************************
Global $vResult = _OL_FolderSize($oOutlook, "*", Default, False, True)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_FolderSize Example Script 1", "Error calculating folder size. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_FolderSize Example Script 1", "Size of your inbox (excluding subfolders) is: " & @CRLF & $vResult & " Bytes (" & Round($vResult / 1024, 2) & " KB).")

; *****************************************************************************
; Example 2
; Calculate size of your Inbox (including all subfolders)
; *****************************************************************************
$vResult = _OL_FolderSize($oOutlook, "*", Default, True, True)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_FolderSize Example Script 2", "Error calculating folder size. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_FolderSize Example Script 2", "Size of your inbox (including subfolders) is: " & @CRLF & $vResult & " Bytes (" & Round($vResult / 1024, 2) & " KB).")

; *****************************************************************************
; Example 3
; Calculate size and number of items of your Inbox (including all subfolders)
; *****************************************************************************
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_FolderSize Example Script 3", "Processing your inbox might take a few minutes!", 5)
$vResult = _OL_FolderSize($oOutlook, "*", Default, True, False)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_FolderSize Example Script 3", "Error calculating folder size. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($vResult, "OutlookEX UDF: _OL_FolderSize Example Script 3", "", 0, Default, "total size & number of items")

; *****************************************************************************
; Example 4
; Calculate size and number of items for a PST (including all subfolders)
; *****************************************************************************
Global $sPath = FileOpenDialog("Please select a PST file to process!", "", "PST files (*.pst)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST))
If @error Then Exit
Global $oPST = _OL_PSTAccess($oOutlook, $sPath)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_FolderSize Example Script 4", "Error accessing PST '" & $sPath & "'. @error = " & @error & ", @extended = " & @extended)
$vResult = _OL_FolderSize($oOutlook, $oPST, Default, True, False)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_FolderSize Example Script 4", "Error calculating folder size. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($vResult, "OutlookEX UDF: _OL_FolderSize Example Script 4", "", 0, Default, "total size & number of items")

_OL_Close($oOutlook)
