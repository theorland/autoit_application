#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get free/busy information for the current user.
; Each character = 30 minutes. 0 = free, 1 = all other states
; *****************************************************************************
Global $sCurrentUser = $oOutlook.GetNameSpace("MAPI").CurrentUser.Name
Global $sFreeBusy = _OL_RecipientFreeBusyGet($oOutlook, $sCurrentUser, _NowCalcDate())
If @error <> 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_RecipientFreeBusyGet Example Script", "Error getting free/busy information for current user. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_RecipientFreeBusyGet Example Script", "Free/Busy information for current user starting today. Each char = 30 minutes, completeformat = False:" & @CRLF & $sFreeBusy)

; *****************************************************************************
; Example 2
; Get free/busy information for the current user.
; Each character = 1 hour. Characters according to the OlBusyStatus constants
; *****************************************************************************
$sFreeBusy = _OL_RecipientFreeBusyGet($oOutlook, $sCurrentUser, _NowCalcDate(), 60, True)
If @error <> 0 Then Exit MsgBox(48, "OutlookEX UDF: _OL_RecipientFreeBusyGet Example Script", "Error getting free/busy information for current user. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_RecipientFreeBusyGet Example Script", "Free/Busy information for current user starting today. Each char = 60 minutes, completeformat = True:" & @CRLF & $sFreeBusy)

_OL_Close($oOutlook)