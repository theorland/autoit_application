#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Get a list of all distribution lists the current user is a member of
; *****************************************************************************
Global $avUser = _OL_ItemRecipientCheck($oOutlook, $oOutlook.Session.CurrentUser.Name) ; Resolve Current User Name
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberOf Example Script", "Error resolving the current user. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_DistListMemberOf($avUser[1][2])
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_DistListMemberOf Example Script", "Error getting distribution lists the current user is a member of. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($Result, "OutlookEX UDF: _OL_DistListMemberOf Example Script", "", 0, "|", "Distribution List Object|Distribution List Name|Distribution List ID")

_OL_Close($oOutlook)