#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all adress lists that are used when resolving an address then
; display all members of the last address list (most of the time the GAL is the
; first in the list and takes a lot of time to display).
; *****************************************************************************
Global $aResult = _OL_AddressListGet($oOutlook)
If @error <> 0 Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_AddressListMemberGet Example Script", "Error " & @error & " when listing address lists!")
$aResult = _OL_AddressListMemberGet($oOutlook, $aResult[$aResult[0][0]][2])
If @error <> 0 Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_AddressListMemberGet Example Script", "Error " & @error & " gettings members of first address lists!")
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_AddressListMemberGet Example Script - All members of the first address list", "", 0, "|", "EMail address|Name|OlAddresEntryUserType|Identifier|Object of the address entry")

_OL_Close($oOutlook)