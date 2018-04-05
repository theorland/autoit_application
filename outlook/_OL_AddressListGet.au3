#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all adress lists that are used when resolving an address
; *****************************************************************************
Global $aResult = _OL_AddressListGet($oOutlook)
If @error <> 0 Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_AddressListGet Example Script", "Error " & @error & " when listing address lists!")
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_AddressListGet Example Script - Lists used to resolve an address", "", 0, "|", "OlAddressListType|Name|Index|Resolving order|Identifier")

; *****************************************************************************
; Example 2
; List all adress lists
; *****************************************************************************
$aResult = _OL_AddressListGet($oOutlook, False)
If @error <> 0 Then _
	Exit MsgBox(16, "OutlookEX UDF: _OL_AddressListGet Example Script", "Error " & @error & " when listing all address lists!")
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_AddressListGet Example Script - All address lists", "", 0, "|", "OlAddressListType|Name|Index|Resolving order|Identifier")

_OL_Close($oOutlook)