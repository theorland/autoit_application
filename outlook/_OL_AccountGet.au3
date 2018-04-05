#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all accounts available for the current profile
; *****************************************************************************
Global $aResult = _OL_AccountGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_AccountGet Example Script", "Error getting list of accounts for the current profile. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_AccountGet Example Script", "", 0, "|", "|AccountType|Displayname|SMTPAddress|Username|Account object|AutoDiscoverConnectionMode|ExchangeConnectionMode|ExchangeMailboxServerName|ExchangeMailboxServerVersion")

_OL_Close($oOutlook)