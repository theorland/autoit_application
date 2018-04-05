#include <OutlookEX.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all stores available for the current profile
; *****************************************************************************
Global $aResult = _OL_StoreGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_StoreGet Example Script", "Error retrieving list of accounts for the current profile. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_StoreGet Example Script", "", 0, Default, _
"display name|store type|path for a .pst or .ost|cached Exchange store|.pst or .ost|Instant Search enabled|store is open|store id|OOF set|Warning quota|Send quota|Receive quota|Current size|Free space|Max submit size")

_OL_Close($oOutlook)