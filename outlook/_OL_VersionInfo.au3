#include <OutlookEX.au3>

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 1 - Display the version info for this UDF
;------------------------------------------------------------------------------------------------------------------------------------------------
Global $aVersionInfo = _OL_VersionInfo()
Global $aResult[9][2] = [[8,2],["Release type", $aVersionInfo[1]],["Major version", $aVersionInfo[2]],["Minor version", $aVersionInfo[3]], _
	["Sub version", $aVersionInfo[4]],["Release date", $aVersionInfo[5]],["AutoIt version required", $aVersionInfo[6]],["Authors", $aVersionInfo[7]], _
	["Contributors", $aVersionInfo[8]]]
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_VersionInfo - Version Info for the UDF")