#cs ----------------------------------------------------
   Created by : theo (lourenzoisthebest@gmail.com)
   Created Time : 2018-04-05
   Archive pst

--------------------------------------------------------



#ce ----------------------------------------------------

#pragma compile(Out, ..\archive_pst.exe)
#pragma compile(Icon, .\archive_pst.ico)
; #pragma compile(ExecLevel, highestavailable)
#pragma compile(Compatibility, win7)
#pragma compile(Compression, 9)
#pragma compile(inputboxres, false)
#pragma compile(UPX, False)
#pragma compile(FileDescription, "Archive PST - Automate Archive Process"  )
#pragma compile(ProductName, "archive_pst")
#pragma compile(ProductVersion, 1.3)
#pragma compile(FileVersion, "1.3.0.0" ) ; The last parameter is optional.
#pragma compile(LegalCopyright, "© Theo Christ" )
#pragma compile(LegalTrademarks, 'Trademark something, and some text in "quotes" etc...')
#pragma compile(CompanyName, 'home')

#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/SO



#include ".\outlook\OutlookEx.au3"
#include <Array.au3>

Opt('MustDeclareVars', 1)




Global $EXCLUDED_FOLDER[] = ['Deleted Items' , 'Conversation Action Settings', 'Calendar', _
 'Tasks', 'RSS Feeds', 'Contacts', 'Notes', 'Journal',  'Quick Step Settings']


Global $OL_MAIN = _OL_Open()
Global Const $IniFile_PATH =@ScriptDir & "\config\pst.ini"

FileDelete(@ScriptDir & "\Outlook_Debug.txt")

$__iOL_Debug = 3



Global $folder_archive = _OL_PSTAccess($OL_MAIN,"D:\ics-user\Documents\Outlook Files\theo@is-indonesia.com.2017.pst")

#cs
   ;Example 1 Works
Local $aResult = _OL_FolderTree($OL_MAIN, "*")
_ArrayDisplay($aResult, "Shows my folder structure")

;Example 2 Works
$aResult = _OL_FolderTree($OL_MAIN, "theo@is-indonesia.com.current\")
_ArrayDisplay($aResult, "Shows my folder structure")
#ce

;Example 3 fails

$info = _OL_PSTGet($OL_MAIN)
_ArrayDisplay($info,"this is get")
Global $EXCEPT_FOLDER = []
Local $aResult = _OL_FolderTree($OL_MAIN, "*")
_ArrayDisplay($aResult, "folder main ")

$aResult = _OL_FolderTree($OL_MAIN, $folder_archive)
_ArrayDisplay($aResult, "folder archive")



