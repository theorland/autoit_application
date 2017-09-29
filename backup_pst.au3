#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <GuiConstants.au3>
#include <File.au3>

#include <Date.au3>


Global Const $IniFile_PATH =@ScriptDir & "\config\scan_pst.ini"
Global Const $DEFAULT_SCANPST_PATH = "C:\Program Files\Microsoft Office\Office15\SCANPST.EXE"
Global Const $DEFAULT_POWER_CFG = "C:\Windows\System32\powercfg.exe"
Global Const $DEFAULT_POWER_HIGH = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
Global Const $DEFAULT_POWER_BALANCE = "381b4222-f694-41f0-9685-ff5bb260df2e"
Global Const $DEFAULT_PST0 = "C:\Users\ics-user\Document\Outlook Files\Outlook.pst"
Global Const $DEFAULT_BACKUP_TARGET = "theo@is-indonesia.com"
Global Const $DEFAULT_BACKUP_SERVER = "\\ICSSERVER\OutlookPST"
Global Const $DEFAULT_BACKUP_APP = "C:\Program Files\TeraCopy\teracopy.exe"
Global Const $DEFAULT_BACKUP_LOG= "\\ICSSERVER\OutlookPST\log"
Global Const $DEFAULT_BACKUP_CONF= "\\ICSSERVER\OutlookPST\conf"


Global $WIN_TITLE = "Microsoft Outlook Inbox Repair Tool"

Global $Backup_Log = $DEFAULT_BACKUP_LOG
Global $Backup_Server = $DEFAULT_BACKUP_SERVER
Global $Backup_Target= $DEFAULT_BACKUP_TARGET
Global $Backup_App = $DEFAULT_BACKUP_APP
Global $Backup_Conf = $DEFAULT_BACKUP_CONF
Global $Cust_Priority = 4
Global $hLog = FileOpen(@ScriptDir & "\ScanPST_" & @YEAR & "_" & @MON & "_" &  @MDAY & ".log", 1)

Global $DELAY_FORCE = 3000

Global $All_PST[0];



IniFile_Backup_Load()

Outlook_Close()

DoBackup()

; Function and Closing

Func IniFile_Backup_Load()
   If Not FileExists ($IniFile_PATH) Then
	  IniFile_WriteDefault()
	  MsgBox($MB_TASKMODAL, "Please Edit Config", _
	  "Please edit the pst.ini first," _
	  & @CR & "  and then run this application again" )
   EndIf


; backup
 $Backup_Server = IniRead($IniFile_PATH, "Backup", "SERVER", _
      $DEFAULT_BACKUP_SERVER)
 $Backup_Target = IniRead($IniFile_PATH, "Backup", "TARGET", _
      $DEFAULT_BACKUP_TARGET)
 $Backup_Log = IniRead($IniFile_PATH, "Backup", "LOG", _
      $DEFAULT_BACKUP_LOG)
 $Backup_Conf = IniRead($IniFile_PATH, "Backup", "CONF", _
      $DEFAULT_BACKUP_CONF)
 $Backup_App = IniRead($IniFile_PATH, "Backup", "APP", _
      $DEFAULT_BACKUP_APP)
EndFunc


Func _WinWaitActivate($title,$text="",$timeout=250)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	Return WinWaitActive($title,$text,$timeout)
EndFunc

Func ChangePower_ToHigh()
	RunWait($Power_Cfg & " /setactive " & $Power_High)
EndFunc

Func ChangePower_ToNormal()
	RunWait($Power_Cfg & " /setactive " & $Power_Balance)
EndFunc

Func Outlook_Close()
  Local $process_name = "outlook.exe"
  SplashTextOn("Scanning Process", _
  "Now Closing '"  & $process_name & "'" , _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
  ProcessClose ( $process_name )
  Sleep($DELAY_FORCE)
  If ProcessExists($process_name) Then
	 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
  EndIf

  SplashTextOn("Scanning Process", _
  'Closing "' & $process_name & '" Completed' , _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
EndFunc

Func DoBackup()
   $Backup_File_Conf = $Backup_Conf & "\" & @WDAY  & ".ini"

   $is_today = IniRead($Backup_File_Conf,"Schedule",@ComputerName,"none")
   If $is_today = "none" Then
	  Return 0
   EndIf
   MsgBox($IDOK,"Debugging", " Load File " & $Backup_File_Conf & " With computer Name " & @ComputerName )
   _FileWriteLog($hLog,"Today is this computer backup")

   ShellExecuteWait($Backup_App,"Copy """ & $Backup_Target & """ """ & $Backup_Server & """ /OverwriteOlder /Close")

   $Backup_File_Log =$Backup_Log & "\" & @YEAR & "_" & @MON & "_" & @MDAY &".ini"
   If IniWrite($Backup_File_Log,"Log",@ComputerName, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC) <> 0 Then
	  _FileWriteLog($hLog,"Success Write Log In Server")
   Else
	  _FileWriteLog($hLog,"Error Write Log In Server")
   EndIf
EndFunc

