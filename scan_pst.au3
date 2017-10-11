#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <GuiConstants.au3>
#include <File.au3>


Global Const $IniFile_PATH =@ScriptDir & "\config\scan_pst.ini"
Global Const $DEFAULT_SCANPST_PATH = "C:\Program Files\Microsoft Office\Office15\SCANPST.EXE"
Global Const $DEFAULT_POWER_CFG = "C:\Windows\System32\powercfg.exe"
Global Const $DEFAULT_POWER_HIGH = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
Global Const $DEFAULT_POWER_BALANCE = "381b4222-f694-41f0-9685-ff5bb260df2e"
Global Const $DEFAULT_PST0 = "C:\Users\ics-user\Document\Outlook Files\Outlook.pst"
Global Const $DEFAULT_BACKUP_EMAIL = "name@is-indonesia.com"
Global Const $DEFAULT_BACKUP_SERVER = "\\ICSSERVER\OutlookPST"
Global Const $DEFAULT_BACKUP_APP = "C:\Program Files\TeraCopy\teracopy.exe"

Global $WIN_TITLE = "Microsoft Outlook Inbox Repair Tool"

Global $ScanPST_PATH = $DEFAULT_SCANPST_PATH
Global $Power_Cfg = $DEFAULT_POWER_CFG
Global $Power_High = $DEFAULT_POWER_HIGH
Global $Power_Balance = $DEFAULT_POWER_HIGH
Global $Backup_Server = $DEFAULT_BACKUP_SERVER
Global $Backup_Name = $DEFAULT_BACKUP_EMAIL
Global $Backup_App = $DEFAULT_BACKUP_APP
Global $Cust_Priority = 4
Global $hLog = FileOpen(@ScriptDir & "\ScanPST_" & @YEAR & "_" & @MON & "_" &  @MDAY & ".log", 1)

Global $DELAY_FORCE = 3000

Global $All_PST[0];


IniFile_Load()
  SplashTextOn("Scan PST loading ini file ", _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
Outlook_Close()
ChangePower_ToHigh()

For $file_pst In $all_PST

   _FileWriteLog($hLog,"Fixing PST """ & $file_pst & """")

   ScanPST_Run($file_pst)
Next

ScanPST_Clear()
ChangePower_ToNormal()

FileClose($hLog)
SplashOff()

; Function and Closing

Func IniFile_Load()
   If Not FileExists ($IniFile_PATH) Then
	  IniFile_WriteDefault()
	  MsgBox($MB_TASKMODAL, "Please Edit Config", _
	  "Please edit the pst.ini first," _
	  & @CR & "  and then run this application again" )
   EndIf

 ;scanPST Path
 $ScanPST_PATH = IniRead($IniFile_PATH,"Config", "SCANPST_PATH", _
 $DEFAULT_SCANPST_PATH )

   $curr_i = 0
   $curr_name = "PST" & $curr_i
   $curr_pst = IniRead($IniFile_PATH,"File",$curr_name,"N/A")
   While $curr_pst <> "N/A"
	  _ArrayAdd( $All_PST, $curr_pst )
	  $curr_i = $curr_i + 1
	  $curr_name = "PST" & $curr_i
	  $curr_pst = IniRead($IniFile_PATH,"File",$curr_name,"N/A")
   WEnd


 $Backup_Server = IniRead($IniFile_PATH, "Backup", "SERVER", _
      $DEFAULT_BACKUP_SERVER)
 $Backup_Name = IniRead($IniFile_PATH, "Backup", "NAME", _
      $DEFAULT_BACKUP_EMAIL)
 $Backup_App = IniRead($IniFile_PATH, "Backup", "APP", _
      $DEFAULT_BACKUP_APP)


 $Power_Cfg = IniRead($IniFile_PATH, "Power", "APP", _
      $DEFAULT_POWER_CFG)
 $Power_High = IniRead($IniFile_PATH, "Power", "HIGH", _
      $DEFAULT_POWER_HIGH)
 $Power_Balance = IniRead($IniFile_PATH, "Power", "BALANCE", _
      $DEFAULT_POWER_BALANCE)


EndFunc


Func IniFile_WriteDefault()
   IniWrite($IniFile_PATH, "Config", "SCANPST_PATH", _
      $DEFAULT_SCANPST_PATH);
   IniWrite($IniFile_PATH, "File", "PST0", _
      $DEFAULT_PST0)
   IniWrite($IniFile_PATH, "Backup", "SERVER", _
      $DEFAULT_BACKUP_SERVER)
   IniWrite($IniFile_PATH, "Backup", "NAME", _
      $DEFAULT_BACKUP_EMAIL)
   IniWrite($IniFile_PATH, "Power", "APP", _
      $DEFAULT_POWER_CFG)
   IniWrite($IniFile_PATH, "Power", "HIGH", _
      $DEFAULT_POWER_HIGH)
   IniWrite($IniFile_PATH, "Power", "BALANCE", _
      $DEFAULT_POWER_BALANCE)
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

Func ScanPST_Clear()
   Local $process_name = "scanpst.exe"
  SplashTextOn("Scanning Process", _
  "Now Closing '"  & $process_name & "'" , _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )
  ProcessClose ( $process_name )
  Sleep(500)
  If ProcessExists($process_name) Then
	 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
  EndIf

  SplashTextOn("Scanning Process", _
  'Closing "' & $process_name & '" Completed' , _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY  )

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

Func GetTextScanPST()
   Local $timeout = 1000, $text = "", $title = $WIN_TITLE
   Local $process_name = "scanpst.exe"
   Local $result_text =""
   WinWait($title,$text,$timeout)
   If Not WinActive($title,$text) Then WinActivate($title,$text)

   Local $hWnd = WinWaitActive($title,$text,$timeout)

   if $hWnd <>0 Then
	  $result_text = WinGetText($hWnd)
   EndIf

   If $result_text="" And ProcessExists ( $process_name ) Then
	  return "still"
   EndIf
   return $result_text
EndFunc

Func ScanPST_Run(ByRef $pst_file)

   _FileWriteLog($hLog,"Clear all previous ScanPST")
   ScanPST_Clear()

   ; Make sure if the window is closed
   local $hWnd
   _FileWriteLog($hLog,"Running ScanPST " & $ScanPST_PATH)
   ShellExecute($ScanPST_PATH)

   $hWnd = _WinWaitActivate($WIN_TITLE)
   ProcessSetPriority($hWnd,$Cust_Priority)
   Send($pst_file)
   Send("!S")
   local $is_run = True
   local $text_process =""

   ; Waiting Scanning Process
   _FileWriteLog($hLog,"Starting Scanning Process")
   Sleep(1000)


   While $is_run
	  Sleep(5000)
	  $text_process = GetTextScanPST()
	 SplashTextOn("Waiting Scanning Process", _
		 "This is phase 1 scanning process of file" &  @CRLF & _
		 $pst_file &  @CRLF & _
		 " Please be Patient " _
		  ,-1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY )

	  If $text_process="" Then
		 $is_run = false
		 _FileWriteLog($hLog,"ERROR: Exit Because Error Waiting Scanning Process")
		 ExitLoop
	  EndIf

	  If $text_process = "still" Then
		 ContinueLoop
	  EndIf


	  If StringLen($text_process)>40 AND Not StringInStr ( $text_process , "phase", _
		 $STR_NOCASESENSEBASIC ) Then
		 $is_run = false
		 _FileWriteLog($hLog,"Error: Exit Because No Phase " & $text_process)
		 ExitLoop
	  EndIf

   WEnd

   sleep(1000)

   local $is_error = 0
   local $is_done = 0
   _FileWriteLog($hLog," Decision Phase Repair")
   $text_process = GetTextScanPST()
   If $text_process <>"" Then
	  If StringInStr($text_process, _
		 "been canceled", $STR_NOCASESENSEBASIC ) Then
		 _FileWriteLog($hLog,"ERROR: User Canceled ")
		 $is_error = 1
	  EndIf
	   If StringInStr($text_process, _
	  "error prevented access", $STR_NOCASESENSEBASIC ) Then
		 _FileWriteLog($hLog,"ERROR: Could not open file")
		 $is_error = 1
	  EndIf
	  If StringInStr($text_process, _
	  "in use by another", $STR_NOCASESENSEBASIC ) Then
		 _FileWriteLog($hLog,"ERROR: File already in use")
		 $is_error = 1
	  EndIf
	  If StringInStr($text_process, _
		 "does not exist", $STR_NOCASESENSEBASIC ) Then
		  _FileWriteLog($hLog,"ERROR: File doesn't exists")
		  $is_error = 1
	  EndIf
	  If StringInStr($text_process, _
	  "does not recognize the file", $STR_NOCASESENSEBASIC ) Then
		  _FileWriteLog($hLog,"ERROR: File Type not recognised")
		  $is_error = 1
	   EndIf
	  If StringInStr($text_process, _
	  "error has occurred", $STR_NOCASESENSEBASIC ) Then
		  _FileWriteLog($hLog,"ERROR: An error has occured")
		  $is_error = 1
	   EndIf
	  If StringInStr($text_process, _
	  "is read-only", $STR_NOCASESENSEBASIC ) Then
		  _FileWriteLog($hLog,"ERROR: PST is read only")
		  $is_error = 1
	   EndIf
	  If StringInStr($text_process, _
	  "No errors were found", $STR_NOCASESENSEBASIC ) Then
		  _FileWriteLog($hLog,"SKIP: No Error")
		  $is_done = 1
	   EndIf
	  If StringInStr($text_process, _
		 "Only minor inconsistencies were found", $STR_NOCASESENSEBASIC ) Then
		 _FileWriteLog($hLog,"SKIP: Only minor")
		 $is_done = 1
	  EndIf
   EndIf
   If $is_error = 1 Then
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		 Send("{ENTER}")
	  Else
		 Return 0
	  EndIf
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		 Send("!C")
	  Else
		 Return 0
	  EndIf
	  WinWaitClose($hWnd)
	  Return 0
   ElseIf $is_done = 1 Then
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		Send("{ENTER}")
	  Else
		 Return 0
	  EndIf
	  WinWaitClose($hWnd)
	  Return 0
   EndIf

   $text_process = GetTextScanPST()
   $need_repair = 0
   If $text_process<>"" Then
	  If StringInStr($text_process, _
	  "To repair these errors", $STR_NOCASESENSEBASIC ) Then
		 $need_repair = 1
	  EndIf
   Else
	  _FileWriteLog($hLog,"ERROR: No Repair Dialog")
	  Return 0
   EndIf


   If $need_repair = 1 Then
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If $hWnd<>0 Then
		 Send("!R")
	  Else
		 _FileWriteLog($hLog,"ERROR: No Repair Dialog")
		 Return 0
	  EndIf
   EndIf
   Sleep(1000)

   $text_process = GetTextScanPST()
   If $text_process<>"" And StringInStr($text_process, _
	  "The backup file", $STR_NOCASESENSEBASIC ) Then
	  _FileWriteLog($hLog,"REPAIR: Overwrite previous backup file")
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If $hWnd<>0 Then
		 Send("!Y")
	  Else
		 Return 0
	  EndIf
	EndIf

   _FileWriteLog($hLog,"Start Repairing Process")
   Sleep(1000)
   $is_run = true
   While $is_run
	  Sleep(5000)
	  $text_process = GetTextScanPST()

	  If $text_process=""  Then
		 _FileWriteLog($hLog,"ERROR: In Waiting Repair Process")
		 $is_run = False
		 ExitLoop
	  EndIf

	  SplashTextOn("Waiting Repairing  Process", _
	  "This is phase 2 repairing process" &  @CRLF & _
	  $pst_file &  @CRLF & _
	  " Please be Patient " _
	  , -1 ,-1, $DLG_NOTITLE  +    $DLG_TEXTVCENTER , -1, -1, "" , 20, $FW_HEAVY )

	  If StringInStr ( $text_process , "Repair complete" ,$STR_NOCASESENSEBASIC ) Then
		 $is_run = False
		 _FileWriteLog($hLog,"Repairing complete")
		 Send("{Enter}")
		 Send("{Space}")
		 ExitLoop
	  EndIf

   WEnd

   _FileWriteLog($hLog,"Process Complete")
   SplashOff()


EndFunc


