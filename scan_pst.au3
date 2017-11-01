#include <AutoItConstants.au3>
#include <FontConstants.au3>
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <GuiConstants.au3>
#include <File.au3>
#Include <ButtonConstants.au3>
#Include <GUIConstantsEx.au3>
#Include <WinAPIEx.au3>

Global Const $IniFile_PATH =@ScriptDir & "\config\pst.ini"
Global Const $DEFAULT_SCANPST_PATH = "C:\Program Files\Microsoft Office\Office15\SCANPST.EXE"
Global Const $DEFAULT_POWER_CFG = "C:\Windows\System32\powercfg.exe"
Global Const $DEFAULT_POWER_HIGH = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
Global Const $DEFAULT_POWER_BALANCE = "381b4222-f694-41f0-9685-ff5bb260df2e"
Global Const $DEFAULT_PST0 = "C:\Users\ics-user\Document\Outlook Files\Outlook.pst"
Global Const $DEFAULT_BACKUP_TARGET = ""
Global Const $DEFAULT_BACKUP_SERVER = "\\ICSSERVER\OutlookPST"
Global Const $DEFAULT_BACKUP_APP = "C:\Program Files\TeraCopy\teracopy.exe"
Global Const $DEFAULT_BACKUP_LOG= "\\ICSSERVER\OutlookPST\log"
Global Const $DEFAULT_BACKUP_CONF= "\\ICSSERVER\OutlookPST\conf"


Global $WIN_TITLE = "Microsoft Outlook Inbox Repair Tool"

Global $ScanPST_PATH = $DEFAULT_SCANPST_PATH

Global $Power_Cfg = $DEFAULT_POWER_CFG
Global $Power_High = $DEFAULT_POWER_HIGH
Global $Power_Balance = $DEFAULT_POWER_HIGH

Global $Backup_Log = $DEFAULT_BACKUP_LOG
Global $Backup_Server = $DEFAULT_BACKUP_SERVER
Global $Backup_Target= $DEFAULT_BACKUP_TARGET
Global $Backup_App = $DEFAULT_BACKUP_APP
Global $Backup_Conf = $DEFAULT_BACKUP_CONF

Global $Cust_Priority = 4
Global $hLog = FileOpen(@ScriptDir & "\ScanPST_" & @YEAR & "_" & @MON & "_" &  @MDAY & ".log", 1)

Global $DELAY_FORCE = 3000

Global $All_PST[0]

Global $do_shutdown = 0

If ((UBound($CmdLine)>1) AND ($CmdLine[1] == "1")) Then
   $do_shutdown = 1
EndIf

Opt('MustDeclareVars', 1)

Cust_Splash("Initialize Ini File")

IniFile_Load()

Cust_Splash("Change Power Profile to High")

ChangePower_ToHigh()

Process_Close("outlook.exe")

For $file_pst In $all_PST
   Cust_Splash("Currently Fixing PST """ & $file_pst & """")

   Process_Close("scanpst.exe")

   ScanPST_Run($file_pst)
Next

Process_Close("scanpst.exe")

Cust_Splash("Change Power Profile to Normal")

ChangePower_ToNormal()

FileClose($hLog)

SplashOff()

Cust_Splash("Start Backup")

Do_Backup()

Cust_Splash("Shutdown")

if ($do_shutdown == 1 ) Then

   Shutdown ($SD_SHUTDOWN)

EndIf


Func Cust_Splash($message=,$title="ScanPST Process Information",$log = 1)

   SplashTextOn($title, _
	  $message , _
	 -1 ,-1, $DLG_NOTITLE  + $DLG_TEXTVCENTER , -1, -1, "" , 16, $FW_HEAVY  )
   if ($log = 1 ) Then
	  _FileWriteLog($hlog, $message)
   EndIf

EndFunc


Func _WinWaitActivate($title,$text="",$timeout=3)
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


Func Process_Close($process_name = "outlook.exe")

  Cust_Splash("Closing " & $process_name & " Start")

  ProcessClose ( $process_name )
  Sleep($DELAY_FORCE)
  If ProcessExists($process_name) Then
	 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
  EndIf

  Cust_Splash("Closing " & $process_name & " Complete")
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

Func Duplicate_File($file)

   Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
   _PathSplit($file, $sDrive, $sDir, $sFileName, $sExtension)

   Local $newFile =$sDrive  & $sDir & $sFileName & ".scan" & $sExtension
   Cust_Splash("We will backup your file " & @CRLF & $file & @CRLF & "For safety" , "Preparing File Copy")

   If (Not FileCopy($file,$newFile, $FC_OVERWRITE )) Then
	  Cust_Splash("ERROR : Backup Failed " , "Preparing File Copy")
	  return 0
   EndIf

   return $newFile
EndFunc

Func Return_File($file_backup, $file_ori)

   Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
   _PathSplit($file_ori, $sDrive, $sDir, $sFileName, $sExtension)

   Local $tempFile =$sDrive  & $sDir & $sFileName & ".temp" & $sExtension

   Cust_Splash("We will return the fixed file to " & @CRLF & $file_ori & @CRLF & "For safety" , "Returning File Copy")

   If (FileMove($file_backup,$tempFile,$FC_OVERWRITE)=0) Then
	  Cust_Splash("ERROR: Cannot  move to temp file" , "Returning File Copy")
	  return 0
   EndIf
   If (FileMove($file_ori,$file_backup,$FC_OVERWRITE)=0) Then
	  Cust_Splash("ERROR: Cannot move to backup file  " , "Returning File Copy")
	  return 0
   EndIf
   If (FileMove($tempFile,$file_ori,$FC_OVERWRITE)=0) Then
	  Cust_Splash("ERROR: Cannot move to ori file  " , "Returning File Copy")
	  return 0
   EndIf


EndFunc


Func ScanPST_Run($ori_file)
   ; Make sure if the window is closed
   local $hWnd

   Cust_Splash("Opening the application " & $ScanPST_PATH _
	  & @CRLF & " For """ & $ori_file   & """" )


   Local $pst_file = Duplicate_File($ori_file)
   if ($pst_file == 0 ) Then
	    Cust_Splash($pst_file)
	  Return 0

   EndIf

   ShellExecute($ScanPST_PATH)

   $hWnd = _WinWaitActivate($WIN_TITLE)

   ProcessSetPriority($hWnd,$Cust_Priority)

   Send($pst_file)
   Send("!s")
   local $is_run = True
   local $text_process =""

   ; Waiting Scanning Process
   Cust_Splash("Starting Scanning Process")

   While $is_run
	  Sleep(5000)
	  $text_process = GetTextScanPST()
	  Cust_Splash("This is phase 1 scanning process of file" &  @CRLF & _
		 $pst_file &  @CRLF & _
		 " Please be Patient ", "Waiting Scanning Process" , 0 )

	  If $text_process="" Then
		 $is_run = false
		 Cust_Splash("ERROR: Exit Because Error Waiting Scanning Process", "Waiting Scanning Process" )
		 ExitLoop
	  EndIf

	  If StringInStr($text_process , "still", _
		 $STR_NOCASESENSEBASIC )  Then
		 ContinueLoop
	  EndIf

	  If StringInStr ( $text_process , "cancel file scan", _
		 $STR_NOCASESENSEBASIC )  Then
		 Send("!n")
	  EndIf

	  If StringLen($text_process)>40 AND Not StringInStr ( $text_process , "phase", _
		 $STR_NOCASESENSEBASIC ) Then
		 $is_run = false
		 Cust_Splash("Error: Exit Because No Phase " & $text_process, "Waiting Scanning Process")
		 ExitLoop
	  EndIf

   WEnd

   sleep(1000)

   local $is_error = 0
   local $is_done = 0
   Cust_Splash("Decision Phase Repair")
   $text_process = GetTextScanPST()

    If StringInStr($text_process, _
	  "been canceled", $STR_NOCASESENSEBASIC ) Then
	  Cust_Splash("ERROR: User Canceled ","Decision Phase Repair")
	  $is_error = 1
   ElseIf StringInStr($text_process, _
   "error prevented access", $STR_NOCASESENSEBASIC ) Then
	  Cust_Splash("ERROR: Could not open file","Decision Phase Repair")
	  $is_error = 1
   ElseIf StringInStr($text_process, _
   "in use by another", $STR_NOCASESENSEBASIC ) Then
	  Cust_Splash("ERROR: File already in use","Decision Phase Repair")
	  $is_error = 1
   ElseIf  StringInStr($text_process, _
	  "does not exist", $STR_NOCASESENSEBASIC ) Then
	   Cust_Splash("ERROR: File doesn't exists","Decision Phase Repair")
	   $is_error = 1
   ElseIf  StringInStr($text_process, _
   "does not recognize the file", $STR_NOCASESENSEBASIC ) Then
	   Cust_Splash("ERROR: File Type not recognised","Decision Phase Repair")
	   $is_error = 1
   ElseIf  StringInStr($text_process, _
   "error has occurred", $STR_NOCASESENSEBASIC ) Then
	   Cust_Splash("ERROR: An error has occured","Decision Phase Repair")
	   $is_error = 1
   ElseIf  StringInStr($text_process, _
   "is read-only", $STR_NOCASESENSEBASIC ) Then
	   Cust_Splash("ERROR: PST is read only","Decision Phase Repair")
	   $is_error = 1
   ElseIf  StringInStr($text_process, _
   "No errors were found", $STR_NOCASESENSEBASIC ) Then
	   Cust_Splash("SKIP: No Error","Decision Phase Repair")
	   $is_done = 1
   ElseIf  StringInStr($text_process, _
	  "Only minor inconsistencies were found", $STR_NOCASESENSEBASIC ) Then
	  Cust_Splash("SKIP: Only minor","Decision Phase Repair")
	  $is_done = 1
   EndIf

   If $is_error = 1 Then
	  Cust_Splash("TAKE DECISION: IS ERROR ","Take Decision")
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		 Send("{ENTER}")
	  Else
		 Return 0
	  EndIf
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		 Send("!c")
	  Else
		 Return 0
	  EndIf
	  WinWaitClose($hWnd)
	  Return 0
   ElseIf $is_done = 1 Then
	  Cust_Splash("TAKE DECISION: IS DONE","Take Decision")
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If ($hWnd <>0) then
		Send("{ENTER}")
	  Else
		 Return 0
	  EndIf
	  WinWaitClose($hWnd)
	  Return 0
   EndIf

   Local $need_repair
   $text_process = GetTextScanPST()
   $need_repair = 0
   If $text_process<>"" Then
	  If StringInStr($text_process, _
	  "To repair these errors", $STR_NOCASESENSEBASIC ) Then
		 $need_repair = 1
	  EndIf
   Else
	  Cust_Splash("ERROR: No Repair Dialog", "After Take Decision")
	  Return 0
   EndIf


   If $need_repair = 1 Then
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If $hWnd<>0 Then
		 Send("!r")
	  Else
		 Cust_Splash("ERROR: No Repair Dialog", "After Take Decision")
		 Return 0
	  EndIf
   EndIf
   Sleep(1000)

   $text_process = GetTextScanPST()
   If $text_process<>"" And StringInStr($text_process, _
	  "The backup file", $STR_NOCASESENSEBASIC ) Then
	  Cust_Splash("REPAIR: Overwrite previous backup file", "After Take Decision")
	  $hWnd = _WinWaitActivate($WIN_TITLE)
	  If $hWnd<>0 Then
		 Send("!y")
	  Else
		 Return 0
	  EndIf
	EndIf

   Cust_Splash("REPAIR: Start Repairing Process", "After Take Decision")
   Sleep(1000)
   $is_run = true
   While $is_run
	  Sleep(5000)
	  $text_process = GetTextScanPST()

	  If $text_process=""  Then
		 Cust_Splash("ERROR: In Waiting Repair Process" & $text_process)
		 $is_run = False
		 ExitLoop
	  EndIf

	  Cust_Splash( _
		 "REPAIR: This is phase 2 repairing process" &  @CRLF & _
		 $pst_file &  @CRLF & _
		 " Please be Patient " _
		 ,"REPAIR: Waiting Repairing Process", 0 )


	  If StringInStr ( $text_process , "Repair complete" ,$STR_NOCASESENSEBASIC ) Then
		 $is_run = False
		 Cust_Splash(" ")
		 Send("{Enter}")
		 Send("{Space}")
		 ExitLoop
	  EndIf

   WEnd

   $hWnd = _WinWaitActivate($WIN_TITLE)

   ProcessClose($hWnd)

   Cust_Splash("Start Return File from Scanning")

   Return_File($pst_file, $ori_file)

   Cust_Splash("Scan and Repair for """ & $pst_file & """ Complete")

   SplashOff()

EndFunc

#cs

   BACKUP SECTION

#ce
#ce

Func Do_Backup()

   Local  $Backup_File_Conf = $Backup_Conf & "\" & @WDAY  & ".ini"

   Local $is_today = IniRead($Backup_File_Conf,"Schedule",@ComputerName,"none")
   Cust_Splash("Checking Backup For" & @ComputerName  & " In file " & $Backup_File_Conf , "BACKUP PROCESS DECITON" )
   If $is_today = "none" Then
	  Cust_Splash("No Backup Today" , "BACKUP PROCESS STARTED" )
	  Return 0
   EndIf
   ; MsgBox($IDOK,"Debugging", " Load File " & $Backup_File_Conf & " With computer Name " & @ComputerName )

   Cust_Splash("Today is this computer backup" , "BACKUP PROCESS STARTED" )

   ShellExecuteWait($Backup_App,"Copy """ & $Backup_Target & """ """ & $Backup_Server & """ /OverwriteOlder /Close")

   Local $Backup_File_Log =$Backup_Log & "\" & @YEAR & "_" & @MON & "_" & @MDAY & ".ini"

   If IniWrite($Backup_File_Log,"Log",@ComputerName, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC) <> 0 Then
	  Cust_Splash("Success Write Log In Server", "BACKUP PROCESS STARTED")
   Else
	  Cust_Splash("Error Write Log In Server", "BACKUP PROCESS STARTED")
   EndIf
EndFunc



#cs

   INI FILE LOAD SECTION

#ce
#ce

Func IniFile_Load()
   IniFile_ScanPST()
   IniFile_Backup()
   IniFile_Power()
EndFunc


Func IniFile_ScanPST()
  ;scanPST Path
   $ScanPST_PATH = IniRead($IniFile_PATH,"Config", "SCANPST_PATH", _
   $DEFAULT_SCANPST_PATH )
   Local $curr_i, $curr_name, $curr_pst

   $curr_i = 0
   $curr_name = "PST" & $curr_i
   $curr_pst = IniRead($IniFile_PATH,"File",$curr_name,"N/A")
   While $curr_pst <> "N/A"
	  _ArrayAdd( $All_PST, $curr_pst )
	  $curr_i = $curr_i + 1
	  $curr_name = "PST" & $curr_i
	  $curr_pst = IniRead($IniFile_PATH,"File",$curr_name,"N/A")
   WEnd
EndFunc

Func IniFile_Backup()
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

Func IniFile_Power()
   $Power_Cfg = IniRead($IniFile_PATH, "Power", "APP", _
	  $DEFAULT_POWER_CFG)
   $Power_High = IniRead($IniFile_PATH, "Power", "HIGH", _
	  $DEFAULT_POWER_HIGH)
   $Power_Balance = IniRead($IniFile_PATH, "Power", "BALANCE", _
	  $DEFAULT_POWER_BALANCE)
EndFunc
