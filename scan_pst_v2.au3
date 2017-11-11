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
#Include <StringConstants.au3>
#Include <LIB_scan_pst_process.au3>
#Include <LIB_scan_pst_backup.au3>

Opt('MustDeclareVars', 1)

Global Const $IniFile_PATH =@ScriptDir & "\config\pst.ini"
Global Const $DEFAULT_SCANPST_PATH = "C:\Program Files\Microsoft Office\Office15\SCANPST.EXE"
Global Const $DEFAULT_PST0 = "C:\Users\ics-user\Document\Outlook Files\Outlook.pst"

Global Const $WIN_TITLE = "Microsoft Outlook Inbox Repair Tool"

Global $ScanPST_PATH = $DEFAULT_SCANPST_PATH
Global $Cust_Priority = $PROCESS_HIGH
Global $Do_Default
Global $Conf_Safety = 1

Global $hLog = FileOpen(@ScriptDir & "\ScanPST_" & @YEAR & "_" & @MON & "_" &  @MDAY & ".log", 1)

Global $DELAY_FORCE = 3000

Global $All_PST[0]
Global $All_Exec[0]

Global $do_shutdown = 0

If ((UBound($CmdLine)>1) AND ($CmdLine[1] == "1")) Then
   $do_shutdown = 1
EndIf

#cs
   MAIN PROCESS START HERE
#ce

Wnd_Create("ScanPST Process Information")

Cust_Splash("Initialize Ini File")

IniFile_Load()

IniWrite($Backup_Log_File,"START", @ComputerName, _
	   @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC)

Wnd_Init()

Cust_Splash("Change Power Profile to High")

ChangePower_ToHigh()

Cust_Process_Close("outlook.exe")

For $file_pst In $all_PST
   If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
	  ExitLoop
   EndIf
   Cust_Splash("Currently Fixing PST """ & $file_pst & """")

   Cust_Process_Close("scanpst.exe")

   ScanPST_Run($file_pst)
Next

Switch $Wnd_Process_Status
Case 0
IniWrite($Backup_Log_File,"PROCESS", @ComputerName, _
	  "STOPPED "  & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC)
Case 1
IniWrite($Backup_Log_File,"PROCESS", @ComputerName, _
	  "RUNNING "  & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC)
Case 2
IniWrite($Backup_Log_File,"PROCESS", @ComputerName, _
	  "SHUTDOWN "  & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC)
EndSwitch

Cust_Process_Close("scanpst.exe")

Cust_Splash("Change Power Profile to Normal")

ChangePower_ToNormal()

Cust_Splash("Start Backup")

Do_Backup()

Cust_Splash("Exec Post Scan PST")

For $file_exec in $All_Exec
   If $file_exec<>"" Then
	  RunWait($file_exec)
   EndIf
Next



FileClose($hLog)

if (($do_shutdown == 1) OR _
   ($Wnd_Process_Status==$Wnd_Process_Status_VALUE_SHUT)) Then

   Cust_Splash("Shutdown")
   Shutdown ($SD_SHUTDOWN )

EndIf

#cs
   END OF MAIN PROCESS
#ce

Func Wnd_Init()
   Switch $Do_Default
   Case 1
	  GuiCtrlSetState($Wnd_GUI_RB_Run, $GUI_CHECKED)
	  $Wnd_Process_Status = $Wnd_Process_Status_VALUE_RUN
   Case 2
	  GuiCtrlSetState($Wnd_GUI_RB_Shut, $GUI_CHECKED)
	  $Wnd_Process_Status = $Wnd_Process_Status_VALUE_SHUT
   Case 3
	  GuiCtrlSetState($Wnd_GUI_RB_Stop, $GUI_CHECKED)
	  $Wnd_Process_Status = $Wnd_Process_Status_VALUE_STOP
   EndSwitch
EndFunc



Func ScanPST_Run($ori_file)
   ; Make sure if the window is closed
   local $hWnd

   Cust_Splash("Opening the application " & $ScanPST_PATH _
	  & @CRLF & " For """ & $ori_file   & """" )

   Run($ScanPST_PATH)

   Local $pst_file = Duplicate_File($ori_file)
   if ($pst_file == 0 ) Then
	  Cust_Splash("Error Backuping " )
	  Return 0
   EndIf

   $hWnd = _WinWaitActivate($WIN_TITLE)

   ProcessSetPriority($hWnd,$Cust_Priority)

   If $hWnd=0 Then
	  Run($ScanPST_PATH)
   EndIf

   Send($pst_file)
   Send("!s")
   local $is_run = True
   local $text_process =""

   _WinWaitActivate($WIN_TITLE)

   ; Waiting Scanning Process
   Cust_Splash("Starting Scanning Process")

   While $is_run

	  Cust_Sleep($DELAY_FORCE)

	  If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
		 Return 0
	  EndIf

	  $text_process = GetText_ScanPST($WIN_TITLE)
	  Cust_Splash("This is phase 1 scanning process of file" &  @CRLF & _
		 $pst_file &  @CRLF & _
		 " Please be Patient ", "Waiting Scanning Process" , 0 )

	  If StringLen($text_process)>40 AND Not StringInStr ( $text_process , "phase", _
		 $STR_NOCASESENSEBASIC ) Then
		 $is_run = false
		 Cust_Splash("Error: Exit Because No Phase " & $text_process, "Waiting Scanning Process")
		 ExitLoop
	  ElseIf StringInStr ( $text_process , "cancel file scan", _
		 $STR_NOCASESENSEBASIC )  Then
		 Send("!n")
	  ElseIf $text_process == "#run#"  Then
		 ContinueLoop
  	  ElseIf $text_process == "#dead#" Then
		 Return 0
	  EndIf


   WEnd

   Cust_Sleep(1000)

   local $is_error = 0
   local $is_done = 0
   Cust_Splash("Decision Phase Repair")
   $text_process = GetText_ScanPST($WIN_TITLE)
   While $text_process == "#run#"
	  Cust_Sleep(1000)
	  If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
		 Return 0
	  EndIf
	  $text_process = GetText_ScanPST($WIN_TITLE)
   WEnd

   If $text_process == "#dead#" Then
	  Return 0
   ElseIf StringInStr($text_process, _
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
	  Cust_Splash("TAKE DECISION: IS ERROR JUST QUIT","Take Decision")
	  Return 0
   ElseIf $is_done = 1 Then
	  Cust_Splash("TAKE DECISION: IS DONE JUST QUIT","Take Decision")
	  Return 0
   EndIf

   Local $need_repair = 0
   $text_process = GetText_ScanPST($WIN_TITLE)
   While $text_process == "#run#"
	  Cust_Sleep(1000)
	  If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
		 Return 0
	  EndIf
	  $text_process = GetText_ScanPST($WIN_TITLE)
   WEnd

   If $text_process == "#dead#" Then
	  Cust_Splash("ERROR: No Repair Dialog", "After Take Decision")
	  Return 0
   ElseIf StringInStr($text_process, _
	  "To repair these errors", $STR_NOCASESENSEBASIC ) Then
		 $need_repair = 1
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
   Cust_Sleep(1000)

   $text_process = GetText_ScanPST($WIN_TITLE)
   While $text_process == "#run#"
	  $text_process = GetText_ScanPST($WIN_TITLE)
   WEnd
   If $text_process == "#dead#" Then
	  Return 0
   ElseIf StringInStr($text_process, _
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
   Cust_Sleep(1000)
   $is_run = true
   While $is_run
 	  Cust_Splash( _
		 "REPAIR: This is phase 2 repairing process" &  @CRLF & _
		 $pst_file &  @CRLF & _
		 " Please be Patient " _
		 ,"REPAIR: Waiting Repairing Process", 0 )
	  Wnd_Sleep(5000)
	  If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
		 Return 0
	  EndIf
	  $text_process = GetText_ScanPST($WIN_TITLE)

	  If $text_process=="#dead#"  Then
		 Cust_Splash("ERROR: In Waiting Repair Process" & $text_process)
		 $is_run = False
		 ExitLoop
	  ElseIf $text_process ="#run#" Then
		 ContinueLoop
	  ElseIf StringInStr ( $text_process , "Repair complete" ,$STR_NOCASESENSEBASIC ) Then
		 $is_run = False
		 Cust_Splash("Close Reparing Window")
		 Send("{Enter}")
		 Send("{Space}")
		 ExitLoop
	  Else

	  EndIf
   WEnd

   Cust_Splash("Start Return File from Scanning")
   Return_File($pst_file, $ori_file)
   Cust_Splash("Scan and Repair for """ & $pst_file & """ Complete")

EndFunc

Func GetText_ScanPST($WIN_TITLE)

   Local $timeout = 1000, $text = "", $title = $WIN_TITLE
   Local $process_name = "scanpst.exe"
   Local $result_text =""

   WinWait($WIN_TITLE,$text,$timeout)
   If Not WinActive($WIN_TITLE,$text) Then WinActivate($WIN_TITLE,$text)
   Local $hWnd = WinWaitActive($WIN_TITLE,$text,$timeout)

   if $hWnd <>0 Then
	  $result_text = WinGetText($hWnd)
   EndIf

   If $result_text=="" Then
	  If ProcessExists( $process_name ) Then
		 return "#run#"
	  Else
		 Return "#dead#"
	  EndIf
   EndIf
   return $result_text
EndFunc


Func Duplicate_File($file)

   if $Conf_Safety == 0 Then
	  return $file
   EndIf

   Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
   _PathSplit($file, $sDrive, $sDir, $sFileName, $sExtension)

   Local $tempDir =$sDrive  & $sDir  & "temp\"
   Local $tempFile = $tempDir & $sFilename & $sExtension
   DirCreate ($tempDir)

   Cust_Splash("We will backup your file " & @CRLF & $file & @CRLF & "For safety" , "Preparing File Copy")

   FileDelete($tempFile)
   ShellExecute($Backup_App,"Copy """ & $file & """ """ & $tempDir & """ /OverwriteOlder /Close")

   Local $text = WinGetTitle("[CLASS:TeraCopy3]");
   Cust_Sleep(1000)
   Local $text = WinGetTitle("[CLASS:TeraCopy3]");
   If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
	  Return 0
   EndIf

   While (ProcessExists("teracopy.exe") And Not StringInStr($text,"100% done",$STR_NOCASESENSEBASIC))
	  Cust_Sleep(1000)
	  If ($Wnd_Process_Status<>$Wnd_Process_Status_VALUE_RUN) Then
		 Return 0
	  EndIf
	  $text = WinGetTitle("[CLASS:TeraCopy3]")
   WEnd
   Cust_Sleep(1000)


   Local $newFile =$sDrive  & $sDir & $sFileName & ".scan" & $sExtension
   If FileExists($tempFile) Then
	  If Not FileMove($tempFile,$newFile,$FC_OVERWRITE) Then
		 Cust_Splash("ERROR : Cannot Move The File " , "Backup Process")
		 Return 0
	  EndIf
	  return $newFile

   EndIf
   Cust_Splash("ERROR : File Backup not Exists '" & $tempFile & "'" , "Backup Process")
   Return 0

EndFunc


Func Return_File($file_backup, $file_ori)

   If $Conf_Safety == 0 Then
	  return $file_ori
   EndIf
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

   $curr_i = 0
   $curr_name = "EXEC" & $curr_i
   $curr_pst = IniRead($IniFile_PATH,"Config",$curr_name,"N/A")
   While $curr_pst <> "N/A"
	  _ArrayAdd( $All_Exec, $curr_pst )
	  $curr_i = $curr_i + 1
	  $curr_name = "EXEC" & $curr_i
	  $curr_pst = IniRead($IniFile_PATH,"Config",$curr_name,"N/A")
   WEnd

   $Do_Default = IniRead($IniFile_Path,"Config","Default",1 )
   Cust_Splash("Default is " & $Do_Default )
   $Conf_Safety = IniRead($IniFile_Path,"Config","Safety",1 )

EndFunc

Func Feedback_Save($text_feedback)

   $text_feedback = StringStripWS ($text_feedback, _
	  BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING, $STR_STRIPSPACES))

   If $text_feedback=="" Then
	  Return 0
   EndIf

   Local $filename_feedback= $Backup_Log_Dir& "\feedback_" & @YEAR & "_" & @MON & "_" &  @MDAY & ".log"
   Local $f_Wnd = FileOpen($filename_feedback,$FO_APPEND)
   FileWriteLine($f_Wnd,"[[[["  & @ComputerName & "]]]] " & @YEAR & "-" & @MON & "-" &  @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC)
   FileWriteLine($f_Wnd,$text_feedback)
   FileClose($f_Wnd)

EndFunc


