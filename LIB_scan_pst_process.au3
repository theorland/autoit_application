#Include <LIB_scan_pst_win.au3>

Global Const $DEFAULT_POWER_CFG = "C:\Windows\System32\powercfg.exe"
Global Const $DEFAULT_POWER_HIGH = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
Global Const $DEFAULT_POWER_BALANCE = "381b4222-f694-41f0-9685-ff5bb260df2e"


Global $Power_Cfg = $DEFAULT_POWER_CFG
Global $Power_High = $DEFAULT_POWER_HIGH
Global $Power_Balance = $DEFAULT_POWER_HIGH

Opt("MustDeclareVars",1)

Global $User_Choose_Sleep_Info
Func Cust_Sleep($sleep)

   Wnd_Sleep($sleep)
   If $User_Choose_Sleep_Info <> $Wnd_Process_Status Then
	  $User_Choose_Sleep_Info = $Wnd_Process_Status
	  If $RUN_WHEN_SHUTDOWN = 1 Then
		 Switch $Wnd_Process_Status
		 Case $Wnd_Process_Status_VALUE_STOP
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "END THE TASK")
			Cust_Process_Close("scanpst.exe")
			Exit(0)
		 Case $Wnd_Process_Status_VALUE_SHUT
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "ONLY SHUTDOWN")
		 Case $Wnd_Process_Status_VALUE_RUN
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "SCAN PROCESS AND SHUT ")
		 case
		 EndSwitch
	  Else
		 Switch $Wnd_Process_Status
		 Case $Wnd_Process_Status_VALUE_STOP
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "END THE TASK")
			Cust_Process_Close("scanpst.exe")
			Exit(0)
		 Case $Wnd_Process_Status_VALUE_SHUT
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "SCAN PROCESS AND OPEN OUTLOOK")
		 Case $Wnd_Process_Status_VALUE_RUN
			Cust_Splash("[USER CHOOSE]" & @CRLF &  @CRLF & "SCAN PROCESS")
		 case
		 EndSwitch
	  EndIf
   EndIf
EndFunc

Func Cust_Splash($message,$title="ScanPST Information",$log = 1)

   GUICtrlSetData($Wnd_GUI_Title,$title)
   GUICtrlSetData($Wnd_GUI_Text,$message)
   if ($log = 1 ) Then
	  _FileWriteLog($hlog, $message)
   EndIf

EndFunc


Func _WinWaitActivate($title,$text="",$timeout=500)
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


Func Cust_Process_Close($process_name = "outlook.exe")

  Cust_Splash("Closing Process " & $process_name )
  ProcessClose ( $process_name )
  Sleep($DELAY_FORCE)
  If ProcessExists($process_name) Then
	 ShellExecuteWait("taskkill.exe", '/F /IM "' & $process_name & '"')
  EndIf

EndFunc

Func IniFile_Power()
   $Power_Cfg = IniRead($IniFile_PATH, "Power", "APP", _
	  $DEFAULT_POWER_CFG)
   $Power_High = IniRead($IniFile_PATH, "Power", "HIGH", _
	  $DEFAULT_POWER_HIGH)
   $Power_Balance = IniRead($IniFile_PATH, "Power", "BALANCE", _
	  $DEFAULT_POWER_BALANCE)
EndFunc