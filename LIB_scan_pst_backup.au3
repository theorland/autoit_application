Global Const $DEFAULT_BACKUP_TARGET = ""
Global Const $DEFAULT_BACKUP_SERVER = "\\ICSSERVER\OutlookPST"
Global Const $DEFAULT_BACKUP_APP = "C:\Program Files\TeraCopy\teracopy.exe"
Global Const $DEFAULT_BACKUP_LOG= "\\ICSSERVER\OutlookPST\log"
Global Const $DEFAULT_BACKUP_CONF= "\\ICSSERVER\OutlookPST\conf"
#cs

   BACKUP SECTION

#ce
#ce

Global $Backup_Log = $DEFAULT_BACKUP_LOG
Global $Backup_Server = $DEFAULT_BACKUP_SERVER
Global $Backup_Target= $DEFAULT_BACKUP_TARGET
Global $Backup_App = $DEFAULT_BACKUP_APP
Global $Backup_Conf = $DEFAULT_BACKUP_CONF



Func Do_Backup()

   Local  $Backup_File_Conf = $Backup_Conf & "\" & @WDAY  & ".ini"

   Local $is_today = IniRead($Backup_File_Conf,"Schedule",@ComputerName,"none")
   Cust_Splash("Checking Backup For '" & @ComputerName  & "' In file " & $Backup_File_Conf , "BACKUP PROCESS DECITON" )
   If $is_today = "none" Then
	  Cust_Splash("No Backup Today" , "BACKUP PROCESS STARTED" )
	  Return 0
   EndIf
   ; MsgBox($IDOK,"Debugging", " Load File " & $Backup_File_Conf & " With computer Name " & @ComputerName )

   Cust_Splash("Today is this computer backup" , "BACKUP PROCESS STARTED" )

   ShellExecute($Backup_App,"Copy """ & $Backup_Target & """ """ & $Backup_Server & """ /OverwriteOlder /Close")
   Cust_Sleep(1000)
   Local $text = WinGetTitle("[CLASS:TeraCopy3]")
   While (ProcessExists("teracopy.exe") And Not StringInStr($text,"100% done",$STR_NOCASESENSEBASIC))

	  Cust_Sleep(1000)
	  If ($Wnd_Process_Status<>3) Then
		 Return 0
	  EndIf

	  $text = WinGetTitle("[CLASS:TeraCopy3]")
   WEnd

   Local $Backup_File_Log =$Backup_Log & "\" & @YEAR & "_" & @MON & "_" & @MDAY & ".ini"

   If IniWrite($Backup_File_Log,"Log",@ComputerName, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":"  & @MIN & ":" & @SEC) <> 0 Then
	  Cust_Splash("Success Write Log In Server", "BACKUP PROCESS STARTED")
   Else
	  Cust_Splash("Error Write Log In Server", "BACKUP PROCESS STARTED")
   EndIf

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
