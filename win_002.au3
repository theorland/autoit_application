#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>


Opt("MustDeclareVars",1)
Opt("GUIOnEventMode", 1)

Global $Window_hWnd
Global $Window_Text
Global $Window_Log
Wnd_Create()
Global $PROCESS_STATUS = 1

While $PROCESS_STATUS = 1
   ConsoleWrite("Fuck Off"  & @CRLF)
   Sleep(1000)
WEnd



Func Wnd_Create()
   Local $Total_w =@DesktopWidth-160, $Total_h = @DesktopHeight-160
   Local $Top_w = Round($Total_w  * 60 / 100)  - 10 , _
		 $Top_h = Round ($Total_h / 2 ) - 5
   Local $Right_w = Round($Total_w * 40 / 100) - 10 , _
		 $Right_h_1 = Round ($Total_h / 2 ) - 10, _
		 $Right_l_1 = 10 + $Left_w

   Local $Obj_Style =  BitOR($WS_DLGFRAME, $WS_POPUPWINDOW )
   $Window_hWnd = GUICreate("SCAN PST",$Total_w ,$Total_h,-1,-1, $Obj_Style )
   GUISetOnEvent($GUI_EVENT_CLOSE, "Wnd_Evt_Close")

   $Obj_Style = BitOr($GUI_SS_DEFAULT_LABEL,$SS_SIMPLE,$SS_CENTER)
   $Window_Text  = GUICtrlCreateLabel("How do you know what will you do for this ",5,5, $Left_w, $Left_h_1,$Obj_Style)
   GUICtrlSetFont($Window_Text,16,$FW_BOLD)


   $Window_Log = GuiCtrlCreateEdit($Right_l_1,


   GUISetState(@SW_SHOW,$Window_hWnd)

EndFunc

Func Wnd_Evt_Close()
   $PROCESS_STATUS = 0
EndFunc

Func Wnd_Confirm($_message,$_title ="Please Confirm",$timeout =0 )
   Local $confirm = MsgBox($MB_YESNO, $_title , $_message ,$timeout,$Window_hWnd)
   If $IDYES = $confirm Then
	  Return True
   EndIf
	  Return False
EndFunc


