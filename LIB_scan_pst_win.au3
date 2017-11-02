#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>
#include <ButtonConstants.au3>
#include <ColorConstants.au3>
#include <AutoItConstants.au3>

Opt("MustDeclareVars",1)
Opt("GUIOnEventMode", 1)

Global $Wnd_GUI_hWnd
Global $Wnd_GUI_Title
Global $Wnd_GUI_Text
Global $Wnd_GUI_RB_Run
Global $Wnd_GUI_RB_Shut
Global $Wnd_GUI_RB_Stop

Global $Wnd_Process_Status = 1

#cs
Wnd_Create()
Global $Wnd_Process_Status = 1
While $Wnd_Process_Status =1
   Wnd_Sleep(2000)
WEnd

#ce

Func Wnd_Sleep($timeout)
   GUISetState(@SW_SHOW,$Wnd_GUI_hWnd)
   WinSetOnTop($Wnd_GUI_hWnd,"",$WINDOWS_ONTOP)
   Sleep($timeout)
   WinSetOnTop($Wnd_GUI_hWnd,"",$WINDOWS_NOONTOP)
EndFunc


Func Wnd_Create($title="Normal Title")
   Local $Total_w =@DesktopWidth-160, $Total_h = @DesktopHeight-160, _
		 $Total_l =5, $Total_t = 5
   Local $Title_t = $Total_t, _
		 $Title_h = 40
   Local $Inside_w = $Total_w  - 10 , _
		 $Text_h = $Total_h - 100 - 40 - 10 , _
		 $Text_t = $Title_h + 10
   Local $Bottom_w = Round(($Total_w -10) / 3) - 10 , _
		 $Bottom_h = 40
   Local $Bottom_1_t = $Total_t + $Title_h + 5+  $Text_h + 10 , _
		 $Bottom_2_t = $Total_t + $Title_h + 5+  $Text_h + 10  + $Bottom_h , _
		 $Bottom_l_1 = $Total_l, _
		 $Bottom_l_2 = $Total_l + $Bottom_w +10, _
		 $Bottom_l_3 = $Total_l + ($Bottom_w +10) * 2


   Local $Obj_Style =  BitOR($WS_DLGFRAME, $WS_POPUPWINDOW )
   $Wnd_GUI_hWnd = GUICreate($title,$Total_w ,$Total_h,-1,-1, $Obj_Style )
   GUISetOnEvent($GUI_EVENT_CLOSE, "Wnd_Evt_Close")

   $Obj_Style = $GUI_SS_DEFAULT_LABEL+$SS_CENTER
   $Wnd_GUI_Title  = GUICtrlCreateLabel("Title", _
		 $Total_l,$Title_t, $Inside_w, $Title_h ,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_Title,30,$FW_HEAVY, $GUI_FONTUNDER)

   $Obj_Style = $GUI_SS_DEFAULT_LABEL+$SS_CENTER
   $Wnd_GUI_Text  = GUICtrlCreateLabel("How do you know what will you do for this " & @CRLF & " N "  & @CR & " N " _
   & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " _
   & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " , _
		 $Total_l,$Text_t, $Inside_w, $Text_h ,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_Text,20,$FW_EXTRABOLD)

   $Obj_Style = $GUI_SS_DEFAULT_LABEL+$SS_CENTER
   Local $temp_label  = GUICtrlCreateLabel("Do you want to ? (Just press and wait for the window closed)", _
	  $Total_l, $Bottom_1_t,$Inside_w, $Bottom_h ,$Obj_Style)
   GUICtrlSetFont($temp_label,20,$FW_BOLD)

   $Obj_Style = BitOr($BS_AUTORADIOBUTTON,$BS_PUSHLIKE)
   $Wnd_GUI_RB_Run = GUICtrlCreateRadio("Continue Scan Process", _
	   $Bottom_l_1,$Bottom_2_t,$Bottom_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Run,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Run,$COLOR_GREEN)

   $Wnd_GUI_RB_Shut = GUICtrlCreateRadio("Just turn off", _
	  $Bottom_l_2,$Bottom_2_t, $Bottom_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Shut,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Shut,$COLOR_BLACK)

   $Wnd_GUI_RB_Stop = GUICtrlCreateRadio("I want to working now", _
	  $Bottom_l_3,$Bottom_2_t, $Bottom_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Stop,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Stop,$COLOR_RED)

   GuiCtrlSetState($Wnd_GUI_RB_Run, $GUI_CHECKED)

   GUICtrlSetOnEvent($Wnd_GUI_RB_Run, "Wnd_Evt_Radio")
   GUICtrlSetOnEvent($Wnd_GUI_RB_Shut, "Wnd_Evt_Radio")
   GUICtrlSetOnEvent($Wnd_GUI_RB_Stop, "Wnd_Evt_Radio")
   GUISetState(@SW_SHOW,$Wnd_GUI_hWnd)
EndFunc

Func Wnd_Evt_Radio()
   If (BitAND(GUICtrlRead($Wnd_GUI_RB_Run), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status = 1
   ElseIf (BitAND(GUICtrlRead($Wnd_GUI_RB_Shut), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status  = 2
   ElseIf (BitAND(GUICtrlRead($Wnd_GUI_RB_Stop), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status  = 0
   EndIf
EndFunc

Func Wnd_Evt_Close()

   $Wnd_Process_Status = 0
EndFunc


Func Wnd_Confirm($_message,$_title ="Please Confirm",$timeout =1000 )
   Local $confirm = MsgBox($MB_YESNO, $_title , $_message ,$timeout,$Wnd_GUI_hWnd)
   If $IDYES = $confirm Then
	  Return True
   EndIf
	  Return False
EndFunc


