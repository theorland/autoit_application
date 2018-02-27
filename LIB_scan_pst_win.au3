#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <EditConstants.au3>
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
Global $Wnd_GUI_Feedback

Global Const $Wnd_Process_Status_VALUE_STOP = 0
Global Const $Wnd_Process_Status_VALUE_RUN = 1
Global Const $Wnd_Process_Status_VALUE_SHUT = 2

Global $Wnd_Process_Status = 1

#cs
	Wnd_Create()
	Wnd_Create_Not_Shutdown();
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

Func Wnd_Create_Not_Shutdown()
	GUICtrlSetData($Wnd_GUI_RB_Shut,"Run n End with Outlook")
EndFunc

Func Wnd_Create($title="Normal Title")
   Local $Total_w =@DesktopWidth-160, $Total_h = @DesktopHeight-160, _
		 $Total_l =5, $Total_t = 5 , $Inside_w = $Total_w  - 10
   Local $Title_t = $Total_t, _
		 $Title_h = 50

   Local $feedback_w = 300, $feedback_h = 400

   Local $Bottom_2_w = Round(($Total_w -10) / 3) - 10 , _
		 $Bottom_h = 40

   Local $Text_w = $Inside_w  ,   _
		 $Text_h = $Total_h - ($Bottom_h +10)*2 - ($Title_h +10) , _
		 $Text_t = $Title_h + 10


   #cs
   Local $feedback_l = $title_feed_l, $feedback_t = $title_feed_t + $title_feed_h + 5
   #ce

   Local $Bottom_1_t = $Total_t + $Title_h +  $Text_h + 10 , _
		 $Bottom_1_w = $Inside_w  - $feedback_w - 10, _
		 $Bottom_2_t = $Bottom_1_t + $Bottom_h +5, _
		 $Bottom_l_1 = $Total_l, _
		 $Bottom_l_2 = $Total_l + $Bottom_2_w +10, _
		 $Bottom_l_3 = $Total_l + ($Bottom_2_w +10) * 2

   Local $title_feed_l = $Inside_w -$feedback_w ,  $title_feed_t = $Bottom_1_t, _
		 $title_feed_w = $feedback_w, $title_feed_h =20

   Local $Obj_Style =  BitOR($WS_DLGFRAME, $WS_POPUPWINDOW )
   $Wnd_GUI_hWnd = GUICreate($title,$Total_w ,$Total_h,-1,-1, $Obj_Style )
   GUISetOnEvent($GUI_EVENT_CLOSE, "Wnd_Evt_Close")

   $Obj_Style = BitOR($GUI_SS_DEFAULT_LABEL,$SS_CENTER)
   $Wnd_GUI_Title  = GUICtrlCreateLabel("Title", _
		 $Total_l,$Title_t, $Inside_w, $Title_h ,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_Title,30,$FW_HEAVY, $GUI_FONTUNDER)

   $Obj_Style = BitOR($GUI_SS_DEFAULT_LABEL,$SS_CENTER)
   $Wnd_GUI_Text  = GUICtrlCreateLabel("How do you know what will you do for this " & @CRLF & " N "  & @CR & " N " _
   & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " _
   & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " & @CR & " N " , _
		 $Total_l,$Text_t, $Inside_w, $Text_h ,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_Text,20,$FW_EXTRABOLD)


   $Obj_Style = BitOR($GUI_SS_DEFAULT_LABEL,$SS_CENTER)
   Local $temp_label  = GUICtrlCreateLabel("Do you want to ? (Just press and wait for the window closed)", _
	  $Total_l, $Bottom_1_t,$Bottom_1_w, $Bottom_h ,$Obj_Style)
   GUICtrlSetFont($temp_label,20,$FW_BOLD)

   $Obj_Style = BitOR($GUI_SS_DEFAULT_LABEL,$SS_CENTER,$SS_NOTIFY)
   Local $title_tb  = GUICtrlCreateLabel("GIVE US YOUR FEEDBACK", _
	  $title_feed_l,$title_feed_t,$title_feed_w,$title_feed_h,$Obj_Style)
   GUICtrlSetFont($title_tb,12,$FW_BOLD,$GUI_FONTUNDER )
   GUICtrlSetColor($title_tb,$COLOR_BLUE)


   #cs
   $Wnd_GUI_Feedback  = GUICtrlCreateEdit("" & @CRLF, _
	  $feedback_l, $feedback_t, $feedback_w, $feedback_h, _
	  BitXOR($GUI_SS_DEFAULT_EDIT,$WS_HSCROLL,$WS_VSCROLL))
   GUICtrlSetFont($Wnd_GUI_Feedback,10,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_Feedback,$COLOR_BLACK)
   #ce

   $Obj_Style = BitOr($BS_AUTORADIOBUTTON,$BS_PUSHLIKE)
   $Wnd_GUI_RB_Run = GUICtrlCreateRadio("Continue Scan Process", _
	   $Bottom_l_1,$Bottom_2_t,$Bottom_2_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Run,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Run,$COLOR_GREEN)

   $Wnd_GUI_RB_Shut = GUICtrlCreateRadio("Just turn off", _
	  $Bottom_l_2,$Bottom_2_t, $Bottom_2_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Shut,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Shut,$COLOR_BLACK)

   $Wnd_GUI_RB_Stop = GUICtrlCreateRadio("I want to working now", _
	  $Bottom_l_3,$Bottom_2_t, $Bottom_2_w, $Bottom_h,$Obj_Style)
   GUICtrlSetFont($Wnd_GUI_RB_Stop,20,$FW_BOLD)
   GUICtrlSetColor($Wnd_GUI_RB_Stop,$COLOR_RED)

   GuiCtrlSetState($Wnd_GUI_RB_Run, $GUI_CHECKED)

   GUICtrlSetOnEvent($Wnd_GUI_RB_Run, "Wnd_Evt_Radio")
   GUICtrlSetOnEvent($Wnd_GUI_RB_Shut, "Wnd_Evt_Radio")
   GUICtrlSetOnEvent($Wnd_GUI_RB_Stop, "Wnd_Evt_Radio")
   GUICtrlSetOnEvent($title_tb,"Wnd_Evt_Feedback")
   GUISetState(@SW_SHOW,$Wnd_GUI_hWnd)
EndFunc

Func Wnd_Evt_Radio()
   If (BitAND(GUICtrlRead($Wnd_GUI_RB_Run), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status = $Wnd_Process_Status_VALUE_RUN
   ElseIf (BitAND(GUICtrlRead($Wnd_GUI_RB_Shut), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status  = $Wnd_Process_Status_VALUE_SHUT
   ElseIf (BitAND(GUICtrlRead($Wnd_GUI_RB_Stop), $GUI_CHECKED) = $GUI_CHECKED) Then
	  $Wnd_Process_Status  = $Wnd_Process_Status_VALUE_STOP
   EndIf
EndFunc

Func Wnd_Evt_Close()
   $Wnd_Process_Status = $Wnd_Process_Status_VALUE_STOP
EndFunc

Func Wnd_Evt_Feedback()
   Local $text = InputBox("Give Us Feedback","We will read it every morning","","",-1,-1,Default,Default,0,$Wnd_GUI_hWnd)
   Feedback_Save($text)
EndFunc

Func Wnd_Confirm($_message,$_title ="Please Confirm",$timeout =1000 )
   Local $confirm = MsgBox($MB_YESNO, $_title , $_message ,$timeout,$Wnd_GUI_hWnd)
   If $IDYES = $confirm Then
	  Return True
   EndIf
	  Return False
EndFunc


