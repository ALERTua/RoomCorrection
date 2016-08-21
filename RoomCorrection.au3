Opt("MouseCoordMode", 2) ;1=absolute, 0=relative, 2=client

$optionsFile 					= "RoomCorrection.ini"
$filepath 						= EnvGet("ProgramW6432")&"\Realtek\Audio\HDA\RtkNGUI64.exe"
$title 							= "Realtek HD Audio Manager"
$speakerConfigurationControl 	= "[CLASS:SysTabControl32; INSTANCE:1]"
$roomCorrectionControl 			= "[CLASS:#32770; INSTANCE:18]"
$roomCorrectionSwitchControl 	= "[CLASS:Button; INSTANCE:41]"
$meterButtonControl				= "[CLASS:Button; INSTANCE:39]"

Const $FL	= "[CLASS:msctls_updown32; INSTANCE:1]"
Const $FLG	= "[CLASS:msctls_updown32; INSTANCE:2]"
Const $FR	= "[CLASS:msctls_updown32; INSTANCE:9]"
Const $FRG	= "[CLASS:msctls_updown32; INSTANCE:10]"
Const $CE	= "[CLASS:msctls_updown32; INSTANCE:3]"
Const $CEG	= "[CLASS:msctls_updown32; INSTANCE:4]"
Const $SWG	= "[CLASS:msctls_updown32; INSTANCE:11]"
Const $RL	= "[CLASS:msctls_updown32; INSTANCE:7]"
Const $RLG	= "[CLASS:msctls_updown32; INSTANCE:8]"
Const $RR	= "[CLASS:msctls_updown32; INSTANCE:14]"
Const $RRG	= "[CLASS:msctls_updown32; INSTANCE:15]"

Const $FLVal	= "[CLASS:Static; INSTANCE:20]"
Const $FLGVal	= "[CLASS:Static; INSTANCE:21]"
Const $FRVal	= "[CLASS:Static; INSTANCE:28]"
Const $FRGVal	= "[CLASS:Static; INSTANCE:29]"
Const $CEVal	= "[CLASS:Static; INSTANCE:22]"
Const $CEGVal	= "[CLASS:Static; INSTANCE:23]"
Const $SWGVal	= "[CLASS:Static; INSTANCE:30]"
Const $RLVal	= "[CLASS:Static; INSTANCE:26]"
Const $RLGVal	= "[CLASS:Static; INSTANCE:27]"
Const $RRVal	= "[CLASS:Static; INSTANCE:33]"
Const $RRGVal	= "[CLASS:Static; INSTANCE:34]"

Local $FLValue
Local $FLGValue
Local $FRValue
Local $FRGValue
Local $CEValue
Local $CEGValue
Local $SWGValue
Local $RLValue
Local $RLGValue
Local $RRValue
Local $RRGValue

Local $hwnd

if not FileExists($filepath) Then
   MsgBox(0, "Room Correction Settings Manager", "Cannot locate "&$filepath&@CRLF&"Exiting")
   Exit
EndIf

if not FileExists($optionsFile) Then
   RunProgram()
   $myBox = MsgBox(1, "Room Correction Settings Manager", "Please move this window away and fill Room Correction parameters."&@CRLF&"You will have to do this just once."&@CRLF&"When you are done, press OK")
   If $myBox == 1 Then
	  StoreSettings()
	  MsgBox(0, "Room Correction Settings Manager", "Settings stored in "&$optionsFile&@CRLF&"Next time you run this program, it will automaticaly fill stored settings"&@CRLF&"You can edit settings file manually"&@CRLF&"You can delete the file to reset settings")
   Else
	  Exit
   EndIf
EndIf

RunProgram()
ReadSettings()

for $i = 1 to 3
   FillValues()
   sleep(250)
Next
MsgBox(0, "Room Correction Settings Manager", "Settings applied"&@CRLF&"Thank you for using Room Correction Settings Manager")




Func Increment($incrementControl, $times)
   local $indentX = $times > 0 ? 45 : 15 ;button position offset. client coords
   local $indentY = 15
   ;Print("times: " & $times & " : " & $indentX & " " & $indentY)
   ControlClick($hwnd, "", $incrementControl, "primary", abs($times) * 2, $indentX, $indentY)
   sleep(50)
EndFunc

Func SetValue($valueControl, $controlsControl, $desiredValue, $name)
   local $currentValue = GetNumValue($valueControl)

   ;skip if value is already set
   if $desiredValue = $currentValue Then
	  ;Print("value already set")
	  return
   EndIf

   ;get steps
   local $step = 0
   Increment($controlsControl, -1)
   $step = Round(Abs($currentValue - GetNumValue($valueControl)), 2)
   Increment($controlsControl, 1)

   ;another algorithm if the value is maxed
   if $step == 0 Then
	  Increment($controlsControl, 1)
	  $step = Round(Abs($currentValue - GetNumValue($valueControl)), 2)
	  Increment($controlsControl, -1)
   EndIf

   ;get increment steps
   local $resultValue = $desiredValue - $currentValue
   local $steps = Round($resultValue / $step, 0)

   ;increment
   PrintMany($name, $desiredValue, $currentValue, $step, $resultValue, $steps)
   Increment($controlsControl, $steps)
EndFunc

Func FillValues()
   SetValue($FLVal	,	$FL		,	$FLValue	, "FL"	)
   SetValue($FLGVal	,	$FLG	,	$FLGValue	, "FLG"	)
   SetValue($FRVal	,	$FR		,	$FRValue	, "FR"	)
   SetValue($FRGVal	,	$FRG	,	$FRGValue	, "FRG"	)
   SetValue($CEVal	, 	$CE		,	$CEValue	, "CE"	)
   SetValue($CEGVal	,	$CEG	,	$CEGValue	, "CEG"	)
   SetValue($SWGVal	,	$SWG	,	$SWGValue	, "SWG"	)
   SetValue($RLVal	, 	$RL		,	$RLValue	, "RL"	)
   SetValue($RLGVal	,	$RLG	,	$RLGValue	, "RLG"	)
   SetValue($RRVal	, 	$RR		,	$RRValue	, "RR"	)
   SetValue($RRGVal	,	$RRG	,	$RRGValue	, "RRG"	)
EndFunc

Func RunProgram()
   if WinExists($title) Then
	  WinClose($title)
   EndIf

   Run($filepath)
   $hwnd = WinWait($title)
   WinActivate($hwnd)
   WinWaitActive($hwnd)

   For $i = 1 to 5
	  ControlCommand($hwnd, "", $speakerConfigurationControl, "TabRight")
  	  sleep(200)
	  if CheckRoomCorrectionTabActive() Then
		 ExitLoop
	  EndIf
   Next
   SwitchRoomCorrection(1)
EndFunc

Func StoreSettings()
   IniWrite($optionsFile	,	"RoomCorrection"	,	"FrontLeft"			,	GetNumValue($FLVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"FrontLeftGain"		,	GetNumValue($FLGVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"FrontRight"		,	GetNumValue($FRVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"FrontRightGain"	,	GetNumValue($FRGVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"Center"			,	GetNumValue($CEVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"CenterGain"		,	GetNumValue($CEGVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"SubWooferGain"		,	GetNumValue($SWGVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"RearLeft"			,	GetNumValue($RLVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"RearLeftGain"		,	GetNumValue($RLGVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"RearRight"			,	GetNumValue($RRVal	))
   IniWrite($optionsFile	,	"RoomCorrection"	,	"RearRightGain"		,	GetNumValue($RRGVal	))
EndFunc

Func ReadSettings()
   $FLValue 	= IniRead($optionsFile	,	"RoomCorrection"	,	"FrontLeft"			,	GetNumValue($FLVal	))
   $FLGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"FrontLeftGain"		,	GetNumValue($FLGVal	))
   $FRValue 	= IniRead($optionsFile	,	"RoomCorrection"	,	"FrontRight"		,	GetNumValue($FRVal	))
   $FRGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"FrontRightGain"	,	GetNumValue($FRGVal	))
   $CEValue		= IniRead($optionsFile	,	"RoomCorrection"	,	"Center"			,	GetNumValue($CEVal	))
   $CEGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"CenterGain"		,	GetNumValue($CEGVal	))
   $SWGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"SubWooferGain"		,	GetNumValue($SWGVal	))
   $RLValue		= IniRead($optionsFile	,	"RoomCorrection"	,	"RearLeft"			,	GetNumValue($RLVal	))
   $RLGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"RearLeftGain"		,	GetNumValue($RLGVal	))
   $RRValue		= IniRead($optionsFile	,	"RoomCorrection"	,	"RearRight"			,	GetNumValue($RRVal	))
   $RRGValue	= IniRead($optionsFile	,	"RoomCorrection"	,	"RearRightGain"		,	GetNumValue($RRGVal	))
EndFunc

Func SwitchRoomCorrection($switch = 1)
   sleep(150)
   ;CheckRoomCorrectionVisible
   if not CheckRoomCorrectionSwitchVisible() Then
	  MsgBox(0, "Room Correction Settings Manager", "Couldn't switch Room Correction. Please enable it manually and try again")
	  Exit
   EndIf

   if ( $switch = 1 and not CheckRoomCorrectionActive() ) or ( $switch = 0 and CheckRoomCorrectionActive() ) then
	  ControlCommand($hwnd, "", $roomCorrectionSwitchControl, "Check")
	  ControlCommand($hwnd, "", $meterButtonControl, "Check")
   EndIf
EndFunc

Func CheckRoomCorrectionTabActive()
   return (ControlCommand($hwnd, "", $roomCorrectionControl, "IsVisible") == 1)
EndFunc

Func CheckRoomCorrectionSwitchVisible()
   return (ControlCommand($hwnd, "", $roomCorrectionSwitchControl, "IsVisible") == 1)
EndFunc

Func CheckRoomCorrectionActive()
   return (ControlCommand($hwnd, "", $meterButtonControl, "IsEnabled") == 1)
EndFunc

Func GetNumValue($control)
   return Number(ControlGetText($hwnd, "", $control))
EndFunc






Func Print($string)
   ConsoleWrite($string&@CRLF)
EndFunc

Func Printmany($s1 = "", $s2 = "", $s3 = "", $s4 = "", $s5 = "", $s6 = "", $s7 = "", $s8 = "", $s9 = "", $s10 = "")
   local $output = ""
   if not String($s1) = "" Then
	  $output = $output & "1: " & String($s1) & @CRLF
   EndIf
   if not String($s2) = "" Then
	  $output = $output & "2: " & String($s2) & @CRLF
   EndIf
   if not String($s3) = "" Then
	  $output = $output & "3: " & String($s3) & @CRLF
   EndIf
   if not String($s4) = "" Then
	  $output = $output & "4: " & String($s4) & @CRLF
   EndIf
   if not String($s5) = "" Then
	  $output = $output & "5: " & String($s5) & @CRLF
   EndIf
   if not String($s6) = "" Then
	  $output = $output & "6: " & String($s6) & @CRLF
   EndIf
   if not String($s7) = "" Then
	  $output = $output & "7: " & String($s7) & @CRLF
   EndIf
   if not String($s8) = "" Then
	  $output = $output & "8: " & String($s8) & @CRLF
   EndIf
   if not String($s9) = "" Then
	  $output = $output & "9: " & String($s9) & @CRLF
   EndIf
   if not String($s10) = "" Then
	  $output = $output & "10: " & String($s10) & @CRLF
   EndIf

   Print($output)
EndFunc