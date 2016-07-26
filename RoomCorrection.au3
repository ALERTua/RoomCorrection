Opt("MouseCoordMode", 2) ;1=absolute, 0=relative, 2=client

$optionsFile 					= "RoomCorrection.ini"
$filepath 						= EnvGet("ProgramW6432")&"\Realtek\Audio\HDA\RtkNGUI64.exe"
$title 							= "Realtek HD Audio Manager"
$roomCorrectionText				= "Enable Room Correction"
$speakerConfigurationControl 	= "[CLASS:SysTabControl32; INSTANCE:2]"
$roomCorrectionControl 			= "[CLASS:#32770; INSTANCE:7]"
$roomCorrectionSwitchControl 	= "[CLASS:Button; INSTANCE:33]"
$meterButtonControl				= "[CLASS:Button; INSTANCE:31]"

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

Const $FLVal	= "[CLASS:Static; INSTANCE:8]"
Const $FLGVal	= "[CLASS:Static; INSTANCE:9]"
Const $FRVal	= "[CLASS:Static; INSTANCE:16]"
Const $FRGVal	= "[CLASS:Static; INSTANCE:17]"
Const $CEVal	= "[CLASS:Static; INSTANCE:10]"
Const $CEGVal	= "[CLASS:Static; INSTANCE:11]]"
Const $SWGVal	= "[CLASS:Static; INSTANCE:18]"
Const $RLVal	= "[CLASS:Static; INSTANCE:14]"
Const $RLGVal	= "[CLASS:Static; INSTANCE:15]"
Const $RRVal	= "[CLASS:Static; INSTANCE:21]"
Const $RRGVal	= "[CLASS:Static; INSTANCE:22]"

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
   ElseIf $myBox == 2 Then
	  ;Exit
   EndIf
   Exit
EndIf

RunProgram()
ReadSettings()

for $i = 1 to 2
   FillValues()
   sleep(1000)
Next
MsgBox(0, "Room Correction Settings Manager", "Settings applied"&@CRLF&"Thank you for using Room Correction Settings Manager")



Func SetValue($valueControl, $controlsControl, $desiredValue)
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

   if $step == 0 Then
	  Increment($controlsControl, 1)
	  $step = Round(Abs($currentValue - GetNumValue($valueControl)), 2)
	  Increment($controlsControl, -1)
   EndIf

   ;get increment steps
   local $resultValue = $desiredValue - $currentValue
   local $steps = $resultValue / $step

   ;increment
   PrintMany($desiredValue, $currentValue, $resultValue, $steps)
   Increment($controlsControl, $steps)
EndFunc

Func FillValues()
   SetValue($FLVal	,	$FL		,	$FLValue	)
   SetValue($FLGVal	,	$FLG	,	$FLGValue	)
   SetValue($FRVal	,	$FR		,	$FRValue	)
   SetValue($FRGVal	,	$FRG	,	$FRGValue	)
   SetValue($CEVal	, 	$CE		,	$CEValue	)
   SetValue($CEGVal	,	$CEG	,	$CEGValue	)
   SetValue($SWGVal	,	$SWG	,	$SWGValue	)
   SetValue($RLVal	, 	$RL		,	$RLValue	)
   SetValue($RLGVal	,	$RLG	,	$RLGValue	)
   SetValue($RRVal	, 	$RR		,	$RRValue	)
   SetValue($RRGVal	,	$RRG	,	$RRGValue	)
EndFunc

Func RunProgram()
   if WinExists($title) Then
	  WinClose($title)
   EndIf

   Run($filepath)
   $hwnd = WinWait($title)
   WinActivate($hwnd)
   WinWaitActive($hwnd)
   ControlCommand($hwnd, "", $speakerConfigurationControl, "TabRight")
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

Func CheckRoomCorrectionVisible()
   sleep(150)
   local $tmp = ControlCommand($hwnd, $roomCorrectionText, $roomCorrectionSwitchControl, "IsVisible")
   local $output = False
   if $tmp = 1 Then
	  $output = True
   EndIf
   return $output
EndFunc

Func SwitchRoomCorrection($switch = 1)
   if not CheckRoomCorrectionVisible() Then
	  return False
   EndIf

   if ( $switch = 1 and not CheckRoomCorrectionActive() ) or ( $switch = 0 and CheckRoomCorrectionActive() ) then
	  PressRoomCorrectionSwitch()
	  return True
   EndIf
   return False
EndFunc

Func PressRoomCorrectionSwitch()
   ControlCommand($hwnd, $roomCorrectionText, $roomCorrectionSwitchControl, "Check")
   ControlCommand($hwnd, $roomCorrectionText, $meterButtonControl, "Check")
EndFunc

Func CheckRoomCorrectionActive()
   local $tmp = ControlCommand($hwnd, "", $meterButtonControl, "IsEnabled")
   local $output = False
   if $tmp = 1 Then
	  $output = True
   EndIf
   return $output
EndFunc

Func GetIncDecCoords($control) ;[0] = incX, [1] = incY, [2] = decX, [3] = decY
   local $incrementPos = ControlGetPos($hwnd, "", $control)
   Local $decrementPos[2] = [$incrementPos[0], $incrementPos[1] + 15]
   local $output[4] = [$incrementPos[0]+3, $incrementPos[1], $decrementPos[0]+3, $decrementPos[1]]
   return $output
EndFunc

Func GetNumValue($control)
   return Number(ControlGetText($hwnd, "", $control))
EndFunc

Func Increment($control, $times)
   local $incDecCoords = GetIncDecCoords($control)
   Local $resultCoords[2]
   if $times > 0 Then
	  $resultCoords[0] = $incDecCoords[0]
	  $resultCoords[1] = $incDecCoords[1]
   Else
	  $resultCoords[0] = $incDecCoords[2]
	  $resultCoords[1] = $incDecCoords[3]
   EndIf
   Click($resultCoords[0], $resultCoords[1], Abs($times))
   sleep(10)
EndFunc

Func Click($x, $y, $times)
   ;Print("Clicking " & $times & " time(s)")
   MouseMove($x, $y, 0)
   sleep(10)
   MouseClick("primary", $x, $y, $times, 0)
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