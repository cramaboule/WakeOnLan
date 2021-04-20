#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoItv11.ico
#AutoIt3Wrapper_Run_After=copy %out% "H:\COMMON\Informatique\Applications\installation\wol"
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;----------------------- Wake On Lan -------------------------
;
; V1.1 Added combo box
; V1.0 Inital release
;
;-------------------------------------------------------------


$debug = 0

$Ver = "V1.1"

$gPort = "7"
$gIP = '10.0.1.254'
$gMask = '255.0.0.0'
If $debug Then

	$gMac = '112233445566'
Else
	$hFile = FileOpen('known-Mac-address.txt', 0)
	If $hFile = -1 Then
		MsgBox(16, "File missing", "Error: file 'known-Mac-address.txt' missing")
		$sgMag = ''
	Else
		$sgMag = FileRead($hFile)
		$sgMag = StringReplace($sgMag, @CRLF, "|")
	EndIf
	FileClose($hFile)
EndIf

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Wake On Lan " & $Ver, 329, 184)
$MAC = GUICtrlCreateCombo('', 96, 62, 217, 21)
$Label1 = GUICtrlCreateLabel("IP or Hostname:", 8, 14, 80, 20)
$IP = GUICtrlCreateInput($gIP, 96, 10, 217, 21)
$Label4 = GUICtrlCreateLabel("Network Mask:", 8, 40, 80, 20)
$InMask = GUICtrlCreateInput($gMask, 96, 36, 217, 21)
$Label2 = GUICtrlCreateLabel("MAC Address:", 8, 66, 71, 20)
$Port = GUICtrlCreateInput($gPort, 96, 88, 217, 21)
$Label3 = GUICtrlCreateLabel("Port :", 8, 92, 69, 20)
$Go = GUICtrlCreateButton("Wake Up", 112, 136, 160, 35, $BS_DEFPUSHBUTTON)
$Status = GUICtrlCreateLabel("Status: Ready", 8, 144, 100, 20)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

; Add additional items to the combobox.
GUICtrlSetData($MAC, $sgMag, '')


While 1
	$nMsg = GUIGetMsg()
	Select
		Case $nMsg = $GUI_EVENT_CLOSE
			Exit
		Case $nMsg = $Go
			$aMac = StringSplit(GUICtrlRead($MAC), '=')
			$sMac = StringRegExpReplace($aMac[1], "(?i)[:-]", '')
			If ($sMac <> '' Or StringLen($sMac) = 12) Then
				GUICtrlSetData($Status, "Status: Sending!")
				_SendWakeUp(_GetBroadcastAddress(GUICtrlRead($IP), GUICtrlRead($InMask)), $sMac, GUICtrlRead($Port))
				If @error <> 0 Then
					GUICtrlSetData($Status, "Status: Error N° " & @error)
				Else
					GUICtrlSetData($Status, "Status: Sent!")
					Sleep(2000) ; 2 sec
					GUICtrlSetData($Status, "Status: Ready")
				EndIf
			Else
				MsgBox(16, "Error MAC address", "Please enter a valid MAC address")
			EndIf
	EndSelect
WEnd ; Target PC Info


Func _SendWakeUp($sIp1, $ssMac, $sPort)
	$mistake = 0
	UDPStartup()
	If @error <> 0 Then $mistake = 1
	$sock = UDPOpen($sIp1, $sPort)
	If @error <> 0 Then $mistake = 2
	$Magic = _WOL_GenerateMagicPacket($ssMac)
	UDPSend($sock, $Magic)
	If @error <> 0 Then $mistake = 3
	UDPCloseSocket($sock)
	If @error <> 0 Then $mistake = 4
	UDPShutdown()
	If @error <> 0 Then $mistake = 5
	If $mistake <> 0 Then SetError($mistake)
EndFunc   ;==>_SendWakeUp

Func _WOL_GenerateMagicPacket($strMACAddress)

	$MagicPacket = ""
	$MACData = ""
	For $p = 1 To 11 Step 2
		$MACData = $MACData & _HexToChar(StringMid($strMACAddress, $p, 2))
	Next
	For $p = 1 To 6
		$MagicPacket = _HexToChar("ff") & $MagicPacket
	Next
	For $p = 1 To 16
		$MagicPacket = $MagicPacket & $MACData
	Next
	For $p = 1 To 6
		$MagicPacket &= _HexToChar("00")
	Next
	Return $MagicPacket
EndFunc   ;==>_WOL_GenerateMagicPacket

; This function convert a MAC Address Byte (e.g. "1f") to a char
Func _HexToChar($strHex)
	Return Chr(Dec($strHex))
EndFunc   ;==>_HexToChar

Func _GetNetworkAddress($sIPAddress, $sSubnetMask)
	Local $aIPAddress = StringSplit($sIPAddress, '.'), $aSubnetMask = StringSplit($sSubnetMask, '.')
	$sIPAddress = ''

	For $i = 1 To $aIPAddress[0]
		$aIPAddress[$i] = BitAND($aIPAddress[$i], $aSubnetMask[$i])
		$sIPAddress &= $aIPAddress[$i] & '.'
	Next
	Return StringTrimRight($sIPAddress, 1)
EndFunc   ;==>_GetNetworkAddress

Func _GetBroadcastAddress($sIPAddress, $sSubnetMask)
	Local $aIPAddress = StringSplit($sIPAddress, '.'), $aSubnetMask = StringSplit($sSubnetMask, '.')
	$sIPAddress = ''
	For $i = 1 To $aSubnetMask[0]
		$aSubnetMask[$i] = _BinToDec(_Convert_To_Binary($aSubnetMask[$i], 1))
		$aIPAddress[$i] = BitOR($aIPAddress[$i], $aSubnetMask[$i])
		$sIPAddress &= $aIPAddress[$i] & '.'
	Next
	Return StringTrimRight($sIPAddress, 1)
EndFunc   ;==>_GetBroadcastAddress

Func _Convert_To_Binary($iNumber, $inverse)
	Local $sBinString = ""
	While Number($iNumber)
		$sBinString = ($inverse) ? Number( Not (BitAND($iNumber, 1))) & $sBinString : BitAND($iNumber, 1) & $sBinString
		$iNumber = BitShift($iNumber, 1)
	WEnd
	If $sBinString = '' Then $sBinString = ($inverse) ? '11111111' : '00000000'
	Return $sBinString
EndFunc   ;==>_Convert_To_Binary


Func _BinToDec($num)
	Local $dec = 0, $len = StringLen($num)
	For $i = 0 To $len - 1
		$dec += 2 ^ $i * StringMid($num, $len - $i, 1)
	Next
	Return $dec
EndFunc   ;==>_BinToDec
