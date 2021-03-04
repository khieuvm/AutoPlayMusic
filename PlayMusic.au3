#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=PlayMusic
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_ProductName=Khieudeptrai
#AutoIt3Wrapper_Res_ProductVersion=1.0.0.1
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © Khieudeptrai
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>
#include <GUIConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>

_Singleton(@ScriptName)

Opt("TrayMenuMode", 3)
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
Opt("GUIOnEventMode", 1) ; Change to OnEvent mode

Global Const $IniFile = "Config.ini"
Global $ScriptName    = "Config"
Global $GUI
Global $AboutGUI
Global $MIM_HourDefault = 15
Global $MIM_MinDefault = 0
Global $MIM_SourceMP3Default = "\\10.116.16.22\public1\RadioChannel\MusicInMotion"

Global $Tidy_HourDefault = 10
Global $Tidy_MinDefault = 0
Global $Tidy_SourceMP3Default = "\\10.116.16.22\public1\RadioChannel\Tidy-up"

Global $DestinationMP3Default = "D:\"
Global $DestinationMP3
Global $MIM_DestinationMP3
Global $Tidy_DestinationMP3

Global $MIM_SourceMP3
Global $MIM_Hour
Global $MIM_Min
Global $MIM_HourDownload
Global $MIM_MinDownLoad

Global $Tidy_SourceMP3
Global $Tidy_Hour
Global $Tidy_Min
Global $Tidy_HourDownload
Global $Tidy_MinDownLoad

Global $bCopied

ReadIniFile()
CreateGUIConfig()
CreateGUIAbout()
Main()

Func Main()
	Local $idAbout = TrayCreateItem("About")
	TrayItemSetOnEvent(-1, "OnAbout")

	Local $idConfig = TrayCreateItem("Config")
	TrayItemSetOnEvent(-1, "OnConfig")

	Local $idExit = TrayCreateItem("Exit")
	TrayItemSetOnEvent(-1, "OnExit")

	TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.
	TraySetClick(16) ; Show the tray menu when the mouse if hovered over the tray icon.

	TraySetToolTip("PlayMusic")
	TraySetIcon("icon.ico")

	; Check if the registry key is already existing, so as not to damage the user's system.
	RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PlayMusic")

	Local $sfilePath = StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 0, -1) - 1)
	; @error is set to non-zero when reading a registry key that doesn't exist.
	If @error Then
		; Write a single REG_SZ value.
		RegWrite("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "PlayMusic", "REG_SZ", $sfilePath)
	Else
		RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PlayMusic")
		RegWrite("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "PlayMusic", "REG_SZ", $sfilePath)
	EndIf

	$bCopied = False

	$MIM_HourDownload = $MIM_Hour
	$MIM_MinDownLoad = $MIM_Min

	$Tidy_HourDownload = $Tidy_Hour
	$Tidy_MinDownLoad = $Tidy_Min

	For $i = 1 To 50 Step +1
	   Send("{VOLUME_DOWN}")
	Next

	SetTime()
	Local $bDisplayNotify = False

	While 1
		Local $Now_Hour = @HOUR
		Local $Now_Min = @MIN

		If $MIM_HourDownload = $Now_Hour Then
			If $bCopied = False Then
				If @MIN >= $MIM_MinDownLoad Then
					FileDelete($MIM_DestinationMP3)
					$bCopied = FileCopy($MIM_SourceMP3 & "\" & @YEAR & "\" & @MON & "\" & @MDAY & "\*", $MIM_DestinationMP3)
					If $bCopied = True Then
					   MsgBox($MB_SYSTEMMODAL, "Notify", "MusicInMotion will start after few minutes with volume 60. Please turn on speaker. Thanks.", 10)
					EndIf
				EndIf
			EndIf

			If $bDisplayNotify = False Then
				If @MIN = $MIM_MinDownLoad + 9 Then
					If $bCopied = True Then
					   MsgBox($MB_SYSTEMMODAL, "Notify", "MusicInMotion will start after 1 minutes with volume 60. Please turn on speaker. Thanks.", 10)
					   $bDisplayNotify = True
					Else
					   MsgBox($MB_SYSTEMMODAL, "Notify", "MusicInMotion will not start because Autodownload Fail. Thanks.", 15)
					   $bDisplayNotify = True
					EndIf
				EndIf
			EndIf

			Sleep(5000)

		EndIf

		If $MIM_Hour = $Now_Hour Then
			If $bCopied = True Then
				If @MIN = $MIM_Min Then

					PlayMusic($MIM_DestinationMP3)

					$bCopied = False
					$bDisplayNotify = False
				EndIf
			EndIf

			Sleep(5000)
		EndIf

		If @WDAY = 6 Then
			If $Tidy_HourDownload = $Now_Hour Then
				If $bCopied = False Then
					If @MIN >= $Tidy_MinDownLoad Then
						FileDelete($Tidy_DestinationMP3)
						$bCopied = FileCopy($Tidy_SourceMP3 & "\" & @YEAR & "\" & @MON & "\" & @MDAY & "\*", $Tidy_DestinationMP3)
						If $bCopied = True Then
						   MsgBox($MB_SYSTEMMODAL, "Notify", "Tidy-up will start after few minutes with volume 60. Please turn on speaker. Thanks.", 10)
						EndIf
					EndIf
				EndIf

				If $bDisplayNotify = False Then
					If @MIN = $Tidy_MinDownLoad + 9 Then
						If $bCopied = True Then
						   MsgBox($MB_SYSTEMMODAL, "Notify", "Tidy-up will start after 1 minutes with volume 60. Please turn on speaker. Thanks.", 10)
						   $bDisplayNotify = True
						Else
						   MsgBox($MB_SYSTEMMODAL, "Notify", "Tidy-up will not start because Autodownload Fail. Thanks.", 15)
						   $bDisplayNotify = True
						EndIf
					EndIf
				EndIf

				Sleep(5000)

			EndIf


			If $Tidy_Hour = $Now_Hour Then
				If $bCopied = True Then
					If @MIN = $Tidy_Min Then

						PlayMusic($Tidy_DestinationMP3)

						$bCopied = False
						$bDisplayNotify = False
					EndIf
				EndIf

				Sleep(5000)

			EndIf
		EndIf

		Sleep(5000)
	WEnd
EndFunc

Func PlayMusic(ByRef $Music)
	Send("{RCTRL}")
	For $i = 1 To 50 Step +1
	   Send("{VOLUME_DOWN}")
	Next
	For $i = 1 To 30 Step +1
	   Send("{VOLUME_UP}")
	   Sleep(50)
	Next

	SoundPlay($Music, 1)
	Send("{RCTRL}")
	For $i = 1 To 50 Step +1
	   Send("{VOLUME_DOWN}")
	Next

	FileDelete($Music)
EndFunc

Func CreateGUIConfig()
	$GUI = GuiCreate($ScriptName, 480, 200)
	GUISetOnEvent($GUI_EVENT_CLOSE, "OnCancel")
	GUISetIcon ("icon.ico")

	GUICtrlCreateLabel ("Auto Play Music", 120, 15)
	GUICtrlCreateLabel ("Hour", 390, 15)
	GUICtrlCreateLabel ("Minute", 430, 15)
	;Time MIM
	Global $MIM_inputhour = GUICtrlCreateInput ($MIM_Hour, 390, 40, 30, 20)
	GUICtrlSetLimit($MIM_inputhour, 2)
	Global $MIM_inputmin = GUICtrlCreateInput ($MIM_Min, 430, 40, 30, 20)
	GUICtrlSetLimit($MIM_inputmin, 2)
	;Time Tidy
	Global $Tidy_inputhour = GUICtrlCreateInput ($Tidy_Hour, 390, 80, 30, 20)
	GUICtrlSetLimit($Tidy_inputhour, 2)
	Global $Tidy_inputmin = GUICtrlCreateInput ($Tidy_Min, 430, 80, 30, 20)
	GUICtrlSetLimit($Tidy_inputmin, 2)

   GUICtrlCreateLabel ("MIM Source Folder", 10, 43)
   Global $MIM_inputSourceFolder = GUICtrlCreateInput ($MIM_SourceMP3, 120, 40, 250, 20)

   GUICtrlCreateLabel ("Tidy-up Source Folder", 10, 83)
   Global $Tidy_inputSourceFolder = GUICtrlCreateInput ($Tidy_SourceMP3, 120, 80, 250, 20)

   GUICtrlCreateLabel ("Destination Folder", 10, 123)
   Global $inputDestFolder = GUICtrlCreateInput ($DestinationMP3, 120, 120, 250, 20)

   $ButtonOK = GuiCtrlCreateButton("OK", 100, 160, 80, 25)
   GUICtrlSetOnEvent(-1, "OnOK")
   $ButtonCancel = GuiCtrlCreateButton("Cancel", 300, 160, 80, 25)
   GUICtrlSetOnEvent(-1, "OnCancel")
EndFunc

Func OnConfig()
    GUISetState(@SW_SHOW, $GUI)
EndFunc

Func OnOK()
   $MIM_Hour = Number(GUICtrlRead($MIM_inputhour))
   $MIM_Min = Number(GUICtrlRead($MIM_inputmin))
   $MIM_SourceMP3 = GUICtrlRead($MIM_inputSourceFolder)

   $Tidy_Hour = Number(GUICtrlRead($Tidy_inputhour))
   $Tidy_Min = Number(GUICtrlRead($Tidy_inputmin))
   $Tidy_SourceMP3 = GUICtrlRead($Tidy_inputSourceFolder)

   $DestinationMP3 = GUICtrlRead($inputDestFolder)

   $MIM_DestinationMP3 = $DestinationMP3 & "/MIM.mp3"
   $Tidy_DestinationMP3 = $DestinationMP3 & "/Tydi-up.mp3"

   GUISetState(@SW_HIDE, $GUI)
   WriteIniFile()
   SetTime()
EndFunc

Func SetTime()
	;MIM
   If $MIM_Min < 10 Then
	  $MIM_HourDownload = $MIM_Hour - 1
	  $MIM_MinDownload = $MIM_Min + 50
   Else
	  $MIM_MinDownload = $MIM_Min - 10
   EndIf

   ;Tidy
   If $Tidy_Min < 10 Then
	  $Tidy_HourDownload = $Tidy_Hour - 1
	  $Tidy_MinDownload = $Tidy_Min + 50
   Else
	  $Tidy_MinDownload = $Tidy_Min - 10
   EndIf

   $bCopied = False
   FileDelete($MIM_DestinationMP3)
   FileDelete($Tidy_DestinationMP3)
EndFunc

Func OnCancel()
   GUISetState(@SW_HIDE, $GUI)
   GUICtrlSetData($MIM_inputhour, $MIM_Hour)
   GUICtrlSetData($MIM_inputmin, $MIM_Min)
   GUICtrlSetData($MIM_inputSourceFolder, $MIM_SourceMP3)

   GUICtrlSetData($Tidy_inputhour, $Tidy_Hour)
   GUICtrlSetData($Tidy_inputmin, $Tidy_Min)
   GUICtrlSetData($Tidy_inputSourceFolder, $Tidy_SourceMP3)

   GUICtrlSetData($inputDestFolder, $DestinationMP3)
EndFunc

Func CreateGUIAbout()
	$AboutGUI = GuiCreate("PlayMusic - About", 300, 150)
	GUISetOnEvent($GUI_EVENT_CLOSE, "OnCancelAbout")
	GUISetIcon ("icon.ico")

	GUICtrlCreateLabel ("Auto Play Music", 69, 17, 250, 50)
	GUICtrlSetResizing(-1,$GUI_DOCKTOP+$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM)
	GUICtrlSetFont(-1, 20, 500, 0, "Cambria")

	GUICtrlCreateLabel ("Version 1.0.0.1", 109, 50, 250, 30)
	GUICtrlSetResizing(-1,$GUI_DOCKTOP+$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM)
	GUICtrlSetFont(-1, 10, 200, 0, "Cambria")

	GUICtrlCreateLabel ("Design by Khieudeptrai", 82, 70, 250, 30)
	GUICtrlSetResizing(-1,$GUI_DOCKTOP+$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM)
	GUICtrlSetFont(-1, 11, 400, 0, "Cambria")

	GUICtrlCreateLabel ("Copyright © 2021 Khieudeptrai", 72, 105, 250, 30)
	GUICtrlSetResizing(-1,$GUI_DOCKTOP+$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM)
	GUICtrlSetFont(-1, 9.5, 200, 0, "Cambria")


	GUICtrlCreateLabel ("All rights reserved", 108, 120, 250, 30)
	GUICtrlSetResizing(-1,$GUI_DOCKTOP+$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM)
	GUICtrlSetFont(-1, 9.5, 200, 0, "Cambria")
EndFunc

Func OnCancelAbout()
	GUISetState(@SW_HIDE, $AboutGUI)
EndFunc

Func OnAbout()
	GUISetState(@SW_SHOW, $AboutGUI)
EndFunc   ;==>About

Func OnExit()
	GUIDelete($AboutGUI)
	GUIDelete($GUI)
    Exit
EndFunc

Func ReadIniFile()
   $MIM_Hour = Number(IniRead($IniFile, "Config", "MIM_Hour", $MIM_HourDefault))
   $MIM_Min = Number(IniRead($IniFile, "Config", "MIM_Min", $MIM_MinDefault))
   $MIM_SourceMP3 = IniRead($IniFile, "Config", "MIM_Source", $MIM_SourceMP3Default)

   $Tidy_Hour = Number(IniRead($IniFile, "Config", "Tidy_Hour", $Tidy_HourDefault))
   $Tidy_Min = Number(IniRead($IniFile, "Config", "Tidy_Min", $Tidy_MinDefault))
   $Tidy_SourceMP3 = IniRead($IniFile, "Config", "Tidy_Source", $Tidy_SourceMP3Default)

   $DestinationMP3 = IniRead($IniFile, "Config", "Dest", $DestinationMP3Default)

   $MIM_DestinationMP3 = $DestinationMP3 & "/MIM.mp3"
   $Tidy_DestinationMP3 = $DestinationMP3 & "/Tydi-up.mp3"
EndFunc

Func WriteIniFile()
   IniWrite($IniFile, "Config", "MIM_Hour", $MIM_Hour)
   IniWrite($IniFile, "Config", "MIM_Min", $MIM_Min)
   IniWrite($IniFile, "Config", "MIM_Source", $MIM_SourceMP3)

   IniWrite($IniFile, "Config", "Tidy_Hour", $Tidy_Hour)
   IniWrite($IniFile, "Config", "Tidy_Min", $Tidy_Min)
   IniWrite($IniFile, "Config", "Tidy_Source", $Tidy_SourceMP3)

   IniWrite($IniFile, "Config", "Dest", $DestinationMP3)
EndFunc