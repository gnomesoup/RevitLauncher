#Persistent
#SingleInstance
#NoTrayIcon
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Perform an install/uninstall of the WHA Revit Launcher Program
programName = WHA Revit Launcher

If !A_IsAdmin
{
	MsgBox, 16,, You do not have administrator privileges.`n`nTo install %programName%, please re-run Setup.exe with administrator privileges.
	ExitApp
}

FileGetVersion, newVersion, %programName%.exe
versionSuffix := " beta2"
installFolder = %A_ProgramFiles%\%programName%%versionSuffix%
installFullPath = %installFolder%\%programName%.exe
; ListVars
upgrade := 0
silent := 0
If silent
	GoSub, Install


Gui, New
Gui, Font, s22cBlack, Arial
Gui, Add, Text, w500 center, Install %programName%
Gui, Font, s10 c666666, Arial
Gui, Add, Text, w500 center yp+35, Version: %newVersion%
Gui, Font, s10 cBlack, Arial
Gui, Add, Text, w500, This little program automates many of the tedious tasks`nthat come along with creating locals and detaching models in Revit
Gui, Add, Text, w500, Nothing will be written to the registry, and no desktop icons will be added.`nA shortcut will be placed in the start menu.
Gui, Add, Text, w500 vInstallText, By proceeding, %programName% will be installed in:
Gui, Font, s14, Arial
Gui, Add, Text, w500 vInstallText2, %installFolder%\
Gui, Font, s18, Arial
Gui, Add, Button, w150 xp+10 yp+100 default vInstallB, &Install
Gui, Add, Button, w150 xp+165 vUninstallB, &Uninstall
Gui, Add, Button, w150 xp+165, &Cancel
Gui, Font, s10, Arial
Gui, Add, Text, w500 xp-340 yp+75, %programName% was specifically developed for Wright Heerema | Architects`nby Michael Pfammatter
GuiControl, Disable, UninstallB
IfExist, %installFullPath%
{
	FileGetVersion, oldVersion, %installFullPath%
	; oldVersion = 0.0.1.0
	If oldVersion >= %newVersion%
	{
		GuiControl, Disable, InstallB
		GuiControl, Enable, UninstallB
		GuiControl, +Default, UninstallB
		GuiControl, Text, InstallText, Congratulations!
		GuiControl, Text, InstallText2, Version %oldVersion% is already installed on your device.
	}
	Else
	{
		GuiControl, Text, InstallText, By proceeding %programName% will be upgrade from:
		GuiControl, Text, InstallText2, Version %oldVersion% to Version %newVersion%
		GuiControl, Text, InstallB, &Upgrade
		GuiControl, Enable, UninstallB
		upgrade := 1
	}
	
}


Gui, Show
Return

GuiClose:
GuiEscape:
ButtonCancel:
ExitApp

ButtonInstall:
; MsgBox, %A_ScriptDir% to %installFolder%`n`n%A_StartMenuCommon%
MsgBox, 33,, Proceed with installing %programName%?
IfMsgBox Cancel
	Return
Gui, Destroy
Install:
FileCopyDir, %A_ScriptDir%, %installFolder%, 1
If ErrorLevel
{
	MsgBox, 16,, There was an issue installing %programName% to %installFolder%. Please see your BIM Coordinator.
	ExitApp
}
IfNotExist, %A_StartMenuCommon%\%programName%%versionSuffix%.lnk
	FileCreateShortcut, %installFullPath%, %A_StartMenuCommon%\%programName%%versionSuffix%.lnk
	If ErrorLevel
		shortcutMessage = `n`nA valid shortcut could not be create in the Start Menu. Go to %installFolder% to locate the program.
MsgBox, 1,, Congratulations! %programName% was successfully installed.%shortcutMessage%
ExitApp

ButtonUninstall:
folderCount := 0
fileCount := 0
MsgBox, 33,, Are you sure you would like to Uninstall %programName%?
IfMsgBox Cancel
	Return
Gui, Destroy
; Remove the main install directory
FileRemoveDir, %installFolder%, 1
If ErrorLevel
	MsgBox, The Program could not be uninstalled!
Else
	folderCount += 1
localFolder = %A_AppData%\%programName%
; Remove any local data
IfExist, %localFolder%
{
	FileRemoveDir, %localFolder%, 1
	If ErrorLevel
		MsgBox, The local data could not be removed. This is not a big deal but files may have been left behind.
	Else
		folderCount += 1
}
IfExist, %localFolder%%versionSuffix%
{
	FileRemoveDir, %localFolder%, 1
	If ErrorLevel
		MsgBox, The local data could not be removed. This is not a big deal but files may have been left behind.
	Else
		folderCount += 1
}

; Remove startup shortcut
FileDelete, %A_Startup%\%programName%.lnk
If !ErrorLevel
	fileCount += 1
FileDelete, %A_Startup%\%programName%%versionSuffix%.lnk
If !ErrorLevel
	fileCount += 1
	

; Remove start menu shortcut
FileDelete, %A_StartMenuCommon%\%programName%.lnk
If !ErrorLevel
	fileCount += 1
FileDelete, %A_StartMenuCommon%\%programName%%versionSuffix%.lnk
If !ErrorLevel
	fileCount += 1

If folderCount > 0
	MsgBox, 48,, %programName% was successfully uninstalled.`n`nDirecotries Deleted:%folderCount%`nShortcuts Deleted:%fileCount%
Else
	MsgBox, 16,, We encountered an error uninstalling %programName%.  We apologize for the inconvenience. 
ExitApp
