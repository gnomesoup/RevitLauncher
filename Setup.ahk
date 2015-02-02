#Persistent
#SingleInstance
#NoTrayIcon
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Perform an install/uninstall of the WHA Revit Launcher Program
programName = WHA Revit Launcher
startVar = %1%
silent := if startVar = "silent" ? true : false
uninstall := if startVar = "uninstall" ? true : false

; Check if setup is run with admin privileges
If !A_IsAdmin
{
	PrettyMsg("You do not have administrator privileges.`n`nTo install " . programName . ", please re-run Setup.exe with administrator privileges.", "exit")
}


FileGetVersion, newVersion, %programName%.exe
versionSuffix := ""
installFolder = %A_ProgramFiles%\%programName%%versionSuffix%
installFullPath = %installFolder%\%programName%.exe
; ListVars
upgrade := 0


; If the script is run in silent mode, check for running process and skip GUI
if silent
{
	FileGetVersion, oldVersion, %installFullPath%
	If oldVersion >= %newVersion%
		ExitApp
	Process, Exist, %programName%.exe
	if ErrorLevel
		Runwait, taskkill /im "%programName%.exe" /f
	upgrade := if FileExist(installFullPath) ? true : false
	gosub, Install
}

; If the script is run in uninstall mode, check for running process and skip GUI
if uninstall
{
	Process, Exist, %programName%.exe
	if ErrorLevel
		Runwait, taskkill /im "%programName%.exe" /f
	gosub, Uninstall
}

; Check if Revit Launcher is running and close it if it is.
Process, Exist, %programName%.exe
if ErrorLevel
{
	if PrettyMsg("Changes cannot be made to " . programName . " while it is running.`n`nWould you like us to close it for you?", "question", 2) == "Yes"
		Runwait, taskkill /im "%programName%.exe" /f
	else
		ExitApp
}
Process, Exist, %programName%.exe
if ErrorLevel
	PrettyMsg("Setup was unable to close " . programName . ". Sorry for the trouble.`n`nSetup will now exit. Please close the program and try running setup again.", "exit")

setupWidth := 450
setupBcount := 3
setupBwidth := (setupWidth - (15 * (setupBcount-1))) / setupBcount
setupBloc := setupBwidth + 15
Gui, Setup:New
Gui, Setup:Font, s22cBlack, Arial
Gui, Setup:Add, Text, xm w%setupWidth% center, Install %programName%
Gui, Setup:Font, s10 c666666, Arial
Gui, Setup:Add, Text, w%setupWidth% center yp+35, Version: %newVersion%
Gui, Setup:Font, s10 cBlack, Arial
Gui, Setup:Add, Text, w%setupWidth%, This little program automates many of the tedious tasks that come along with creating locals and detaching models in Revit. 
Gui, Setup:Add, Text, w%setupWidth%, This program writes a setting to your registry in order to allow right click creation of a Dated Folder. A shortcut will also be placed in the start menu.
Gui, Setup:Add, Text, w%setupWidth% vInstallText, By proceeding, %programName% will be installed in:
Gui, Setup:Font, s14, Arial
Gui, Setup:Add, Text, w%setupWidth% vInstallText2, %installFolder%\
Gui, Setup:Font, s18, Arial
Gui, Setup:Add, Button, w%setupBwidth% xm yp+100 default vInstallB, &Install
Gui, Setup:Add, Button, w%setupBwidth% xp+%setupBloc% vUninstallB, &Uninstall
Gui, Setup:Add, Button, w%setupBwidth% xp+%setupBloc%, &Cancel
Gui, Setup:Font, s10, Arial
Gui, Setup:Add, Text, w%setupWidth% xm yp+75, %programName% was specifically developed by Michael Pfammatter for:
whaLogo(setupWidth, "Setup")
Gui, Setup:Font, s10, Arial
GuiControl, Setup:Disable, UninstallB
IfExist, %installFullPath%
{
	FileGetVersion, oldVersion, %installFullPath%
	; oldVersion = 0.0.1.0
	If oldVersion >= %newVersion%
	{
		GuiControl, Setup:Disable, InstallB
		GuiControl, Setup:Enable, UninstallB
		GuiControl, Setup:+Default, UninstallB
		GuiControl, Setup:Text, InstallText, Congratulations!
		GuiControl, Setup:Text, InstallText2, Version %oldVersion% is already installed on your device.
	}
	Else
	{
		GuiControl, Setup:Text, InstallText, By proceeding %programName% will be upgrade from:
		GuiControl, Setup:Text, InstallText2, Version %oldVersion% to Version %newVersion%
		GuiControl, Setup:Text, InstallB, &Upgrade
		GuiControl, Setup:Enable, UninstallB
		upgrade := 1
	}
	
}


Gui, Setup:Show
Return

SetupGuiClose:
SetupGuiEscape:
SetupButtonCancel:
ExitApp

SetupButtonInstall:
proceed := PrettyMsg("Proceed with installing " . programName . "?", "question")
if proceed = no
	ExitApp
else if !proceed
	return

gosub, progressGUI

Gui, Setup:Destroy

Install:
FileCopyDir, %A_ScriptDir%, %installFolder%, 1
If ErrorLevel
{
	PrettyMsg("There was an issue installing " . programName . " to " . installFolder . ". Please see your BIM Coordinator.", "exit")
}
IfNotExist, %A_StartMenuCommon%\%programName%%versionSuffix%.lnk
{
	FileCreateShortcut, %installFullPath%, %A_StartMenuCommon%\%programName%%versionSuffix%.lnk
	If ErrorLevel
		shortcutMessage = `n`nA valid shortcut could not be create in the Start Menu. Go to %installFolder% to locate the program.
}

; Write Dated Folder values to registry
RegWrite, REG_SZ, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, Icon, "%installFolder%\Support\wha.ico"
addError := ErrorLevel
RegWrite, REG_SZ, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, MUIVerb, WHA Dated Folder
addError += ErrorLevel
RegWrite, REG_SZ, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, Position, Bottom
addError += ErrorLevel
RegWrite, REG_SZ, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder\command,, "%installFolder%\WHA Dated Folder Creator.exe" "`%v"
addError += ErrorLevel
Gui, progressGUI:Destroy
if !silent
	PrettyMsg(programName . " was successfully installed.", "success")
ExitApp

SetupButtonUninstall:
folderCount := 0
fileCount := 0
proceed := PrettyMsg("Are you sure you would like to Uninstall " . programName . "?", "question")
if proceed = no
	ExitApp
else if !proceed
	return
Gui, Setup:Destroy

Uninstall:
gosub, ProgressGUI
GuiControl, progressGUI:, progressText, Please wait while %programName% is uninstalled.
; Remove the main install directory
FileRemoveDir, %installFolder%, 1
If ErrorLevel
	PrettyMsg("The Program could not be uninstalled!, Please see your BIM Coordinator.")
Else
	folderCount += 1
localFolder = %A_AppData%\%programName%
; Remove any local data
IfExist, %localFolder%
{
	FileRemoveDir, %localFolder%, 1
	If ErrorLevel
		PrettyMsg("The local data could not be removed. This is not a big deal but files may have been left behind.")
	Else
		folderCount += 1
}
IfExist, %localFolder%%versionSuffix%
{
	FileRemoveDir, %localFolder%, 1
	If ErrorLevel
		PrettyMsg("The local data could not be removed. This is not a big deal but files may have been left behind.")
	Else
		folderCount += 1
}

; Remove WHA Dated Folder registry values
RegDelete, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, Icon
RegDelete, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, MUIVerb
RegDelete, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder, Position
RegDelete, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder\command
RegDelete, HKEY_CLASSES_ROOT, Directory\Background\shell\WHADatedFolder

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
Gui, progressGUI:Destroy

if !uninstall
{
	If folderCount > 0
		PrettyMsg(programName . " was successfully uninstalled.`n`nDirecotries Deleted:" . folderCount . "`nShortcuts Deleted:" . fileCount, "success")
	Else
		PrettyMsg("We encountered an error uninstalling " . programName . ".  We apologize for the inconvenience.", "exit")
}


ExitApp

ProgressGUI:
progressGUIwidth := 450
Gui, progressGUI:new
whaLogo(progressGUIwidth, "progressGUI")
Gui, progressGUI:Font, s12 cBlack, Arial
Gui, progressGUI:Add, Text, w%progressGUIwidth% xm center vprogressText, Please wait while %programName% is installed
Gui, progressGUI:Show
return