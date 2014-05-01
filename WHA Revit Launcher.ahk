#Persistent
#SingleInstance, Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; ©2014 Wright Heerema | Architects
; Created by Michael Pfammatter

; This utility will launch a Revit file for you using the appropriate version of the software as well as
; automate a few things you don't want to be doing over and over and over again.  

; ### Set Variables ###
SetTitleMatchMode, RegEx ;Allows for less precise window finding

; Set the name of the program. This should be the same as the named exe file
; I don't grab this using A_ variables in case a user changes the filename
If A_IsCompiled
	programName = WHA Revit Launcher
Else
	programName = Revit Launcher Dev

; Location where icons and images are kept
supportFolder = %A_ScriptDir%\Support

; Name of the bimGuy. I'm trying to keep it generic but this may need to 
; change in the future
bimGuy = your BIM Coordinator

; Global variable need to be set later in a function
mouseX =
mouseY =

; Allow for some color consistency in menus
guiColor1 = 666666
guiColor2 = 9f1d21
guiColor3 = ffffff

; Set the path of the file with local settings.
; If a user ini file doesn't exist, create one.
iniPathLocal = %A_AppData%\%programName%\%programName%.ini
IfNotExist %iniPathLocal%
{
	IfNotExist %A_AppData%\%programName%
		FileCreateDir, %A_AppData%\%programName%
	FileCopy, %A_ScriptDir%\%programName%.ini, %iniPathLocal%
	If ErrorLevel ; If copying the file goes wrong, notify and exit
	{
		MsgBox, 1, %programName%, We were unable to properly initialize %programName%.  Please see %bimGuy% for additional information.
		ExitApp
	}
	; As long as we are setting things up, make sure the program starts when the user logs including
	; This can be turned off in settings
	linkFile = %A_Startup%\%programName%.lnk
	IfNotExist %LinkFile%
		FileCreateShortcut, %A_ScriptFullPath%, %LinkFile%
}

;load up the local settings and favorites
iniLocal := class_EasyIni(iniPathLocal) 
If (!iniLocal.Settings.iniPathCentral) ;Check if settings file is good
{
	MsgBox, 16, %programName%, A file named "%iniPathLocal%" either could not be found in the same directory as this executable or it was manually altered.  This program will now exit.`n`nPlease contact your friendly BIM Coordinator.
	ExitApp
}

; Load up the global project list
iniPathCentral := iniLocal.Settings.iniPathCentral
iniCentral := class_EasyIni(iniPathCentral)

; Get some default settings
detach := iniLocal.Settings.Detach
workset := iniLocal.Settings.Workset
debug := iniLocal.Settings.Debug

; Debug function to make finding problems easier
DebugMe("Start")

; By default we assume Revit is not running
revitRunning := 0

; Create the main tray menu that will be the main interface
GoSub, TrayMenu

; This is all that is done when the program is started
; Any additional features will need to be initialized by the user
Return
; ### End of Set Variables ###




; ### Create the right click menu ###
TrayMenu:

; Read list of favorites
favList := iniLocal.GetSections(, "C")

; Turn favList into an array
StringSplit, favList, favList, `n

; Set a variable to count the number of favorites a user has
favN := 0

; Go through a users ini file and set variables for all the favorites
Loop, %favList0%
{
	favSection := favList%A_Index%
	if instr(favSection, "Favorite")
	{
		favN += 1
		fav%favN% := iniLocal [favSection].ProjectID
		Fav%favN%Name := iniLocal [favSection].Name
		favVersion%favN% := iniCentral [fav%favN%].Version
	}
}
	
;Populate right click menu including favorites
Menu, tray, NoStandard ;Get rid of default menu items
Menu, tray, Icon,  %supportFolder%\wha.ico
Menu, tray, Tip, Wright Heerema | Architects`nRevit Launcher
;Give different menu options if the S drive is not available
IfNotExist %iniPathCentral%
{
	Menu, tray, Icon,  %supportFolder%\alert.ico
	Menu, tray, Add, Error! Central Project List Not Found, TrayReset
	Menu, tray, Icon, Error! Central Project List Not Found, %supportFolder%\alert.ico
}
Else
{
	Menu, tray, Add, Find Project and Launch, FindLaunch
	Menu, tray, icon, Find Project and Launch, %supportFolder%\search.ico
	Menu, tray, Add
	Menu, tray, Add, Quick Launch Add/Remove, AddRemove
	Menu, tray, icon, Quick Launch Add/Remove, %supportFolder%\star.ico
	Loop, % FavN ;search through list of favorites and make menu items for them
	{
		favTitle := fav%A_Index%Name
		favIcon := favVersion%A_Index%
		Menu, tray, Add, %favTitle%, favSub
		If detach
			Menu, tray, Icon, %favTitle%, %supportFolder%\detach.ico, 1, 16
		Else
			Menu, tray, Icon, %favTitle%, %supportFolder%\Revit%favIcon%file.ico, 1, 16
		; MsgBox, Added %A_Index%
	}
}
Menu, tray, Add
Menu, tray, Add, Detach Models, DetachSub
If detach
	Menu, tray, Check, Detach Models
Menu, tray, Add, Specify Worksets, WorksetSub
if workset
	Menu, tray, Check, Specify Worksets
Menu, tray, Add
Menu, tray, Add, Settings, SettingsSub
Menu, tray, Icon, Settings, %supportFolder%\settings.ico
Menu, tray, Add, %programName% Help, HelpSub
Menu, tray, Add, Exit, ExitSub
OnMessage(0x404, "AHK_NOTIFYICON") ;Allow left-click of tray icon
return

; What to do if the tray icon is left-clicked
RightShow:
Menu, tray, Show
Return

; Assign project variables and run the main routine
FavSub: 
menuID := A_ThisMenuItemPos - 3
projectID := fav%menuID%
GoSub, MainRoutine
return

; If the server is not available, check again when the user requests
TrayReset:
MsgBox, 53, %programName%, The centralized list of projects could not be found.`nThis may be due to the network not being available.  Please check the network and try again.  Consult %bimGuy% should this problem persist.
IfMsgBox Retry
	ReloadMe()
return

; Swap out the icons and prepare to launch all models detached
DetachSub:
detach := MenuCheck(detach, "Detach Models")
; iniLocal.Settings.Detach := detach
; iniLocal.Save()
ReloadMe()
return

; Prepare to launch all models with the "Specify Worksets" dialog box in Revit
WorksetSub:
workset := MenuCheck(workset, "Specify Worksets")
; iniLocal.Settings.Workset := workset
; iniLocal.Save()
TrayOpen()
return

ExitSub: ;Leaving so soon?
ExitApp
return
; ### End of right click menu ###



; ### Create the Search and Launch Menu ###
; Menu that is launched when a user wants to open a project they don't access frequently. 
FindLaunch:
; Load up the global project list
iniCentral := class_EasyIni(iniPathCentral)
; Get all of the projects
centralSections := iniCentral.GetSections(,"C")
; Turn list of projects into an array for looping
StringSplit, centralSections, centralSections, `n
; Create the actual menu
Gui, FindLaunch:New,, %programName%
Gui, FindLaunch:Font, s24 , Arial
Gui, FindLaunch:Add, Text, center w500, Select a project to Launch:
Gui, FindLaunch:Font, s10 , Arial
Gui, FindLaunch:Add, ListView, AltSubmit r10 w500 gFindLaunchList -Multi, Number|Name
; Add projects to the listview
loop, %centralSections0%
{
	LV_Add(,iniCentral [centralSections%A_Index%].Number, iniCentral [centralSections%A_Index%].Name)
}
LV_ModifyCol(1)
Gui, FindLaunch:Font, s10 c%guiColor1%, Arial
Gui, FindLaunch:Add, Text, w500 center yp+225, A local file will be created where appropriate and`nopened with the correct version of Revit
Gui, FindLaunch:Font, s18 cBlack, Arial
; Change the button to reflect if we are opening detached or specifed worksets
If detach
	Gui, FindLaunch:Add, Button,  w150 yp+50 xp+75 gFindLaunchFind, &Detach
Else If workset
	Gui, FindLaunch:Add, Button,  w150 yp+50 xp+75 gFindLaunchFind, &Specify Workset
Else
	Gui, FindLaunch:Add, Button,  w150 yp+50 xp+75 gFindLaunchFind, &Launch
Gui, FindLaunch:Add, Button,  wp xp+200 Default gFindLaunchGuiClose, &Cancel
GuiControl, FindLaunch:Disable, Button1
Gui, FindLaunch:Show
Return

; Enable the launch button once the user selects a project
FindLaunchList:
If A_GuiEvent = I
	GuiControl, FindLaunch:Enable, Button1
If A_GuiEvent = DoubleClick
	GoSub, FindLaunchFind
Return

; Get the project selected, set project variables, and run the main routine
FindLaunchFind:
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
projectID := centralSections%RowNumber%
DebugMe("FindLaunchFind")
Gui, Destroy
GoSub, MainRoutine
Return
; ### End of the Search and Launch Menu ###



; ### Create the Add/Remove Quick Launch Menu ###
; Menu that allows users to add or remove programs they frequent to the tray menu
; First we simply ask if they want to add to the list or remove from the list
AddRemove:
Gui, AddRemove:New,, %programName%
Gui, AddRemove:Font, s24, Arial
Gui, AddRemove:Add, Text, center w650, What would you like to do?
Gui, AddRemove:Font, s18, Arial
Gui, AddRemove:Add, Button, Default w300 gAddProject, &Add a Project
Gui, AddRemove:Add, Button, wp xp+350 gRemoveProject, &Remove a Project
Gui, AddRemove:Show
return

; Menu to add a project to the quick launch list
AddProject: 
; Load up the global list of projects
iniCentral := class_EasyIni(iniPathCentral)
; Get all of the projects in the list
centralSections := iniCentral.GetSections(,"C")
; Create an array of the projects for looping
StringSplit, centralSections, centralSections, `n
; Create the add project menu
Gui, AddProject:New,, %programName%
Gui, AddProject:Font, s18 , Arial
Gui, AddProject:Add, Text, center w500, Select a project to add to your project list:
Gui, AddProject:Font, s10 , Arial
Gui, AddProject:Add, ListView, AltSubmit r10 w500 gAddProjectList -Multi, Number|Name
; Add all of the projects to the listview
loop, %centralSections0%
{
	LV_Add(,iniCentral [centralSections%A_Index%].Number, iniCentral [centralSections%A_Index%].Name)
}
LV_ModifyCol(1)
Gui, AddProject:Font, s10 c%guiColor1%, Arial
Gui, AddProject:Add, Text, w500 center yp+250, The selected project will be added to your "Quick Launch" menu.`nA local file will be created and opened using the correct version of Revit.`nNeed a project added? Let %bimGuy% know.
Gui, AddProject:Font, s18 cBlack, Arial
Gui, AddProject:Add, Button,  w150 yp+50 xp+75 gAddProjectAdd, &Add to list
Gui, AddProject:Add, Button,  wp xp+200 Default gAddProjectGuiClose, &Cancel
GuiControl, Disable, &Add to list
Gui, AddProject:Show
; Destroy the menu that asks if we want to add or remove
Gui, AddRemove:Destroy
return

; Menu to remove a program from the Quick Launch list
RemoveProject:
; Create the Remove from Quick Launch menu
Gui, RemoveProject:New,, %programName%
Gui, RemoveProject:Font, s18, Arial
Gui, RemoveProject:Add, Text, center w500, Select a project to remove:
Gui, RemoveProject:Font, s10, Arial
Gui, RemoveProject:Add, ListView, AltSubmit r10 w500 gRemoveProjectList -Multi, Name
; Add a list of all the projects in the Quick Launch list to the list view
loop, %favList0%
{
	LV_Add(,iniLocal [favList%A_Index%].Name)
}
Gui, RemoveProject:Font, s10 c%guiColor1%, Arial
Gui, RemoveProject:Add, Text, w500 center yp+250, The selected project will be removed from your "Quick Launch" menu.
Gui, RemoveProject:Font, s18 cBlack, Arial
Gui, RemoveProject:Add, Button,  w250 yp+50 xp+25 gRemoveProjectRemove, &Remove from list
Gui, RemoveProject:Add, Button,  w150 xp+300 Default gRemoveProjectGuiClose, &Cancel
GuiControl, Disable, &Remove from list
Gui, RemoveProject:Show
; Destroy the menu that asks if we want to add or remove
Gui, AddRemove:Destroy
return

; Exit out of a whole bunch of menus
AddRemoveGuiEscape:
AddRemoveGuiClose:
AddProjectGuiEscape:
AddProjectGuiClose:
RemoveProjectGuiEscape:
RemoveProjectGuiClose:
FindLaunchGuiEscape:
FindLaunchGuiClose:
Gui, Destroy
; Once the menu is destroyed, reopen the tray menu to let users know it is still down there
TrayOpen()
return

; Enable the remove button once a users selects a project
RemoveProjectList:
If A_GuiEvent = I
	GuiControl, RemoveProject:Enable, Button1
If A_GuiEvent = DoubleClick
	GoSub, RemoveProjectRemove
Return

; This does the hard work of actually REMOVING the project from the Quick Launch list
RemoveProjectRemove:
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
MsgBox, 33, Remove Project?, % "Are you sure you want to remove the following project from your list?`n`n" . iniLocal [favList%RowNumber%].Name
IfMsgBox OK
{
	iniLocal.DeleteSection(favList%RowNumber%)
	iniLocal.Save()
	Gui, RemoveProject:Destroy
	ReloadMe()
}
return

; Enable the add button once a user selects a project
AddProjectList:
If A_GuiEvent = I
	GuiControl, AddProject:Enable, Button1
If A_GuiEvent = DoubleClick
	GoSub, AddProjectAdd
Return

; This does the hard work of actually ADDING the project to the Quick Launch list
AddProjectAdd:
; Make sure you zero out the RowNumber otherwise you could see some weird behaviour
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
newFavID := centralSections%RowNumber%
newFavSection = Favorite%newFavID%
newFavNumber := iniCentral [newFavID].Number
newFavName := iniCentral [newFavID].NameShort

if instr(favList, newFavSection)
{
	MsgBox, 64, %programName%, %newFavNumber% %newFavName% is already in your list of projects
}
else
{
	MsgBox, 33, %programName%, You are sure you want to add the following project to your list?`n`n%newFavNumber% - %newFavName%
	IfMsgBox OK
	{
		iniLocal.AddSection(newFavSection, "ProjectID", newFavID)
		iniLocal.AddKey(newFavSection, "Name", newFavNumber . " " . newFavName)
		iniLocal.AddKey(newFavSection, "Detach", "0")
		iniLocal.AddKey(newFavSection, "Workset", "0")
		iniLocal.save()
		Gui, AddProject:Destroy
		ReloadMe()
	}
}
return
; ### End of Add/Remove Quick Launch Menu ###



; ### Create the Settings Menu ###
SettingsSub:
detachDefault := iniLocal.Settings.Detach
worksetDefault := iniLocal.Settings.Workset
linkFile = %A_Startup%\%programName%.lnk
IfExist %linkFile%
	startupExist := 1
Else
	startupExist := 0
Gui, Settings:New,, %programName%
Gui, Settings:Font, s24 cBlack, Arial
Gui, Settings:Add, Text, w500, %programName% Settings:
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Text, yp+80 section, Defaults
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vdetachCheck gSettingsSubDetach, Detach by default
GuiControl,, detachCheck, %detachDefault%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+30 w400, Controls whether "Detach Models" is enabled when you first open %programName%.  This option is perfect Project Managers who only want to query a model.  You can still temporarily disable it if you need to update drawings.
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+90 section, 
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vworksetCheck gSettingsSubWorkset, Specify worksets by default
GuiControl,, worksetCheck, %worksetDefault%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+30 w400, Controls whether the "Specify Worksets" dialog box is opened by default.  Choose this option if you are working on larger Revit models.
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+60 section, Startup
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vstartupCheck gSettingsSubStartup, "Startup on Windows Login"
GuiControl,, startupCheck, %startupExist%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+30 w400, If checked, %programName% will open automatically when you log in to Windows.  Just set it and forget it. :)
Gui, Settings:Show
return

SettingsGuiEscape:
SettingsGuiClose:
ReloadMe()
return

SettingsSubDetach:
; msgBox, % !detachDefault
iniLocal.Settings.Detach := !detachDefault
iniLocal.Save()
detachDefault := iniLocal.Settings.Detach
GuiControl,, detachCheck, %detachDefault%
detach := detachDefault
return

SettingsSubWorkset:
; msgBox, % !detachDefault
iniLocal.Settings.Workset := !worksetDefault
iniLocal.Save()
worksetDefault := iniLocal.Settings.Workset
GuiControl,, worksetCheck, %worksetDefault%
workset := worksetDefault
return

SettingsSubStartup:
DebugMe("Settings Startup Check")
IfNotExist, %LinkFile%
{
	FileCreateShortcut, %A_ScriptFullPath%, %LinkFile% 
	If ErrorLevel
	{
		MsgBox, 1, %programName%, There was a problem creating a shortcut in your startup folder, %A_Startup%.  Please contact %bimGuy% for additional information.
	}
}
Else
{
	FileDelete, %LinkFile%
	If ErrorLevel
	{
		MsgBox, 1, %programName%, There was a problem removing the startup shortcut from your startup folder, %A_Startup%. Please contact %bimGuy% for additional information.
	}
}
IfExist %linkFile%
	startupExist := 1
Else
	startupExist := 0
GuiControl,, startupCheck, %startupExist%
Return
; ### End of Settings Menu ###



; ### Create the Help Menu ###
HelpSub:
MsgBox, Help is coming soon...
TrayOpen()
return
; ### End of Help Menu ###



; ### Main Routine ###
; Routine that actually does the work of creating backups and opening locals
MainRoutine:
iniCentral := class_EasyIni(iniPathCentral)
pNumber := iniCentral [projectID].Number
pName := iniCentral [projectID].Name
pNameShort := iniCentral [projectID].NameShort
pVersion := iniCentral [projectID].Version
pCentral := iniCentral [projectID].Central
IfNotExist, %pCentral% ;Check if project is setup correctly
	{
		MsgBox, 16, Launcher Error, There is an issue locating the project associated with %projectID%.`nPlease see your friendly BIM manager for additional information.`n`nHave a nice day!
		TrayOpen()
	}
revitUser = %A_Username% ;Users computer username
localFolder = %A_MyDocuments%\Revit ;Users folder where centrals will be saved
projectFolder = %localFolder%\%projectID% %pNameShort% ;Project folder name
localFile = %projectID% %pNameShort% LOCAL %revitUser% ;Local file name no extension
localPath = %projectFolder%\%localFile%.rvt ;Local file name & path

FindRevit:
;Find the path to the correct Revit flavor depending on the version read from the ini file
IfWinExist, %localFile%.rvt ;Check to see if the file is already open
{
	WinActivate
	WinMaximize	
	MsgBox, 16, %programName%, Revit is already running with the specified local file:`n%localFile%.rvt
	ReloadMe()
}
IfEqual, pVersion, 2013
{
	revitPath = %A_ProgramFiles%\AutoDesk\Revit 2013\Program\Revit.exe
	revitTitle = Autodesk Revit 2013
	IfNotExist %revitPath%
	{
		revitPath = %A_ProgramFiles%\Autodesk\Revit Architecture 2013\Program\Revit.exe
		revitTitle = Autodesk Revit Architecture 2013
	}
}
Else
{
	revitPath = %A_ProgramFiles%\AutoDesk\Revit %pVersion%\Revit.exe
	revitTitle = Autodesk Revit %pVersion%
	IfNotExist %revitPath%
	{
		revitPath = %A_ProgramFiles%\Autodesk\ Revit Architecture %pVersion%\Revit.exe
		revitTitle = Autodesk Revit Architecture %pVersion%
	}
}
IfNotExist, %revitPath%
{
	MsgBox, 16, %programName%, Revit %pVersion% cannot be found on this computer.  Please see %bimGuy% for additional information.
	ReloadMe()
}

FindMonitor: 
;Check for a Worksharing Monitor for the right version of Revit
monitorPath = %A_ProgramFiles% (x86)\Autodesk\Worksharing Monitor for Autodesk Revit %pVersion%\WorksharingMonitor.exe
monitorTitle = Worksharing Monitor for Autodesk Revit %pVersion%
IfNotExist, %monitorPath%
{
	MsgBox, 64, Worksharing Monitor, The Worksharing Monitor could not be found on your system.  Please notify your BIM manager so you can play well with others., 7
	monitorPath =
}

DebugMe("LaunchSplash")
Gosub, LaunchSplash
If !detach
{
	launchStatus = Copying local...
	GuiControl,, LaunchText, %launchStatus%
	DebugMe("CreateDir")
	;Check if Directory Exists
	IfNotExist, %projectFolder%
	{
		DebugMe("project folder not exist")
		FileCreateDir, %projectFolder% ;Create a directory if it doesn't
		If Errorlevel
		{
			MsgBox, 16, Could not create directory, A directory structure could not be created in %localFolder%.
			ExitApp
		}
	}

	LocalBackup:
	DebugMe("LocalBackup")
	;Create a backup of the last local created
	IfExist, %projectFolder%\%localFile%.1.rvt ;backup of a backup
		fileMove, %projectFolder%\%localFile%.1.rvt, %projectFolder%\%localFile%.2.rvt, 1
	IfExist, %localPath% ;backup of the local
		fileMove, %localPath%, %projectFolder%\%localFile%.1.rvt, 1

	DebugMe("LocalCreate")
	FileCopy, %pCentral%, %localPath%, 1
	if ErrorLevel ;Check to make sure everything copied alright
	{
		MsgBox, 16, Local Creation Error, The local file could not be copied to your computer.  Please see your BIM manager for additional information.
		ExitApp
	}
}

DebugMe("Launch")
launchStatus = Launching Revit %pVersion%
GuiControl,, LaunchText, %launchStatus%
;Open Revit with the correct version
If monitorPath ;Open the worksharing monitor if it exists
	IfWinNotExist, %monitorTitle%
		Run, %monitorPath%
IfWinExist, ^%revitTitle% ;Check to see if Revit is already running
{
	WinActivate
	WinMaximize
}
Else
{
	Run, %revitPath%, Max, %localPath%
	WinWait, ^%revitTitle%
	WinMaximize
}
WinActivate
Send ^o
WinWait, Open
Sleep 300
WinGet, openID, ID, Open
ControlGet, fileHwnd, Hwnd,, Edit1, ahk_id %openID%
ControlGet, folderHwnd, Hwnd,, SysListView321, ahk_id %openID%
ControlGet, openHwnd, Hwnd,, Button1, ahk_id %openID%
; MsgBox Open Window: %openID%`nFile name: %fileHwnd%`nFolderView:%folderHwnd%`nOpen Button: %openHwnd%
If detach
	fName := pCentral
Else
	fName := localPath
If workset
{
	SplitPath, fName, fName, fPath
	; MsgBox, 1, Workset Flow, Workset: %workset% File: %fPath%
	Control, EditPaste, %fPath%, , ahk_id %fileHwnd%
	ControlClick, Button1, Open,, L, 2, NA
	Sleep 600
	ControlSend, SysListView321, %fName%, Open
	ControlSend, Button1, {Down 6}{Enter}, Open
}
Else
{
	; msgbox, 1, Regular Flow, Workset: %workset% File: %fName%
	Control, EditPaste, %fName%, Edit1, Open
}
If detach
{
	ControlClick, Button5, Open,, L
}
ControlClick, Button1, Open,, L, 2, NA
Sleep 200
If workset
{
	WinWait, Opening Worksets,, 60
	Gui, Launch:Destroy
	Sleep 200
	WinWait, Copied Central Model,, 60
	ControlClick, Button1, Copied Central Model,, L, 2, NA
}
Else If detach
{
	Sleep 300
	WinWait, Detach Model from Central,, 60
	ControlClick, Button1, Detach Model from Central,, L
	Gui, Launch:Destroy
}
Else
{
	WinWait, Copied Central Model,, 60
	ControlClick, Button1, Copied Central Model,, L, 2, NA
	Gui, Launch:Destroy
}
WinActivate, ^%revitTitle%
Return

LaunchSplash:
Gui, Launch:New
If debug
	Gui, Launch:-Caption ; AlwaysOnTop 
Else
	Gui, Launch:-Caption AlwaysOnTop 
Gui, Launch:Margin, 50, 50
Gui, Launch:Color, DCDCDC
Gui, Launch:Add, Picture, AltSubmit w700 h-1, %supportFolder%\WRIGHT-HEEREMA-logo.png
Gui, Launch:Font, cBlack s36, Arial
Gui, Launch:Add, Text,yp+150 center w700 vLaunchText, Launching...
Gui, Launch:Font, c%guiColor1% s18, Arial
Gui, Launch:Add, Text,yp+70 center w700, %pNumber% %pName%
If detach
{
	Gui, Launch:Font, c%guiColor2% s24, Arial
	Gui, Launch:Add, Text, yp+70 w700 center, Detached
}
Gui, Launch:Show
return
; ### End of Main Routine ###



; ### Functions ###
AHK_NOTIFYICON(wParam, lParam)
{
    Global
	CoordMode, Mouse, Screen
	MouseGetPos, mouseX, mouseY
	CoordMode, Mouse, Window
	if (lParam = 0x202) ; WM_LBUTTONUP
    {
        SetTimer, RightShow, -1
        return 0
    }
}

MenuCheck(mcVar, mcTitle)
{
	mcVar := !mcvar
	if mcvar
		Menu, tray, Check, %mcTitle%
	else
		Menu, tray, Uncheck, %mcTitle%
	Return mcVar
}

trayOpen() ;open the tray menu in the lower right corner
{
	Global
	CoordMode, Menu, Screen
	Menu, Tray, Show, %mouseX%, %mouseY%
	CoordMode, Menu, Window
	Return True
}

ReloadMe()
{
	Global
	Menu, Tray, DeleteAll
	Gui, Destroy
	GoSub, TrayMenu
	CoordMode, Menu, Screen
	Menu, Tray, Show, %mouseX%, %mouseY%
	CoordMode, Menu, Window
	Exit
	Return True
}

MouseCleanClick(x, y)
{
MouseGetPos, px, py
; MsgBox, %x%,%y%`n%px%,%py%
Click %x%, %y%
Click %px%, %py%, 0
}

DebugMe(debugText)
{
	Global
	If debug
	{
		ListVars
		MsgBox, 1, Debug, %debugText%
		IfMsgBox Cancel
		ReloadMe()
	}
}

BetterMsgBox(message)
{

}
Return