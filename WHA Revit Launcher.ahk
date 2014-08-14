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
{
	programName = WHA Revit Launcher
	FileGetVersion, scriptVersion, %A_ScriptFullPath%
}
Else
{
	programName = Revit Launcher Dev
	scriptVersion = LaunchDev
}

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

; Set the path of the file with local settings".
; If a user ini file doesn't exist, create one.
iniPathLocal = %A_AppData%\%programName%\%programName%.ini
IfNotExist %iniPathLocal%
{
	IfNotExist %A_AppData%\%programName%
		FileCreateDir, %A_AppData%\%programName%
	FileCopy, %A_ScriptDir%\%programName%.ini, %iniPathLocal%
	If ErrorLevel ; If copying the file goes wrong, notify and exit
	{
		MsgBox, 16, %programName%, We were unable to properly initialize %programName%.  Please see %bimGuy% for additional information.
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

; Set location where local files will be saved
If (iniLocal.Settings.LocalFolder) ;Check if ini file has setting
{
	localFolder := iniLocal.Settings.LocalFolder
}
Else ;add a setting with the default value to the ini file
{
	localFolder = %A_MyDocuments%\Revit
	iniLocal.AddKey("Settings", "LocalFolder", localFolder)
	iniLocal.save()
} 

; Set whether an explorer window will be opened to the project folder
If (iniLocal.Settings.ExploreDefault) ;Check if ini file has setting
{
	exploreLaunch := iniLocal.Settings.ExploreDefault
}
Else ;add a setting with the default value to the ini file
{
	exploreLaunch := 0
	iniLocal.AddKey("Settings", "ExploreDefault", exploreLaunch)
	iniLocal.save()
} 

; Debug function to make finding problems easier
DebugMe("Start")

; By default we assume Revit is not running
revitRunning := 0

; Create a log to keep track of comings and goings
; logCheck := 1
If (iniLocal.Settings.LocalLog) ;Check if ini file has log setting
{
	localLog := iniLocal.Settings.LocalLog
}
Else ;add a setting with the default value
{
	SplitPath, iniPathLocal,, settingsFolder
	localLog := settingsFolder . "\Log"
	iniLocal.AddKey("Settings", "LocalLog", localLog)
	iniLocal.save()
}
IfNotExist, %localLog%
	FileCreateDir, %localLog%
If (iniLocal.Settings.ExitReason) ;Check if ini file has setting
{
	lastExit := iniLocal.Settings.ExitReason
}
Else ;add a setting with the default value to the ini file
{
	iniLocal.AddKey("Settings", "ExitReason", "Initial")
	iniLocal.save()
}
logMe("Program", "Opened", lastExit)

; Check if user can write to the global list of projects "iniPathCentral"
SplitPath, iniPathCentral, , iniCentralDir, , iniCentralName
globalLog = %iniCentralDir%\%iniCentralName%.log
If !LogMe("ManageList", "Write check")
{
	iniCentralEdit := 0
	LogMe("Program", "Write check", "Fail")
}
Else
{
	iniCentralEdit := 1
	LogMe("Program", "Write check", "Success")
}

noNetwork := 0
networkLogged := 0
SetTimer, NetworkCheck, 5000
GoSub, NetworkCheck
if !noNetwork
	GoSub, TrayMenu

OnExit, ExitSub
; This is all that is done when the program is started
; Any additional features will need to be initialized by the user
Return

; Check if the network is available
NetworkCheck:
IfNotExist, %iniPathCentral%
{
	noNetwork := 1
	if !networkLogged
	{
		networkLogged := LogMe("Program", "Network Check", "Fail")
		ReloadMe("noshow")
	}
}
IfExist, %iniPathCentral%
{
	if noNetwork
	{
		noNetwork := !logMe("Program", "Network Check", "Success")
		ReloadMe("noshow")
	}
	noNetwork := 0
	return
}
return

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
		; Assign the project ID from the local ini file
		fav%favN% := iniLocal [favSection].ProjectID
		; Assign the project name from the local ini file
		fav%favN%Name := iniLocal [favSection].Name
	}
}
	
;Populate right click menu including favorites
Menu, tray, NoStandard ;Get rid of default menu items
Menu, tray, Icon,  %supportFolder%\wha.ico
Menu, tray, Tip, Wright Heerema | Architects`nRevit Launcher
;Give different menu options if the S drive is not available
If noNetwork
{
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
		; Get the current project version from the Global Project List
		favVersion%A_Index% := iniCentral [fav%A_Index%].Version
		; Check if project still exists
		if !(favVersion%A_Index%)
		{
			MsgBox, % "Uh oh, The project:`n`n" . fav%A_Index%Name . "`n`ncould not be found in the Global Project List. It will be removed from your quick launch menu.`n`nYou may want to try adding it again. Please see " . bimGuy . " should this problem persist."
			iniLocal.DeleteSection(favList%A_Index%)
			iniLocal.Save()
			LogMe("Favorite", "Remove", favList%A_Index%, "No longer in global list")
		}
		else
		{
			favTitle := fav%A_Index%Name
			favIcon := favVersion%A_Index%
			Menu, tray, Add, %favTitle%, favSub
			IfExist, %supportFolder%\Revit%favIcon%file.ico
				Menu, tray, Icon, %favTitle%, %supportFolder%\Revit%favIcon%file.ico, 1, 16
			If detach
				Menu, tray, Icon, %favTitle%, %supportFolder%\detach%favIcon%.ico, 1, 16
		}
	}
}
Menu, tray, Add
Menu, tray, Add, Detach Models, DetachSub
If detach
	Menu, tray, Check, Detach Models
Menu, tray, Add, Specify Worksets, WorksetSub
If workset
	Menu, tray, Check, Specify Worksets
Menu, tray, Add
Menu, tray, Add, Settings, SettingsSub
If iniCentralEdit
	Menu, tray, Add, Manage Global Project List, ManageListSub
Menu, tray, Icon, Settings, %supportFolder%\settings.ico
Menu, tray, Add, %programName% Help, HelpSub
If !A_IsCompiled
{
	Menu, tray, Add, ListVars, ListVarsSub
	Menu, tray, Add, KeyHistory, KeyHistorySub
}
Menu, tray, Add, Exit, TrayExit
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
LogMe("MenuTray", "No Network")
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

ListVarsSub:
ListVars
return

KeyHistorySub:
KeyHistory
return

TrayExit:
ExitApp

ExitSub: ;Leaving so soon?
iniLocal := class_EasyIni(iniPathLocal)
iniLocal.Settings.ExitReason := A_ExitReason
iniLocal.Save()
LogMe("Program", "Closed", A_ExitReason)
ExitApp
return
; ### End of right click menu ###
































UserListView:
loop, %centralSections0%
{
	LV_Add(,iniCentral [centralSections%A_Index%].Number, iniCentral [centralSections%A_Index%].Name, iniCentral [centralSections%A_Index%].Version, centralSections%A_Index%)
}
;Set column widths and hide the LaunchID column
LV_ModifyCol(1, "SortDesc")
LV_ModifyCol(1, 95)
LV_ModifyCol(2, 335)
LV_ModifyCol(3, 47)
LV_ModifyCol(4, 0)
;Sort by the first column
Return


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
Gui, FindLaunch:Add, ListView, AltSubmit r10 w500 gFindLaunchList -Multi, Number|Name|Version|LaunchID
; Add projects to the listview
GoSub, UserListView
Gui, FindLaunch:Font, s10 c%guiColor1%, Arial
Gui, FindLaunch:Add, Text, w500 center yp+230, Choose a project.  A local file will be created and opened using the correct version of Revit.  Choosing detach will open the project disconnected from the central file.`nDon't see your project listed? Contact %bimGuy%.
Gui, FindLaunch:Font, s18 cBlack, Arial
; Change the button to reflect if we are opening detached or specifed worksets
Gui, FindLaunch:Add, Button,  w150 yp+55 xp+10 gFindLaunchFind, &Launch
Gui, FindLaunch:Add, Button,  w150 yp+0 xp+165 gFindLaunchDetach, &Detach
Gui, FindLaunch:Add, Button,  wp yp+0 xp+165 Default gFindLaunchGuiClose, &Cancel
GuiControl, FindLaunch:Disable, Button1
GuiControl, FindLaunch:Disable, Button2
Gui, FindLaunch:Show
Return



; Enable the launch button once the user selects a project
FindLaunchList:
If A_GuiEvent = I 
{
	If !detach
		GuiControl, FindLaunch:Enable, Button1
	GuiControl, FindLaunch:Enable, Button2
}
If A_GuiEvent = DoubleClick
	if detach
		GoSub, FindLaunchDetach
	else
		GoSub, FindLaunchFind
Return

FindLaunchDetach:
detach := 1
; Get the project selected, set project variables, and run the main routine
FindLaunchFind:
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
LV_GetText(rowText, rowNumber, 4)
projectID := rowText
DebugMe("FindLaunchFind")
LogMe("FindLaunch", projectID)
Gui, Destroy
GoSub, MainRoutine
Return
; ### End of the Search and Launch Menu ###




























; ### Create the Add/Remove Quick Launch Menu ###
; Menu that allows users to add or remove programs they frequent to the tray menu
; First we simply ask if they want to add to the list or remove from the list
AddRemove:
guiTitle := "What would you like to do?"
guiWidth := 550
bWidth := (guiWidth - 10) / 2
bLoc := bWidth + 10
tFont := GetFontMax(guiTitle, guiWidth)
Gui, AddRemove:New,, %programName%
Gui, AddRemove:Font, %tFont%, Arial
Gui, AddRemove:Add, Text, center w%guiWidth%, %guiTitle%
Gui, AddRemove:Font, s18, Arial
Gui, AddRemove:Add, Button, Default w%bWidth% gAddProject, &Add a Project
Gui, AddRemove:Add, Button, wp xp+%bLoc% gRemoveProject, &Remove a Project
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
guiTitle := "Add a project to your Quick Launch Menu:"
guiWidth := 500
bWidth := (guiWidth - 10) / 2
bLoc := bWidth + 10
tFont := GetFontMax(guiTitle, guiWidth)
Gui, AddProject:New,, %programName%
Gui, AddProject:Font, %tFont% , Arial
Gui, AddProject:Add, Text, center w%guiWidth%, %guiTitle%
Gui, AddProject:Font, s10 , Arial
Gui, AddProject:Add, ListView, AltSubmit r10 w500 gAddProjectList -Multi, Number|Name|Version|LaunchID
; Add all of the projects to the listview
GoSub, UserListView

Gui, AddProject:Font, s10 c%guiColor1%, Arial
Gui, AddProject:Add, Text, w500 center yp+230, Choose a project and it will be added to your "Quick Launch" menu.`nDon't see your project listed? Contact %bimGuy%.
Gui, AddProject:Font, s18 cBlack, Arial
Gui, AddProject:Add, Button,  w%bWidth% yp+60 xm gAddProjectAdd, &Add to list
Gui, AddProject:Add, Button,  wp xp+%bLoc% Default gAddProjectGuiClose, &Cancel
GuiControl, Disable, &Add to list
Gui, AddProject:Show
; Destroy the menu that asks if we want to add or remove
Gui, AddRemove:Destroy
return



; Menu to remove a program from the Quick Launch list
RemoveProject:
; Create the Remove from Quick Launch menu
guiTitle := "Select a project to remove:"
guiWidth := 500
bWidth := (guiWidth - 10) / 2
bLoc := bWidth + 10
tFont := GetFontMax(guiTitle, guiWidth)
Gui, RemoveProject:New,, %programName%
Gui, RemoveProject:Font, %tFont% , Arial
Gui, RemoveProject:Add, Text, center w%guiWidth%, %guiTitle%
Gui, RemoveProject:Font, s10, Arial
Gui, RemoveProject:Add, ListView, AltSubmit r10 w%guiWidth% gRemoveProjectList -Multi, Name|ID
; Add a list of all the projects in the Quick Launch list to the list view
loop, %favList0%
{
	LV_Add(,iniLocal [favList%A_Index%].Name, favList%A_Index%)
}
LV_ModifyCol(1, 498)
LV_ModifyCol(2, 0)
Gui, RemoveProject:Font, s10 c%guiColor1%, Arial
Gui, RemoveProject:Add, Text, w500 center yp+230, The selected project will be removed from your "Quick Launch" menu.
Gui, RemoveProject:Font, s18 cBlack, Arial
Gui, RemoveProject:Add, Button, w%bWidth% xm yp+60 gRemoveProjectRemove, &Remove from list
Gui, RemoveProject:Add, Button,  w%bWidth% xp+%bLoc% Default gRemoveProjectGuiClose, &Cancel
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



RemoveProjectList:
; Enable the remove button once a users selects a project
If A_GuiEvent = I
	GuiControl, RemoveProject:Enable, Button1
If A_GuiEvent = DoubleClick
	GoSub, RemoveProjectRemove
Return



RemoveProjectRemove:
; This does the hard work of actually REMOVING the project from the Quick Launch list
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
LV_GetText(rowText, rowNumber, 2)
MsgBox, 33, Remove Project?, % "Are you sure you want to remove the following project from your list?`n`n" . iniLocal [favList%RowNumber%].Name
IfMsgBox OK
{
	iniLocal.DeleteSection(favList%RowNumber%)
	iniLocal.Save()
	Gui, RemoveProject:Destroy
	LogMe("Favorite", "Remove", favList%RowNumber%)
	ReloadMe()
}
return



AddProjectList:
; Enable the add button once a user selects a project
If A_GuiEvent = I
	GuiControl, AddProject:Enable, Button1
If A_GuiEvent = DoubleClick
	GoSub, AddProjectAdd
Return



AddProjectAdd:
; This does the hard work of actually ADDING the project to the Quick Launch list
; Make sure you zero out the RowNumber otherwise you could see some weird behaviour
RowNumber := 0
RowNumber := LV_GetNext(RowNumber)
LV_GetText(rowText, rowNumber, 4)
newFavSection = Favorite%rowText%
newFavNumber := iniCentral [rowText].Number
newFavName := iniCentral [rowText].NameShort

if instr(favList, newFavSection)
{
	MsgBox, 64, %programName%, %newFavNumber% %newFavName% is already in your list of projects
}
else
{
	MsgBox, 33, %programName%, You are sure you want to add the following project to your list?`n`n%newFavNumber% - %newFavName%
	IfMsgBox OK
	{
		iniLocal.AddSection(newFavSection, "ProjectID", rowText)
		iniLocal.AddKey(newFavSection, "Name", newFavNumber . " " . newFavName)
		iniLocal.AddKey(newFavSection, "Detach", "0")
		iniLocal.AddKey(newFavSection, "Workset", "0")
		iniLocal.save()
		Gui, AddProject:Destroy
		LogMe("Favorite", "Add", rowText, newFavNumber, newFavName, "Detach 0", "Workset 0")
		ReloadMe()
	}
}
return
; ### End of Add/Remove Quick Launch Menu ###























; ### Create the Settings Menu ###
SettingsSub:
detachDefault := iniLocal.Settings.Detach
worksetDefault := iniLocal.Settings.Workset
exploreDefault := iniLocal.Settings.ExploreDefault
linkFile = %A_Startup%\%programName%.lnk
IfExist %linkFile%
	startupExist := 1
Else
	startupExist := 0

guiTitle := programName . " Settings:"
guiWidth := 500
bWidth := (guiWidth - 10) / 2
bLoc := bWidth + 10
tFont := GetFontMax(guiTitle, guiWidth)
Gui, Settings:New,, %programName%
Gui, Settings:Font, %tFont% cBlack, Arial
Gui, Settings:Add, Text, center w%guiWidth%, %guiTitle%

;Start of default settings section
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Text, yp+80 section, Defaults
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vdetachCheck gSettingsSubDetach, Detach by default
GuiControl,, detachCheck, %detachDefault%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400, Controls whether "Detach Models" is enabled when you first open %programName%.  This option is perfect for Project Managers who only need to query a model.  You can temporarily disable it before launching if you need to save changes to a model.
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+75 section,
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vworksetCheck gSettingsSubWorkset, Specify worksets by default
GuiControl,, worksetCheck, %worksetDefault%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400, Controls whether the "Specify Worksets" dialog box is opened by default.  Choose this option if you frequently work on larger Revit models.
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+50 section,
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vexploreCheck gSettingsSubExplore, Open Project Folder in Explorer
GuiControl,, exploreCheck, %exploreDefault%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400, With this setting checked, a Windows Explorer window will be opened to the projects working directory when the project is launched.

;Start of behavior at startup section
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+75 section, Startup
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Checkbox, xs+100 ys-5 vstartupCheck gSettingsSubStartup, Startup on Windows Login
GuiControl,, startupCheck, %startupExist%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400, If checked, %programName% will open automatically when you log in to Windows.  Just set it and forget it. :)

;Start of file locations settings section
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xp-100 yp+75 section, Locations
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Text, xs+100 ys-5, Local File Save Location
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400 vlocalLocation, %localFolder%
Gui, Settings:Add, Button, w75 xs+100 ys+45 gSettingsSubLocal, Change
Gui, Settings:Add, Button, w75 xs+185 ys+45 gSettingsSubLocalDefault, Default
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xm yp+50 section,
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Text, xs+100 ys-5, Log File
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400 vlogLocation, %localLog%
Gui, Settings:Add, Button, w75 xs+100 ys+45 gSettingsSubLog, View

;Start of about section
Gui, Settings:Font, s12 cBlack, Arial
Gui, Settings:Add, Text, xm yp+75 section, About
Gui, Settings:Font, s18, Arial
Gui, Settings:Add, Text, xs+100 ys-5, %programName%
Gui, Settings:Font, s9 c%guiColor1%, Arial
Gui, Settings:Add, Text, xs+100 ys+25 w400, Version: %scriptVersion%
Gui, Settings:Add, Text, xs+100 yp+18 w400, %programName% was created by Michael Pfammatter and is copyright by Wright Heerema | Architects in 2014
Gui, Settings:Show
DebugMe("SettingsSub")
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
LogMe("Settings", "detachDefault", detach)
return



SettingsSubWorkset:
; msgBox, % !detachDefault
iniLocal.Settings.Workset := !worksetDefault
iniLocal.Save()
worksetDefault := iniLocal.Settings.Workset
GuiControl,, worksetCheck, %worksetDefault%
workset := worksetDefault
LogMe("Settings", "worksetDefault", workset)
return



SettingsSubExplore:
;Allows user to set whether an explorer window will be opened to the project folder after the model is launched.
DebugMe("SettingsSubExplore")
iniLocal.Settings.ExploreDefault := !exploreDefault
iniLocal.Save()
exploreDefault := iniLocal.Settings.ExploreDefault
GuiControl,, exploreCheck, %exploreDefault%
exploreLaunch := exploreDefault
LogMe("Settings", "exploreLaunch", exploreLaunch)
Return



SettingsSubStartup:
DebugMe("Settings Startup Check")
IfNotExist, %LinkFile%
{
	FileCreateShortcut, %A_ScriptFullPath%, %LinkFile% 
	If ErrorLevel
	{
		MsgBox, 16, %programName%, There was a problem creating a shortcut in your startup folder, %A_Startup%.  Please contact %bimGuy% for additional information.
		LogMe("Error", "Shortcut", A_ScriptFullPath, LinkFile)
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
LogMe("Settings", "startupExist", startupExist)
Return



SettingsSubLocal:
;Allows user to set location where local files will be saved. It also give the option to set it back to the default location.
FileSelectFolder, selectedFolder, , 3, Choose a location for your local files to reside:
If ErrorLevel
	Return
localFolder := selectedFolder
iniLocal.Settings.LocalFolder := localFolder
iniLocal.save()
GuiControl, , localLocation, %localFolder%
LogMe("Settings", "localLocation", localFolder)
Return



SettingsSubLocalDefault:
IfEqual, localFolder, %A_MyDocuments%\Revit
	MsgBox, 16, %programName%, The local file save location is currently set to the default location.
Else
{
	MsgBox, 17, %programName%, This will return the local file save location to the default setting.  Are you sure you would like to proceed?
	IfMsgBox, Cancel
		Return
	localFolder = %A_MyDocuments%\Revit
	iniLocal.Settings.LocalFolder := localFolder
	iniLocal.save()
	GuiControl, , localLocation, %localFolder%
	LogMe("Settings", "localLocation", localFolder, "Default")
}
Return

SettingsSubLog:
Run, %localLog%
Return
; ### End of Settings Menu ###

























; ### Start of Manage Project List Menu ###
ManageListSub:
; Load up the global list of projects
iniCentral := class_EasyIni(iniPathCentral)
; Set a variable to use the Manage Add menu for both adding and editing
maEdit := 0
; Create the add project menu
Gui, Manage:New,, %programName%
Gui, Manage:Font, s24 , Arial
Gui, Manage:Add, Text, center w680, Select a project to manage:
Gui, Manage:Font, s18 , Arial
Gui, Manage:Add, Button, w680 gManageProjectAdd, &Add a New Project
Gui, Manage:Font, s10 , Arial
Gui, Manage:Add, ListView, AltSubmit r15 w680 gManageProjectList -Multi, Launcher ID|Project Number|Name|Version
GoSub, ManageListUpdate
LV_ModifyCol(1, 100)
LV_ModifyCol(2, 100)
LV_ModifyCol(3, 400)
LV_ModifyCol(4, 59)
Gui, Manage:Font, s10 c%guiColor1%, Arial
Gui, Manage:Font, s18 cBlack, Arial
Gui, Manage:Add, Button, w200 yp+350 xm gManageProjectEdit, &Edit Project
Gui, Manage:Add, Button, wp xp+240 gManageProjectRemove, &Remove Project
Gui, Manage:Add, Button, wp xp+240 Default gManageGuiClose, &Exit
GuiControl, Manage:Disable, Button2
GuiControl, Manage:Disable, Button3
Gui, Manage:Show
Return



ManageListUpdate:
; Add all of the projects to the listview
maEdit := 0
Gui, Submit, NoHide
Gui, Manage:Default
LV_Delete()
; Get all of the projects in the list
centralSections := iniCentral.GetSections(,"C")
; Create an array of the projects for looping
StringSplit, centralSections, centralSections, `n
loop, %centralSections0%
{
	LV_Add(, centralSections%A_Index%, iniCentral [centralSections%A_Index%].Number, iniCentral [centralSections%A_Index%].Name, iniCentral [centralSections%A_Index%].Version)
}
GuiControl, Manage:Disable, Button2
GuiControl, Manage:Disable, Button3
Return



ManageProjectList:
If A_GuiEvent = I
{
	GuiControl, Manage:Enable, Button2
	GuiControl, Manage:Enable, Button3
}
If A_GuiEvent = DoubleClick
	GoSub, ManageProjectEdit
Return



ManageProjectAdd:
GuiControl, Manage:Disable, Button2
GuiControl, Manage:Disable, Button3
Gui, ManageAdd:New
Gui, ManageAdd:+OwnerManage
Gui, ManageAdd:Font, s24 , Arial
Gui, ManageAdd:Add, Text, vmaTitle, Add a project to the global list:
Gui, ManageAdd:Font, s10 , Arial
Gui, ManageAdd:Add, Text, w95 Right, Central File:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 vmaCentral w540 gFileCheck
Gui, ManageAdd:Add, Button, w30 xp+550 yp-3 gManageBrowse ,...
Gui, ManageAdd:Add, Text, xm w95 Right, Project Number:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 gManageID vmaProject w200
Gui, ManageAdd:Add, Text, xp+340 yp+0 w95 Right vmaVersionText, Revit Version:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 vmaVersion w100
Gui, ManageAdd:Add, Text, xm w95 Right, Project Name:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 vmaName w540
Gui, ManageAdd:Add, Text, xm w95 Right, Short Name:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 gManageSubmitEnable vmaShort w540
Gui, ManageAdd:Add, Text, xm w95 Right, Project Folder:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 vmaWorkingFolder w540
Gui, ManageAdd:Add, Text, xm w95 Right, Launcher ID:
Gui, ManageAdd:Add, Edit, xp+100 yp+0 vmaLauncherID w540
Gui, ManageAdd:Font, s18 , Arial
Gui, ManageAdd:Add, Button, xm+360 w150 vmaAddButton gManageProjectNew, &Add
Gui, ManageAdd:Add, Button, wp xp+170 yp+0 gManageAddGuiClose, &Cancel
If maEdit ;Change the title and fill the form only if we are editing a project
{
	GuiControl,, maAddButton, &Edit
	GuiControl,, maTitle, Edit a project in the global list:
	GuiControl,, maCentral, %oldCentral%
	GuiControl,, maProject, %oldProject%
	GuiControl,, maVersion, %oldVersion%
	GuiControl,, maName, %oldName%
	GuiControl,, maShort, %oldShort%
	GuiControl,, maWorkingFolder, %oldWorkingFolder%
	GuiControl,, maLauncherID, %rowText%
}
Else ;If we are adding a new project, disable edits until central is selected
{
	GuiControl, ManageAdd:Disable, maName
	GuiControl, ManageAdd:Disable, maShort
	GuiControl, ManageAdd:Disable, maWorkingFolder
	GuiControl, ManageAdd:Disable, maVersion
}
GuiControl, ManageAdd:Disable, maProject
GuiControl, ManageAdd:Disable, maLauncherID
GuiControl, ManageAdd:Disable, Button2
Gui, ManageAdd:Show
Return

ManageBrowse:
FileSelectFile, selectedFile, 1,X:\
If ErrorLevel
	Return
maCentral := selectedFile
GuiControl,, maCentral, %maCentral%
Return

ManageSubmitEnable:
Gui, Submit, NoHide
GuiControl, Enable, Button2
Return




ManageProjectNew:
Gui, ManageAdd:Submit, NoHide ;Don't hide the window until the data has been checked
;Check validity - We don't need to verify the Central File or ProjectID as they have already been checked
;Check that version is in right format
If !(RegExMatch(maVersion, "P)^20\d\d$"))
{
	MsgBox, 64, %programName%, You have not entered a valid Revit Version.  We are looking for the year of the release only.  For example, if the project was last saved in Revit 2014 or Revit Architecture 2014, enter "2014" for the version.
	Gui, Font, s10 cRed, Arial
	GuiControl, Font, maVersionText
	Return
}
Else
{
	Gui, Font, s10 cBlack, Arial
	GuiControl, Font, maVersionText
}

;Remove trailing slashes from the working folder
maWorkingFolder := RegExReplace(maWorkingFolder, "\\$")
Gui, ManageAdd:Destroy

;Create new Section in the global ini file
If maEdit
{
	MsgBox, 49, %programName%, Are you sure you want to update this project?`n`nNumber: %maProject%`nName: %maName%`nShort Name: %maShort%`nVersion: %maVersion%`nCentral: %maCentral%`nProject Folder: %maWorkingFolder%
	iniCentral.DeleteSection(rowText)
	LogMe("ManageList", "Pre-edit", rowText, oldProject, oldName, oldShort, oldVersion, oldCentral, oldWorkingFolder)
	LogMe("ManageList", "Edit", maLauncherID, maProject, maName, maShort, maVersion, maCentral, maWorkingFolder)
}
Else
	LogMe("ManageList", "Add", maLauncherID, maProject, maName, maShort, maVersion, maCentral, maWorkingFolder)
iniCentral.AddSection(maLauncherID)
iniCentral.AddKey(maLauncherID, "Number", maProject)
iniCentral.AddKey(maLauncherID, "Name", maName)
iniCentral.AddKey(maLauncherID, "NameShort", maShort)
iniCentral.AddKey(maLauncherID, "Version", maVersion)
iniCentral.AddKey(maLauncherID, "Central", maCentral)
iniCentral.AddKey(maLauncherID, "WorkingFolder", maWorkingFolder)
;Update the list of projects in the manage window. This happens before the save so there is no delay on the user side.
GoSub, ManageListUpdate
;Save our work
iniCentral.save()
Return


ManageProjectEdit:
;Edit a project that is already in the global list
maEdit := 1
rowNumber := 0
rowNumber := LV_GetNext(RowNumber)
LV_GetText(rowText, rowNumber)
oldProject := iniCentral [rowText].Number
oldName := iniCentral [rowText].Name
oldShort := iniCentral [rowText].NameShort
oldVersion := iniCentral [rowText].Version
oldCentral := iniCentral [rowText].Central
oldWorkingFolder := iniCentral [rowText].WorkingFolder
GoSub, ManageProjectAdd
Return



ManageProjectRemove:
rowNumber := 0
rowNumber := LV_GetNext(RowNumber)
LV_GetText(rowText, rowNumber)
delPNumber := iniCentral [rowText].Number
delPName := iniCentral [rowText].Name
MsgBox, 49, %programName%, Are you sure you want to delete project:`n%delPNumber% %delPName%`n`nThis will effect all users in the office.
IfMsgBox, Cancel
	Return
iniCentral.DeleteSection(rowText)
LogMe("ManageList", "Remove", rowText, delPNumber, delPName)
GoSub, ManageListUpdate
iniCentral.Save()
; newFavID := centralSections%RowNumber%
Return



FileCheck: 
Gui, Submit, NoHide
If (FileExist(maCentral)) && (InStr(maCentral, ".rvt", -4) > 0)
{ ; Change text color if central file does not exist or is not a valid Revit file
	SplitPath, maCentral, maCentralFile, maWorkingFolder
	StringSplit, maSplit, maWorkingFolder, "\"
	maSpaceLocale := InStr(maSplit3, A_Space)
	maProject := SubStr(maSplit3, 1, maSpaceLocale - 1)
	If !maEdit 
	{ ;Check if central file is already listed but only if we are adding a new project
		Loop, %centralSections0%
		{
			centralTest := iniCentral [centralSections%A_Index%].Central		
			If (centralTest = maCentral) 
			{
				MsgBox, 16, %programName%, % "This central file is already included in the global list under Laucher ID: " . centralSections%A_Index%
				Gui, ManageAdd:Destroy
				Return
			}
		}
	}
	maName := SubStr(maSplit3, maSpaceLocale + 1)	
	manageEdits := 1 ; Parameter to quickly set whether field are editable
	Gui, Font, s10 cBlack, Arial
	If !maEdit
	{
		GuiControl,, maProject, %maProject%
		GuiControl,, maVersion, 2014
		GuiControl,, maName, %maName%
		GuiControl,, maWorkingFolder, %maWorkingFolder%
	}
}
else
{
	manageEdits := 0
	Gui, Font, s10 cRed, Arial
}
GuiControl, Font, maCentral
GuiControl, Enable%manageEdits%, maName
GuiControl, Enable%manageEdits%, maShort
GuiControl, Enable%manageEdits%, maWorkingFolder
GuiControl, Enable%manageEdits%, maVersion
If !maEdit
	GuiControl, Enable%manageEdits%, maProject
Return



ManageID: 
;Create a unique ID from the Project Number
If !maEdit
{
	Gui, Submit, NoHide
	maIDCount := 0
	maProjectClean := SubStr(maProject, 1, 8)
	Loop, %centralSections0%
	{
		sectionTest := SubStr(centralSections%A_Index%, 1, InStr(centralSections%A_Index%, "-",, 0) - 1)
		If sectionTest = %maProjectClean%
			maIDCount += 1
	}
	If (StrLen(maIDCount) < 2)
		maIDCount := "0" . maIDCount
	maLauncherID := SubStr(maProject, 1, 8) . "-" . maIDCount
	GuiControl,, maLauncherID, %maLauncherID%
}
Return



ManageAddGuiEscape:
ManageAddGuiClose:
maEdit := 0
Gui, ManageAdd:Destroy
Return



ManageGuiEscape:
ManageGuiClose:
ReloadMe()
Return
; ### End of Manage Project List Menu ###

















; ### Create the Help Menu ###
HelpSub:
helpFile := "https://wharchs.sharepoint.com/_layouts/15/start.aspx#/BIM%20Standards/WHA%20Revit%20Launcher.aspx"
Run, %helpFile%
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
workingFolder := iniCentral [projectID].WorkingFolder
IfNotExist, %pCentral% ;Check if project is setup correctly
	{
		MsgBox, 16, Launcher Error, There is an issue locating the project associated with %projectID%.`nPlease see your friendly BIM manager for additional information.`n`nHave a nice day!
		LogMe("Launcher", "Error", "Locating Central", projectID, pCentral)
		ReloadMe()
	}
revitUser = %A_Username% ;Users computer username
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
	LogMe("Launcher", "Error", "FindRevit", revitPath, projectID, pVersion)
	ReloadMe()
}

FindMonitor: 
;	Check for a Worksharing Monitor for the right version of Revit
; Because Workingshing monitor becomes 64bit with version 2015, we have to change folder locations
; ? in variable assignment is a Ternary operator to avoid complicated if/else statement
mon64 := pVersion >= 2015 ? "" : " (x86)"
adesk := pVersion >= 2015 ? "" : " Autodesk"
monitorPath = %A_ProgramFiles%%mon64%\Autodesk\Worksharing Monitor for%adesk% Revit %pVersion%\WorksharingMonitor.exe
monitorTitle = Worksharing Monitor for Autodesk Revit %pVersion%
IfNotExist, %monitorPath%
{
	MsgBox, 64, Worksharing Monitor, The Worksharing Monitor could not be found on your system.  Please notify your BIM manager so you can play well with others., 7
	monitorPath =
	LogMe("Launcher", "Error", "FindMonitor", monitorPath, projectID, pVersion)
}
DebugMe("LaunchSplash")
Gosub, LaunchSplash
DebugMe("exploreLaunch")
;If the user has the the setting turned on, open the working directory of the project.
If exploreLaunch
{
	If workingFolder
		Run, %workingFolder%
}
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
	IfExist, %projectFolder%\%localFile%.4.rvt
	{
		FileGetTime, backupFour, %projectFolder%\%localFile%.4.rvt, C
		weekTwo := A_Now
		weekTwo += -14, D
		If backupFour > weekTwo
		{
			FileMove, %projectFolder%\%localFile%.3.rvt, %projectFolder%\%localFile%.4.rvt
		}
	}
	Else
	{
		IfExist, %projectFolder%\%localFile%.3.rvt
			FileMove, %projectFolder%\%localFile%.3.rvt, %projectFolder%\%localFile%.4.rvt
	}
	IfExist, %projectFolder%\%localFile%.3.rvt
	{
		FileGetTime, backupThree, %projectFolder%\%localFile%.3.rvt, C
		weekOne := A_Now
		weekOne += -7, D
		If backupThree > weekOne
		{
			FileMove, %projectFolder%\%localFile%.2.rvt, %projectFolder%\%localFile%.3.rvt
			FileSetTime,, %projectFolder%\%localFile%.3.rvt, C
		}
	}
	Else
	{
		IfExist, %projectFolder%\%localFile%.2.rvt
		{
			FileMove, %projectFolder%\%localFile%.2.rvt, %projectFolder%\%localFile%.3.rvt
			FileSetTime,, %projectFolder%\%localFile%.3.rvt, C
		}
	}
	IfExist, %projectFolder%\%localFile%.1.rvt ;backup of a backup
		fileMove, %projectFolder%\%localFile%.1.rvt, %projectFolder%\%localFile%.2.rvt, 1
	IfExist, %localPath% ;backup of the local
		fileMove, %localPath%, %projectFolder%\%localFile%.1.rvt, 1

	DebugMe("LocalCreate")
	FileCopy, %pCentral%, %localPath%, 1
	if ErrorLevel ;Check to make sure everything copied correctly
	{
		MsgBox, 16, Local Creation Error, The local file could not be copied to your computer.  Please see your BIM manager for additional information.
		ExitApp
		LogMe("Launcher", "Error", "LocalCreate", projectID, localPath, pCentral)
	}
}

DebugMe("Launch")
launchStatus = Launching Revit %pVersion%
GuiControl,, LaunchText, %launchStatus%
;Open Revit with the correct version
If monitorPath ;Open the worksharing monitor if it exists
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
WinWait, Project Not Saved Recently,, 1
if !Errorlevel
{
	Gui, Launch:Destroy
	MsgBox, 48, %programName%, Please save your current project/family before launching a new one.
	LogMe("Launcher", "Error", "Project Not Saved Dialog", projectID, revitPath, revitTitle, localPath)
	ReloadMe("noshow")
}
WinWait, Open,, 15
If ErrorLevel
{
	Gui, Launch:Destroy
	MsgBox, 48, %programName%, There seems to be an issue launching your project. Check Revit for any opened dailog boxes and try launching again.`n`nThanks.
	LogMe("Launcher", "Error", "OpenWait", projectID, revitPath, revitTitle, localPath)
	ReloadMe("noshow")
}
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
	Sleep 300
	Control, EditPaste, %fPath%, , ahk_id %fileHwnd%
	ControlClick, Button1, Open,, L, 2, NA
	Sleep 300
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
	LogMe("Launcher", "workset", projectID, fName, fPath)
	WinWait, Opening Worksets,, 60
	Gui, Launch:Destroy
	Sleep 200
	WinWait, Copied Central Model,, 30
	ControlClick, Button1, Copied Central Model,, L, 2, NA
}
Else If detach
{
	LogMe("Launcher", "detach", projectID, fName)
	Sleep 300
	WinWait, Detach Model from Central,, 30
	If !ErrorLevel
		ControlClick, Button1, Detach Model from Central,, L
	Gui, Launch:Destroy
}
Else
{
	LogMe("Launcher", "standard", projectID, fName)
	WinWait, Copied Central Model,, 30
	Sleep 300
	If !ErrorLevel
		ControlClick, Button1, Copied Central Model,, L, 2, NA
	Gui, Launch:Destroy
}
WinActivate, ^%revitTitle%

; Set detach back to global setting
detach := iniLocal.Settings.Detach
ReloadMe("noshow")

Return

LaunchSplash:
Gui, Launch:New
If debug
	Gui, Launch:-Caption ; AlwaysOnTop 
Else
	Gui, Launch:-Caption
Gui, Launch:Margin, 50, 50
Gui, Launch:Color, DCDCDC
whaLogo(700, "Launch")
Gui, Launch:Font, cBlack s36, Arial
Gui, Launch:Add, Text,xm yp+150 center w700 vLaunchText, Launching...
Gui, Launch:Font, c%guiColor1% s18, Arial
Gui, Launch:Add, Text,yp+70 center w700, %pNumber% %pName%
Gui, Launch:Font, c%guiColor2% s24, Arial
if detach or workset
{
	Gui, Launch:Add, Text, yp+70 w700 vlaunchSub center,
	if detach & workset
		GuiControl, Launch:Text, launchSub, Detach and Specify
	else if workset
		GuiControl, Launch:Text, launchSub, Specify Worksets
	else
		GuiControl, Launch:Text, launchSub, Detach
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

ReloadMe(sx = "show")
{
	Global
	Menu, Tray, DeleteAll
	Gui, Destroy
	GoSub, TrayMenu
	If !(sx = "noshow")
	{
		CoordMode, Menu, Screen
		Menu, Tray, Show, %mouseX%, %mouseY%
		CoordMode, Menu, Window
	}
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

LogMe(category, params*)
{
	global localLog
	global globalLog
	global scriptVersion
	FormatTime, yearWeek, , yWeek
	FormatTime, logTime, , yyyy-MM-dd HH:mm:ss
	sep := ", "
	logLine := logTime . sep . category

	for index, param in params
	{
		logLine .=  sep . param
	}
	logFile := localLog . "\" . A_Username . "-" . yearWeek . ".log"
	FileAppend, %logLine%`n, %logFile%
	If (category = "ManageList")
	{
		FileAppend, % logLine . ", user:" . A_Username . ", version:" . scriptVersion . "`n", %globalLog%
	}
	If ErrorLevel
		Return 0
	Else
		Return 1
}
Return
