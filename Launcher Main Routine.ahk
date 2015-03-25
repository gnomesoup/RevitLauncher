MainRoutine:
gosub, GetProjectInfo
if audit
	gosub, GetAuditInfo
gosub, FindRevit
gosub, FindMonitor
gosub, LaunchSplash
gosub, ExploreLaunch

if detach
{
	LaunchUpdate("Detaching the Revit " . pVersion . " Model")
	gosub, OpenRevit
	gosub, OpenDialog
	gosub, SelectCentral
	gosub, DetachClick
	gosub, OpenModel
}
else if specify
{
	LaunchUpdate("Creating local...")
	gosub, LocalCreate
	LaunchUpdate("Launching Revit " . pVersion . " with Specify")
	gosub, OpenRevit
	gosub, OpenDialog
	gosub, SelectLocal
	gosub, SpecifyClick
	gosub, OpenModel
	gosub, SpecifyWait
	gosub, LocalDismiss
}
else if audit
{
	LaunchUpdate("Creating central backup...")
	gosub, AuditBackup
	LaunchUpdate("Auditing a Revit " . pVersion . " Model")
	gosub, OpenRevit
	gosub, OpenDialog
	gosub, SelectCentral
	gosub, AuditClick
	gosub, SelectCentral
	gosub, OpenModel
	gosub, CloseModel
	gosub, SaveAsCentralLaunchUpdate("Creating local...")
	gosub, LocalCreate
	LaunchUpdate("Launching the Revit " . pVersion . " Model")
	gosub, SelectLocal
	gosub, OpenModel
	gosub, LocalDismiss
	gosub, CloseModel
	gosub, SyncRelease
}
else
{
	LaunchUpdate("Creating local...")
	gosub, LocalCreate
	LaunchUpdate("Launching the Revit " . pVersion . " Model")
	gosub, SelectLocal
	gosub, OpenModel
	gosub, LocalDismiss
}

return

GetProjectInfo:
iniCentral := class_EasyIni(iniPathCentral)
pNumber := iniCentral [projectID].Number
pName := iniCentral [projectID].Name
pNameShort := iniCentral [projectID].NameShort
pVersion := iniCentral [projectID].Version
pCentral := iniCentral [projectID].Central
workingFolder := iniCentral [projectID].WorkingFolder
IfNotExist, %pCentral% ;Check if project is setup correctly
	{
		PrettyMsg("The central file does not exist:`n`n" . pCentral . "`n`nPlease see " . bimGuy . " for assistance")
		LogMe("Launcher", "Error", "Locating Central", projectID, pCentral)
		ReloadMe()
	}
revitUser = %A_Username% ;Users computer username
projectFolder = %localFolder%\%projectID% %pNameShort% ;Project folder name
localFile = %projectID% %pNameShort% LOCAL %revitUser% ;Local file name no extension
localPath = %projectFolder%\%localFile%.rvt ;Local file name & path
return

GetAuditInfo:
SplitPath, pCentral, centralFileName, centralFileDir, , centralBackupName
centralBackupName := centralBackupName . "_backup"
FormatTime, auditDate, , yyyy-MM-dd
auditArchiveDir := workingFolder . "\Archive\" . auditDate . " Pre-audit Backup"
centralFileBackupDir := centralFileDir . "\" . centralBackupName
auditError := 0
LogMe("Launcher", "Audit", pCentral)	
return

FindRevit:
;Find the path to the correct Revit flavor depending on the version read from the ini file
if !detach
{
	IfWinExist, %localFile%.rvt ;Check to see if the file is already open
	{
		WinActivate
		WinMaximize	
		if audit
		{
			PrettyMsg("You have a local file open for the central model you are trying to audit. Please close """ . localFile . ".rvt"" and attempt the audit again.")
			Gui, AuditModel:Show
			Gui, Launch:Destroy
			LogMe("Launcher", "Error", "Audit Fail", "Open Local File", localFile, pCentral)
			return
		}
		PrettyMsg("Revit is already running with the specified local file:`n`n" . localFile . ".rvt")
		ReloadMe()
	}
}

IfWinExist, %CentralFileName%
{
	WinActivate
	WinMaximize
	PrettyMsg("Open Central File!`n`n" . pCentral . " is currently open. Please close this file immediately and try again.", "alert")
	if audit
	{
		Gui, AuditProject:Show
		Gui, Launch:Destroy
		LogMe("Launcher", "Error", "Audit Fail", "Open Central File", pCentral)
		return
	}
	LogMe("Launcher", "Error", "Open Central File", pCentral)
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
	PrettyMsg("Revit " . pVersion . " cannot be found on this computer.  Please see " . bimGuy . " for additional information.", , 1)
	LogMe("Launcher", "Error", "FindRevit", revitPath, projectID, pVersion)
	ReloadMe()
}
return

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
	PrettyMsg("The Worksharing Monitor could not be found on your system.  Please notify " . bimGuy . " so you can play well with others.")
	monitorPath =
	LogMe("Launcher", "Error", "FindMonitor", monitorPath, projectID, pVersion)
}
return

ExploreLaunch:
If exploreLaunch
{
	If workingFolder
		Run, "%workingFolder%"
}
return

LaunchSplash:
launchWidth := 500
SysGet, pMonitor, MonitorPrimary
SysGet, pMonitor, MonitorWorkArea, %pMonitor%
Gui, Launch:New
Gui, Launch:-Caption AlwaysOnTop
Gui, Launch:Margin, 50, 25
Gui, Launch:Color, DCDCDC
whaLogo(launchWidth, "Launch")
Gui, Launch:Font, cBlack s24, Arial
Gui, Launch:Add, Text, xm yp+45 center w%launchWidth% vLaunchText, Launching...
Gui, Launch:Font, c%guiColor1% s12, Arial
Gui, Launch:Add, Text, yp+40 center w%launchWidth% vLaunchSub, %pNumber% %pName%
Gui, Launch:Font, c%guiColor2% s18, Arial
launchHeight := 150
launchY := pMonitorBottom - launchHeight
Gui, Launch:Show, y%launchY% h%launchHeight%
return

LocalCreate:
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
	LogMe("Launcher", "Error", "LocalCreate", projectID, localPath, pCentral)
	PrettyMsg("The local file could not be copied to your computer.  Please see your BIM manager for additional information.", "alert", 1)
	ReloadMe()
}


AuditBackup:
auditArchiveDirTest := auditArchiveDir
n := 0
While FileExist(auditArchiveDirTest)
{
	n += 1
	auditArchiveDirTest := auditArchiveDir . A_Index
}
if (n > 12)
{
	if !PrettyMsg("Warning!`n`nThere have been at least " . n . " audit backup folders created already. It is not typical to need to audit your project this frequently. Press ""OK"" if you are sure you want to continue", "alert")
		ReloadMe()
}
FileCreateDir, %auditArchiveDirTest%
If ErrorLevel
{
	PrettyMsg("Error`n`nWe were unable to create the archive directory for the file being audited`n`n" . pCentral . "`n`nPlease check your project settings in the global project list and try again.", "alert", 1)
	LogMe("Launcher", "Audit", "Error creating backup", auditArchiveDir)
}
FileCopy, %pCentral%, %auditArchiveDirTest%
auditError := If ErrorLevel ? auditError + ErrorLevel : auditError
FileCopyDir, %centralFileBackupDir%, %auditArchiveDirTest%\%centralBackupName%
auditError := If ErrorLevel ? auditError + ErrorLevel : auditError
if auditError
{
	PrettyMsg("Error`n`nWe were unable to archive the file being audited`n`n" . pCentral . "Contact " . bimGuy . " for additional information", "alert", 1)
	ReloadMe()
}
return

OpenRevit:
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
	WinWait, ^%revitTitle% - \[Recent Files\]
	WinMaximize
}
WinActivate
return

OpenDialog:
IfWinExist, Open, &Detach from Central
	WinActivate, Open, &Detach from Central
else
{
	Send ^o
	WinWait, Project Not Saved Recently,, 1
	if !Errorlevel
	{
		Gui, Launch:Destroy
		PrettyMsg("Save your work!`n`nPlease save your current project/family before launching a new one.")
		LogMe("Launcher", "Error", "Project Not Saved Dialog", projectID, revitPath, revitTitle, localPath)
		ReloadMe("noshow")
	}
	WinWait, Open, &Detach from Central, 30
	If ErrorLevel
	{
		Gui, Launch:Destroy
		PrettyMsg("There seems to be an issue launching your project. Check Revit for any opened dailog boxes and try launching again.`n`nThanks.")
		LogMe("Launcher", "Error", "OpenWait", projectID, revitPath, revitTitle, localPath)
		ReloadMe("noshow")
	}
}
attempt := 0
gosub, CheckOpenID
if attempt > 5
{
	PrettyMsg(programName . " could not correctly identify your session of Revit. Please try again.`n`nFeel free to contact " . bimGuy . " should this problem persist.")
	LogMe("Launcher", "Error", "openID", "projectID", projectID, "revitPath", revitPath, "revitTitle", revitTitle, "localPath", localPath)
	ReloadMe()
	return
}
if openID = 0x0
{
	gosub, CheckOpenID
	return
}

ControlGet, fileHwnd, Hwnd,, Edit1, ahk_ID %openID%
ControlGet, folderHwnd, Hwnd,, SysListView321, ahk_ID %openID%
ControlGet, openHwnd, Hwnd,, Button1, ahk_ID %openID%
if (detach or audit)
	fName := pCentral
Else
	fName := localPath
If workset
{
	StringGetPos, fSplit, fName, `\, r1
	fPath := SubStr(fName, 1, fSplit + 1)
	fNameShort := SubStr(fName, fSplit + 2)
	Control, EditPaste, %fPath%, , ahk_ID %fileHwnd%
	ControlClick, Button1, ahk_ID %openID%,, L, 2, NA
	Sleep 100
	gosub, SpecifyCheck
	ControlSend, Button1, {Down 6}{Enter}, ahk_ID %openID%
}
Else
{
	ControlSend, Edit1, {Ctrl Down}a{Ctrl Up}, ahk_ID %openID%
	Control, EditPaste, %fName%, Edit1, ahk_ID %openID%
}
If detach
{
	attempt := 0
	gosub, DetachRecheck
}
If audit
{
	attempt := 0
	gosub, AuditRecheck
}
attempt := 0
gosub, OpenProject
If workset
{
	LogMe("Launcher", "workset", projectID, fName, fPath)
	WinWait, Opening Worksets,, 60
	WinWaitClose, Opening Worksets
	If detach
	{
		WinWait, Detach Model from Central,, 30
		If ErrorLevel
		{
			PrettyMsg("This project may not have detach properly. Please check to make sure you are not working in the central model.`n`nPlease contact " . bimGuy . " should this problem persist.", "alert", 1)
			LogMe("Launcher", "Error", "detach", "specify", "detach message never found", "projectID", projectID, "fName", fName, "fPath", fPath)
			ReloadMe("noshow")
		}
		else
		{	
			WinActivate, Detach Model from Central
			dmcID := WinActive("Detach Model from Central")
			ControlClick, Button1, ahk_id %dmcID%,, L
			IfWinExist, Detach Model from Central ; In case the first method doesn't work
			{
				ControlClick, Button1, Detach Model from Central,, L
			}
		}
	}
	else
	{
		WinWait, Copied Central Model,, 30
		ControlClick, Button1, Copied Central Model,, L, 2, NA
	}
	Gui, Launch:Destroy
}
Else If detach
{
	
	LogMe("Launcher", "detach", projectID, fName)
	WinWait, Detach Model from Central,, 30
	dmcID := WinActive("Detach Model from Central")
	If ErrorLevel
	{
		PrettyMsg("This project did not detach properly.`n`nPlease contact " . bimGuy . " should this problem persist.")
		LogMe("Launcher", "Error", "detach", "projectID", projectID, "fName", fName)
		ReloadMe("noshow")
	}
	else
	{
		ControlClick, Button1, ahk_id %dmcID%,, L
	}
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
return

DetachClick:
; Click the detach button on the open dialog box
; Repeat until it is actually clicked
Button5State := 0
While !Button5State
{
	if A_Index > 5
		launchError = DetachButton
	WinActivate, Open, &Detach from Central
	ControlClick, Button5, ahk_ID %openID%,, L
	ControlGet, Button5State, Checked,, Button5, ahk_ID %openID%
}
return

AuditClick:
While !Button4State
{
	if A_Index > 5
		launchError = AuditButton
	WinActivate, Open, &Detach from Central
	ControlClick, Button4, ahk_ID %openID%,, L
	WinWait, Audit Warning,, 5
	if ErrorlLevel
		return
	ControlClick, Button1, Audit Warning
	ControlGet, Button4State, Checked,, Button4, ahk_ID %openID%
}
return

; Main Routine Functions
LaunchUpdate(sText)
{
	GuiControl, Launch:, LaunchText, %sText%
}