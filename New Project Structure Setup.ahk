#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force ; Only allow one process at a time
#NoTrayIcon
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;~ Initial variables
programName := "Wright Heerema | Architects"
guiWidth := 600
initialized := 0
sourceFolder := "S:\03 OFFICE TEMPLATES\NEW PROJECT FOLDER SET UP"
destinationDrive := if A_IsCompiled ? "X:\" : A_Desktop . "\Temporary\Test\"
FormatTime, currentYear, , yyyy
tooOld := currentYear - 2
tooNew := currentYear + 2

;~ Start of main user interface
Gui, Main:New,, %programName%
Gui, Main:Font, s32, Arial
Gui, Main:Add, Text, center w%guiWidth%, Project Structure Setup
Gui, Main:Font, s18, Arial
textWidth := 120
Gui, Main:Add, Text, center w%guiWidth%, Welcome to the Wright Heerema | Architects project setup. Please provide the following information to continue.
Gui, Main:Font, s12, Arial
Gui, Main:Add, Text, w%textWidth%, Project Number:
Gui, Main:Font, s18, Arial
editWidth := guiWidth - textWidth
Gui, Main:Add, Edit, w%editWidth% yp-10 xp+%textWidth% vpNum gCheck,
Gui, Main:Font, s10 cGray, Arial
Gui, Main:Add, Text, vnumHelp xp yp+40 w%editWidth%,
Gui, Main:Font, s12 cBlack, Arial
Gui, Main:Add, Text, w%textWidth% xm yp+50, Project Name:
Gui, Main:Font, s18 cBlack, Arial
Gui, Main:Add, Edit, w%editWidth% yp-10 xp+%textWidth% vpName gCheck,
Gui, Main:Font, s10 cGray, Arial
Gui, Main:Add, Text, vnameHelp xp yp+40 w%editWidth%, This field must be completed to continue. Keep it short but descriptive.
buttonWidth := (guiWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Main:Font, s24 cBlack, Arial
Gui, Main:Add, Button, xm w%buttonWidth%, &Cancel
Gui, Main:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth%, &OK
GuiControl, Main:Disable, Button2
Gui, Main:Show

;~ Make sure users are putting in good information
goto, Check
return


MainGuiClose:
MainGuiEscape:
MainButtonCancel:
ExitApp


Check:
Gui, Submit, NoHide
dateTest := false
pNum := Trim(pNum)
RegExMatch(pNum, "O)^\d\d\d\d(xxxx|\d\d\d\d|\d\d\d\d\.\d+)$", match)
matchYear := SubStr(match.Value, 1, 4)
if match
{
	if matchYear <= %tooOld%
	{
		dateIssue := match.Value . " is an old project. Please enter a project newer than " . tooOld . "."
		dateTest := false
	}
	else if matchYear >= %tooNew%
	{
		dateIssue := "While I'm all for planning ahead, please enter a project from before " . tooNew . "."
		dateTest := false
	}
	else
	{
		dateIssue := if InStr(match.Value, "xxxx") ? "Try to get a project number assigned rather than using 'xxxx'" : " "
		dateTest := true
	}
}

if pName and dateTest
{
	GuiControl, Main:Enable, Button2
	Gui, Font, cBlack s18
	GuiControl, Main:Font, pNum
}
else
{
	GuiControl, Disable, Button2
	if !dateTest and StrLen(pNum) >= 8
	{
		Gui, Font, cRed s18
	}
	else
	{
		Gui, Font, cBlack s18
	}
	GuiControl, Main:Font, pNum
}
Gui, Font, cBlack s18
GuiControl, Main:Font, pName
dateIssue := if dateIssue ? dateIssue : "Please enter an 8 digit project number with no dashes or spaces, e.g.(" . currentYear . "0520)."
GuiControl,, numHelp, %dateIssue%
dateIssue := false
return

MainButtonOK:
folderName := destinationDrive . SubStr(pNum, 1, 4) . "\" . Trim(pNum) . " " . Trim(pName)
folderName := Trim(folderName)
IfExist, %folderName%
{
	PrettyMsg("A project with the folder '" . folderName . "' already exists. Please check the project name and number and try again.", "alert", 1)
	Gui, Font, cRed s18
	GuiControl Main:Font, pNum
	GuiControl Main:Font, pName
	return
}
Gui, Main:Hide
Gui, Confirm:New,, %programName%
Gui, Confirm:Font, s32, Arial
Gui, Confirm:Add, Text, center w%guiWidth%, Project Structure Setup
Gui, Confirm:Font, s18, Arial
Gui, Confirm:Add, Text, w%guiWidth% center, The project folder structure will be created in the following location:
fontS := GetFontMax(folderName, guiWidth)
Gui, Confirm:Font, %fontS% cGreen, Arial
Gui, Confirm:Add, Text, w%guiWidth% center, %folderName%
Gui, Confirm:Font, s18 cBlack, Arial
Gui, Confirm:Add, Text, w%guiWidth% center, Please ensure the location is correct before proceeding
buttonWidth := (guiWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Confirm:Font, s24
Gui, Confirm:Add, Button, xm w%buttonWidth% yp+75, &Cancel
Gui, Confirm:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth%, &OK
GuiControl, Main:Disable, Button2
Gui, Confirm:Show
return

ConfirmButtonOK:
Gui, Splash:Destroy
Gui, Splash:New,, %programName%
Gui, Font, s32, Arial
Gui, Add, Text, center w%guiWidth%, Project Structure Setup
Gui, Font, s18, Arial
Gui, Add, Text, center w%guiWidth%, Please stand by...`n`nProject folder struture generation in progress...
Gui, Show
FileCopyDir, %sourceFolder%, %folderName%, 0
Gui, Destroy
if ErrorLevel
	PrettyMsg("There was an issue creating the folder structure. Please see your BIM Coordinator for additional information")
else
{
	PrettyMsg("Thank you for using the WHA New Project Structure Setup. Your project has been created and can be found in the following location:`n`n" . folderName, "success")
	ExitApp
}
return

ConfirmGuiClose:
ConfirmGuiEscape:
ConfirmButtonCancel:
Gui, Main:Show
Gui, Confirm:Destroy
return

CopyFilesAndFolders(SourcePattern, DestinationFolder, DoOverwrite = false)
; Copies all files and folders matching SourcePattern into the folder named DestinationFolder and
; returns the number of files/folders that could not be copied.
{
    ; First copy all the files (but not the folders):
	FileCopy, %SourcePattern%, %DestinationFolder%, %DoOverwrite%
	ErrorCount := ErrorLevel
    ; Now copy all the folders:
    Loop, %SourcePattern%, 2  ; 2 means "retrieve folders only".
    {
        FileCopyDir, %A_LoopFileFullPath%, %DestinationFolder%\%A_LoopFileName%, %DoOverwrite%
        ErrorCount += ErrorLevel
		if ErrorLevel  ; Report each problem folder by name.
			MsgBox Could not copy %A_LoopFileFullPath% into %DestinationFolder%.

    }
    return ErrorCount
}