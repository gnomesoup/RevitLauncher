#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#NoTrayIcon
#InstallKeybdHook
#KeyHistory
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; If the file path is passed on to the script, then use it
if %0% > 0
	mainDir = %1%
else ; otherwise just use the scripts own directory
	mainDir := A_ScriptDir

; Default location for support
supportDir := "S:\04 SOFTWARE\WHA Revit Launcher Settings"
; Default location of the ini file that has the rest of the settingss
iniPath := supportDir . "\Email Signature Settings.ini"
IfNotExist, %iniPath%
	PrettyMsg("The settings file for the email signature creator could not be found. Please check for the following file:`n`n" .  iniPath, "exit")
FormatTime, archiveDate, , yyyy-MM-dd
; Grab settings from an ini file to allow things to change in the future
gosub, IniSettingCheck

; Default name to show for all GUI's
programName := "WRIGHT HEEREMA | ARCHITECTS"

; Load samples to create sigs from
FileRead, html, %sampleDir%\sample.htm
FileRead, rtf, %sampleDir%\sample.rtf
FileRead, txt, %sampleDir%\sample.txt
FileRead, htmlReply, %sampleDir%\sample_reply.htm
FileRead, rtfReply, %sampleDir%\sample_reply.rtf
FileRead, txtReply, %sampleDir%\sample_reply.txt


MainInterface:
; Set up the main GUI

; Set up som variables for the GUI
guiTitle := "Email Signature Maker"
guiMainWidth := 450
initialized := 0
lWidth := 100
tLoc := lWidth + 15
tagLoc := 100
cVert := 24
tWidth := guiMainWidth - tLoc
issuedFont := "s10"
marginX := 40
oldcBoxY := 0
mInitial = (optional)
phone2 = (optional)
phone1 = 312.356.

; Prefill settings for debuging
; because I don't want to type it in each time
; comment this out for production
;~ if !A_IsCompiled
;~ {
	;~ first = Billy
	;~ last = Franklin
	;~ mInitial = B.
	;~ cred = P.O.O.P.
	;~ title = Chief Executive Manure Scooper
	;~ phone1 = 312.867.5309
	;~ phone2 = 312.531.9009
;~ }

Gui, Main:New, +Hwnd, %programName%
Gui, Main:Margin, %marginX%, 20
whaLogo(guiMainWidth, "Main")
Gui, Main:Font, s33 cBlack, Arial
Gui, Main:Add, Text, center w%guiMainWidth% xm, %guiTitle%
Gui, Main:Font, s12, Arial
; Ask for single signature update details
; First name, required
Gui, Main:Add, Text, right w%lWidth% yp+75 xm, First:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vfirst gCheck Limit, %first%
; Middle name or initial, optional
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Initial:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vmInitial gCheck Limit, %mInitial%
; Last name, required
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Last:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vlast gCheck Limit, %last%
; Credentials after name like "AIA" or "LEED AP BD+C", optional
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Credentials:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vcred gCheck Limit, %cred%
; User's title, required
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Title:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vtitle gCheck Limit, %title%
; Direct ring central phone number, required
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Direct Phone:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vphone1 gCheck Limit, %phone1%
; Cell number, optional
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Mobile Phone:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vphone2 gCheck Limit, %phone2%

; Add buttons to the user interface
buttonWidth := (guiMainWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Main:Font, s24 cBlack, Arial
Gui, Main:Add, Button, xm yp+70 w%buttonWidth% vcancelButton, &Cancel
Gui, Main:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth% vcreateButton, &Create
Gui, Main:Add, Button, xm w%guiMainWidth% vbatchButton gBatchCreateGui, &Batch Create
Gui, Main:Add, Button, xm w%guiMainWidth% vsettingsButton, &Settings

; Disable all buttons by default so we can wait for the program to finish loading.
GuiControl, Main:Disable, Button1
GuiControl, Main:Disable, Button2
GuiControl, Main:Disable, Button3
GuiControl, Main:Disable, Button4

; Create an unframed GUI with some thinking dots to the the user know that the
; program is loading.
Gui, Thinking:new
Gui, Thinking:Margin, %marginX%, 20
Gui, Thinking:-SysMenu -Caption
Gui, Thinking:+OwnerMain
Gui, Main:+Disabled
guiThinkingWidth := guiMainWidth + marginX + marginX
guiThinkingHeight := 275
thinkingHalf := (guiThinkingWidth / 2) - 60
thinkingHeight := (guiThinkingHeight / 2) - 60
Gui, Thinking:Font, cBlack s72, Arial
Gui, Thinking:Add, Text, w40 x%thinkingHalf% y%thinkingHeight% vthinkingText1, •
Gui, Thinking:Font, cGray s72, Arial
Gui, Thinking:Add, Text, wp yp xp+40 vthinkingText2, •
Gui, Thinking:Add, Text, wp yp xp+40 vthinkingText3, •

; Show the main interface
Gui, Main:Show, AutoSize
; subroutine to check for invalid input into the single signature info above
gosub, Check

; Get the position of the main gui so we can show the thinking dots in the same place
WinGetPos, mainX, mainY, , , A
thinkingX := mainX + 5
thinkingY := mainY + 150
; overlay the thinking dots on the main gui
Gui, Thinking:Show, w%guiThinkingWidth% h%guiThinkingHeight% x%thinkingX% y%thinkingY%

; Set a timer to make the dots move while thinking
; thinking state is started at 2 to make sure that the middle ball will
; be the next to change
thinkingState = 2
SetTimer, ThinkingTextChange, 700

; Set up the signature to give options for a single phone or two phones
gosub, phoneLineSingle
Gui, Main:-Disabled

; Enable all the buttons now that we are done setting up signatures
; Some sleeps are interjected to make a nice slow release of the GUI
; Looks a lot better then just activating everything at once.
GuiControl, Main:Enable, Button1
Sleep, 100
GuiControl, Main:Enable, Button3
Sleep, 100
GuiControl, Main:Enable, Button4
Sleep, 400
Gui, Thinking:Destroy
SetTimer, ThinkingTextChange, Off
return

; Subroutine to make the thinking balls change colors to give user
; feedback that the application is processing
ThinkingTextChange:
if thinkingState = 1
{
	; set ball to left black
	FontChange("thinkingText1", "Black", "Thinking", "72")
	FontChange("thinkingText2", "Gray", "Thinking", "72")
	FontChange("thinkingText3", "Gray", "Thinking", "72")
	ThinkingState = 2
}
else if thinkingState = 2
{
	; set ball in middle black
	FontChange("thinkingText1", "Gray", "Thinking", "72")
	FontChange("thinkingText2", "Black", "Thinking", "72")
	FontChange("thinkingText3", "Gray", "Thinking", "72")
	ThinkingState = 3
}
else if thinkingState = 3
{
	; set ball to right black
	FontChange("thinkingText1", "Gray", "Thinking", "72")
	FontChange("thinkingText2", "Gray", "Thinking", "72")
	FontChange("thinkingText3", "Black", "Thinking", "72")
	ThinkingState = 1
}
return

MainGuiClose:
MainButtonCancel:
MainGuiEscape:
; Handles what happens when user tries to close gui
ExitApp

BatchCreateGui:
; Create a GUI to confirm that we want to make a whole bunch of signatures at once
batch := 1
guiBatchTitle = Batch Signature Maker
batchFileDir := fileDir
Gui, Batch:New, , %programName%
Gui, Batch:+OwnerMain
Gui, Batch:Margin, %marginX%, 20
whaLogo(guiMainWidth, "Batch")
Gui, Batch:Font, s32 cBlack, Arial
Gui, Batch:Add, Text, center w%guiMainWidth% xm, %guiBatchTitle%
Gui, Batch:Font, s12, Arial
Gui, Batch:Add, Text, w%guiMainWidth% xm, You are about to update signatures for all users based off of the comma seperated employee list. The signatures will be placed in:
Gui, Batch:Font, s10, Arial
; Allow the user to temporarily change where the batched signatures
; will be saved
Gui, Batch:Add, Edit, w%guiMainWidth% xm r1 vbatchFileDir gBatchCheck, %batchFileDir%
Gui, Batch:Font, s12, Arial
Gui, Batch:Add, Button, w100 xm yp+30 gBatchViewFile, &View
Gui, Batch:Add, Button, wp xp+110 yp gBatchChangeFile, &Change
Gui, Batch:Font, s12, Arial
Gui, Batch:Add, Text, w%guiMainWidth% xm, This process cannot be undone. Click "Edit Settings" to verify the information in the employee list. All overwritten signatures will be saved in the archive save location. Click "Batch" to proceed.
Gui, Batch:Font, s24 cBlack, Arial
Gui, Batch:Add, Button, xm yp+70 w%buttonWidth%, &Cancel
Gui, Batch:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth% gBatchCreate, &Batch
Gui, Batch:Add, Button, xm w%guiMainWidth% vbatchButton gMainButtonSettings, &Edit Settings
; get the position of the main window so all of our windows stick together
WinGetPos, mainX, mainY, , , A
Gui, Batch:Show, x%mainX% y%mainY%
Gui, Main:Hide
return

BatchGuiClose:
BatchButtonCancel:
BatchGuiEscape:
; what to do when user tries to cancel the batch GUI
Gui, Main:Show
Gui, Batch:Destroy
batch := 0
return

BatchCheck:
; Check the validity of the path specified in the the Batch GUI
Gui, Batch:Submit, NoHide
IfNotExist %batchFileDir%
	FontChange("batchFileDir", "red", "Batch", "10")
else
	FontChange("batchFileDir", "black", "Batch", "10")
return

BatchChangeFile:
; Create a user interface to select a new path for the batched signatures
FileSelectFolder, tempFile, *%batchFileDir%, 3
if tempFile =
	tempFile := batchFileDir
GuiControl, Batch:, batchFileDir, %tempFile%
return

BatchViewFile:
; Let a user open an explorer window showing the batched signature save location
Run %batchFileDir%
return

BatchCreate:
; Do the hard work of creating the batch files, but make sure we are up to it first
if !PrettyMsg("Are you sure you would like to update all signatures in the following folder?`n`n" . batchFileDir, , 2)
	return
Gui, Batch:Submit
; Start counting the number of signatures
n := 0
batchError := ""
; Set up a GUI that shows the progress of the script
gosub, GuiConfirm
; Read through the CSV file that has all of the employee information and
; create a signature from the data.
Loop, Read, %employeeList%
{
	batchErrorCount := 0
	If A_Index < 3
		Continue
	If A_Index > 2000
		break
	Loop, parse, A_LoopReadLine, CSV
	{
		if (A_Index = 1)
			username = %A_LoopField%
		if (A_Index = 2)
			first = %A_LoopField%
		if (A_Index = 3)
			mInitial = %A_LoopField%
		if (A_Index = 4)
			last = %A_LoopField%
		if (A_Index = 5)
			cred = %A_LoopField%
		if (A_Index = 6)
			title = %A_LoopField%
		if (A_Index = 7)
			phone1 = %A_LoopField%
		if (A_Index = 8)
			phone2 = %A_LoopField%
	}
	if !username ; make sure the line is not blank or mis-formated
		batchErrorCount += 1
	if !first
		batchErrorCount += 1
	if !last
		batchErrorCount += 1
	if !title
		batchErrorCount += 1
	if !phone1
		batchErrorCount += 1
	if batchErrorCount > 0
	{
		batchError := batchError . "Line " . A_Index . ": " . username
		continue
	}
	gosub, CheckDetails ; correctly format the data for the signature
	gosub, CreateSig ; actually create the signature for a single user
	n += 1 ; put a notch in our belt and advance the signature count
	GuiControl, Confirm:, statusText, Please wait. %n% signatures created so far.
}
; Let us feel good about getting some work done.
PrettyMsg("Complete!`nCreated " . n . " signatures")
If batchError
	PrettyMsg("The following lines of the Employee List were missing critical information to create a signature. If signatures are required for these users, please check their information and try again`n`n" . batchError, "alert", 1)
; Show us all the pretty signatures and exit
Run %batchFileDir%
ExitApp

MainButtonCreate:
; Do all the work of creating a single signature
username := Substr(first, 1, 1) . last
StringLower, username, username

gosub, CheckDetails ; reformat the data to look good in the signature
batch := 0
createAbort := 0

gosub, CreateSig ; actually make the signature
if createAbort
	return

; add or update the information from the single signature in the employee list
gosub, CsvAddName 

; confirm that the signature was made and show them the files
PrettyMsg("Your signature can be found in `n" . sigDir, "success")
Run, %sigDir%
ExitApp


; These functions were included in the Check subroutine
;~ OptionCheck:
;~ Gui, Submit, NoHide
;~ mInitial := Trim(mInitial)
;~ phone2 := Trim(phone2)

;~ if mInitial = (optional)
;~ {
	;~ FontChange("mInitial", "gray")
;~ }
;~ else 
	;~ FontChange("mInitial", "black")

;~ if phone2 = (optional)
;~ {
	;~ FontChange("phone2", "gray")
;~ }
;~ else FontChange("phone2", "black")
;~ return

Check:
; Check that all of the data entered into the single signature line is valid
Gui, Submit, NoHide

; Get rid of any extra spaces that may have been accidentally entered.
first := Trim(first)
mInitial := Trim(mInitial)
last := Trim(last)
cred := Trim(cred)
title := Trim(title)
phone1 := Trim(phone1)
phone2 := Trim(phone2)

; Make the text for the middle inital field grey to start
if mInitial = (optional)
{
	FontChange("mInitial", "gray")
}
else 
	FontChange("mInitial", "black")

; run through a check to make sure all required variables are present and
; correctly formated
allGood := 0
if (first = "")
	allGood += 1
if (last = "")
	allGood += 1
if (title = "")
	allGood += 1
if !RegExMatch(phone1, "^(\d{3}\.){2}\d{4}")
	allGood += 1

; this regex statement will keep the phone variables valid as long as they are
; being filled in correctly
phoneSlowMatch :="^\d{1,3}$|^\d{3}\.?$|^\d{3}\.\d{1,3}$|^\d{3}.\d{3}\.$|^\d{3}\.\d{3}\.\d{1,4}$|\d{3}\.\d{3}\.\d{4}"

; check the phone numbers to validity
if !RegExMatch(phone1, phoneSlowMatch)
	FontChange("phone1", "red")
else
	FontChange("phone1", "black")

; if phone two is in it's default optional state, make it grey
if phone2 = (optional)
{
	FontChange("phone2", "gray")
}
else ; check phone two for validity if filled out
{
	if !RegExMatch(phone2, "^$|^\d{3}\.\d{3}\.\d{4}")
		allGood += 1
	if !RegExMatch(phone2, phoneSlowMatch)
		FontChange("phone2", "red")
	else
		FontChange("phone2", "black")
}

; enable the create button if everything looks good
if allGood > 0
	GuiControl, Disable, Button2
else
	GuiControl, Enable, Button2
return

FontChange(guiLabel, guiColor, guiName = "Main", guiFontSize = "12")
{
	; change the color of a named variable
	; guiLabel = the "v" label added to the gui edit element
	; guiColor = the color you would like the font changed to
	; guiName = the name of the GUI the edit field exists in
	; guiFontSize = the size of the font without the "S"
	global
	Gui, %guiName%:Font, c%guiColor% s%guiFontSize%
	GuiControl, %guiName%:Font, %guiLabel%
}

TextInput(sigFile) {
	; swap out the placeholders in the sample signature with the desired
	; employee information
	; sigFile = the particular signature file format that needs the info swapped
	global
	StringReplace, %sigFile%, %sigFile%, `%first`%, %first%, , All
	If ErrorLevel
		PrettyMsg("Error 457`n`nStringReplace failed: " . sigFile . " - " . ErrorLevel, "alert")
	StringReplace, %sigFile%, %sigFile%, `%mInitial`%, %mInitial%, , All
	StringReplace, %sigFile%, %sigFile%, `%last`%, %last%, , All
	StringReplace, %sigFile%, %sigFile%, `%cred`%, %cred%, , All
	StringReplace, %sigFile%, %sigFile%, `%title`%, %title%, , All
	StringReplace, %sigFile%, %sigFile%, `%phone1`%, %phone1%, , All
	StringReplace, %sigFile%, %sigFile%, `%phone2`%, %phone2%, , All
	StringReplace, %sigFile%, %sigFile%, `%username`%, %username%, , All
}

CreateSig:
; This is where is all happens.
; all the placeholder data will get swapped, and all of the files that make
; up the signature will be written to the specified directory

; we run the txt version of the signature first and show it to the user
; this only happens for single signatures
TextInput("txtProof")
if (batch = 0) {
	if !PrettyMsg("Review the signature:`n`n" . txtProof, "question", 2)
	{
		createAbort := 1
		return
	}
	Gui, Main:Cancel
}
; now we run all the other versions
TextInput("htmlProof")
TextInput("rtfProof")
TextInput("txtReplyProof")
TextInput("htmlReplyProof")
TextInput("rtfReplyProof")

; Assign the signature directory
; Different directories may be required if we are running in batch mode.
sigDir := if batch ? batchFileDir . "\" . username : fileDir . "\" . username

; Check if the signature already exists
; If it does, then archive the first new one created today
; We don't bother archiving subsequent versions created on the same day
; as they are most likely corrections of mistakes.
IfExist, %sigDir%
{
	StringSplit, fromFolder, sigDir, `\
	fromFolderN := fromFolder0 - 1
	fromFolder := fromFolder%fromFolderN%
	IfNotExist, %archiveDir%\%archiveDate% %fromFolder%
		FileCreateDir, %archiveDir%\%archiveDate% %fromFolder%
	FileMoveDir, %sigDir%, %archiveDir%\%archiveDate% %fromFolder%\%username%
	If ErrorLevel
		FileRemoveDir, %sigDir%, 1
}

; Create the directory for the signature and the sigantures files
FileCreateDir, %sigDir%
FileCopyDir, %sampleDir%\sample_files, %sigDir%\%username%_files, 1
FileCopyDir, %sampleDir%\sample_reply_files, %sigDir%\%username%_reply_files, 1
FileAppend, %txtProof%, %sigDir%\%username%.txt
FileAppend, %htmlProof%, %sigDir%\%username%.htm
FileAppend, %rtfProof%, %sigDir%\%username%.rtf
FileAppend, %txtReplyProof%, %sigDir%\%username%_reply.txt
FileAppend, %htmlReplyProof%, %sigDir%\%username%_reply.htm
FileAppend, %rtfReplyProof%, %sigDir%\%username%_reply.rtf
return

CheckDetails:
; Format the employee data to make it fit well in the sample signature.
; We also create a line for the signature in the correct format for the 
; batch employee list in case we want them added for later use.
batchAddLine := username . "," . first
if (mInitial = "(optional)") or (mInitial = "")
{
	mInitial := ""
	batchAddLine := batchAddLine . "," . mInitial
}
else
{
	batchAddLine := batchAddLine . "," . mInitial
	mInitial := " " . mInitial
}
batchAddLine := batchAddLine . "," . last
batchAddLine := batchAddLine . "," . cred
if !(cred = "")
	cred := ", " . cred
batchAddLine := batchAddLine . "," . title
batchAddLine := batchAddLine . "," . phone1
if RegExMatch(phone1, "^\d\d\d\.\d\d\d.\d\d\d\d$")
	phone1 := phone1 . " (direct)" 

if (phone2 = "(optional)") or (phone2 = "") 
{
	phone2 := ""
	batchAddLine := batchAddLine . "," . phone2
	txtProof := txt1
	txtReplyProof := txt1Reply
	htmlProof := html1
	htmlReplyProof := html1Reply
	rtfProof := rtf1
	rtfReplyProof := rtf1Reply
}
else 
{
	batchAddLine := batchAddLine . "," . phone2
	if RegExMatch(phone2, "^\d\d\d\.\d\d\d.\d\d\d\d$")
		phone2 := phone2 . " (cell)"
	txtProof := txt
	txtReplyProof := txtReply
	htmlProof := html
	htmlReplyProof := htmlReply
	rtfProof := rtf
	rtfReplyProof := rtfReply
}
return

phoneLineSingle:
txt1 := TextRemove(txt, "phone2")
txt1Reply := TextRemove(txtReply, "phone2")
rtf1:= TextRemove(rtf, "phone2")
rtf1Reply := TextRemove(rtfReply, "phone2")
html1 := TextRemove(html, "phone2")
html1Reply := TextRemove(htmlReply, "phone2")
return

TextRemove(sHayStack, sNeedle)
{
	stack := ""
	Loop, Parse, sHayStack, `n
	{
		
		if !RegExMatch(A_LoopField, sNeedle)
			stack := stack . A_LoopField . "`n"
	}
	return stack
}


GuiConfirm:
Gui, Confirm:New, +Hwnd, %programName%
Gui, Confirm:+OwnerMain
Gui, Main:+Disabled
Gui, Confirm:Margin, %marginX%, 20
Gui, Confirm:Font, s33 cBlack, Arial
Gui, Confirm:Add, Text, center w%guiMainWidth%, %guiTitle%
Gui, Confirm:Font, s18, Arial
Gui, Confirm:Add, Text, center w%guiMainWidth% vstatusText, Please wait. %n% signatures created so far.
Gui, Confirm:Show
return

MainButtonSettings:
gosub, IniSettingCheck
Gui, Settings:New, , %programName%
Gui, Settings:+OwnerMain
Gui, Settings:Margin, %marginX%, 20
whaLogo(guiMainWidth, "Settings")
Gui, Settings:Font, s33 cBlack, Arial
Gui, Settings:Add, Text, center w%guiMainWidth% xm, %guiTitle%
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Text, w%guiMainWidth% yp+75 xm, Default Finished Signature Save Location:
Gui, Settings:Font, s10, Arial
Gui, Settings:Add, Edit, w%guiMainWidth% yp+25 vfileDir gSettingsCheck, %fileDir%
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Button, w100 xm yp+30 gviewFile, View
Gui, Settings:Add, Button, wp xp+110 yp gchangeFile, Change
Gui, Settings:Add, Text, w%guiMainWidth% yp+50 xm, Sample Signature Files Location:
Gui, Settings:Font, s10, Arial
Gui, Settings:Add, Edit, w%guiMainWidth% yp+25 vsampleDir gSettingsCheck, %sampleDir%
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Button, w100 xm yp+30 gviewSample, View
Gui, Settings:Add, Button, wp xp+110 yp gchangeSample, Change
Gui, Settings:Add, Text, w%guiMainWidth% yp+50 xm, Employee List (CSV File):
Gui, Settings:Font, s10, Arial
Gui, Settings:Add, Edit, w%guiMainWidth% yp+25 vemployeeList gSettingsCheck, %employeeList%
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Button, w100 xm yp+30 gviewEmployee, View
Gui, Settings:Add, Button, wp xp+110 yp gchangeEmployee, Change
Gui, Settings:Add, Text, w%guiMainWidth% yp+50 xm, Signature Archive Location:
Gui, Settings:Font, s10, Arial
Gui, Settings:Add, Edit, w%guiMainWidth% yp+25 varchiveDir gSettingsCheck, %archiveDir%
Gui, Settings:Font, s12, Arial
Gui, Settings:Add, Button, w100 xm yp+30 gviewArchive, View
Gui, Settings:Add, Button, wp xp+110 yp gchangeArchive, Change
buttonWidth := (guiMainWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Settings:Font, s24 cBlack, Arial
Gui, Settings:Add, Button, xm yp+70 w%buttonWidth% vsettingsCancelButton, &Cancel
Gui, Settings:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth% vsettingsOKButton, &OK 
WinGetPos, mainX, mainY, , , A
Gui, Settings:Show, x%mainX% y%mainY%
if Batch
	Gui, Batch:Destroy
else
	Gui, Main:Hide

return

SettingsButtonCancel:
SettingsGuiEscape:
Gui, Main:Show
Gui, Settings:Destroy
return

SettingsButtonOK:
if Batch
{
	Gui, Batch:-Disabled
	GuiControl, Batch:, fileDirBatch, %fileDir%
}
else
	Gui, Main:-Disabled
Gui, Settings:Submit
settingUpdate := 0
fileDir := RegExReplace(fileDir, "\\$")
sampleDir := RegExReplace(sampleDir, "\\$")
archiveDir := RegExReplace(archiveDir, "\\$")
if !(iniSettings.settings.fileDir = fileDir)
{
	iniSettings.settings.fileDir := fileDir
	settingUpdate += 1
}
if !(iniSettings.settings.sampleDir = sampleDir)
{
	iniSettings.settings.sampleDir := sampleDir
	settingUpdate += 1
}
if !(iniSettings.settings.employeeList = employeeList)
{
	iniSettings.settings.employeeList := employeeList
	settingUpdate += 1
}
if !(iniSettings.settings.archiveDir = archiveDir)
{
	iniSettings.settings.archiveDir := archiveDir
	settingUpdate += 1
}
if (settingUpdate > 0)
	iniSettings.Save()
return



changeFile:
FileSelectFolder, tempFile, *%fileDir%, 3
if tempFile =
	tempFile := fileDir
GuiControl, Settings:, fileDir, %tempFile%
return

viewFile:
Run, %fileDir%
return

changeSample:
FileSelectFolder, tempFile, *%sampleDir%, 3
if tempFile =
	tempFile := sampleDir
GuiControl, Settings:, sampleDir, %tempFile%
return

viewSample:
Run, %sampleDir%
return

changeArchive:
FileSelectFolder, tempFile, *%archiveDir%, 3
if tempFile =
	tempFile := archiveDir
GuiControl, Settings:, archiveDir, %tempFile%
return

viewArchive:
Run, %archiveDir%
return

changeEmployee:
FileSelectFile, tempFile, 1, %employeeList%, , Comma Seperated Value (*.csv)
if tempFile =
	tempFile := employeeList
GuiControl, Settings:, employeeList, %tempFile%
return

viewEmployee:
Run, %employeeList%
return

SettingsCheck: 
; Change text color if folder does not exist
Gui, Settings:Submit, NoHide
;~ GuiControl, Settings:Disable, RenameFiles
IfExist, %fileDir%
{ 
	Gui, Settings:Font, Arial s10 cBlack 
	GuiControl, Font, fileDir
}
else
{
	Gui, Settings:Font, cRed Arial s10
	GuiControl, Settings:Font, fileDir
}
IfExist, %sampleDir%
{ 
	Gui, Settings:Font, Arial s10 cBlack 
	GuiControl, Font, sampleDir
}
else
{
	Gui, Settings:Font, cRed Arial s10
	GuiControl, Settings:Font, sampleDir
}
IfExist, %employeeList%
{ 
	Gui, Settings:Font, Arial s10 cBlack 
	GuiControl, Font, employeeList
}
else
{
	Gui, Settings:Font, cRed Arial s10
	GuiControl, Settings:Font, employeeList
}

return

IniSettingCheck:
iniSettings := class_EasyIni(iniPath)
fileDir := iniSettings.settings.fileDir
sampleDir := iniSettings.settings.sampleDir
employeeList := iniSettings.settings.employeeList
archiveDir := iniSettings.settings.archiveDir
supportDir := iniSettings.settings.supportDir
return

CsvAddName:

FileRead, allEmployees, %employeeList%
if ErrorLevel
{
	PrettyMsg("Error 795`n`nThe employee list could not be found in:`n`n" . employeeList, "alert")
	return
}
else
{
	userRegLine := "m)^" . username . ",.*$"
	if (RegExMatch(allEmployees, userRegLine, userLine))
	{
		if (PrettyMsg("The user """ . username . """ is already in the batch employee list. Would you like to update their information?", "question", 2) = "Yes")
			StringReplace, allEmployees, allEmployees, %userLine%`r`n
		else
			return
	}
	else
	{
		if (PrettyMsg("Would you like " . username . " added to the batch signature creation list?", "question", 2) = "No")
		return
	}
	allEmployees := allEmployees . batchAddLine . "`r`n"
	employeesOnly := ""
	csvHeader := ""
	Loop, Parse, allEmployees, `n
	{
		if (A_Index <= 2)
			csvHeader := csvHeader . A_LoopField . "`n"
		else if A_LoopField
			employeesOnly := employeesOnly . A_LoopField . "`n"
	}
	;~ Sort, employeesOnly
	allEmployees := csvHeader . employeesOnly
	FormatTime, fileTime, , yyyy-MM-dd-hhmmss
	FileMove, %employeeList%, %supportDir%/EmployeeList%fileTime%.csv, 1
	If ErrorLevel
	{
		PrettyMsg("Error 829`n`nThere was a problem adding """ . username . """ to the employee list located in:`n`n" . employeeList, "alert")
		return
	}
	;~ FileAppend, %allEmployees%, *%employeeList%
	file := FileOpen(employeeList, "rw")
	file.Write(allEmployees)
	file.Close
	allEmployees := ""
	employeesOnly := ""
	csvHeader := ""
}
return

;~ batchDataCheck(sVar)
;~ {
	;~ global
	;~ if !sVar ; make sure the line is not blank or mis-formated
	;~ {
		;~ batchError := batchError . "Line " . A_Index . ": " . username
		;~ continue
	;~ }
;~ }