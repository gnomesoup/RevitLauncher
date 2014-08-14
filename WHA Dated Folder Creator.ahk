#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#NoTrayIcon
#InstallKeybdHook
#KeyHistory
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; If the file path is passed on to the script, then use it
if %0% > 0
	mainDir = %1%
else ; otherwise just use the scripts own directory
	mainDir := A_ScriptDir

programName := "WRIGHT HEEREMA | ARCHITECTS"

FormatTime, fYear, , yyyy
FormatTime, fMonth, , MM
FormatTime, fDay, , dd
cAbbr := []
cValue := []
; Get a list of the current recipients
Loop, %mainDir%\*.search-ms
{
	n := InStr(A_LoopFileName, "(")
	cAbbr[A_Index] := Trim(SubStr(A_LoopFileName, n+1))
	cAbbr[A_Index] := SubStr(cAbbr[A_Index], 1, InStr(cAbbr[A_Index], ")")-1)
	cValue[A_Index] := Trim(SubStr(A_LoopFileName, 1, n-1))
}

defaultAbbr := ["OWN", "TEN", "PM", "GC", "CIV", "MEP", "STR", "LIT", "FUR", "COD", "ACO", "ELE", "FS", "HW", "LND", "IT", "SEC", "SPC", "SUS", "TRF"]

defaultValue := ["Owner", "Tenant", "Project Manager", "General Contractor", "Civil", "MEP-FP", "Structural", "Lighting", "Furniture", "City-Code Consultant", "Acoustic", "Elevator", "Food Service", "Hardware", "Landscape", "Low Voltage-IT-AV", "Security", "Spec Writer", "Sustainability", "Traffic"]

;~ PrettyMsg(cValue[1], "exit")
;~ Start of main user interface

MainInterface:
guiTitle := "Dated Folder Creator"
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

Gui, Main:New, +Hwnd, %programName%
Gui, Main:Margin, %marginX%, 20
Gui, Main:Font, s34 cBlack, Arial
Gui, Main:Add, Text, center w%guiMainWidth%, %guiTitle%
Gui, Main:Font, s12, Arial
Gui, Main:Add, Text, right w%lWidth% yp+70, Date:
Gui, Main:Add, Edit, center w65 yp-5 xp+%tLoc% vfYear gFolderUpdate Limit4 Number, %fYear%
Gui, Main:Add, Text,  xp+70, -
Gui, Main:Add, Edit, center w35 xp+12 vfMonth gFolderUpdate Limit2 Number, %fMonth%
Gui, Main:Add, Text,  xp+40, -
Gui, Main:Add, Edit, center w35 xp+12 vfDay gFolderUpdate Limit2 Number, %fDay%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Description:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vfDesc gFolderUpdate Limit, %fDesc%
Gui, Main:Add, GroupBox, w%guiMainWidth% vcBoxGroup h175 yp+40 xm
x := lWidth - 70
Gui, Main:Add, Text, right w70 yp xm+%x%, Issued to:
;------------Upper--------------

;---------Check Boxes-----------
Gui, Main:Font, %issuedFont%, Arial
addButtonY := 135
Loop, % cAbbr.MaxIndex()
{
	Gui, Main:Add, Checkbox, xm ym vcBox%A_Index% gFolderUpdate
		, % cAbbr[A_Index]
	Gui, Main:Add, Text, xm ym vcLabel%A_Index%, % cValue[A_Index]
}
; If there is nothing in the list, adjust some things
addButtonY := if (addButtonY = 135) ? 141 : addButtonY 
Gui, Main:Font, s12, Arial
n := addButtonY+5
Gui, Main:Add, button, xm ym vcBoxNew gAbbrNew, +Add a Recipient
;---------Check Boxes-----------

;------------Lower--------------
Gui, Main:Add, Text, right w%lWidth% ym xm vfolderLabel, Folder Name:
Gui, Main:Font, s24, 
Gui, Main:Add, Text, yp+40 center w%guiMainWidth% vfolderText xm
buttonWidth := (guiMainWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Main:Font, s24 cBlack, Arial
Gui, Main:Add, Button, xm yp+70 w%buttonWidth% vcancelButton, &Cancel
Gui, Main:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth% vcreateButton, &Create
GuiControl, Focus, fDesc
gosub, CheckArrange
Gui, Main:Show, AutoSize
folderNameWidth := guiMainWidth

gosub, FolderUpdate
return

CheckArrange:
checkCount := cAbbr.MaxIndex()
upperY := 162
; If Checkbox count < 7 then center and stack
if checkCount < 7
{
	cBoxX := marginX + tLoc
	cLabelX := cBoxX + 60
	cBoxY := upperY + 30
	Loop, % checkCount
	{
		GuiControl, Main:MoveDraw, cBox%A_Index%, x%cBoxX% y%cBoxY%
		GuiControl, Main:MoveDraw, cLabel%A_Index%, x%cLabelX% y%cBoxY%
		cBoxY += cVert
	}
	GuiControl, Main:MoveDraw, cBoxNew, x%cBoxX% y%cBoxY%
	cBoxY += cVert
}
; Otherwise Put them into two stacks but don't forget the add button
else
{
	cBoxY := upperY + 30
	cBoxRightY := cBoxY
	cBoxX := marginX + 20
	cBoxRightX := marginX + (guiMainWidth / 2) + 20
	Loop, % checkCount
	{
		if (A_Index < (checkCount + 3) / 2)
		{
			x := cBoxX
			y := cBoxY
		}
		else
		{
			x := cBoxRightX
			y := cBoxRightY
		}
		labelX := x + 60
		GuiControl, Main:MoveDraw, cBox%A_Index%, x%x% y%y%
		GuiControl, Main:MoveDraw, cLabel%A_Index%, x%labelX% y%y%
		if (A_Index < (checkCount + 3) / 2)
			cBoxY += cVert
		else
			cBoxRightY += cVert
		GuiControl, Main:MoveDraw, cBoxNew, x%cBoxRightX% y%cBoxRightY%
		cBoxRighttY += cVert
	}
}
; Move the lower portion to fit the checks
cBoxHeight := cBoxY - upperY + cVert
GuiControl, Main:MoveDraw, cBoxGroup, h%cBoxHeight%
lowerY := cBoxY - oldcBoxY + cVert
RelativeMove(0, lowerY, "Main", "folderText", "folderLabel", "cancelButton", "createButton")
return

FolderUpdate:
; Grab the information from the form but don't hide it
Gui, Main:Submit, NoHide

; Reset some variables
sentTo := ""

; ###Check the validity of the information###

; Check the year
; With the 'validEntry' function, we have to pass the variable
; name as character string so it can know it's own name.

; Regex explaination:
; ^ = The begining of the string
; backslash 'd' = any single digit
; $ = The end of the string
; (?!00) is a look ahead that says the value cannot equal 00
; [1-9] will match the numbers 1 to 9, so no zero or letters or characters
; the veritical pipe is "or"

fYear := validEntry("fYear", "^20\d\d$", "Main")

; Check the month. If it is too short, add a zero
if (StrLen(fMonth) = 1)
	fMonth := 0 . fMonth
fMonth := validEntry("fMonth", "(?=(0[1-9]|1[012]))^\d\d$", "Main")

; Check the day. If it is too short, add a zero
if (StrLen(fDay) = 1)
	fDay := 0 . fDay
fDay := validEntry("fDay", "(?=(0[1-9]|[1-2]\d|3[0-1]))^\d\d$", "Main")

; Group up the checked variables and combine them with a '+'
Loop, % cAbbr.MaxIndex()
{
	if cBox%A_Index%
	{
		sentTo := sentTo . cAbbr[A_Index] . "+"
	}
}


; Remove the last plus from the end of the list
sentTo := SubStr(sentTo, 1, StrLen(sentTo) - 1)

; If the description is used, add a space to the end.
if fDesc
	fDesc := Trim(fDesc) . " "

; Compile the final folder name and remove any extra spaces
folderName := fYear . "-" . fMonth . "-" . fDay . " " . fDesc . sentTo
folderName := Trim(folderName)
; Find the maximum size of the folder name so it displays nicely
folderNameFont := GetFontMax(folderName, folderNameWidth, 24, 8)

; Update the window
Gui, Main:Font, %folderNameFont%, Arial
GuiControl, Main:Font, folderText
GuiControl, Main:, folderText, %folderName%

; Don't allow the folder to be created without the correct information
If InStr(folderName, "??") || StrLen(folderName) < 11
	GuiControl, Main:Disable, createButton
else
	GuiControl, Main:Enable, createButton
return


MainGuiClose:
MainButtonCancel:

ExitApp

MainButtonCreate:
; Check if the desired folder name already exists
IfExist, %folderName%
{
	whereTo := PrettyMsg("The folder:`n`n'" . folderName . "'`n`nAlready exists. Would you like it opened?", "question",, guiMainWidth)
	if whereTo = Yes
		run %mainDir%\%folderName%
	else if !whereTo
		return
	ExitApp
}
FileCreateDir, %mainDir%\%folderName%
ExitApp
return

AbbrNew: ; Create a new window to add a recipient
oldcBoxY := cBoxY + cVert
guiAbbrWidth := 300
guiTitle := "Add a Recipient"
bWidth := (guiAbbrWidth - 15) / 2
bLoc := bWidth + 15
Gui, Main:+Disabled
Gui, Abbr:New, +HwndlistboxWindow +OwnerMain, %programName%
Gui, Abbr:Margin, 40, 20
gFont := GetFontMax(guiTitle, guiAbbrWidth)
Gui, Abbr:Font, %gFont%, Arial
Gui, Abbr:Add, Text, w%guiAbbrWidth%, %guiTitle%

; Create list of abbreviations
DDLlist := []
Loop, % defaultAbbr.MaxIndex()
{
	DDLlist := DDLlist . defaultAbbr[A_Index] . " - " . defaultValue[A_Index] . "|"
}

; Remove items that are already listed
Loop, % cAbbr.MaxIndex()
	DDLlist := RegExReplace(DDLlist, cAbbr[A_Index] . " - [^\|]*\|")
DDLlist := DDLlist . "«Manual Entry»"
;~ gFont := GetFontMax(RegExReplace(DDLlist, "\|", "`n"), guiAbbrWidth - 30, , 10)
Gui, Abbr:Font, s12
Gui, Abbr:Add, ListBox, vabbrChoice w%guiAbbrWidth% r18 Choose1 Multi gabbrBox, % DDLlist
Gui, Abbr:Font, s24
Gui, Abbr:Add, Button, vabbrButtonOk Default w%guiAbbrWidth% gAbbrSubmit, &OK
Gui, Abbr:Show
n := 1
return

abbrBox:
; If the user did not double click, return to the menu
if A_GuiControlEvent <> DoubleClick
	return 

AbbrSubmit:
Gui, Abbr:Submit, NoHide
if InStr(abbrChoice, "«Manual Entry»")
	gosub, ManualAdd
else
	gosub, AbbrParse
return


ManualAdd:
addText := ""
Loop, Parse, abbrChoice, |
{
	if InStr(A_LoopField, "-")
		addText := addText . Trim(SubStr(A_LoopField, 1, 3)) . "+"
}
GuiControl, Abbr:Hide, abbrChoice
GuiControl, Abbr:Hide, abbrButtonOk
Gui, Abbr:Font, s18
Gui, Abbr:Add, Text, w%guiAbbrWidth% ym+75 xm center vaddAbbrPreview, %addText%???
Gui, Abbr:Font, s12
Gui, Abbr:Add, Text, w95 right xm ym+125, Abbreviation:
Gui, Abbr:Add, Edit, w75 xp+100 yp-5 vaddAbbr Limit3 Uppercase gAbbrManualCheck, 
Gui, Abbr:Add, Text, w95 right xm yp+45, Name:
Gui, Abbr:Add, Edit, w200 xp+100 yp-5 vaddValue Limit25 gAbbrManualCheck,
Gui, Abbr:Font, s24, Arial
Gui, Abbr:Add, Button, w%bWidth% xm ym+200, &Cancel
Gui, Abbr:Add, Button, w%bWidth% xp+%bLoc% gAbbrManualSubmit vabbrManualOK Default, &OK
GuiControl, Abbr:Disable, abbrManualOK
Gui, Abbr:Show, AutoSize
return

AbbrManualCheck:
Gui, Abbr:Submit, NoHide
if addAbbr && addValue
	GuiControl, Abbr:Enable, abbrManualOK
if addAbbr
	GuiControl, Abbr:, addAbbrPreview, %addText%%addAbbr%
return

AbbrManualSubmit:
Gui, Abbr:Submit, NoHide
Loop, % cAbbr.MaxIndex()
{
	If cAbbr[A_Index] = addAbbr
	{
		PrettyMsg(addAbbr . " has already been assigned.`nPlease use a different abbreviation.", "alert", 1, guiMainWidth)
		return
	}
}
manualReplace := addAbbr . " - " . addValue
StringReplace, abbrChoice, abbrChoice, «Manual Entry», %manualReplace%

abbrParse:
newEntry := ""
abbrMax := if cAbbr.MaxIndex() ? cAbbr.MaxIndex() : 0
addText := ""
; Split up the returned string to check for multi-selection
Loop, Parse, abbrChoice, |
{
	; Use hyphen as delimiter to separate the name/value
	n := InStr(A_LoopField, "-")
	addAbbr := (Trim(SubStr(A_LoopField, 1, n - 1)))
	addValue := (Trim(SubStr(A_LoopField, n + 1)))
	cAbbr.Insert(addAbbr)
	cValue.Insert(addValue)
	abbrMax += 1
	cBox%abbrMax% := true
	; Add check boxes to the main window
	Gui, Main:Font, %issuedFont% cBlack, Arial
	Gui, Main:Add, Checkbox, xm+%tLoc% ym+%addButtonY% vcBox%abbrMax% gFolderUpdate Checked, % cAbbr[abbrMax]
	Gui, Main:Add, Text, xp+%tagLoc% vcLabel%abbrMax%, % cValue[abbrMax]
	addButtonY += cVert
	; Add a search file in the main folder 
	addName := ""
	searchFullPath = %mainDir%\%addValue% (%addAbbr%).search-ms
	searchText =
	(
	<?xml version="1.0"?>
	<persistedQuery version="1.0"><viewInfo iconSize="32" stackIconSize="0" displayName="Search Results in Test" autoListFlags="0"><visibleColumns><column viewField="System.ItemNameDisplay"/><column viewField="System.DateModified"/><column viewField="System.ItemTypeText"/><column viewField="System.Size"/><column viewField="System.ItemFolderPathDisplayNarrow"/></visibleColumns><sortList><sort viewField="System.Search.Rank" direction="descending"/><sort viewField="System.DateModified" direction="descending"/><sort viewField="System.ItemNameDisplay" direction="ascending"/></sortList></viewInfo><query><conditions><condition type="leafCondition" property="System.Generic.String" operator="wordmatch" propertyType="string" value="%addAbbr%" localeName="en-US"><attributes><attribute attributeID="{9554087B-CEB6-45AB-99FF-50E8428E860D}" clsid="{C64B9B66-E53D-4C56-B9AE-FEDE4EE95DB1}" chs="1" sqro="585" timestamp_low="1360499943" timestamp_high="30385846"><condition type="leafCondition" property="System.Generic.String" operator="wordmatch" propertyType="string" value="%addAbbr%" localeName="en-US"><attributes><attribute attributeID="{9554087B-CEB6-45AB-99FF-50E8428E860D}" clsid="{C64B9B66-E53D-4C56-B9AE-FEDE4EE95DB1}" chs="1" sqro="585" timestamp_low="592413142" timestamp_high="30385846"><condition type="leafCondition" property="" operator="imp" propertyType="string" value="%addAbbr%" localeName="en-US"><attributes><attribute attributeID="{9554087B-CEB6-45AB-99FF-50E8428E860D}" clsid="{C64B9B66-E53D-4C56-B9AE-FEDE4EE95DB1}" chs="0" parsedString="%addAbbr%" localeName="en-US" timestamp_low="592413142" timestamp_high="30385846"/></attributes></condition></attribute></attributes></condition></attribute></attributes></condition></conditions><kindList><kind name="item"/></kindList><scope><include path="::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\%mainDir%" attributes="1887437183"/></scope></query><properties><author Type="string">Dated Folder Creator</author></properties></persistedQuery>
	)
	FileDelete, %searchFullPath%
	FileAppend, %searchText%, %searchFullPath%
}
gosub, AbbrFinish
return

AbbrFinish: ; Move the remaining items down to make room for the new check boxes
Gui, Main:-Disabled
Gui, Abbr:Submit
gosub, CheckArrange
Gui, Main:Show, AutoSize
gosub, FolderUpdate
return


AbbrGuiEscape:
AbbrButtonCancel:
AbbrGuiClose:
Gui, Main:-Disabled
Gui, Abbr:Hide
return

validEntry(haystack, needle, guiName)
{
	global
	If RegExMatch(%haystack%, "O)" . needle, match)
	{
		Gui, Font, cDefault s12
		GuiControl, %guiName%:Font, % haystack
		return match.Value
	}
	else
	{
		Gui, Font, cRed s12
		GuiControl, %guiName%:Font, % haystack
		Gui, Font, cDefault
		if(haystack = "fYear")
			return "????"
		return "??"
	}
	
}

RelativeMove(sX=0, sY=0, sGuiName="", sControl*)
{
	; This function will move any controls passed to it by the amount specified in the first
	; two variables.
	; In order to control the elements, sGuiName and sControl must be passed as strings
	Loop, % sControl.MaxIndex()
	{
		value := sControl[A_Index]
		GuiControlGet, %value%, %sGuiName%:Pos
		newX := %value%X + sX
		newY := %value%Y + sY
		GuiControl, %sGuiName%:MoveDraw, %value%, x%newX% y%newY%
	}
}