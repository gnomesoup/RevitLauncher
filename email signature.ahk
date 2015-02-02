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

programName := "WRIGHT HEEREMA | ARCHITECTS"
fileDir = %A_ScriptDir%\emailsig
FileRead, html, support\sample\sample.htm
FileRead, rtf, support\sample\sample.rtf
FileRead, rtf2, support\sample\samplePhone2.rtf
FileRead, txt, support\sample\sample.txt
txtProof =
htmlProof =
rtfProof =
txtP2 := "`%phone2`%`r`n"
htmlP2 =
(
<p class=MsoNormal><span style='font-size:8.0pt;font-family:`"Arial`",sans-serif;color:gray;mso-themecolor:background1;mso-themeshade:128;mso-bidi-font-weight:bold'>`%phone2`%<o:p></o:p></span></p>
)
MainInterface:
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


Gui, Main:New, +Hwnd, %programName%
Gui, Main:Margin, %marginX%, 20
Gui, Main:Font, s33 cBlack, Arial
Gui, Main:Add, Text, center w%guiMainWidth%, %guiTitle%
Gui, Main:Font, s12, Arial
Gui, Main:Add, Text, right w%lWidth% yp+75 xm, First:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vfirst gCheck Limit, %first%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Initial:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vmInitial gOptionCheck Limit, %mInitial%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Last:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vlast gCheck Limit, %last%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Title:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vtitle gCheck Limit, %title%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Direct Phone:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vphone1 gCheck Limit, %phone1%
Gui, Main:Add, Text, right w%lWidth% yp+40 xm, Mobile Phone:
Gui, Main:Add, Edit, yp-5 xp+%tLoc% w%tWidth% vphone2 gOptionCheck Limit, %phone2%

buttonWidth := (guiMainWidth - 15) / 2
buttonLocation := buttonWidth + 15
Gui, Main:Font, s24 cBlack, Arial
Gui, Main:Add, Button, xm yp+70 w%buttonWidth% vcancelButton, &Cancel
Gui, Main:Add, Button, Default xp xm+%buttonLocation% w%buttonWidth% vcreateButton, &Create
gosub, Check
gosub, OptionCheck

Gui, Main:Show, AutoSize
return
MainGuiClose:
MainButtonCancel:
MainGuiEscape:

ExitApp

MainButtonCreate:
txtProof := txt
htmlProof := html

username := Substr(first, 1, 1) . last
StringLower, username, username

if (mInitial = "(optional)") or (mInitial = "")
	mInitial := ""
else
	mInitial := " " . mInitial

if RegExMatch(phone1, "^\d\d\d\.\d\d\d.\d\d\d\d$")
	phone1 := phone1 . " (direct)" 

if (phone2 = "(optional)") or (phone2 = "") {
	phone2 := ""
	StringReplace, txtProof, txtProof, %txtP2%, 
	If ErrorLevel
		PrettyMsg("Error removing phone2 from TXT`nErrorLevel:" . ErrorLevel, "exit")
	StringReplace, htmlProof, htmlProof, %htmlP2%, 
	If ErrorLevel 
		PrettyMsg("Error removing phone2 from HTML`nErrorLevel:" . ErrorLevel, "exit")
	rtfProof := rtf
}
else {
	if RegExMatch(phone2, "^\d\d\d\.\d\d\d.\d\d\d\d$")
		phone2 := phone2 . " (mobile)"
	rtfProof := rtf2
}

TextInput("txtProof")
if !PrettyMsg("Review the signature:`n`n" . txtProof, "question", 2)
	return
Gui, Main:Cancel
TextInput("htmlProof")
TextInput("rtfProof")
fileDir = %fileDir%\%username%
IfExist, %fileDir%
	FileRemoveDir, %fileDir%, 1
FileCreateDir, %fileDir%
FileCopyDir, support\sample\sample_files, %fileDir%\%username%_files, 1
FileAppend, %txtProof%, %fileDir%\%username%.txt
FileAppend, %htmlProof%, %fileDir%\%username%.htm
FileAppend, %rtfProof%, %fileDir%\%username%.rtf
PrettyMsg("Your signature can be found in `n" . fileDir, "success")
Run, %fileDir%
ExitApp

OptionCheck:
Gui, Submit, NoHide
mInitial := Trim(mInitial)
phone2 := Trim(phone2)

if mInitial = (optional)
{
	FontChange("mInitial", "gray")
}
else 
	FontChange("mInitial", "black")

if phone2 = (optional)
{
	FontChange("phone2", "gray")
}
else FontChange("phone2", "black")
return

Check:
Gui, Submit, NoHide
allGood := 0
if (first = "")
	allGood += 1
if (last = "")
	allGood += 1
if (title = "")
	allGood += 1
if !RegExMatch(phone1, "^\d\d\d\.\d\d\d.\d\d\d\d")
	allGood += 1
if allGood > 0
	GuiControl, Disable, Button2
else
	GuiControl, Enable, Button2
return

FontChange(guiLabel, guiColor)
{
	global
	Gui, Font, c%guiColor% s12
	GuiControl, Main:Font, %guiLabel%
}

TextInput(sigFile) {
	global
	StringReplace, %sigFile%, %sigFile%, `%first`%, %first%
	StringReplace, %sigFile%, %sigFile%, `%mInitial`%, %mInitial%
	StringReplace, %sigFile%, %sigFile%, `%last`%, %last%
	StringReplace, %sigFile%, %sigFile%, `%title`%, %title%
	StringReplace, %sigFile%, %sigFile%, `%phone1`%, %phone1%
	StringReplace, %sigFile%, %sigFile%, `%phone2`%, %phone2%
	StringReplace, %sigFile%, %sigFile%, `%username`%, %username%
}