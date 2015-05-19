#Persistent
#SingleInstance, Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Include ..\Class_CtlColors\Sources\Class_CtlColors.ahk
#Persistent


guiColor1 = 666666
guiColor2 = 9f1d21
guiColor3 = ffffff
employeeList := "S:\03 OFFICE TEMPLATES\Email Signatures\o365users-signatures.csv"
employeeIni := A_ScriptDir "\phonelist.ini"
emp := class_EasyINI(employeeIni)
;~ myarray := []
;~ Loop Read, % employeeList
;~ {
	;~ subarray := StrSplit(A_LoopReadLine, "csv")
	;~ myarray.Insert(subarray)
;~ }

;~ Loop, % myarray.MaxIndex()
;~ {
	;~ if A_Index <= 2
		;~ continue
	;~ username := myarray[A_Index][1]
	;~ if emp.AddSection(username)
	;~ {
		;~ emp[username].first := myarray[A_Index][2]
		;~ emp[username].mInitial := myarray[A_Index][3]
		;~ emp[username].last := myarray[A_Index][4]
		;~ emp[username].cred := myarray[A_Index][5]
		;~ emp[username].title := myarray[A_Index][6]
		;~ emp[username].phone1 := myarray[A_Index][7]
		;~ emp[username].phone2 := myarray[A_Index][8]
	;~ }
;~ }
;~ emp.Save()

;~ myarray := ""
;~ StrSplit(InputVar, Delimiter="", OmitChars="") {
	;~ array := []
	;~ Loop Parse, InputVar, %Delimiter%, %OmitChars%
		;~ array.Insert(A_LoopField)
	;~ return array
;~ }
;GUI Variables
guiMainWidth := 300
guiMainLeft := guiMainWidth / 2
guiMainRight := guiMainWidth - guiMainLeft - 20
guiRightLoc := guiMainLeft + 20
SysGet, work, MonitorWorkArea
guiMainHeight := workBottom
guiListHeight := guiMainHeight - 30 - 30 - 47
Gui, Main:New, +HwndMainGui -Caption, Phone List
guiMainLocX := workRight - guiMainWidth - 20
guiMainLocY := workBottom - 38

sections := emp.GetSections(, "C")
names := ""
Loop, Parse, sections, `n
{
	max := A_Index
	mInitial := if emp[A_LoopField].mInitial ? emp[A_LoopField].mInitial . " " : ""
	name := emp[A_LoopField].first . " " . mInitial . emp[A_LoopField].last 
		. "|" . emp[A_LoopField].phone1
	names := if A_Index = 1 ? name : names . "`n" . name
}
Sort, names
guiEmpHeight := ((workBottom - 18) / max)
guiEmpLoc := guiEmpHeight
pointSize := Round((guiEmpHeight-8) * .75, 0)
Gui, Main:Font, s%pointSize%, Arial
idNames := Array()
Loop, Parse, names, `n
{
	StringSplit, thisLine, A_LoopField, |
	textOptions := if A_Index = 1 ? "Section" : "yp+" guiEmpLoc
	Gui, Main:Add, Text
		, %textOptions% Right w%guiMainLeft% h%guiEmpHeight% gListPop hwndidName
		, %thisLine1%
	idNames[A_Index] := idName
}
Loop, Parse, names, `n
{
	StringSplit, thisLine, A_LoopField, |
	textOptions := if A_Index = 1 ? "Section xm+" guiRightLoc " ym" : "yp+" guiEmpLoc
	Gui, Main:Add, Text, %textOptions% w%guiMainRight% h%guiEmpHeight% vphone%A_Index%
		, %thisLine2%
}
debug := true

if debug
	Gui, Main:Add, Text, vmouseStatus xm w%guiMainWidth% center, Your Status Here
else
	WhaLogo(guiMainWidth, "Main")
;~ guiMainLocX := workHRight - guiMainWidth - 30 - 30 - 5
Gui, Main:Show, h%guiMainHeight% x%guiMainLocX%
;~ SetTimer, ActiveCheck, 200
OnMessage(0x200, "MouseOver")
PixelGetColor, menuColor, 5, 5
;~ ListVars
backToBlack := []
backToBlack.Push("Test")
return

ListPop:
return

MainGuiEscape:
MainGuiClose:
ExitApp

ActiveCheck:
currentID := WinActive("A")
if (currentID <> HoverGUI) and (currentID <> MainGUI)
	ExitApp


MouseOver(wParam, lParam, Msg, HWND)
{
	Critical
	Global mainGui
	Global guiColor2
	Global idNames
	Global guiMainLocX
	Global pointSize
	Static empInfoLine
	Static oldHWND
	Static empInfoHWND
	Global backToBlack
	SetTimer, ActiveCheck, Off
	MouseGetPos, , , thisWin, thisControl, 2
	Loop, % idNames.MaxIndex()
	{
		inList := if (hwnd = idNames[A_Index]) ? true : false
		if inList
			break
	}
	if (hwnd = empInfoHWND)
		return
	else if (inList and (thisWin = mainGui))
	{
		gosub, EmpInfo
	}
	else
	{
		gosub, EmpClose
	}
	return
	
	EmpInfo:
	gosub, FontReturn
	if (hwnd = oldHWND)
		return
	oldHWND := hwnd
	ControlGetPos, , empY, , , , ahk_id %hwnd%
	Gui, Main:Font, s%pointSize% c%guiColor2%, Arial
	GuiControl, Main:Font, %hwnd%
	backToBlack.Push(hwnd)
	SetTimer, FontReturn, -200
	Gui, EmpInfo:New, -Caption -Parent hwndempInfoHWND
	Gui, EmpInfo:Add, Text, w200 vempInfoLine, Test
	empX := guiMainLocX - 200
	Sleep, 100
	Gui, EmpInfo:Show, x%empX% y%empY%
	return
	
	EmpClose:
	Sleep, 100
	Gui, EmpInfo:Destroy
	gosub, FontReturn
	SetTimer, ActiveCheck, 200
	return
	
	FontReturn:
	for i in backToBlack
	{
		MouseGetPos, , , , currentHWND, 2
		if (currentHWND != i)
		{
			;~ PrettyMsg(backToBlack.MaxIndex())
			Gui, Main:Font, s%pointSize% cBlack, Arial
			GuiControl, Main:Font, %i%
			backToBlack.Remove(A_Index)
		}
		else
			SetTimer, FontReturn, -200
	}
	return
}

/*
oldMouseOver(wParam, lParam, Msg, HWND)
{
	;~ Global guiColor2
	Global idNames
	Static underOn
	Global oldHWND := HWND
	Global oldControl := A_GuiControl
	Static thisY
	Critical
	SetTimer, ActiveCheck, Off
	mouseStatus := "Hover: " hwnd " : " A_GuiControl 
	GuiControl, Main:, mouseStatus, %mouseStatus%
	inList := false
	Loop, % idNames.MaxIndex()
	{
		if (hwnd = idNames[A_Index])
		{
			inList := true
			break
		}
	}
	if inList
	{
		if !underOn
		{
			SetTimer, OverState, -100
			SetTimer, OffState, off
		}
		underOn := true
		SetTimer, UnderSwitch, -200
	}
	else if !underOn
	{
		SetTimer, OffState, -100
		SetTimer, OverState, Off
	}
	return
	
	OverState:
	;~ DebugUpdate("", "OverState Called")
	ControlGetPos, thisX, thisY, , , , ahk_id %oldhwnd%
	Gui, Hover:New, +HWNDHoverGUI -Parent -Caption, Inspect
	Gui, Hover:Font, s10 cBlack, Arial
	Gui, Hover:Add, Text, w150 vhoverStatus, %oldControl% %thisY% %underMouse%
	Gui, Hover:Show, y%thisY%
	return
	
	OffState:
	;~ DebugUpdate("", "OffState Called")
	Gui, Hover:Destroy
	return
	
	UnderSwitch:
	MouseGetPos, , , , underMouse, 2
	inList := false
	Loop, % idNames.MaxIndex()
	{
		if (underMouse = idNames[A_Index])
		{
			inList := true
			break
		}
	}
	GuiControl, Hover:, hoverStatus, %underMouse% vs. %oldHwnd%
	if !inList
	{
		SetTimer, UnderSwitch, Off
		SetTimer, OffState, -1
		underOn := false
	}
	else
		underOn := true
	return
	
}