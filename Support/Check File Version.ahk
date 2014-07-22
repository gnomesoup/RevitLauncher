file := FileOpen("S:\AutoCAD & Revit Standards\Revit Standards\Support\Automation\Working Files\Revit Launcher Test 2015\20150000 Revit Launcher Test-CENTRAL.rvt", "r")

File.Pos := 10
FileSize := File.Length
line := RTrim(file.ReadLine(), "`n")
ListVars
MsgBox, % line

Loop, Read, S:\AutoCAD & Revit Standards\Revit Standards\Support\Automation\Working Files\Revit Launcher Test 2015\20150000 Revit Launcher Test-CENTRAL.rvt
{
	MsgBox, % A_LoopReadLine
	if A_Index > 20
		break
}