#SingleInstance, Force
#Include %A_ScriptDir%\Demo Menu Function.ahk
M := []
M := ["ATIMenu","AllMenu"]
c := 1
Build_Menus()
Gui, Menu, % M[c]
Gui, Add, Button, x200 y200 gChange, Change Up! 
Gui, Show, w479 h379, Add to Inventory
Return

Change:
  C += 1
  Gui, Menu
  Gui, Menu, % M[c]
  Return

UndoChange:
OHSelAll:
Backup:
MsgBox, %A_ThisMenuItem%  
Return