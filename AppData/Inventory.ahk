Menu, tray, Icon , %A_ScriptDir%\Images\01.ico, 1, 1
Version := "1.0.0.0"
#SingleInstance, Force
#NoEnv
#Include %A_ScriptDir%\lib\Functions.ahk
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
loc := "Location", mod := "Model", lblnm := "Label", itm := "Item", i := 1, CT := 7, CE := 1, voipyn := "N", editing := "N", TIItems :=0      ;Default Values
global XML := A_ScriptDir . "\XML\Import"
global ini := A_ScriptDir . "\info.ini"
ArchiveDir := A_ScriptDir . "\Archive"
csvF := A_ScriptDir . "\Inventory.csv", 
FileCopy, %csvF%, %ArchiveDir%\Inventory.%A_YYYY%%A_MM%%A_DD%.%A_Hour%%A_Min%%A_Sec%.Startup.csv, 1
iniread, ArNum, %ini%, Archived Records to Keep
CleanArchive(ArchiveDir,ArNum)
var := "VoIP"

iniread, RegisteredItems, %ini%, RegisteredItems


StringTrimRight, EXEDir, A_ScriptDir, 8

global HRDir := EXEDir . "\Hand Receipts"
global HRTemp := HRDir . "\a2062.pdf"

Gui,  +resize
Gui, Font, s10 ;SegoiUI
Gui, Add, Tab2, x10 y10 w780 h540 BOTTOM vTabArea AltSubmit gsubtab, Add to Inventory|Full Inventory|On-Hand Inventory|Turn-In    ;|Issued

RightClickMenus:
{
  Menu, GridContext, Add, Edit, EditGui
  Menu, GridContext, Add, Delete, DeleteItem

  Menu, CSVContext, Add, Edit, EditCSVGui
  Menu, CSVContext, Add, Delete, DeleteCSV

  Menu, NormalContext, Add, Copy, CopyCell

  Menu, OHContext, Add, Change Quantity, ChgQty
}

TabMenus:
{
  Menu, ItemsMenu, Add, &Register Item Number, RegisterItem
}

Menu, MenuBar, Add, &Items, :ItemsMenu
Gui, Menu, MenuBar


Lists := "LocList|LblList|IList|MList|ITList|NoSNIList|CBIlist"
FLists := "FLocList|FLblList|FIList|FMList|FITList|FNoSNIList"
FFLists := "FLocFList|FLblFList|FIFList|FMFList|FITFList|FSFList"


Fields=
(
LocField
LblField
ItemField
ModField
QtyField
SnField
MacField
IssField
)

Controls=
(
MakeChanges
ExpMAC
ExpXL
FLoc
FLbl
FI   
FM
FIT
FReset
DA2062
)
;~ FLoc
;~ FLbl
;~ FI   
;~ FM       ;Put these back in %Controls% when they work
;~ FIT
;~ FReset
;~ DA2062

Filters := "FLoc|FLbl|FI|FM|FIT|FS"
SmFilters := "FLoc|FLbl|FI|FM|FS"
Combos := "loc|itm|Mod"
global ColumnList := "Location|Label|Item|Model|Qty|Serial|MAC Address|Issued To|Status"
ScanGuiCols := "Line|Location|Label|Item|Model|QTY|Serial"

GVColsF := "1|2|3|4|5|6|7|8|9"
GVColsOH := "1|2|3|4|5|6|7|8"
GVColsTI := "1|2|3|4|5|6|7"

Firstus=
(
Filter Location|
Filter Label Set|
Filter Item|
Filter Model|
Filter Issued To|
FilterStatus|
)

Top := "Filter Location|,Filter Label Set|,Filter Item|,Filter Model|,Filter Issued To|,FilterStatus|"


;~ sort, LocList, U D|
;~ sort, LblList, U D|
;~ sort, IList, U D|
;~ sort, MList, U D|

;~ If (SubStr(IList,1,1)="|")
  ;~ StringTrimLeft, IList, Ilist, 1
gosub, BuildInventory

;~ TAB 1: Add to Inventory
{
Gui, Tab, Add
;~ Gui, Add, Button, xp+10 yp+10 h20 w100 gAddCol, Add Column
;~ Gui, Add, Button, xp+10 yp+10 h50 w150 gScan , Scan
Gui, Add, Picture, xp+10 yp+10 h50 w150 gScan , %A_ScriptDir%\Images\Scan.png

   
Gui, Add, Button, xp+0 y+5 h40 w150 gNoSNAddGui, Add Item Without Serial Number
Gui, Add, Button, xp+0 y+5 h40 w150 gAddIL, Add Issuable Location

Gui, Font, Bold
Gui, Add, Text, xp+0 y+10 h15 w150, Location:
Gui, Font, Norm
Gui, Add, ComboBox, xp+0 yp+20 R20 w150 vloc glblchangeloc, Select or Type Here|%LocList%

Gui, Font, Bold
Gui, Add, Text, xp+0 yp+30 h15 w150, Item:
Gui, Font, Norm
Gui, Add, ComboBox, xp+0 yp+20 R5 w150 vitm gitm, Select or Type Here|%IList%
GuiControl, Text, itm, Item

Gui, Font, Bold
Gui, Add, Text, xp+0 yp+30 h15 w150, Model:
Gui, Font, Norm
Gui, Add, ComboBox, xp+0 yp+20 R5 w150 vMod, Select Item First
GuiControl, disable, Mod

if (i < 10)
  j := "0" . i
else
  j := i
lbl := lblnm . "–" . j  ;Counter Increases Label Number, Label = LabelName–j
Gui, Font, Bold
Gui, Add, Text, xp+0 yp+30 h15 w45, Label:
Gui, Font, Norm
Gui, Add, Button, x70 yp-2 h20 w100 gcglbl, Change Label
Gui, Add, Edit, x20 yp+25 h20 w150 vlbl, %lbl%
GuiControl, Disable, Lbl

Gui, Add, Button, x20 y+5 h40 w150 gundo, Undo
Gui, Add, Button, x20 y+5 h40 w150 gclear, Clear    ;Grid`n(Keep Selections)
;~ Gui, Add, Button, x20 y+5 h40 w150 gReset, Reset

;~ Gui, Add, Button, x20 y+50 h50 w150 gAddtoInventory, Add To Inventory
;~ Gui, Add, Button, x20 y+10 h60 w150 gAddtoInventory, Add To Inventory
;~ Gui, Add, Picture , x20 y+10 h60 w150 gAddtoInventory, %A_ScriptDir%\Add To Inventory.png
Gui, Add, Picture , x20 y+10 h60 w160 gAddtoInventory, %A_ScriptDir%\Images\Add To Inventory.png

Gui, Add, Listview, grid x180 y11 h515 w610 vScanGUIAdd, Line|Location|Label|Item|Model|QTY|Serial|MAC Address|Issued To|Status
}

;~ TAB 2: Full Inventory
{
Gui, Tab, Full

;~ Gui, Add, Listview, grid x180 y51 h465 w610 AltSubmit checked vInventory gCheckMeOut,Location|Label|Item|Model|QTY|Serial|MAC Address|Issued To
Gui, Add, Listview, grid x180 y71 h465 w610 vInventory -Multi AltSubmit gSubLV HwndHLV1 , %ColumnList%
;~ gosub, BuildInventory

Gui, Add, Text, x182 y20 w115 h20, Filter Location:
Gui, Add, Text, x+5 y20 w115, Filter Label Set:
Gui, Add, Text, x+5 y20 w115, Filter Item:
Gui, Add, Text, x+5 y20 w200, Filter Model:
Gui, Add, Text, x+5 y20 w115, Filter Issued To:
Gui, Add, Text, x+5 y20 w115, Filter Status:

Gui, Add, DropdownList, x180 y40 w115 Choose1 vFLoc gFLoc, %LocList%
Gui, Add, DropdownList, x+5 y40 w115 Choose1 vFLbl gFLbl, %LblList%
Gui, Add, DropdownList, x+5 y40 w115 Choose1 vFI gFI, %IList%
Gui, Add, DropdownList, x+5 y40 w200 Choose1 vFM gFM, %MList%
Gui, Add, DropdownList, x+5 y40 w115 Choose1 vFIT gFIT hwndhwndvar, %ITList%
Gui, Add, DropdownList, x+5 y40 w115 Choose1 vFS gFS, %StatList%
Gui, Add, Button,  x+5 y25 w115 h40 gFReset vFReset, Reset Filters



Gui, Add, Text, x20 y20 w150 h20, Search:
Gui, Add, Edit, x20 y+0 w100 h25 vSearchFor gSearchForChange -Multi,
Gui, Add, Button,  x120 y20 h47 w50 default gSearch, Search (F3)   ;h25


Gui, Add, Button,  x20 y+30 h50 w150 gMakeChanges vMakeChanges, Make Changes
Gui, Add, Button,  x20 y+30 h50 w150 vDA2062 gFull2062, Generate Complete`nDA Form 2062    ;
Gui, Add, Button,  x20 y+30 h50 w150 vExpToXLGUI gExpToXLGUI, Export to Excel
Gui, Add, Button,  x20 y+30 h50 w150 vExpMAC gExpMAC, Export MAC Address List for VoIP Phones to Excel
Gui, Add, Button,  x20 y+30 h50 w150 vApply gApplyChanges, Apply
Gui, Add, Button,  x20 y+5 h50 w150 vCancel gCancelChanges, Cancel
Guicontrol, Hide, Apply
Guicontrol, Hide, Cancel
}

;~ TAB 3: On-Hand Inventory
{
Gui, Tab, On-Hand Inventory,,Exact

Gui, Add, Listview, grid x180 y11 h515 w610 AltSubmit checked vOHInventory gCheckMeOut NoSortHdr HwndHLV2,Item|Location|Label|Model|On-Hand QTY|QTY to be Issued|Serial|MAC Address
;~ Gui, Add, Listview, grid x180 y51 h465 w610 vInventory,Location|Label|Item|Model|QTY|Serial|MAC Address|Issued To
;~ gosub, BuildInventory
LV_ModifyCol(5,"Integer")
LV_ModifyCol(6,"Integer")
Gui, Add, Button,  x20 y20 h50 w150 gIssue vIssue, Issue
Gui, Add, ListBox, x20 y90 w150 ReadOnly R20 vSelectedI, Selected:
Gui, Add, Button,  x20 y+20 h50 w150 vClearChk gClearChk, Clear
Gui, Add, Button,  x20 y+20 h50 w150 vVerify gVerify, Verify Issue
;~ GuiControl, disable, SelectedI

}

;~ TAB 4: Turn-In Inventory
{
Gui, Tab, Turn-In,,Exact

Gui, Add, Listview, grid x180 y11 h515 w610 ReadOnly checked vTIInventory Sort , Serial|Location|Label|Item|Model|QTY|MAC Address
;~ Gui, Add, Listview, grid x180 y51 h465 w610 vInventory,Location|Label|Item|Model|QTY|Serial|MAC Address|Issued To
;~ gosub, BuildInventory
Gui, Add, DropDownList,  x20 y20 R5 w150 gCustomerSelect vCustomerSelect, Choose Customer|%ITList%
Gui, Add, Button,  x20 y+10 h40 w150 vTIScan gTIScan, Scan
Gui, Add, Button,  x20 y+10 h50 w150 vTINonSN gTINonSNGui, Turn-In Non-Serialized Items
Gui, Add, ListBox, x20 y+10 w150 ReadOnly R12 vScanned, Serials Scanned:
Gui, Add, Button,  x20 y+10 h40 w150 vTIClear gTIClear, Clear Entries
Gui, Add, Button,  x20 y+10 h40 w150 vTIApply gTIApply, Turn-In
Gui, Add, Text, x20 y+10 w150 vTIRunning, Serialized Items to Turn-In: %TIItems%
GuiControl, Disable, TIScan
GuiControl, Disable, TIApply
GuiControl, Disable, TINonSN
GuiControl, Disable, TIInventory     ;RE-ENABLE!!!!
GuiControl, Disable, TIClear
GuiControl, Choose, CustomerSelect, 1
}
gosub, BuildInventory

SetComboBoxes:
{
;~ Loop, Parse, Flists, |
  ;~ GuiControl, Choose, %A_Loopfield%, 1

GuiControl,, itm, |Select or Type Here|%CBIList%
GuiControl,, Loc, |Select or Type Here|%LocList%
GuiControl,, Mod, |Select or Type Here|%MList%
loop, Parse, Combos, |
  GuiControl, Choose, %A_Loopfield%, 1
 }


Gui, Show, w1200 h600, Inventory
Gui, ListView, Inventory
qCount := LV_GetCount()
if qCount
  GuiControl, Choose,TabArea, 2
else
  GuiControl, Choose,TabArea, 1
  
SplashTextOff
return

^#a::
;~ FileDelete, 
FileCopy, %A_ScriptDir%\Inventory - Copy.csv, %A_ScriptDir%\Inventory.csv, 1
Reload

^#z::
gosub, BuildInventory
Return


	NoSNAddGui:
	Gui, 5:Default
	Gui, Font, s12
	;~ Gui, Font, Bold
	;~ Gui, Add, Text, x10 y10, Location:
	;~ Gui, Font, Norm
	;~ Gui, Add, ComboBox, xp+0 yp+20 R20 w200 vNOSNloc, Select or Type Here|%LocList%

	Gui, Font, Bold
	Gui, Add, Text, x10 y10, Item: (ex. Power Supply)
	Gui, Font, Norm
	Gui, Add, ComboBox, xp+0 y+5 R5 w200 vNOSNitm gNoSNitm, Select or Type Here|%NoSNIList%

	Gui, Font, Bold
	Gui, Add, Text, xp+0 y+30, Description: (ex. Power Supply, Laptop)
	Gui, Font, Norm
	Gui, Add, ComboBox, xp+0 y+5 R5 w200 vNOSNMod, Select Item First

	Gui, Font, Bold
	Gui, Add, Text, xp+0 y+30, Quantity to Add:
	Gui, Font, Norm
	Gui, Add, Edit, xp+0 y+5 w90  vNOSNQty, 0
	Gui, Add, UpDown, ;x+150 yp+0

	Gui, Add, Button, xp+0 y+30 h50 w100 gNOSNAddtoGrid, Add
	Gui, Add, Button, x290 yp+0 h50 w100 gNOSNClear, Clear


	GuiControl, Disable, NOSNMod
	GuiControl, Choose, NOSNloc, 1
	GuiControl, Choose, NOSNitm, 1
	GuiControl, Choose, NOSNMod, 1
	Gui, Show, w400, Add Non-Serialized Items
  Gui, 5: +owner1
  Gui, 1: +Disabled
	Return

  5GuiClose:
  Gui, 5:Destroy
  Gui, 1: -Disabled
  Gui, 1: Default
  Gui, Show
	Return
  
	NoSNitm:
  Gui, 5:Default
  GuiControlGet, NOSNitm, ,NOSNitm
  Gui, 1: Default
	Gui, ListView, Inventory
		/*
		Find: What do you want to be in "Criteria" Column?
		Criteria: What Column do you want to find "Find"?
		Col: What Column will do you want to build list of if "Find" is found?
		LV: Which ListView? (Name)
		ColumnList: ALWAYS ColumnList
		*/
  
	NoSNModList := CompileList(NOSNitm,  "Item", "Model", "Inventory", ColumnList)
  Gui, 5: Default
	GuiControl,, NOSNMod,|Select or Type Here|%NoSNModList%
	GuiControl, enable, NOSNMod
  GuiControl, Choose, NOSNMod, 1
	Return

	NOSNAddtoGrid:
  Gui, 5:Submit
  Gui, 1: -Disabled
  Gui, 1:Default
  
	Line += 1
	Gui, ListView, ScanGuiAdd
	LV_Add("",Line, NOSNItm, , NOSNItm, NOSNMod, NOSNQty)
	LV_Modify(Line,"Col10","On-Hand")
  loop %ct%
  {
    if (A_Index <= 7)
      LV_ModifyCol(A_Index,"AutoHdr")
    if (A_Index = 8 and voipyn = "Y")
      LV_ModifyCol(A_Index,"AutoHdr")
    if (A_Index = 8 and voipyn = "N")
      LV_ModifyCol(A_Index,0)
    if (A_Index > 8)
      LV_ModifyCol(A_Index,0)
  }
  Gui, 5:Destroy
	Return

	NOSNClear:
	Gui, 5:Destroy
	gosub, NoSNAddGui
	Return	






ExpToXLGUI:
Gui, 4: Destroy
Gui, 4: Default
Gui, Add, Text, ,Please Select Columns to Include:

Gui, 1: Default
Gui, ListView, Inventory
Loop, % LV_GetCount("Column")   ;Loop Number of Columns
{
  LV_GetText(Text, "0", A_Index)
  StringReplace, Text, Text, %A_Space%, , ALL 
  hdr .= Text . "`|"
}
StringTrimRight, hdr, hdr, 1
Gui, 4: Default
Loop, Parse, hdr, |
  Gui, Add, Checkbox, Checked v%A_Loopfield%, %A_Loopfield%
Gui, Add, Button, gExpToXL, Export
Gui, +AlwaysOnTop
Gui, Show,,%A_Space%
GuiControlGet, MAVar, Pos, MACAddress
GuiControlGet, ITVar, Pos, IssuedTo
GuiControlGet, MAVar, Pos, MACAddress
GuiControlGet, ITVar, Pos, IssuedTo
;~ var := "VoIP"
;~ If var not in %IListC%
if IListC not contains VoIP
{
  GuiControl, Hide, MACAddress
  GuiControl, , MACAddress, 0
  GuiControl, MoveDraw, IssuedTo, x%MAVarX% y%MAVarY%
  GuiControl, MoveDraw, Status, x%ITVarX% y%ITVarY%
}
LV := "Inventory"
Return

ExpToXL:
Gui, 4:Submit
Gui, ListView, %LV%
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
TtlC := 0
SelCols=
Loop, Parse, hdr, |
{
  if (%A_Loopfield% = 0)
    Continue
  SelCols .= A_Index . "|"
  TtlC += 1
}
StringTrimRight, SelCols, SelCols, 1
Gui, 4: Destroy
Issue := "NO"

Export:
Export :=  GrabView(SelCols,LV,Issue)
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Add                            ;add a new workbook
Columns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I")
EoL := Columns[TtlC]
Xl.Range("A:" . EoL).NumberFormat := "@"
Loop, Parse, Export, `n, `r
{
  RowNumber := A_Index    ; + 1
  if(A_Loopfield = "")
    Break
  
  Loop, Parse, A_Loopfield, csv
  {
    if (A_Index = 10)
      Break
    StringReplace, Field, A_Loopfield, `~, `,, All
    Xl.Range(Columns[A_Index] . RowNumber).Value := Field
  }
}

Xl.ActiveSheet.ListObjects.Add(1, Xl.Range("A1").CurrentRegion,, 1).Name := "Table1"
Xl.ActiveSheet.ListObjects("Table1").TableStyle := "TableStyleMedium2"
Xl.Columns("A:" . EoL).EntireColumn.AutoFit
;~ Xl.Worksheets("Sheet2").Delete()		; delete Sheet2
;~ Xl.Worksheets("Sheet3").Delete()		; delete Sheet3
Xl.Visible := True                          ;by default excel sheets are invisible
SelCols=
hdr=
SplashTextOff                               
Return





CustomerSelect:
GuiControlGet, CustomerSelect
Gui, ListView, TIInventory
LV_Delete()
GuiControl, Text, SelectedI, |
if (CustomerSelect = "Choose Customer")
{
  GuiControl, Disable, TIScan
  GuiControl, Disable, TIApply
  GuiControl, Disable, TINonSN
  GuiControl, Disable, TIClear
  Return
}
gosub, BuildInventory
LV_ModifyCol(3,"Sort")

if (CustomerSelect = "a")
  GuiControl, Enable, TIInventory
GuiControl, Enable, TIScan
GuiControl, Enable, TIApply
TINOSNList := ""
FileRead, Inv, %csvF%
Loop, Parse, Inv, `n, `r
{
  StringSplit, List, A_Loopfield, `,
  if(List8 = CustomerSelect and List6 = "")
    TINOSNList .= "|" . List4
  loop, 9
  {
    List := "List" . A_Index
    %List% := ""
  }
}
StringTrimLeft, TINOSNList, TINOSNList, 2
StringReplace, TINOSNList, TINOSNList, `~, `,, All
If(TINOSNList <> "")
  GuiControl, Enable, TINonSN
GuiControl, Enable, TIClear
Return


SearchForChange:
SRow := 0
Return

F3::
Search:
Gui, ListView, Inventory
GuiControlGet, SearchFor, , SearchFor
if(SearchFor = "")
  Return
;~ MsgBox, %RowN% %SearchFor%
RowN := Search(RowN, SearchFor, "All", "Inventory")
if (RowN = 0)
{
  MsgBox, 4116, Invalid Entry, No matches found! Start from Beginning?
  IfMsgBox, Yes
  {
    RowN := 0
    gosub, Search
  }
  else
    return
}
GuiControl, Focus, Inventory
LV_Modify(RowN, "vis select focus")
Return

GuiControl, Focus, Inventory
LV_Modify(SRow, "Vis Select Focus")
;~ GuiControl, Text, SearchFor,
Return




GuiClose:
ExitApp

If (WhatHappened = "RightClick") {
      CopiedCell=
      Row=
      Column=
      Row := A_EventInfo
      Column := LV_SubItemHitTest(HLV1)
      MouseGetPos, VarX, VarY
      Menu, NormalContext, Show, %VarX%, %VarY%
   }


ChgQty:
;~ LV_GetText(
Return


CheckMeOut:
Gui, ListView, OHInventory
Critical
RowNumber := A_EventInfo
WhatHappened := A_GuiEvent
WhatChanged := ErrorLevel

GuiControl, -AltSubmit, OHInventory
;~ If (WhatHappened = "RightClick") 
;~ {
    ;~ CopiedCell=
    ;~ Row=
    ;~ Column=
    ;~ Row := RowNumber
    ;~ Column := LV_SubItemHitTest(HLV2)
    ;~ MouseGetPos, VarX, VarY
    ;~ Menu, OHContext, Show, %VarX%, %VarY%
 ;~ }
if (WhatHappened = "I")
{
  LV_GetText(Model, RowNumber, 4)
  LV_GetText(OHQty, RowNumber, 5)
  LV_GetText(IQty, RowNumber, 6)
  if (OHQty = "")
    OHQty = 0
  if (IQty = "")
    IQty = 0
  if (InStr(WhatChanged, "c", true))
  {

    qty := OHQty + IQty
    LV_Modify(RowNumber, "Col5", QTY)
    LV_Modify(RowNumber, "Col6", "")
    gosub, CountChecked
  }
  if (InStr(WhatChanged, "C", true))
  {
    ;~ Gui +LastFound
    ;~ Gui, Listview, OHInventory
    ;~ SendMessage, 4140, RowNumber - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE.  0xF000 is LVIS_STATEIMAGEMASK.
    ;~ IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.

    if (OHQty + IQty > 1)
    {
      Loop 1
      {
        Gui +owndialogs
        if (IQty > 0)
          qty := IQty
        else
          qty := OHQty
        InputBox, QTY, Quantity, How many %Model%s would you like to Issue from this location?, , , , , , , , %Qty%
        if (QTY > OHQty)
        {
          MsgBox, 4112, Invalid Entry, You cannot issue more than you have!
          Continue
        }
      }
      if (QTY <= 0 or if errorlevel)
      {
        LV_Modify(RowNumber, "-Check")
        gosub, CountChecked
        Return
      }
    }
    else
      Qty := OHQty
    
    LV_Modify(RowNumber, "Col6", QTY)
    OHQty -= Qty
    if (OHQty = 0)
      OHQty := ""
    LV_Modify(RowNumber, "Col5", OHQTY)    
  }
}

CountChecked:
CheckedItems := ""
Selected := "|Selected:|"
CntChkd := LV_GetNext(0, "Checked")
if CntChkd > 0
{
  GuiControl, enable, Issue
  CheckedItems=
  SortedItems=
  RowNumber = 0  ; This causes the first loop iteration to start the search at the top of the list.
  Loop
  {
    RowNumber := LV_GetNext(RowNumber, "C")  ; Resume the search at the row after that found by the previous iteration.
    if not RowNumber  ; The above returned zero, so there are no more selected rows.
        break
    LV_GetText(Text, RowNumber)
    LV_GetText(TextQ, RowNumber, 6)
    if (TextQ = "")
      TextQ = 0

    CheckedItems .= Text . "|" . TextQ . "`n"
    SortedItems .= Text . "`n"
  }
  Sort, CheckedItems
  Sort, SortedItems, U
  loop, parse, CheckedItems, `n, `r
  {
    if (A_Loopfield = "")
      Break
    loop, parse, A_Loopfield, | ,`n
    {
      if (A_Index = 1)
      {
        txt := RegExReplace(A_Loopfield, "\W", "_")
        CItem := txt
      }
      else If (A_Index = 2)
        %CItem% += abs(A_Loopfield)
    }
  }
  
  loop, parse, SortedItems, `n, `r
  {
    if (A_Loopfield = "")
      Break
    ;~ regexReplace(CheckedItems, A_Loopfield, A_Loopfield, count )
    txt := RegExReplace(A_Loopfield, "\W", "_")
    Selected .= A_Loopfield . ": " . %txt% . "|"
  }
  GuiControl, , SelectedI, %Selected%
}
else
{
  GuiControl, disable, Issue
  GuiControl, , SelectedI,|Selected:|
}  
loop, parse, CheckedItems, `n, `r
{
  if (A_Loopfield = "")
    Break
  loop, parse, A_Loopfield, | ,`n
  {
    if (A_Index = 1)
    {
      txt := RegExReplace(A_Loopfield, "\W", "_")
      %txt% := 0
    }
  }
}

GuiControl, Focus, SelectedI
GuiControl, +AltSubmit, OHInventory
Return

ClearChk:
Gui, ListView, OHInventory
RowNumber := 0
Loop
{
  RowNumber := LV_GetNext(RowNumber,"Checked")
  if not RowNumber  ; The above returned zero, so there are no more selected rows.
        break
  LV_Modify(RowNumber, "-Check")
  
}
gosub, CheckMeOut
Return

Issue:
NOSN := "NO"
iniread, FromName, %ini%, HR, FromName, Enter "From" Field
InputBox, FromName, Please Enter From Field, Who is Issuing This?,,,,,,,,%FromName%
IniWrite, %FromName%, %ini%, HR, FromName
IssueTo := CustGui("Issuing")
if (IssueTo = -1)
  Return
;~ InputBox, IssueTo, Please Enter Customer, Who is this being issued to?
  SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
Gui, 1:Default

gosub, BuildLVInventory
;~ gosub, BuildInventory
RowNumber = 0  ; This causes the first loop iteration to start the search at the top of the list.
Loop
{
  Gui, ListView, OHInventory
  RowNumber := LV_GetNext(RowNumber, "Checked")  ; Resume the search at the row after that found by the previous iteration.
  if not RowNumber  ; The above returned zero, so there are no more selected rows.
      break
  LV_GetText(Serial,RowNumber,7)
  if(Serial = "")
  {
    gosub, IssueNOSN
    Continue
  }
  GuiControl, Choose, TabArea, 2
  Gui, ListView, Inventory
  RowN := Search("0",Serial,"6","Inventory")
  LV_Modify(RowN,"Col8",IssueTo)
  LV_Modify(RowN,"Col9","Issued") 
}
CoordMode, Mouse
MouseMove, 0,0,0

;~ Compile master list of checked: Generates as csv with no trailing , or `n; NO HEADERS
GuiControl, Choose, TabArea, 3
Gui, ListView, OHInventory
MasterIssueList := ""
RowNumber := 0
Loop
{
  RowNumber := LV_GetNext(RowNumber, "Checked")
  ;~ MsgBox, %RowNumber%
  if not RowNumber
    break
  TtlC := LV_GetCount("Column")
  Loop %TtlC%
  {
    LV_GetText(Col, RowNumber, A_Index)
    StringReplace, Col, Col, `,, `~, ALL
    a .= Col
    if (A_Index <> TtlC)
      a .= "`,"
  }
  if (A_Index = 1)
    MasterIssueList .= a
  else
    MasterIssueList .= "`n" . a
  a := ""
}
S := "NO",Q := "NO",MA := "NO",L := "NO"
loop, parse, MasterIssueList, `n, `r
{
  if (S = "YES" and Q = "YES" and MA = "YES" and L = "YES")
    Break
  StringSplit, Col, A_Loopfield,`,
  if (Col2 <> "")
    L := "YES"
  if (Col7 <> "")
    S := "YES"
  if (Col7 = "")
    Q := "YES"
  if (Col8 <> "")
    MA := "YES"
}
SelCols := "2|1|4"
TtlC := 3

if (L = "YES")
{
  SelCols .= "|3"
  TtlC += 1
}
if (q = "YES")
{
  SelCols .= "|6"
  TtlC += 1
}
if (s = "YES")
{
  SelCols .= "|7"
  TtlC += 1
}
if (MA = "YES")
{
  SelCols .= "|8"
  TtlC += 1
}



MsgBox, %MasterIssueList%
LV := "OHInventory"
;~ TtlC := 5
;~ SelCols := "1|4|3|7|8"
Issue := "YES"
gosub, Export
;~ gosub, GenerateXML         ;RE-ENABLE FOR XML
;~ IssuePDF(CustGrid, From, To)
IssuePDF(MasterIssueList, FromName, IssueTo)
gosub, ApplyChanges
gosub, BuildInventory
GuiControl,,SelectedI, |Selected:|
GuiControl,,CustomerSelect,|Choose Customer|%ITList%
GuiControl, Choose, CustomerSelect, 1
gosub, CustomerSelect
SplashTextOff
MsgBox, 131072, Done, You can now resume normal operation.
Return










IssueNOSN:
;~ MsgBox, SHOULD NOT BE HERE`n`n`n`n`n`n`n`n`n`n
Gui, 1:Default
NOSN := "YES"
LV_GetText(Item,RowNumber,1)
LV_GetText(Desc,RowNumber,4)
LV_GetText(OHQty,RowNumber,5)
LV_GetText(IQty,RowNumber,6)
Gui,ListView,Inventory
RowN := DblSearch(Desc, "Model", "On-Hand", "Status", "Inventory", ColumnList)
Gui,ListView,Inventory
if(OHQty = "" or OHQty = 0)
  LV_Delete(RowN)
else
  LV_Modify(RowN,"Col5",OHQty)
LV_Add("",Item,,Item,Desc,IQty,,,IssueTo,"Issued")


Return





itm:
GuiControlGet, itm, ,itm
Gui, ListView, OHInventory
if itm Contains VoIP
  gosub, VoIP
else if (itm = "Select or Type Here")
{
  GuiControl, , Mod, |Select Item First
  GuiControl, Choose, Mod, 1
  GuiControl, disable, Mod
}
if itm in %IListC%
{
  ModList := CompileList(itm, "Item", "Model", "Inventory", ColumnList)
  GuiControl, , Mod, |Select or Type Here|%ModList%
  GuiControl, Choose, Mod, 1
  GuiControl, enable, Mod
}
if itm not in %IListC%
{
  ;~ ModList := CompileList(itm, "Item", "Model", "Inventory", ColumnList)
  GuiControl, , Mod, |Select or Type Here       ;|%ModList%
  GuiControl, Choose, Mod, 1
  GuiControl, enable, Mod
}
Return

GuiSize:
{
th := A_GuiHeight - 20
tw := A_GuiWidth - 20
lvh := th - 25
ilvh := lvh - 60
lvw := tw - 170
ilvw := lvw + 150
GuiControl, MoveDraw, TabArea, h%th% w%tw%
GuiControl, MoveDraw, ScanGUIAdd, h%lvh% w%lvw%
GuiControl, MoveDraw, Inventory, h%ilvh% w%lvw%
GuiControl, MoveDraw, OHInventory, h%lvh% w%lvw%
GuiControl, MoveDraw, TIInventory, h%lvh% w%lvw%
Return
}

lblchangeloc:
GuiControlGet, loc, ,loc
lblnm := loc
if(loc = "Select or Type Here")
  lblnm := "Label"
gosub, lblchange
Return

lblchange:
ifinstring,FullLblList,%lblnm%–01
{
  Highest := 0
  Loop, parse, FullLblList, csv
  {
    IfNotInString, A_Loopfield, %lblnm%–
      continue
    StringLen, Len, lblnm
    Len += 2
    StringMid, Num, A_Loopfield, %Len%, 255
    Num := RegExReplace(Num, "\D")
    If(Num > Highest)
      Highest := Num
    Highest += 1
  }
}

Gui, Submit, NoHide
if (Highest > 0)
{
  i := Highest
  Highest := 0
}
Else
  i := 1
if (i < 10)
  j := "0" . i
else
  j := i
lbl := lblnm . "–" . j
GuiControl, Text, lbl, %lbl%
Return


cglbl:
gui, +owndialogs
InputBox, lblnm, Change Label, Change Label?, , , , , , , , %lblnm%
if (ErrorLevel <> 0)
  return
;~ StringReplace, lblnm, lblnm, `-,,All
;~ StringReplace, lblnm, lblnm, `,,,All
;~ StringReplace, lblnm, lblnm, %A_Space%,,All
;~ i := 1
;~ if (i < 10)
  ;~ j := "0" . i
;~ else
  ;~ j := i
;~ lbl := lblnm . "–" . j
;~ GuiControl, Text, lbl, %lbl%
gosub, lblchange
Return


Scan:
Gui, Submit, NoHide
Gui, ListView, ScanGuiAdd
GuiControlGet, Mod,,Mod
GuiCols := LV_GetCount("Column")
AddColList=
loop, %GuiCols%
{
  Gui, ListView, ScanGuiAdd
  LV_GetText(Text,"0",A_Index)
  AddColList .= Text
  if (A_Index <> GuiCols)
    AddColList .= "|"
}
SSList := SList . "|"
SSList .= CompileList("All", "Serial", "Serial", "ScanGuiAdd", AddColList)
OBreak := "NO"
BacktoOuter := "NO"

Outer:
Loop
{
  BacktoOuter := "NO"
  gosub, ColumnLoop
  if (OBreak = "YES")
    Break
  if (BacktoOuter = "YES")
    Continue
  Line += 1
  sofar=
  Col9=
  SSList .= "|" . Col7 . "|"
  LV_Add("",Line, Loc, lbl, Itm, Mod, "1", Col7,Col8)
  LV_Modify(Line,"Col10","On-Hand")
  CBIList .= "|" . itm
  CBIList .= "|" . itm
  sort, CBIList, U D|
  GuiControl, ,itm, %IList%
  GuiControl, ,FI, %IList%
  GuiControl, Text, itm, %itm%
  ;~ ct := LV_GetCount("Column")
  loop %ct%
  {
    if (A_Index <= 7)
      LV_ModifyCol(A_Index,"AutoHdr")
    if (A_Index = 8 and voipyn = "Y")
      LV_ModifyCol(A_Index,"AutoHdr")
    if (A_Index = 8 and voipyn = "N")
      LV_ModifyCol(A_Index,0)
    if (A_Index > 8)
      LV_ModifyCol(A_Index,0)
  }
  i += 1
  if (i < 10)
    j := "0" . i
  else
    j := i
  lbl := lblnm . "-" . j
  GuiControl, Text, lbl, %lbl%
  SoundPlay, %A_ScriptDir%\Sounds\Success.wav
}
Return

ColumnLoop:

Col7 := ""
Col8 := ""
Gui,+LastFound
WinGetPos,gx,gy
Gui +owndialogs

ScanMe:
Loop
{
  GuiControlGet, loc
  GuiControlGet, Mod
  GuiControlGet, Itm
  
  Prompt = Location: %loc%`nItem: %itm%`nModel: %Mod%`nLabel: %lbl%
  StringReplace, Prompt, Prompt, Select or Type Here, %A_Space%, All
  Inputbox, Col7, Scan Serial, %Prompt%`nScan Serial Number,, 200, 200, %gx%, %gy%
  If (Col7 = "")
  {
    OBreak := "YES"
    Return
  }
  Srch := "`|" . Col7 . "`|"
  IfNotInString, SSList, %Srch%
    break, ScanMe
  Else
  {
    SoundPlay, %A_ScriptDir%\Sounds\Error.wav
    MsgBox, 4118, Serial Already Exists, This serial number already exists in your inventory!
    IfMsgBox TryAgain
      Continue, ScanMe
    IfMsgBox Cancel
    {
      OBreak := "YES"
      Return
    }
    IfMsgBox Continue
      Break, ScanMe
  }
}
Check := Col7
gosub, CheckScan
;~ MsgBox, %OBreak% %BacktoOuter%
if (OBreak = "YES" or BacktoOuter = "YES")
  Return
;~ MsgBox, %OBreak% %BacktoOuter%
IfInString, itm, VoIP
  VoIPNow := "YES"
else
  VoIPNow := "NO"
  
if (VoIPNow = "YES")
{
  Inputbox, Col8, Scan MAC Address, %Prompt%`nSerial: %Col7%`nScan MAC Address,, 200, 220, %gx%, %gy%
  Check := Col8
  gosub, CheckScan
  ;~ MsgBox, %OBreak% %BacktoOuter%
  ;~ if (OBreak = "YES" or BacktoOuter = "YES")
    ;~ Return
}
Return


CheckScan:
ifinstring, RegisteredItems, %Check%=
{
  iniread, InfoLookup, %ini%, RegisteredItems, %Check%
  if (infolookup <> "" and infolookup <> "ERROR")
  {
    StringSplit, Info, infolookup, \
    itm := Info1
    Mod := Info2
    GuiControl, Text, itm, %Itm%
    GuiControl, Text, Mod, %Mod%
    if itm Contains VoIP
      Gosub, VoIP
    SoundPlay, %A_ScriptDir%\Sounds\Change.wav
    BacktoOuter := "YES"
    Return
  }
}
if SubStr(Check, 1, 2) = "I-"
{
  StringGetPos, NxD, Check, `-, L2
  StringLen, Len, Check
  1D := NxD - 2
  2D := NxD + 2
  Itm := SubStr(Check, 3, 1D)
  Mod := SubStr(Check, 2D)
  GuiControl, Text, itm, %Itm%
  GuiControl, Text, Mod, %Mod%
  if itm Contains VoIP
    Gosub, VoIP
  SoundPlay, %A_ScriptDir%\Sounds\Change.wav
  BacktoOuter := "YES"
  Return
}
if SubStr(Check, 1, 2) = "L-"
{
  StringTrimLeft, Check, Check, 2
  GuiControl, Text, Loc, %Check%
  gosub, lblchangeloc
  SoundPlay, %A_ScriptDir%\Sounds\Change.wav
  BacktoOuter := "YES"
  Return
}

LookUp := Check
if SubStr(Check, 1, 3) = "IL-"
{
  Loop, parse, Check,`-
  {
    if A_Index = 1
      Continue
    if A_Index = 2
      Type := A_Loopfield
    if A_Index = 3
      Loc := A_Loopfield
  }
  
  IfNotInString, SSList, %Loc%
  {
    Line += 1
    LV_Add("", Line, Loc, Loc, "Issuable Location", type, "1", Loc)
    LV_Modify(Line, "Col10", "On-Hand")
    loop %ct%
      LV_ModifyCol(A_Index,"AutoHdr")
  }
  GuiControl, Text, loc, %Loc%
  gosub, lblchangeloc
  SoundPlay, %A_ScriptDir%\Sounds\Change.wav ;, Wait
  ;~ SoundPlay, %A_ScriptDir%\Sounds\Success.wav  ;, Wait
  BacktoOuter := "YES"
  Return
}
if(Loc = "Select or Type Here" or Mod = "Select or Type Here" or Itm = "Select or Type Here")
{
  Error := ""
  if(Loc = "Select or Type Here")
    Error .= A_Tab . "Location" . "`n"
  if(Itm = "Select or Type Here")
    Error .= A_Tab . "Item" . "`n"
  if(Mod = "Select or Type Here")
    Error .= A_Tab . "Model"
  SoundPlay, %A_ScriptDir%\Sounds\Error.wav
  MsgBox, 4112, Problem, Please scan or select the following:`n%Error%
  OBreak := "YES"
  Return
}
;~ if (A_Index > 1)
  ;~ sofar .= "`n"
;~ sofar .= Check
Return 
      
      



Undo:
{
TTL := LV_GetCount()
LV_Delete(TTL)
i -= 1
if (i < 10)
  j := "0" . i
else
  j := i
Line -= 1
lbl := lblnm . "–" . j
GuiControl, Text, lbl, %lbl%
Return
}

VoIP:
{
Gui, ListView, ScanGUIAdd
if (voipyn = "Y")
  return
else
{
  CE += 1
  CT += 1
  LV_ModifyCol(8,"AutoHdr")
  voipyn := "Y"
}
Return 
}

GuiContextMenu:
if (editing = "N")
{
  if(A_GuiControl <> "ScanGUIAdd" and A_GuiControl <> "Inventory")
    Return
}
else if (editing = "Y")
{
  if(A_GuiControl <> "Inventory")
    Return
}
if(A_EventInfo = 0)
  Return
Row := A_EventInfo

if (editing = "Y")
  Menu, CSVContext, Show, %A_GuiX%, %A_GuiY%
else if (editing = "N" and A_GuiControl = "ScanGUIAdd")
  Menu, GridContext, Show, %A_GuiX%, %A_GuiY%
;~ else if (editing = "N" and A_GuiControl = "Inventory")
  ;~ Menu, NormalContext, Show, %A_GuiX%, %A_GuiY%
Return



AddtoInventory:
{
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
Gui, Listview, ScanGuiAdd
Rows := LV_GetCount()
if not Rows
  Return
gosub, BuildInventory
Gui, ListView, ScanGuiAdd
TTL := LV_GetCount()
;~ CT := LV_GetCount("Column")
OuterAdd:
loop %TTL%
{
  RowNum := A_Index
  LV_GetText(Serial, RowNum,7)
  if (Serial = "")
  {
    gosub, addNoSN
    continue, OuterAdd
  } 
  Loop 10
  {
    if (A_Index = 1)
      Continue
    LV_GetText(Text, RowNum, A_Index)
    ;~ if(A_Index = 2 and Text = "")

    StringReplace, Text, Text, `, , `~, All
    adding .= Text . ","
  }
  adding .= "`n"
}

InvTtl := LV_GetCount()
GuiControl, Choose, TabArea, 2
Gui, ListView, Inventory
Loop, Parse, adding, `n, `r
{
  if (A_Loopfield = "")
    Break
  Loop, parse, A_Loopfield, csv
  {
    StringReplace, Text, A_Loopfield, `~, `, , All
    if (A_Index = 1)
      LV_Add("",Text)
    else
    {
      RowNumber := LV_GetCount()
      Col := "Col" . A_Index
      LV_Modify(RowNumber, Col, Text)
    }
    LV_Modify(RowNumber, "vis")
  }
}



adding=
gosub, ApplyChanges
gosub, ClearMe
SplashTextOff
Return
}

AddNoSN:
Gui, ListView, ScanGuiAdd
LV_GetText(Item, RowNum, 4)
LV_GetText(Desc, RowNum, 5)
LV_GetText(Qty, RowNum, 6)
RowNumber := Search(0,Desc,"4","Inventory")
Gui, ListView, Inventory
if (RowNumber = 0)
{
  LV_Add("", Item,"",Item,Desc,Qty)
  Ttl := LV_GetCount()
  LV_Modify(Ttl,"Col9","On-Hand")
}
else
{
  LV_GetText(OldQty, RowNumber,5)
  NewQty:= OldQty + Qty
  LV_Modify(RowNumber, "Col5", NewQty)
}
Return

;~ Return


Clear:
{
Gui, ListView, ScanGuiAdd
MsgBox, 4131, Keep Selections?, Keep your current selections? (Location`, Item`, Model`, Label)
IfMsgBox, No
{
  Gosub, Reset
  Return
}
ifMsgBox, Yes
{
  Gosub, ClearMe
  Return
}
ifMsgBox, Cancel
  Return
}

ClearMe:
Gui, ListView, ScanGuiAdd
LV_Delete()
Line := 0
gosub, lblchangeloc
Return

Reset:
Gui, ListView, ScanGuiAdd
LV_Delete()
Line := 0
GuiControl,, loc, |Select or Type Here|%LocList%
GuiControl,, itm, |Select or Type Here|%IList%
GuiControl,, Mod, |Select Item First
GuiControl, Disable, Mod
Loop, Parse, Combos, |
  GuiControl, Choose, %A_Loopfield%, 1

gosub, lblchangeloc
return  


BuildInventory:
{
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
GuiControlGet, CustomerSelect,,CustomerSelect
LocList=
LblList=
IList=
MList=
ITList=
Inv=
CustSNList=

Gui, ListView, Inventory
LV_Delete()
Gui, ListView, OHInventory
LV_Delete()
Gui, ListView, TIInventory
LV_Delete()
FileRead, Inv, %csvF%
Sort, Inv
StringTrimRight, Inv, Inv, 2
Loop, Parse, Inv, `n, `r
{
  RowNum := A_Index
  Loop, Parse, A_Loopfield, csv
  {
    Column := "Col" . A_Index
    %Column% := A_Loopfield
    StringReplace, %Column%, %Column%, `~, `,, All
    if (A_Index = 1)
    {
      LocList .= "`|" . A_Loopfield
      FLocList .= "`|" . A_Loopfield
    }
    if (A_Index = 2)
    {
      FullLblList .= "`," . A_Loopfield
      IfInString, A_Loopfield, `-
      {
        StringGetPos, DL, A_Loopfield, `-, R
        StringMid, label, A_Loopfield, 1, %DL%
      }
      else
        label := A_Loopfield
      LblList .= "`|" . label
      FLblList .= "`|" . label
    }
    if (A_Index = 3)
    {
      if (A_Loopfield <> "Issuable Location")
        CBIList .= "`|" . A_Loopfield
      IList .= "`|" . A_Loopfield
      FIList .= "`|" . A_Loopfield
    }
    if (A_Index = 4)
    {
      MList .= "`|" . A_Loopfield
      FMList .= "`|" . A_Loopfield
    }
    if (A_Index = 6)
    {
      SList .= "|" . A_Loopfield . "|"
      if (A_Loopfield = "")
      {
        NoSNIList .= "|" . Col3
        FNoSNIList .= "|" . Col3
      }
    }
    if (A_Index = 8)
    {
      ITList .= "|" . A_Loopfield
      FITList .= "|" . A_Loopfield
    }
  }
  Gui, ListView, Inventory
  LV_Add("",Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9)
  if (Col8 = "" or Col9 = "Turned In")
  {
    Gui, ListView, OHInventory
    LV_Add("",Col3,Col1,Col2,Col4,Col5,"",Col6,Col7)
  }
  if (Col8 = CustomerSelect and CustomerSelect <> "Choose Customer" and CustomerSelect <> "")
  {
    Gui, ListView, TIInventory
    LV_Add("",Col6,Col1,Col2,Col3,Col4,Col5,Col7)
    CustSNList .= "|" . Col6 . "|"
  }
    
}
Loop 8
{
  Gui, ListView, Inventory
  LV_ModifyCol(A_Index,"AutoHdr")
  Gui, ListView, OHInventory
  LV_ModifyCol(A_Index,"AutoHdr")
  Gui, ListView, TIInventory
  LV_ModifyCol(A_Index,"AutoHdr")
}
gosub, FReset

Gui, ListView, OHInventory
LV_ModifyCol(3,"SortLogical")
LV_ModifyCol(1,"SortLogical")
;~ StringReplace, IListC, IList, `,, ~, ALL
StringReplace, IListC, IListC, |, `,, ALL
IfNotInString, IListC, VoIP
{
  Gui, ListView, Inventory
  LV_ModifyCol(7,0)
  Gui, ListView, OHInventory
  LV_ModifyCol(8,0)
  Gui, ListView, TIInventory
  LV_ModifyCol(7,0)
  GuiControl, Hide, ExpMAC
}
else
  GuiControl, Show, ExpMAC
Gui, ListView, ScanGUIAdd
if(voipyn = "Y")
  LV_ModifyCol(8,"AutoHdr")
else
  LV_ModifyCol(8,0)
LV_ModifyCol(9,0)
LV_ModifyCol(10,0)

Gui, ListView, OHInventory
LV_ModifyCol(1,"Sort")

;~ StringSplit, Top, Top, `,
;~ StringSplit, FLists, FLists, `|
;~ Loop, Parse, Filters, |
;~ {
  ;~ tp := "Top" . A_Index
  ;~ First := %tp%
  
  ;~ tp := "FLists" . A_Index
  ;~ FList := %tp%
  
  ;~ GuiControl, , %A_Loopfield%, % "|" . First . "|" . %FList%
  ;~ GuiControl, Choose, %A_Loopfield%, 1
;~ }
;~ GoSub, BuildSetFilters
SplashTextOff
Inv := ""
Return
}


Full2062:
Customer := CustGui("Generating")
Gui, 1: Default
if (Customer = -1)
  Return
gosub, FReset
iniread, FromName, %ini%, HR, FromName, Enter "From" Field
if (Customer = "MULTIPLE")
{
  InputBox, FromName, Please Enter From Field, Who is Issuing This?,,,,,,,,%FromName%
  InputBox, ToName, Please Enter To Field, What name should go in the "To" field?
  GuiControl, ChooseString, FIT, |Multiple...
}
else
{
  InputBox, FromName, Please Enter From Field, Who is Issuing This?,,,,,,,,%FromName%
  ToName := Customer
  GuiControl, ChooseString, FIT, |%Customer%
}
Sleep, 250
CustGrid := GrabView("3|1|2|4|5|5|6","Inventory","NO")

StringGetPos, Pos, CustGrid, `n
pos += 2
StringMid, CustGrid, CustGrid, %pos%

IssuePDF(CustGrid, FromName, ToName)
SplashTextOff
Return




FReset:  
  Gui, 1:Default
  loop, parse, filters, |                         
   Guicontrol, , %A_Loopfield%, |
  

Filtration:
FLoc:
FLbl:
FI:
FM:
FIT:
FS:
  MultiCustList := ""
  MAC := "NO"
  Gui, 1:Default
  gui, listview, Inventory
  Gui, Submit, NoHide   ;grab the current state of the drop-downs
  if (FIT = "Multiple...")
    gosub, MultiCustGui ;Pop up gui w/ checkboxes for all customers
  Gui, 1:Default
  gui, listview, Inventory
  ;~ MsgBox, %MultiCustList%
  if (MultiCustList = "-1")
  {
    PostMessage, 0x014E, -1, 0,,ahk_id %hwndvar%
    ;~ PostMessage, 0x185, 0, -1, FIT
    ;~ GuiControl, ChooseString, FIT, |  
    Return
  }
  LV_Delete()
  fileread, FInv, inventory.csv
  StringReplace, FInv, FInv, `,, |, All
  StringReplace, FInv, FInv, `~, `,, UseErrorLevel
  FInv := StrReplace(FInv, "~", "`,")
  FOuter:
  Loop, parse, FInv, `r, `n
  {
    if not A_Loopfield
      Break
    Stringsplit, Field, A_Loopfield, |
    
    if (FLoc <> "" and FLoc <> Field1)
      Continue, FOuter
    FLPos := RegExMatch(Field2, "`–")
    FLPos -= 1
    StringLeft, LblCheck, Field2, %FLPos%
    if (FLbl <> "" and FLbl <> LblCheck)
      Continue, FOuter
    if (FI <> "" and FI <> Field3)
      Continue, FOuter
    if (FM <> "" and FM <> Field4)
      Continue, FOuter
    if (FIT = "Multiple...")
    {
      if Field8 not in %MultiCustList%
        Continue, FOuter
    }
    else if (FIT <> "" and FIT <> Field8)
      Continue, FOuter
    
    if (FS <> "" and FS <> Field9)
      Continue, FOuter
  
    FLocFList .= "|" . Field1
    StringGetPos, LblPos, Field2, `–
    if(LblPos = -1)
      LblSet := Field2
    else
      StringLeft, LblSet, Field2, %LblPos%         
    FLblFList .= "|" . LblSet
    FIFList .=  "|" . Field3
    FMFList .=   "|" . Field4
    ;~ if (Field8 = "")
      ;~ IssuedTo := "Not Issued"
    ;~ else
      ;~ IssuedTo := Field8
    ;~ FITFList .=   "|" . IssuedTo
    FITFList .=   "|" . Field8
    FSFList .=   "|" . Field9
    LV_Add("", Field1, Field2, Field3, Field4, Field5, Field6, Field7, Field8, Field9)
    if (Field7 <> "")
      MAC := "YES"
  }
  
  Sort, FLocFList, D| U
  Sort, FLblFList, D| U
  Sort, FIFList, D| U
  Sort, FMFList, D| U
  Sort, FITFList, D| U
  Sort, FSFList, D| U
  
  ;~ MsgBox, %FITFList%`n%FIFList%
  
  loop, parse, filters, |
    GuiControl, ,A_Loopfield, |
  
  Guicontrol, , FLoc, %FLocFList%
  StringReplace, tmp, FLocFList, |,,UseErrorLevel
  if (Errorlevel = 1)
    GuiControl, Choose, FLoc, 1
  
  GuiControl,, FLbl, %FLblFList%
  StringReplace, tmp, FLblFList, |,,UseErrorLevel
  if (Errorlevel = 1)
    GuiControl, Choose, FLbl, 1

  Guicontrol, , FI, %FIFList%
  StringReplace, tmp, FIFList, |,,UseErrorLevel
  if (Errorlevel = 1)
    GuiControl, Choose, FI, 1
  
  Guicontrol, , FM, %FMFList%
  StringReplace, tmp, FMFList, |,,UseErrorLevel
  if (Errorlevel = 1)
    GuiControl, Choose, FM, 1
    
  Guicontrol, , FIT, %FITFList%
  StringReplace, tmp, FITFList, |,,UseErrorLevel
  if (Errorlevel > 1)
    Guicontrol, , FIT, Multiple...
  else if (ErrorLevel = 1)
    GuiControl, Choose, FIT, 1
  
  Guicontrol, , FS, %FSFList%
  StringReplace, tmp, FSFList, |,,UseErrorLevel
  if Errorlevel = 1
    GuiControl, Choose, FS, 1

  loop, parse, FFLists, |
    %A_LoopField% := ""
  
  Loop 8
    LV_ModifyCol(A_Index,"AutoHdr")
  
  if (MAC = "NO")
    LV_ModifyCol(7,0)
  
  FInv := ""
  Gui, ChkCustList: Destroy
Return

MultiCustGui:
Gui, 1:Default
CustList := GrabView("8","Inventory","NO")
StringReplace, CustList, CustList, `n, `,,All
;~ StringReplace, CustList, CustList, `r,,All
StringReplace, CustList, CustList, Issued To`,

Sort, CustList, D`, U

loop, Parse, CustList, `,
  TtlC := A_Index

TtlC -= 2

if (TtlC >= 15)
  R := 15
else
  R := TtlC

Gui, 1: +Disabled
Gui, ChkCustList: New
Gui, ChkCustList: Default
Gui, ChkCustList: +owner1
Gui, +LastFound +HwndGuiHWND
;~ Gui, +OwnDialogs    [or maybe something to do with parent? look at rest of code for answer]
Gui, Add, Text,, Please select all Customers to be included:
Gui, Add, ListView, w300 R%R% cBlack Grid Checked NoSortHdr vLV_MultiCust, Customer
;~ Gui, ListView, SysListView321
Gui, Add, Button, gMultiCust, OK
Gui, Show,, Select Multiple Customers

;~ Gui, ChkCustList:ListView, LV_MultiCust
loop, Parse, CustList, `,
{
  If not A_Loopfield
    Continue
  ;~ MsgBox, %A_Loopfield%
  ;~ MsgBox, LV_Add("", %A_Loopfield%)
  LV_Add("", A_Loopfield)
  ;~ MsgBox, %test%
}
Gui, Show,, Select Multiple Customers
WinWaitClose, ahk_id %GuiHWND%
Return

ChkCustListGuiClose:
ChkCustListGuiEscape:
MultiCustList := "-1"
Gui, ChkCustList: Destroy
Gui, 1: -Disabled
Gui, 1: Default
;~ Gui, 1: +owner1
WinActivate, ahk_class AutoHotkeyGUI
Return


MultiCust:
Gui, ChkCustList: Default
;~ Gui, Submit, NoHide
RowNumber := 0
MultiCustList := ""
Loop
{
  RowNumber := LV_GetNext(RowNumber, "Checked")
  if not RowNumber
    Break
  LV_GetText(Text, RowNumber)
  MultiCustList .= Text . "`,"
}
if (MultiCustList = "")
  MultiCustList := "-1"
Gui, ChkCustList: Submit
Gui, 1: -Disabled
Gui, 1: Default
WinActivate, ahk_class AutoHotkeyGUI
Return



ListCleanup:
Loop, Parse, Lists, |
{
  if (A_Loopfield = "")
    Break
  StringReplace, %A_Loopfield%, %A_Loopfield%, `~, `,, All
  sort, %A_Loopfield%, U D|
  If (SubStr(%A_Loopfield%,1,1)="|")
    StringTrimLeft, %A_Loopfield%, %A_Loopfield%, 1
}
Loop, Parse, FLists, |
{
  if (A_Loopfield = "")
    Break
  StringReplace, %A_Loopfield%, %A_Loopfield%, `~, `,, All
  sort, %A_Loopfield%, U D|
  If (SubStr(%A_Loopfield%,1,1)="|")
    StringTrimLeft, %A_Loopfield%, %A_Loopfield%, 1
}
StringReplace, FullLblList, FullLblList, `~, `,, All
sort, FullLblList, U D`,
If (SubStr(FullLblList,1,1)=",")
StringTrimLeft, FullLblList, FullLblList, 1
StringSplit, FList, FLists, |
Loop, Parse, Filters, |
{
  List := "FList" . A_Index
  FList := %List%
  GuiControl, ,%A_Loopfield%, |Blank
  GuiControl, ,%A_Loopfield%, % "|" . %FList%
  GuiControl, Choose, %A_Loopfield%, 1
}

StringReplace, FIListC, FIList, |, `,, ALL
IfNotInString, FIListC, VoIP
{
  Gui, ListView, Inventory
  LV_ModifyCol(7,0)
  GuiControl, Hide, ExpMAC
}
else
  GuiControl, Show, ExpMAC
Return


;~ FReset:
;~ FList=
;~ gosub, BuildInventory
;~ GuiControl, Choose, FS, 1
;~ return


EditGUI:
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...Gui, Listview, ScanGUIAdd
LV_GetText(Count, Row, "1")
LV_GetText(txtLoc, Row, "2")
LV_GetText(txtLbl, Row, "3")
LV_GetText(txtItm, Row, "4")
LV_GetText(txtMdl, Row, "5")
LV_GetText(txtSN, Row, "6")
IF (voipyn = "Y")
LV_GetText(txtMac, Row, "7")


G2LblWid = 100
G2EditWid = 240
G2Wid := 20 + G2LblWid + G2EditWid
G2BtnX := G2Wid//2-50

Gui, 2:Default
Gui, Font, S9 CDefault, SegoiUI
Gui, Add, Text, x20 y20 w%G2LblWid% h30 , Location:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Label:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Item:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Model:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Serial:
IF (voipyn = "Y")
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , MAC Address:


Gui, Add, Edit, x%G2LblWid%+20 y20 w%G2EditWid% h30 -Multi vtxtLoc, %txtLoc%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtLbl,%txtLbl%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtItm, %txtItm%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtMdl, %txtMdl%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtSN, %txtSN%
IF (voipyn = "Y")
  Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtMac, %txtMac%

if (editing = "N")
  Gui, Add, Button, x%G2BtnX% y+30 w100 h30 gEdit, Edit Item
else if (editing = "Y")
  Gui, Add, Button, x%G2BtnX% y+30 w100 h30 gEditInv, Edit Item
Gui, Show, w%G2Wid% , Edit Item
SplashTextOff
return

Edit:
Gui, 2:Submit, NoHide
Gui, 2:Hide
Gui, 1:Default
LV_Modify(Row, "", Count, txtLoc, txtLbl, txtItm, txtMdl, txtSN)
if (voipyn = "Y")
  LV_Modify(Row, "Col7", txtMac)
LV_ModifyCol(0,"AutoHdr")
Gui, 2:Destroy
Return




DeleteItem:
LV_GetText(RowLabel, Row, 3)
Loop, Parse, RowLabel, –
{
  if A_Index = 1
  {
    DefLabel := A_LoopField
    Break
  }  
}
Ttl := LV_GetCount()
if(Ttl = Row)
{
  gosub, undo
  Return
}
LV_Delete(Row)
Ttl := LV_GetCount()
Loop
{
  LV_Modify(Row, "Col1", Row)
  LV_GetText(LabelChange, Row, 3)
  Loop, Parse, LabelChange, –
  {
      if A_Index = 1
      {
        if (A_Loopfield <> DefLabel)
          Break
      }  
      if A_Index = 2
      {
        lblCount := RegExReplace(A_LoopField, "\D")
        lblCount -= 1
        if (lblCount < 10)
          LC := "0" . lblCount
        else
          LC := lblCount
        LV_Modify(Row, "Col3", DefLabel . "–" . LC)
      }
  }
  Row += 1
  if(Row > Ttl)
    Break
}
Line -= 1
Return

2GuiClose:

Gui, 2: Destroy
Return

EditCSVGUI:
Gui, Listview, Inventory
LV_GetText(txtLoc, Row, 1)
LV_GetText(txtLbl, Row, 2)
LV_GetText(txtItm, Row, 3)
LV_GetText(txtMdl, Row, 4)
LV_GetText(txtQty, Row, 5)
LV_GetText(txtSN, Row, 6)
LV_GetText(txtMac, Row, 7)
LV_GetText(txtIss, Row, 8)
LV_GetText(txtStat, Row, 9)

G2LblWid = 100
G2EditWid = 240
G2Wid := 20 + G2LblWid + G2EditWid
G2BtnX := G2Wid//2-50

Gui, 3:Default
Gui, Font, S9 CDefault, SegoiUI
Gui, Add, Text, x20 y20 w%G2LblWid% h30 , Location:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Label:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Item:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Model:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Quantity:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Serial:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , MAC Address:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Issued To:
Gui, Add, Text, xp+0 y+30 w%G2LblWid% h30 , Status:


Gui, Add, Edit, x%G2LblWid%+20 y20 w%G2EditWid% h30 -Multi vtxtLoc, %txtLoc%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtLbl,%txtLbl%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtItm, %txtItm%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtMdl, %txtMdl%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtQty, %txtQty%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtSN, %txtSN%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtMac, %txtMac%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtIss, %txtIss%
Gui, Add, Edit, xp+0 y+30 w%G2EditWid% h30 -Multi vtxtStat, %txtStat%

Gui, Add, Button, x%G2BtnX% y+30 w100 h30 gEditCSV, Edit Line
Gui, Show, w%G2Wid% , Edit Line
return

EditCSV:
Gui, 3:Submit, NoHide
Gui, 3:Hide
Gui, 1:Default
Gui, Listview, Inventory
LV_Modify(Row, "", txtLoc, txtLbl, txtItm, txtMdl, txtQty, txtSN, txtMac, txtIss, txtStat)
LV_ModifyCol(0,"AutoHdr")
Gui, 3:Destroy
Return

DeleteCSV:
Gui, Listview, Inventory
LV_Delete(Row)
Return

3GuiClose:
Gui, 3:Destroy

Controls:
Loop, parse, Controls, `n, `r
  GuiControl, %EnDis%, %A_Loopfield%
Return

MakeChanges:
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
Gui, ListView, Inventory
gosub, BuildInventory
EnDis := "Disable"
gosub, Controls
Guicontrol, Show, Apply
Guicontrol, Show, Cancel
editing := "Y"
SplashTextOff
Return

subtab:
;~ gosub, ChangeMenu        ;RE-ENABLE
if (editing = "Y")
{
  GuiControlGet, TabArea      ; get the new selected tab
  If (TabArea <> 2)            ; tab #2 is "disabled" so ...
    GuiControl, Choose, %A_GuiControl%, 2 ; ... reset the selection to the previous selected tab
}

Return

CancelChanges:
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
Gui, ListView, Inventory
gosub, BuildInventory
EnDis := "Enable"
gosub, Controls
Guicontrol, Hide, Apply
Guicontrol, Hide, Cancel
editing := "N"
SplashTextOff
Return

ApplyChanges:
;~ FileDelete,%A_ScriptDir%\Inventory.old.csv 
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
Gui, ListView, Inventory
Ttl := LV_GetCount()
adding := ""
Loop %Ttl%
{
  AddRow := ""
  RowNumber := A_Index
  
  Loop 9
  {
    LV_GetText(Text, RowNumber, A_Index)
    StringReplace, Text, Text, `, , ~, ALL
    AddRow .= Text . "`,"
  }
  Adding .= AddRow . "`n"
}
FileCopy, %csvF%, %ArchiveDir%\Inventory.%A_YYYY%%A_MM%%A_DD%.%A_Hour%%A_Min%%A_Sec%.Changed.csv, 1
CleanArchive(ArchiveDir,ArNum)
Clipboard := ""
FileAppend, %ClipboardAll%, %csvF%
;~ Loop
;~ {
  ;~ IfNotExist, %csvF%
    ;~ Break
  ;~ FileDelete, %csvF%
  
  
FileAppend, %Adding%, %csvF%
SplashTextOff
Adding=
gosub, CancelChanges
Return

TIScan:
Gui, ListView, TIInventory
Loop
{
  Gui,+LastFound
  WinGetPos,gx,gy
  Gui +owndialogs
  Inputbox, SSerial, Scan Serial, Scan Serial Number,, 185, 100, %gx%, %gy% 
  If (errorlevel <> 0 or SSerial = "")
    Return
  if SubStr(SSerial, 1, 3) = "IL-"
  {
    Loop, parse, SSerial,`-
    {
      if (A_Index = 1 of A_Index = 2)
        Continue
      if A_Index = 3
        SSerial := A_Loopfield
    }
  }
  Search := "|" . SSerial . "|"
  IfNotInString, CustSNList, %Search%
  {
    SoundPlay, %A_ScriptDir%\Sounds\Error.wav
    MsgBox, 4113, Wrong Serial Number, Serial Number %SSerial% was not issued to this customer!
    IfMsgBox OK
      Continue
    IfMsgBox Cancel
      return
  }
  Loop
  {
    Row := A_Index
    LV_GetText(Text, Row, 1)
    if (Text = SSerial)
      Break
  }
  LV_Modify(Row, "Vis Focus Select Check")
  ScannedSerials .= SSerial . "|"
  GuiControl,,Scanned, %SSerial%
  TIItems += 1
  GuiControl,,TIRunning, Serialized Items to Turn-In: %TIItems%
  SoundPlay, %A_ScriptDir%\Sounds\Success.wav
}
Return


TINonSNGui:
Gui, 6:Default
gui, font, s12
Gui, Add, Text, x10 ,Choose Quantities to turn in: (Can Be 0)
Loop, Parse, TINOSNList, |
{
  if (A_Loopfield = "")
    Break
  Desc := A_Loopfield
  gui, font, Bold
  Gui, Add, Text, x10 y+20,%A_Loopfield%
  gui, font, Norm
  Gui, Add, Text, x10,Quantity to Turn-In:
  Gui, 1:Default
  RowNumber := Search("0",Desc,"5","TIInventory")
  Gui, ListView, TIInventory
  LV_GetText(TIQTY, RowNumber, 6)
  Gui, 6:Default
  txt := RegExReplace(A_Loopfield, "\W", "_")
  varname := txt . "Q"
  Gui, Add, Edit, x+10 yp+0  w50 v%varname%, %TIQTY%
}
Gui, Add, Button, y+20 gTINonSN,Turn-In
Gui, Show,,Items without Serial Numbers
Return


6GuiClose:
Gui, 6:Destroy
Return

TINonSN:
Gui, 6:Default
Gui, Submit
FileRead, Inv, %csvF%
Gui, 1:Default
Loop, Parse, TINOSNList, |
{
  if not A_LoopField
    Break
  txt := RegExReplace(A_Loopfield, "\W", "_")
  varname := txt . "Q"
  TIQTY := %varname%
  if not TIQTY
    continue
  TIModel := A_Loopfield
  StringReplace, TINOSNSelection, TINOSNSelection, `,, `~, All
  TIItem := ""
  
  Desc := A_Loopfield
  RowNumber := Search("0",Desc,"5","TIInventory")
  Gui, ListView, TIInventory
  LV_GetText(TIItem, RowNumber, 4)

  ;~ Loop, Parse, Inv, `n, `r
  ;~ {
    ;~ StringSplit, List, A_Loopfield, `,
    ;~ MsgBox, CustSelect: %CustomerSelect% / TINOSNSelection: %TINOSNSelection%`nList8: %List8%  / List4: %List4%
    ;~ if(List8 = CustomerSelect and List4 = TINOSNSelection)
    ;~ {
      ;~ TIItem := List3
      ;~ Break
    ;~ }
    ;~ loop, 9
    ;~ {
      ;~ List := "List" . A_Index
      ;~ %List% := ""
    ;~ }
  ;~ }
  ;~ MsgBox, %CustomerSelect% / TINOSNSelection: %TINOSNSelection%
  ;~ MsgBox, TIItem (SHOULD NOT BE BLANK): %TIItem%

  RowN := DblSearch(TIModel, "Model", "On-Hand", "Status", "Inventory", ColumnList)
  Gui, ListView, Inventory
  if (RowN = "")
    LV_Add("",TIItem,,TIItem,TIModel,TIQTY,,,,"On-Hand")
  else
  {
    LV_GetText(OldQty, RowN, 5)
    NewQty := OldQty + TIQTY
    LV_Modify(RowN,"Col5",NewQty)
  }
  RowN := DblSearch(TIModel, "Model", CustomerSelect, "Issued To", "Inventory", ColumnList)
  LV_GetText(OldQty, RowN, 5)
  NewQty := OldQty - TIQTY
  if (NewQty = 0)
    LV_Delete(RowN)
  Else
    LV_Modify(RowN,"Col5",NewQty)
}
gosub, ApplyChanges

Gui, 6:Destroy
TINOSNList := ""
FileRead, Inv, %csvF%
Loop, Parse, Inv, `n, `r
{
  StringSplit, List, A_Loopfield, `,
  if(List8 = CustomerSelect and List6 = "")
    TINOSNList .= "|" . List4
  loop, 9
  {
    List := "List" . A_Index
    %List% := ""
  }
}
StringTrimLeft, TINOSNList, TINOSNList, 1
StringReplace, TINOSNList, TINOSNList, `~, `,, All
If(TINOSNList = "")
  GuiControl, Disable, TINonSN
Else If(TINOSNList <> "")
  GuiControl, Enable, TINonSN
Gui, ListView, TIInventory
Ttl := LV_GetCount()
if(Ttl = 0 or Ttl = "")
{
  gosub, BuildInventory
  GuiControl,,CustomerSelect,|Choose Customer|%ITList%
  GuiControl, Choose, CustomerSelect, 1
  gosub, CustomerSelect
}
Inv :=
Return





TIApply:
;~ SplashTextOn, 200, 100, Please Wait, Please wait...
Gui, ListView, TIInventory
RowNumber = 0  ; This causes the first loop iteration to start the search at the top of the list.
Loop
{
  Gui, ListView, TIInventory
  RowNumber := LV_GetNext(RowNumber, "Checked")  ; Resume the search at the row after that found by the previous iteration.
  if not RowNumber  ; The above returned zero, so there are no more selected rows.
      break
  LV_GetText(Text, RowNumber)
  if (Text = "")
    gosub, TINOnSN
  TISerials .= Text . "`,"
}
;~ StringTrimRight, TISerials, TISerials, 1
;~ Gosub, TIClear
gosub, BuildInventory
;~ GuiControl, Choose, TabArea, 2
Gui, ListView, Inventory
Ttl := LV_GetCount()
Loop %Ttl%
{
  LV_GetText(Text, A_Index, 6)
  if Text in %TISerials%
  {
    LV_Modify(A_Index, "Col8","")
    LV_Modify(A_Index, "Col9","On-Hand")
    ;~ Text := Text . "`,"
    ;~ StringReplace, TISerials, TISerials, %Text%,,All
  }
}
gosub, ApplyChanges
;~ gosub, BuildInventory
GuiControlGet,CustomerSelect
gosub, CustomerSelect
Gui, ListView, TIInventory
Ttl := LV_GetCount()
if(Ttl = 0 or Ttl = "")
{
  ;~ gosub, BuildInventory
  GuiControl,,CustomerSelect,|Choose Customer|%ITList%
  GuiControl, Choose, CustomerSelect, 1
  gosub, CustomerSelect
}

TISerials=

SplashTextOff
return

TIClear:
Gui, ListView, TIInventory
LV_Delete()
GuiControl,Choose,CustomerSelect,1
TIItems := 0
GuiControl,,TIRunning, Serialized Items to Turn-In: %TIItems%
GuiControl, Disable, TIScan
GuiControl, Disable, TIApply
GuiControl, Disable, TINonSN
GuiControl, Disable, TIClear
GuiControl,,Scanned, |Serials Scanned:
Return


SubLV:
If (A_GuiEvent = "RightClick") {
      CopiedCell=
      Row=
      Column=
      Row := A_EventInfo
      Column := LV_SubItemHitTest(HLV1)
      MouseGetPos, VarX, VarY
      Menu, NormalContext, Show, %VarX%, %VarY%
   }
Return

CopyCell:
Gui, ListView, Inventory
LV_GetText(CopiedCell, Row, Column)
Clipboard := CopiedCell
Return




BuildLVInventory:
{
LocList=
LblList=
IList=
MList=
ITList=
Inv=
CustSNList=

Gui, ListView, Inventory
LV_Delete()
FileRead, Inv, %csvF%
Sort, Inv
StringTrimRight, Inv, Inv, 2
Loop, Parse, Inv, `n, `r
{
  RowNum := A_Index
  Loop, Parse, A_Loopfield, csv
  {
    Column := "Col" . A_Index
    %Column% := A_Loopfield
    StringReplace, %Column%, %Column%, `~, `,, All
    if (A_Index = 1)
    {
      LocList .= "`|" . A_Loopfield
      FLocList .= "`|" . A_Loopfield
    }
    if (A_Index = 2)
    {
      FullLblList .= "`," . A_Loopfield
      IfInString, A_Loopfield, –
      {
        StringGetPos, DL, A_Loopfield, –, R
        StringMid, label, A_Loopfield, 1, %DL%
      }
      else
        label := A_Loopfield
      LblList .= "`|" . label
      FLblList .= "`|" . label
    }
    if (A_Index = 3)
    {
      IList .= "`|" . A_Loopfield
      FIList .= "`|" . A_Loopfield
    }
    if (A_Index = 4)
    {
      MList .= "`|" . A_Loopfield
      FMList .= "`|" . A_Loopfield
    }
    if (A_Index = 6)
    {
      SList .= "|" . A_Loopfield . "|"
      if (A_Loopfield = "")
      {
        NoSNIList .= "|" . Col3
        FNoSNIList .= "|" . Col3
      }
    }
    if (A_Index = 8)
    {
      ITList .= "|" . A_Loopfield
      FITList .= "|" . A_Loopfield
    }
  }
  Gui, ListView, Inventory
  LV_Add("",Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9)

    
}
Loop 8
{
  Gui, ListView, Inventory
  LV_ModifyCol(A_Index,"AutoHdr")
}

Return
}


^!a::
{
  GuiControl, -AltSubmit, OHInventory
  Gui, ListView, TIInventory

  Ttl := LV_GetCount()
  Loop %Ttl%
  {
    LV_GetText(Text, A_Index, 1)
    if (Text <> "")
      LV_Modify(A_Index,"Check")
  }
  
  Gui, ListView, OHInventory
  Ttl := LV_GetCount()
  Loop %Ttl%
  {
    LV_GetText(Text, A_Index, 7)
    if (Text <> "")
      LV_Modify(A_Index,"Check")
  }
  GuiControl, +AltSubmit, OHInventory
}
Return

^!z::
{
  Gui, ListView, TIInventory
  Ttl := LV_GetCount()
  {
    LV_GetText(Text, A_Index, 1)
    if (Text <> "")
      LV_Modify(A_Index,"-Check")
  }
  
  Gui, ListView, OHInventory
  Ttl := LV_GetCount()
  Loop %Ttl%
    LV_Modify(A_Index,"-Check")
}
Return



ExpMAC:

LookFor := CustGUI("Generating")
E := 8
gosub, BuildInventory
;~ gosub, Filtration
E := 7
LookFor := "All"
;~ gosub, Filtration

LV := "Inventory"
SelCols := "2|4|6|7"
TtlC := 4
Issue := "NO"
gosub, Export
gosub, BuildInventory
Return

AddIL:
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
ILList=
(
ILLocList
ILLblList
ILModList
)
Gui, ListView, ScanGuiAdd
ILLocList := CompileList("All", "Location", "Location", "ScanGuiAdd", ScanGuiCols)
ILLblList := CompileList("All", "Label", "Label", "ScanGuiAdd", ScanGuiCols)
ILModList := CompileList("All", "Model", "Model", "ScanGuiAdd", ScanGuiCols)
Loop, Parse, ILList, `n,`r
{
  StringReplace, %A_Loopfield%, %A_Loopfield%, `,, ~, All
  if (%A_Loopfield% <> "")
    %A_Loopfield% .= "|"
}
StringReplace, ILLocList, ILLocList, `, , |, All
StringReplace, ILModList, ILModList, `, , |, All
StringReplace, ILLocList, ILLocList, ~, `, , All
StringReplace, ILModList, ILModList, ~, `, , All
FileRead, Inv, %csvF%
StringReplace, Inv, Inv, `, , \, All
StringReplace, Inv, Inv, ~, `, , All
Sort, Inv, \
Loop, Parse, Inv, `n, `r
{
  StringSplit, Col, A_Loopfield, \
  if (Col3 = "Issuable Location")
  {
    ILLocList .= Col1 . "`,"
    ILLblList .= Col2 . "`,"
    ILModList .= Col4 . "|"
  }
}
Inv :=""
Loop, Parse, ILList, `n, `r
{
  Modding := %A_LoopField%
  StringTrimRight, Modding, Modding, 1
  Sort, Modding, D| U
  %A_Loopfield% := Modding
}
;~ StringReplace, Inv, Inv, ~, `, , All
Gui, ILGui:New  ; Creates a new GUI Named ILGui.
Gui, font, s12
Gui, Add, Text, x10, Enter this location's name: (i.e. Box 01)
Gui, Add, Edit, x+5 yp+0 h20 w300 gILChngLbl vILLocName, 
Gui, Add, Text, x10 y+20, Enter this location's Label:
Gui, Add, Edit, x+5 yp+0 h20 w300 vILLabel,
Gui, Add, Text, x10 y+20, Enter or select this location's description: (i.e Box, Pelican)
Gui, Add, ComboBox, R5 x+5 yp+0 h20 w300 vILDesc,%ILModList%
Gui, Add, Button, x10 y+10 gILAddToQueue, Add to Queue

Gui, Show
SplashTextOff
Return
;--------------------------
ILChngLbl:
Gui, ILGui: Default
GuiControlGet, ILLocName 
GuiControl, , ILLabel, %ILLocName%
Return

ILAddToQueue:
Gui, ILGui: Submit
ErrMsg := ""
if ILLocList contains %ILLocName%
;~ IfInString, ILLocList, %ILLocName%
  ErrMsg := "There is already a Location by that name in your inventory!`n"
if ILLblList contains %ILLabel%
;~ IfInString, ILLabel, %ILLblList%
  ErrMsg .= "That Label is already taken!"

if (ErrMsg <> "")
{
  MsgBox, 4112, Invalid Entry, %ErrMsg%
  Gui, ILGui: Show
  Return
}
Gui, ILGui: Destroy
Gui, 1: Default
Gui, ListView, ScanGuiAdd
Line += 1
LV_Add("", Line, ILLocName, ILLabel, "Issuable Location", ILDesc, "1", ILLocName)
LV_Modify(Line, "Col10", "On-Hand")
loop %ct%
  LV_ModifyCol(A_Index,"AutoHdr")
GuiControl, Text, loc, %ILLocName%
GuiControl, Text, lbl, %ILLocName%–01
SoundPlay, %A_ScriptDir%\Sounds\Change.wav
Return




;-------
ILGuiGuiEscape:
ILGuiGuiClose:
  Gui, ILGui:Destroy
  Gui, 1:Default
return 


ChangeMenu:
GuiControlGet, TabNum, , TabArea
;~ MsgBox, %CurrTab% %TabNum%
if (CurrTab = TabNum)
  Return
CurrTab := TabNum
;~ Return
Gui, Menu,

;~ Tab 1: ScanGui Add to Inventory
if (TabNum = 1)
{
  Menu, MenuBar, Add, &Items, :ItemsMenu
}


;~ Tab 2: Full Inventory
if (TabNum = 2)
  Return
;~ {
  ;~ Menu, MenuBar, Add, &Display Name, :MenuName
;~ }

;~ Tab 3: On-Hand Inventory
if (TabNum = 3)
  Return
;~ {
  ;~ Menu, MenuBar, Add, &Display Name, :MenuName
;~ }

;~ Tab 4: Turn-In
if (TabNum = 4)
  Return
;~ {
  ;~ Menu, MenuBar, Add, &Display Name, :MenuName
;~ }

;, MenuBar, Add, &Items, :ItemsMenu
Gui, Menu, MenuBar

Return

RegisterItem:
SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
gosub, BuildInventory
RIList := CompileList("NonBlank", "Serial", "Item", "Inventory", ColumnList)
;~ Retrieves list of all Items. Separated by "|", showing commas, no trailing "|", no header
Sort, RIList, D| U
Sort, RIList
StringReplace, RIList, RIList, Issuable Location,,All
StringReplace, RIList, RIList, ||, |,All
if (SubStr(RIList, 1, 1) = "|")
  StringTrimLeft, RIList, RIList, 1
if (SubStr(RIList, 0, 1) = "|")
  StringTrimRight, RIList, RIList, 1

Gui, RI:New
Gui, font, s12 Bold
Gui, Add, Text, x10, This enables you to register a scanned Item Number or Part Number`,`n so the program automatically changes the Item and Model values`n(Like when you scan a label that says "I-Laptop-Dell-E5500")
Gui, font,, Norm
;~ Gui, font, s12
Gui, Add, Text, x10 y+30, Enter or Scan Item Number to Register: ;(Special Characters and spaces are removed automatically)
Gui, Add, Edit, x+5 yp+0 h20 w300 gOldInfoFill vINum, 
Gui, Add, Text, x10 y+20, Enter or select this Item's Item Type: (i.e. VoIP Phone)
Gui, Add, ComboBox, R5 x+5 yp+0 h20 w300 vIType, %RIList%
Gui, Add, Text, x10 y+20, Enter this Item's Model: (i.e. Cisco 7945)
Gui, Add, Edit, x+5 yp+0 h20 w300 vIMod,
Gui, Add, Button, x10 y+10 gRI, Register

Gui, Show, ,Register Item
SplashTextOff
Return

;-------
RIGuiEscape:
RIGuiClose:
  Gui, RI:Destroy
  Gui, 1:Default
return 
;--------
OldInfoFill:
;~ MsgBox, its looking!
Gui, RI:Submit, NoHide

;~ GuiControlGet, INum
;~ Gui submit, NoHide
iniread, Test, %ini%, RegisteredItems, %INum%,
If (Test = "ERROR" or Test = "")
  Return
;~ MsgBox, %Test%
StringSplit, Info, Test, `\
OldItem := Info1
OldModel := Info2
GuiControl, Text, IType, %OldItem%
GuiControl, Text, IMod, %OldModel%
Return

RI:
Gui, RI:Submit
;~ Gui, Submit
;~ INum := RegExReplace(INum, "\W", "_")
;~ StringReplace, INum, INum, %A_Space%,,All
if (INum = "")
{
  MsgBox, 4112, Invalid Entry, This Item Number is already registered!

}
iniread, Test, %ini%, RegisteredItems, %INum%,
;~ MsgBox, %Test%
if (Test <> "ERROR")
{
  ;~ MsgBox, 4112, Invalid Entry, This Item Number is already registered!
  MsgBox, 4132, Invalid Entry, This Item Number is already registered!`nUpdate the values?
  IfMsgBox, No
    Return
  ;~ Gui, RI:Show
  ;~ Return
}

iniwrite, %IType%\%IMod%, %ini%, RegisteredItems, %INum%
;~ gosub, RIGuiClose
;~ Gui, RI:Destroy
;~ Gui, 1:Default
;~ MsgBox, Are Variables dumped? %IMod%
RIList .= "|" . IType
Sort, RIList, D| U
Sort, RIList
StringReplace, RIList, RIList, Issuable Location,,All
StringReplace, RIList, RIList, ||, |,All
if (SubStr(RIList, 1, 1) = "|")
  StringTrimLeft, RIList, RIList, 1
if (SubStr(RIList, 0, 1) = "|")
  StringTrimRight, RIList, RIList, 1
iniread, RegisteredItems, %ini%, RegisteredItems
GuiControl, Text, itm, %IType%
GuiControl, Text, Mod, %IMod%
;~ MsgBox, %Test2%
GuiControl, Text, INum, %A_Space%
GuiControl, , IType, |%RIList%
GuiControl, Text, IMod, %A_Space%
Gui, RI:Show
Return

Insert::Pause

Verify:
iniread, PrevHR, %ini%, HR,HRNum
PrevHR := "0000" . PrevHR
StringRight, PrevHR, PrevHR, 5
InputBox, VHRNum, Enter Hand Receipt Number, Enter Hand Receipt Number to be verified:,,,,,,,,%PrevHR%
Gui, Verify:New
iniread, VSerials, %ini%, IssuedSerials, %VHRNum%
if (SubStr(VSerials, 1, 1) = "|")
  StringTrimLeft, VSerials, VSerials, 1

StringReplace, CSVSerials, VSerials, `,, ~, all 
StringReplace, CSVSerials, CSVSerials, |, `,, all 
Gui, add, text, x10 y10, Scan Serial:
Gui, Add, Edit, w200 vVScan, 
Gui, add, ListBox, vVerifySerials x+10 y10 w200 h500 ReadOnly, %VSerials%
gui, add, Button, default gVScanEntry x10 y100 vButton,OK 
Gui, Show, ,Verify Issue
GuiControl, Hide, Button
ControlFocus, VScan
Return

VScanEntry:
GuiControlGet, VScan
StringReplace, ChkVScan, VScan, `,, ~, All
if VSerials not contains %ChkVScan%
{
  MsgBox, 4112, Invalid Entry, Serial Number: %VScan% is not supposed to be issued with this order!
  ControlFocus, VScan
  GuiControl, Text, VScan,
  Return
}
StringReplace, VSerials, VSerials, %VScan%,,All
StringReplace, VSerials, VSerials, ||, |,All

GuiControl, Text, VerifySerials, |%VSerials%
ControlFocus, VScan
GuiControl, Text, VScan,
Return

#Z::         ;DELETE THIS AFTER TESTING
BuildSetFilters:
;~ SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
SelCols := "1|2|3|4|7|8"
FInfo := GrabView(SelCols,"Inventory","NO")   
  /*
  Returns "," separated list from current view based on criteria, commas are ~
  SelCols: Which Columns to grab, formatted 1|2|5|8
  LV: Which Listview, "ScanGuiAdd", "Inventory", "OHInventory" or "TIInventory"
  Issue: "YES" will only collect Checked Items, "NO" grabs all
  */
;~ Firstus is the Header on each Filter, `n separated lines, all with trailing "|" 
StringReplace, FInfo, FInfo, `,, |, All
StringReplace, FInfo, FInfo, `~, `,, All
loop, Parse, Filters, |
  GuiControl, , %A_Loopfield%, |
;~ Return
StringSplit, FName, Filters, |
StringSplit, FHdr, Firstus, `n, `r
Loop, Parse, FInfo
{
  StringSplit, Col, FInfo, |
  Loop, 6
  {
    FNmTemp := "FName" . A_Index
    FName := %FNmTemp%
    FTemp := "Col" . A_Index
    FCol := %FTemp%
    MsgBox, %FName%
    %FName% .= %FCol% . "|"
    
  }
  
}
;~ MsgBox, %Floc% %FIT% 
MsgBox, %FLoc%`n%FLbl%`n%FI%`n%FM%`n%FIT%`n%FS%
Return
loop, Parse, Filters, |
  Sort, %A_Loopfield%, D| U

loop, Parse, Filters, |


FHdr := "FHdr" . A_Index
  


SplashTextOff
Return
