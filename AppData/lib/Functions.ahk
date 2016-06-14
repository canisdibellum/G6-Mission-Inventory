CompileList(Find, Criteria, Col, LV, ColumnList)
{
  /*
  Find: What do you want to be in "Criteria" Column? (All means ALL)
  Criteria: What Column do you want to find "Find"?
  Col: What Column will do you want to build list of if "Find" is found?
  LV: Which ListView? (Name)
  ColumnList: ALWAYS ColumnList
  */
  ;~ Gui, 1: Default
  Gui, ListView, %LV%
  loop, parse, ColumnList, `|
  {
    if (A_Loopfield = Col)
    ;~ {
      A := A_Index
      ;~ break
    ;~ }
  ;~ }
  ;~ loop, parse, ColumnList, `|
  ;~ {
    if (A_Loopfield = Criteria)
    ;~ {
      B := A_Index
      ;~ break
    ;~ }
  }
  ;~ MsgBox, % Find . " / " . Criteria . " " . B . " / " . Col . " " . A . " / " . LV . " / " . ColumnList
  NList=
  RowNumber = 0
  Gui, ListView, %LV%
  Ttl := LV_GetCount()
  Loop %Ttl%
  {
    LV_GetText(Text, A_Index, B)
    ;~ MsgBox, %Text%
    if (Find = "NonBlank" and Text <> "")
      Text := "NonBlank"
    
    if (Text <> Find and Find <> "All" and Find <> "NONE")
      Continue
    if (Text = "" and Find = "NONE")
    {
      LV_GetText(Found, A_Index, A)
      if (NList = "")
        NList .= Found
      else
        NList .= "|" . Found
      Continue
    }
    LV_GetText(Found, A_Index, A)
    if (NList = "")
      NList .= Found
    else
      NList .= "|" . Found
  }
  Sort, NList, D| U
  ;~ MsgBox, %NList%
  Return NList
}

Search(ByRef SRow,SearchFor,Col,LV)
{
  Gui, Listview, %LV%
  Ttl := LV_GetCount()
  TtlC := LV_GetCount("Column")
  SearchLoop:
  Loop
  {
    SRow +=1
    if (SRow > Ttl)
      Break, SearchLoop
    if(Col = "All")
    {
      Loop, %TtlC%
      {
        LV_GetText(SText, SRow, A_Index)
        if (SText = SearchFor)
          Break, SearchLoop
      }
    }
    else
    {
      LV_GetText(SText, SRow, Col)
      
      if (SText = SearchFor)
          Break, SearchLoop
    }
   }
  if (SText <> SearchFor)
    SRow := 0
Return SRow
}

DblSearch(Find, Criteria, Find2, Criteria2, LV, ColumnList)
{
  Gui, ListView, %LV%
  loop, parse, ColumnList, `|
  {
    if (A_Loopfield = Criteria)
    {
      A := A_Index
      break
    }
  }
  loop, parse, ColumnList, `|
  {
    if (A_Loopfield = Criteria2)
    {
      B := A_Index
      break
    }
  }
  ;~ MsgBox, %Find% / %Criteria% %A%/ %Find2% / %Criteria2% %B%/ %LV% / %ColumnList%
  RowNumber = 0
  Gui, ListView, %LV%
  Ttl := LV_GetCount()
  Loop %Ttl%
  {
    LV_GetText(Text, A_Index, A)
    LV_GetText(Text2, A_Index, B)
    if (Text = Find and Text2 = Find2)
      Found := A_Index
  }
  ;~ if (Found = "")
    ;~ Found := 0
  Return Found
}

CleanArchive(Directory,ArNum)
{
  ;~ Keeps the number of archived inventories below the value defined in the ini (Default is 100)
  ;~ See "Archived File Name Explanation.txt" in the archive folder for more info and how to change that value if so desired.
  Files := ""
  Loop %Directory%\*.csv
    TtlA := A_Index
  if (TtlA > ArNum)
  {
    TtlA -= ArNum
    Loop %Directory%\*.csv
      Files .= A_LoopFileLongPath . "|"
    Sort, Files, D| \
    Loop, parse, files, |
    {
      if (A_Index <= TtlA)
        FileDelete, %A_Loopfield%
      else
        Break
    }
  }
  return
}


GrabView(SelCols,LV,Issue)   
{
  /*
  Returns "," separated list from current view based on criteria WITH HEADERS, commas are ~ 
  SelCols: Which Columns to grab, formatted 1|2|5|8
  LV: Which Listview, "ScanGuiAdd", "Inventory", "OHInventory" or "TIInventory"
  Issue: "YES" will only collect Checked Items, "NO" grabs all
  */
  
  Grid := ""
  Gui, 1: Default
  Gui, ListView, %LV%
  Ttl := LV_GetCount()
  Ttl += 1
  Loop, Parse, SelCols, |
    TtlC := A_Index
  RowNumber := 0
  Loop, %Ttl%
  {
    if (Issue = "YES" and A_Index <> 1)
    {
      RowNumber := LV_GetNext(RowNumber, "Checked")
      If Not RowNumber
        Break
    }
    Loop, Parse, SelCols, |         ;Parse through Selected Column Headers (SelCols)
    {
      LV_GetText(Text, RowNumber, A_Loopfield)
      StringReplace, Text, Text, `,, `~, ALL 
      GridRow .= Text
      if (A_Index <> TtlC)
        GridRow .= "`,"
    }
    Grid .= GridRow
    GridRow=
    if (RowNumber <> Ttl)
      Grid .= "`n"
    if (Issue = "No")
      RowNumber += 1
  }
  Return Grid
}

CompileItemLists(M,C,List)  ;(MasterIssueList,CompiledList,Desired Thing to be listed("SI"[Serialized Items] {or "ILM"[Issuable Location Models]} or "QI"[Quantified Items])
{
;~ C := "StartMe"
C := ""
Loop, parse, M, `n, `r
{
  Loop, parse, A_Loopfield, csv
  {
    if (A_Index = 1)
      I := A_Loopfield  ;Item
    if (A_Index = 4)
      M := A_Loopfield  ;Model
    if (A_Index = 7)
      S := A_Loopfield  ;Serial Number
  }
  If(List = "SI")
  {
    if (S <> "" and I <> "Issuable Location")
      C .= I . "`,"
      ;~ C .= "`n" . Item
  }
  /*
  If (List = "ILM")
  {
    if (I = "Issuable Location")
      C.Push(M)
      ;~ C .= "`n" . Item
  }
  */
  If (List = "QI")
  {
    if (S = "")
      C .= I . "`,"
      ;~ C .= "`n" . Item
  }
}
StringReplace, C, C, `n,,1
Sort, C, U D`,
Return C
}

CompileSerialLists(M,P,C,SLorQ)        ;(MasterIssueList, ParseList(ItemList), CompiledList, Serialized, Labeled or Quantified)
{
  ;~ XMLAddSerials      DELETEME
Loop, Parse, P, csv
{
  DI := A_Loopfield                   ;Desired Item
  Loop, parse, M, `n, `r
  {
    If (A_Loopfield = "")
      Break
    Loop, parse, A_Loopfield, csv
    {
      if (A_Index = 1)
        I := A_Loopfield  ;Item
      if (A_Index = 3)
        L := A_Loopfield  ;Label
      if (A_Index = 4)
        Mo := A_Loopfield  ;Model
      if (A_Index = 6)
        Q := A_Loopfield  ;Qty
      if (A_Index = 7)
        S := A_Loopfield  ;Serial Number
    }

    If (I <> DI)
      Continue
    
    C .= "`n"
    
    if (SLorQ = "S")
      C .= I . "|" . L . "|" . S
    if (SLorQ = "L")
      C .= Mo . "|" . S
    if (SLorQ = "Q")
    {
      if (S = "")
        C .= I . "|" . Mo . "|" . Q
    }
  }
}
StringReplace, C, C, `n,
Sort, C
Return C
}

LV_SubitemHitTest(HLV) {
   ; https://autohotkey.com/board/topic/80265-solved-which-column-is-clicked-in-listview/
   ; To run this with AHK_Basic change all DllCall types "Ptr" to "UInt", please.
   ; HLV - ListView's HWND
   Static LVM_SUBITEMHITTEST := 0x1039
   VarSetCapacity(POINT, 8, 0)
   ; Get the current cursor position in screen coordinates
   DllCall("User32.dll\GetCursorPos", "Ptr", &POINT)
   ; Convert them to client coordinates related to the ListView
   DllCall("User32.dll\ScreenToClient", "Ptr", HLV, "Ptr", &POINT)
   ; Create a LVHITTESTINFO structure (see below)
   VarSetCapacity(LVHITTESTINFO, 24, 0)
   ; Store the relative mouse coordinates
   NumPut(NumGet(POINT, 0, "Int"), LVHITTESTINFO, 0, "Int")
   NumPut(NumGet(POINT, 4, "Int"), LVHITTESTINFO, 4, "Int")
   ; Send a LVM_SUBITEMHITTEST to the ListView
   SendMessage, LVM_SUBITEMHITTEST, 0, &LVHITTESTINFO, , ahk_id %HLV%
   ; If no item was found on this position, the return value is -1
   If (ErrorLevel = -1)
      Return 0
   ; Get the corresponding subitem (column)
   Subitem := NumGet(LVHITTESTINFO, 16, "Int") + 1
   Return Subitem
}
/*
typedef struct _LVHITTESTINFO {
  POINT pt;
  UINT  flags;
  int   iItem;
  int   iSubItem;
  int   iGroup;
} LVHITTESTINFO, *LPLVHITTESTINFO;
*/

CustGUI(Status)
{
  ;~ Status can be either "Issuing" or "Generating" or "AddIL"
  global Customer
  ITList := CompileList("Issued", "Status", "Issued To", "Inventory", ColumnList)
  if (ITList = "" and Status = "Generating")
  {
    MsgBox, 4112, %A_Space%, Nothing is Issued!
    gosub, CustGuiGuiEscape
    Return
  }

  Sort, ITList, D| U
  if (Status <> "AddIL")
  {
    Loop, Parse, ITList, |
      R := A_Index
  }
  else
  {
    Loop, Parse, ILList, |
      R := A_Index
  }
  if (R > 5)
    R := 5
  if (ITList = "")
    R := 1
  Gui, CustGui:New  ; Creates a new unnamed and unnumbered GUI.
  Gui, +LastFound +HwndGuiHWND
  Gui, Font, s12
  if (Status = "Issuing")
    Gui, Add, Text, ,Select or enter customer:
  if (Status = "Generating")
    Gui, Add, Text, ,Select customer:
  if (Status = "AddIL")
    Gui, Add, Text, ,Select or enter Description:
  

  if (Status = "Issuing")
    Gui, Add, ComboBox, R%R% w200 vCustomer, %ITList%
  if (Status = "Generating")
    Gui, Add, DropDownList, R%R% w200 vCustomer, %ITList%
  if (Status = "AddIL")
    Gui, Add, ComboBox, R%R% w200 vCustomer, %ILList%
  
  Gui, Add , Button, Default gButtonOK, OK
  if (Status = "Generating")
    Gui, Add , Button, gMultiple, Multiple Customers
  Gui, Show,,%A_Space%

  WinWaitClose, ahk_id %GuiHWND%  ;--waiting for gui to close
  return Customer               ;--returning value
;-------
ButtonOK:
  GuiControlGet, Customer
  Gui, CustGui:Destroy
return

Multiple:
  Customer := "MULTIPLE"
  Gui, CustGui:Destroy
return 

;-------
CustGuiGuiEscape:
CustGuiGuiClose:
  Customer = -1
  Gui, CustGui:Destroy
return 
  
}












;NewFunction
IssuePDF(ByRef CustGrid, FromName, IssueTo)
{
  SplashTextOn, 400, 200, Please Wait, Please wait...`n`nHave some digital patience...
FoxIt := "C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe"
StoredSerials := ""

GenerateXML:
FileDelete, %A_ScriptDir%\XML\*.xml

iniread, HRNum, %ini%, HR, HRNum, 0
HRNum += 1
IniWrite, %HRNum%, %ini%, HR, HRNum

HRNum := "00000" . HRNum
StringRight, HRNum, HRNum, 5
PG := "A"
AddList :=
gosub, XMLAddIncrease
%XMLAdd% .= "<FROM>" . FromName . "</FROM><TO>" . IssueTo . ":</TO><RECPTNR>" . HRNum . "</RECPTNR>"




gosub, XMLAddSerials
gosub, XMLAddNoSN
gosub, XMLAddIL

if (Row <> "")
  Tail := PG . "_" . Row
else
  Tail := PG
Field := "ITEMDES" . Tail
AddEndLine := "YES"
if (PG = "B" and Row = "")
  AddEndLine := "NO"
if(AddEndLine = "YES")
  %XMLAdd% .= "`<" . Field . "`>-------------------------------Nothing Follows-------------------------------`<`/" . Field . "`>"
 

Loop, parse, AddList, `n
{
  if(A_Loopfield = "")
    break
  TtlPgs := A_Index
}
if(PG = "B")
  TtlPgs += 1
if (PG = "B" and Row = 1)
  TtlPgs -= 1
Loop, parse, AddList, `n, `r
{
  if (A_Loopfield = "")
    Break
  PGNum := A_Index + 1
  if(A_Index = 1)
    %A_Loopfield% .= "`<PAGE`>1`</PAGE`>"
  else
    %A_Loopfield% .= "`<PAGE`>0`</PAGE`>"
  %A_Loopfield% .= "<OFPG>" . TtlPgs . "`</OFPG`>`<PAGEA`>" . PGNum . "`</PAGEA`><OFPGA>" . TtlPgs . "`</OFPGA`></form1>"
  Filenm := XML . A_Index . "`.xml"
  Append := %A_Loopfield%
  FileAppend, %Append%, %Filenm%      ;, UTF-8
  TtlXML := A_Index
}

IfNotExist, %FoxIt%
{
  Loop, parse, AddList, `n, `r
  {
    if (A_Loopfield = "")
      Break
    %A_Loopfield% := ""
  }
  AddList=
  CustGrid := ""
  SplashTextOff
  MsgBox, No FoxIt, You will have to manually import the XMLs from the XML Folder!
  Return
}

Gosub, ExportPDF


Loop, parse, AddList, `n, `r
{
  if (A_Loopfield = "")
    Break
  %A_Loopfield% := ""
}

iniwrite, %StoredSerials%, %ini%, IssuedSerials, %HRNum%

CustGrid := ""
Return

XMLAddSerials:
;~ Compile list of Serialized Items:
SerialItemList := CompileItemLists(CustGrid,SerialItemList,"SI")

if not SerialItemList
  Return

;~ Compile lists of Serialized, grouped by Item then sorted by Model:
ModSer := CompileSerialLists(CustGrid,SerialItemList,ModSer, "S")
;right now each line of ModSer looks like Item|Label|Serial


a := ""
loop, Parse, ModSer, `n, `r
{
  Loop, Parse, A_Loopfield, |
  {
    if (A_Index = 1)
      a .= RegExReplace(A_Loopfield,"[^\w]","") . ","
  }
}

SIssItemList := a
a=
sort, SIssItemList, U D`,

loop, parse, SIssItemList, csv
{
  if (A_Loopfield = "")
    Break
  %A_Loopfield% := ""
}

loop, Parse, ModSer, `n, `r
{
  if (A_Loopfield = "")
    Break
  Loop, Parse, A_Loopfield, |
  {    
    if (A_Index = 1)
      Item := A_LoopField
    if (A_Index = 2)
      Label := A_LoopField
    if (A_Index = 3)
      Serial := A_LoopField
  }
  ValVar := RegExReplace(Item,"[^\w]","")
  ValVarI := ValVar . "I"
  %ValVarI% := Item
  %ValVar% .= Label . "|" . Serial . "`n"
}


loop, parse, SIssItemList, csv
{
  qty := "Q" . A_Loopfield
  if (A_Loopfield ="")
    Break  
  Loop, Parse, %A_Loopfield%, `n, `r
  {
    if (A_Loopfield = "")
      Break
    %qty% := A_Index
  }
}


loop, parse, SIssItemList, csv
{
  if (A_Loopfield ="")
    Break
  StillGoing := "YES"
  ValVarI := A_Loopfield . "I"
  
  ItemDes := ""
  SNPerLine := 0
  RLen := 4
  SecLine := "NO"
  if (Row <> "")
    Tail := PG . "_" . Row
  else
    Tail := PG
  qty := "Q" . A_Loopfield
  StringReplace, R, %ValVarI%, `~, `,, All
  Field := "STOCKNR" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . R . "`<`/" . Field . "`>"
  Field := "QTYAUTH" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . %Qty% . "`<`/" . Field . "`>"
  Field := "QTYA" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . %Qty% . "`<`/" . Field . "`>"
  Field := "ITEMDES" . Tail
  %XMLAdd% .= "`<" . Field . "`>S/N:"
  NextLoop := %A_Loopfield%
  Loop, Parse, NextLoop, `n, `r
  {
    RunningQty := A_Index
    if (A_Loopfield = "")
      Break 
    QtyAdded := A_Index
    Loop, Parse, A_Loopfield, `|, `n
    {
      if (A_Loopfield = "")
      Break 
      if (A_Index = 1)
        Continue
      if (A_Index = 2)
      {
        AddSerial := A_Loopfield
        SNPerLine += 1
        NLen := StrLen(AddSerial)
        SLen := NLen + 1
        RLen += SLen
        if (RLen > 56 and SecLine = "NO")
        {
          %XMLAdd% .= "`n"
          SecLine := "YES"
          RLen := SLen
        }
        if (SecLine = "YES" and RLen > 56 or if (SNPerLine = 11))
        {
          if (Row <> "")
            Tail := PG . "_" . Row
          else
            Tail := PG
          Field := "ITEMDES" . Tail
          %XMLAdd% .= "`</" . Field . "`>"
          ;~ Row += 1
          ;~ if (PG = "A" and Row > 15)
          ;~ {
            ;~ PG := "B"
            ;~ Row := ""
          ;~ }
          ;~ if (PG = "B" and Row > 20)
            ;~ gosub, XMLAddIncrease
          
          gosub, RowIncrease
          if (Row <> "")
            Tail := PG . "_" . Row
          else
            Tail := PG
          Field := "ITEMDES" . Tail
          %XMLAdd% .= "`<" . Field . "`>"
          SecLine := "NO"
          RLen := SLen
          SNPerLine := 1
        }
        %XMLAdd% .= AddSerial
        StoredSerials .= "|" . AddSerial . "|"
        IF (QtyAdded <> %QTY%)
          %XMLAdd% .= "`,"
      }
    }
  }
  
  if (Row <> "")
    Tail := PG . "_" . Row
  else
    Tail := PG
  Field := "ITEMDES" . Tail
  %XMLAdd% .= "`</" . Field . "`>"
  if (RunningQty = %QTY%)
    StillGoing := "NO"
  gosub, RowIncrease
}
StillGoing := "NO"
SIssItemList := ""
SerialItemList := ""
ModSer := ""
Return

XMLAddNoSN:
;~ Compile List of Quantified (non-serialized) Items:
QItemList := CompileItemLists(CustGrid,QItemList,"QI")

;~ Compile List of Quantified & QTYs:
ModQty := CompileSerialLists(CustGrid,QItemList, ModQty, "Q")
;right now each line of ModQty looks like Item|Description|Qty


Loop, Parse, ModQty, `n, `r
{
  if (A_Loopfield = "")
    Break
  StringReplace, R, A_Loopfield, `~, `,, All
  Loop, parse, R, `|,
  {
    if (A_Loopfield = "")
      Break
    if (A_Index = 1)
      I := A_Loopfield
    if (A_Index = 2)
      D := A_Loopfield
    if (A_Index = 3)
      Q := A_Loopfield
  }
  if (Row <> "")
    Tail := PG . "_" . Row
  else
    Tail := PG
  Field := "STOCKNR" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . I . "`<`/" . Field . "`>"
  Field := "QTYAUTH" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . Q . "`<`/" . Field . "`>"
  Field := "QTYA" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . Q . "`<`/" . Field . "`>"
  Field := "ITEMDES" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . D . "`<`/" . Field . "`>"
  gosub, RowIncrease
}
Return


XMLAddIL:
;~ Compile Lists of Issuable Location Serials,  grouped by Model:
ModLbl := CompileSerialLists(CustGrid,"Issuable Location",ModLbl, "L")
;right now each line of ModLbl looks like Model|Label
if (ModLbl = "")
  Return

StockNrList := ""
loop, Parse, ModLbl, `n, `r
{
  if (A_Loopfield = "")
    Break
  Loop, Parse, A_Loopfield, |
  {
    if (A_Index = 1)
      StockNrList .= A_Loopfield . "`,"
  }
}

sort, StockNrList, U D`,


Loop, parse, StockNrList, csv
{
  ValVar := RegExReplace(A_Loopfield,"[^\w]","")
  ;~ StringReplace, ValVar, A_Loopfield, %A_Space%,,All
  ;~ StringReplace, ValVar, ValVar, `~,,All
  if (%ValVar% <> "")
    %ValVar% := ""
}

ValVars := ""

loop, Parse, ModLbl, `n, `r
{
  if (A_Loopfield = "")
    Break
  Loop, Parse, A_Loopfield, |
  {
    if (A_Index = 1)
      Model := A_Loopfield
    if (A_Index = 2)
      Label := A_Loopfield
  }
  StringReplace, ValVar, Model, %A_Space%,,All
  StringReplace, ValVar, ValVar, `~,,All
  ValVars .= ValVar . "`,"
  %ValVar% .= Label . "`,"
  
}


Loop, parse, ValVars, csv
  sort, A_Loopfield, D`,

loop, parse, ValVars, csv
{
  qty := "Q" . A_Loopfield
  if (A_Loopfield ="")
    Break  
  Loop, Parse, %A_Loopfield%, csv
  {
    if (A_Loopfield = "")
      Break
    %qty% := A_Index
  }

}


loop, parse, StockNrList, csv
{
  if (A_Loopfield ="")
    Break
  StillGoing := "YES"
  StringReplace, ValVar, A_Loopfield, %A_Space%,,All
  StringReplace, ValVar, ValVar, `~,,All
  ItemDes := ""
  LblPerLine := 0
  RLen := 0
  SecLine := "NO"
  if (Row <> "")
    Tail := PG . "_" . Row
  else
    Tail := PG
  qty := "Q" . ValVar
  StringReplace, R, A_Loopfield, `~, `,, All
  Field := "STOCKNR" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . R . "`<`/" . Field . "`>"
  Field := "QTYAUTH" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . %Qty% . "`<`/" . Field . "`>"
  Field := "QTYA" . Tail
  %XMLAdd% .= "`<" . Field . "`>" . %Qty% . "`<`/" . Field . "`>"
  Field := "ITEMDES" . Tail
  %XMLAdd% .= "`<" . Field . "`>"
  NextLoop := %ValVar%
  Loop, Parse, NextLoop, csv
  {
    if (A_Loopfield = "")
      Break 
    RunningQty := A_Index
    AddLabel := A_Loopfield
    LblPerLine += 1
    NLen := StrLen(AddLabel)
    SLen := NLen + 1
    RLen += SLen
    if (RLen > 56 and SecLine = "NO")
    {
      %XMLAdd% .= "`n"
      SecLine := "YES"
      RLen := SLen
    }
    if (SecLine = "YES" and RLen > 56 or if (LblPerLine = 11))
    {
      if (Row <> "")
        Tail := PG . "_" . Row
      else
        Tail := PG
      Field := "ITEMDES" . Tail
      %XMLAdd% .= "`</" . Field . "`>"
      ;~ Row += 1
      ;~ if (PG = "A" and Row > 15)
      ;~ {
        ;~ PG := "B"
        ;~ Row := ""
      ;~ }
      ;~ if (PG = "B" and Row > 20)
        ;~ gosub, XMLAddIncrease
      gosub, RowIncrease
      if (Row <> "")
        Tail := PG . "_" . Row
      else
        Tail := PG
      Field := "ITEMDES" . Tail
      %XMLAdd% .= "`<" . Field . "`>"
      SecLine := "NO"
      RLen := SLen
      LblPerLine := 1
    }
    %XMLAdd% .= AddLabel
    StoredSerials .= "|" . AddLabel . "|"
    if(A_Index <> %QTY%)
      %XMLAdd% .= "`,"
  }
  if (Row <> "")
    Tail := PG . "_" . Row
  else
    Tail := PG
  Field := "ITEMDES" . Tail
  %XMLAdd% .= "`</" . Field . "`>"
  if (RunningQty = %QTY%)
    StillGoing := "NO"
  gosub, RowIncrease
}
Loop, parse, StockNrList, csv
{
  ValVar := RegExReplace(A_Loopfield,"[^\w]","")
  ;~ StringReplace, ValVar, A_Loopfield, %A_Space%,,All
  ;~ StringReplace, ValVar, ValVar, `~,,All
  if (%ValVar% <> "")
    %ValVar% := ""
}
StillGoing := "NO"

StockNrList := ""
ModLbl := ""
Return


XMLAddIncrease:
Row := ""
XA += 1
XMLAdd := "XMLAdd" . XA
AddList .= XMLAdd . "`n"
;~ %XMLAdd% := "`<`?xml version`=""1.0"" encoding`=""UTF-8""`?`>`n`<form1 xmlns`:xfa`=""http`:`/`/www`.xfa`.org`/schema`/xfa`-data`/1`.0`/""`>"
%XMLAdd% := "`<`?xml version`=""1.0"" encoding`=""UTF-8""`?`>`n`<form1 xmlns`:xfa`=""http`:`/`/www`.xfa`.org`/schema`/xfa`-data`/1`.0`/""`>"
if (StillGoing = "YES")
{
  Field := "STOCKNR" . PG
  %XMLAdd% .= "`<" . Field . "`>" . R . "`<`/" . Field . "`>"
}
Return

RowIncrease:
Row += 1

if (PG = "A" and Row > 15)
{
  PG := "B"
  Row := ""
  if (StillGoing = "YES")
  {
    Field := "STOCKNR" . PG
    %XMLAdd% .= "`<" . Field . "`>" . R . "`<`/" . Field . "`>"
  }
}
if (PG = "B" and Row > 20)
  gosub, XMLAddIncrease

if (Row <> "")
  Tail := PG . "_" . Row
else
  Tail := PG
Return




ExportPDF:
FileDir := HRDir . "\" . HRNum . " - " . IssueTo
FileCreateDir, %FileDir%
SetTitleMatchMode, 2
Loop, %TtlXML%
{
  XMLImport := XML . A_Index . "`.xml"
  PDFName := HRNum . " - " . IssueTo
  if (A_Index > 1)
    PgNum := A_Index + 1
  else
    PgNum := A_Index
  if(TtlXMl > 1)
    PDFName .= " `. Pg" . PgNum
  PDFName .= ".pdf"      ;reenable
  PDFPath := FileDir . "\" . PDFName 
  run, "%FoxIt%" "%HRTemp%"
  WinWaitActive, a2062.pdf
  Sleep, 1000
  ;~ Send, !mu
  SendMessage, 0x111 , 46061, , , ahk_class classFoxitReader  ;Import Command
  ;~ WinWaitActive, Open
  SetKeyDelay, 50, 50
  ;~ Clipboard := XMLImport
  ;~ Sleep, 250
  ;~ ControlSend, Edit1, {CtrlDown}v{CtrlUp}, Open
  Win := WinExist()
  ;~ ControlSetText, Edit1, %XMLImport%, ahk_id %Win%
  ControlSetText, Edit1, %XMLImport%, Open
  Sleep, 500
  ;~ Win := WinExist()
  ControlSend, Button1, {Enter}, Open
  WinWaitClose, Open
  ;~ WinWaitClose, ahk_id %Win%
  Win := ""
  WinWaitActive, ahk_class #32770
  Sleep, 500
  WinGet, Win, ID, ahk_class #32770, OK
  ;~ Win := WinExist("Foxit Reader","OK")
  ;~ ControlSend, Button1, {Enter}, ahk_class #32770
  ControlSend, Button1, {Enter}, ahk_id %Win%
  ;~ ControlSend, Button1, {Enter}, FoxIt Reader
  ;~ WinWaitClose, ahk_class #32770
  WinWaitClose, ahk_id %Win%
  Sleep, 500
  ;~ ControlSend,,{CtrlDown}{ShiftDown}s{ShiftUp}{CtrlUp}, a2062.pdf
  SendMessage, 0x111 , 2264, , , ahk_class classFoxitReader ;SaveAs Command
  ;~ WinWaitActive, ahk_class #32770
  ;~ Clipboard := PDFPath
  ControlSend, Edit1, {Space}{Space}{Space}, ahk_class #32770
  ControlSetText, Edit1, %PDFPath%, ahk_class #32770
;~ pause           ;disable or delete
  ;~ ControlSend, Edit1, {CtrlDown}v{CtrlUp}, ahk_class #32770
  ;~ Sleep, 500
  ControlSend, Button3, {Enter}, ahk_class #32770
  WinWait, %PDFName%
  Sleep, 500
  Send, !{F4}
  WinWaitClose, %PDFName%,,1
  if errorlevel
  {
    IfWinActive, ahk_class #32770
    ControlSend, Button2, u, ahk_class #32770
    ;~ WinClose, %PDFName%
    ;~ WinKill, %PDFName%
    ;~ WinWaitClose, %PDFName%
  }
  Sleep, 1000
}
SetBatchLines, 10ms
Run, "%FileDir%"
Loop, %FileDir%\*.*
{
  run, "%FoxIt%" "%A_LoopFileLongPath%"
  Sleep, 500
}
Return



Return
}

ActiveControlIsOfClass(Class) {
    ControlGetFocus, FocusedControl, A
    ControlGet, FocusedControlHwnd, Hwnd,, %FocusedControl%, A
    WinGetClass, FocusedControlClass, ahk_id %FocusedControlHwnd%
    return (FocusedControlClass=Class)
}