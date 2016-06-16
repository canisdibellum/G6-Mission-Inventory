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
;~ StringReplace, IListC, IListC, |, `,, ALL
IfNotInString, IList, VoIP
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