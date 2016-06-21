<<<<<<< HEAD
#  List of all COM commands I find  
(A lot come from tutorial posted at [http://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/](http://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/ "Basic AHK_L COM Tutorial for Excel"))
## Start Here ##
### Grab existing Excel instance
careful, tricky if you have multiple instances  

    Xl := ComObjActive("Excel.Application")  
>gets a handle to the Application object, not the sheet object. Also, how do you define active? Actually, it would be the first excel application object/process registered on the Running Object Table. That's why I use the following:  

    ControlGet, hwnd, hwnd, , Excel71, ahk_class XLMAIN
    window := Acc_ObjectFromWindow(hwnd, -16)
    xlBook := window.parent
    xlApp := window.application
    xlSheet := window.activesheet  
    
>However, this will cause a Com Error if a cell is being modified       

### To create new instance of Excel  
    Xl := ComObjCreate("Excel.Application")  
        
## Open or Create New Worksheets ##
### Open Worksheet ###
can use variable or "C:\Directory\File.xlsx"  

    Xl.Workbooks.Open(Path)  
    
### Create New Worksheet ###
Creates new spreadsheet file  

    Xl.Workbooks.Add
    
### Make Visible:
Can be put in at any time, but this is set to false by default  
`Xl.Visible := True`  

## Input text ##
### set cell 'A1' to a string  
    Xl.Range("A1").Value := "hello world!
    
### set cell to a variable  

    helloworld := "hello world!"
    Xl.Range("A1").Value := helloworld
    
### loop through the cells in a column

    while (Xl.Range("A" . A_Index).Value != "") 
    {  
    	Xl.Range("A" . A_Index).Value := value 
    }  

### loop through a row  
Original Post  

    Row := "1" 
    Columns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q") ;array of column letters 
    For Key, Value In Columns 
    XL.Range(Value . Row).Value := value ;set values of each cell in a row
Less tedious method Using ASCII Characters to build "columns" array  

    Row := 1
    Columns := []
    loop 17		;Only through "Q" like other example
    {
      AscCode := A_Index + 64 ;ASCII Character for "A" is 65
      Letter := Chr(AscCode)
      Columns.Push(Letter)  ;adds letter to array with index = A_Index, so Columns[1] = "A"
    }
    For Key, Value In Columns
    XL.Range(Value . Row).Value := value ;set values of each cell in a row
        
### Get Value from cell ###
Sets cell value to variable  

    Xl.Range("A1").Value := helloworld
        
### Save Methods ###
Quick-save  

    Xl_Workbook.Save() ;quick save already existing file
    
Save As  

    XL.ActiveWorkbook.SaveAs(BookName) ;'bookname' is a variable with the path and name of the file you desire

2nd Save As Method  

    Xl_Workbook := Xl.Workbooks.Open(Path) ;handle to specific workbook 
    Xl_Workbook.Save() ;quick save already existing file
_____
    
### Miscellaneous ###
Copy cell to clipboard  

    Xl.Range("A:A").Copy
    
Paste as values  

    Xl.Range("A:A").PasteSpecial(-4163) ;'-4163' is the constant for values only
    
*I'm assuming Normal paste is  

    Xl.Range("A:A").Paste

Change the column number format to 'text'  

    Xl.Range("A:A").NumberFormat := "@"
    
Deselect cells  

    Xl.CutCopyMode := False ;use this with copy
    
Sort sheet by column  

    Xl.Range("A1:Q100").Sort(Xl.Columns(1), 1) ;sort sheet by data in the 'a' column
    
Possible Cell Loading method (taken from forum, not tested)  

    Columns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q") ;array of column letters 
    Columns := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"] ;array of column letters
    Columns := SplitStr("ABCDEFGHIJKLMNOPQ") ;array of column letters
    
    
    SplitStr(str, delim="", omit="") {
    	array := []
    	Loop, Parse, str, %delim%, %omit%
    		array.insert(A_LoopField)
    	return, array
    }



                           
--------------------------
To delete extra sheets from Excel Export  

    Count := Xl.Worksheets().Count
    if (Count > 1)
    {
	    Loop, %Count%
	    {
	      if (A_Index = 1)
	    Continue
	      ws := "Sheet" . A_Index
	      Xl.Worksheets(ws).Delete()
	    }
    }
    

To print excel and keep (somewhat) hidden  

    Xl := ComObjCreate("Excel.Application")
    Xl.Workbooks.Add			;add a new workbook
    Xl.Range("A1").Value := "Some Text"
    XL.Application.Dialogs(8).Show
    Xl.Visible := False
    XL.ActiveWorkbook.Saved := True
    XL.ActiveWorkbook.Close()
=======
#  List of all COM commands I find  
(A lot come from tutorial posted at [http://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/](http://autohotkey.com/board/topic/69033-basic-ahk-l-com-tutorial-for-excel/ "Basic AHK_L COM Tutorial for Excel"))
## Start Here ##
### Grab existing Excel instance
careful, tricky if you have multiple instances  
`Xl := ComObjActive("Excel.Application")`  
    

### To create new instance of Excel
`Xl := ComObjCreate("Excel.Application")`  
    
## Open or Create New Worksheets ##
### Open Worksheet ###
can use variable or "C:\Directory\File.xlsx"  
`Xl.Workbooks.Open(Path)`  

### Create New Worksheet ###
Creates new spreadsheet file  
`Xl.Workbooks.Add`

### Make Visible:
Can be put in at any time, but this is set to false by default  
`Xl.Visible := True`  

## Input text ##
### set cell 'A1' to a string  
    Xl.Range("A1").Value := "hello world!
    
### set cell to a variable  

    helloworld := "hello world!"
    Xl.Range("A1").Value := helloworld
    
### loop through the cells in a column

    while (Xl.Range("A" . A_Index).Value != "") 
    {  
    	Xl.Range("A" . A_Index).Value := value 
    }  

### loop through a row  
    Row := "1" 
    Columns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q") ;array of column letters 
    For Key, Value In Columns 
    XL.Range(Value . Row).Value := value ;set values of each cell in a row
    
### Get Value from cell ###
Sets cell value to variable  
    Xl.Range("A1").Value := helloworld
        



To delete extra sheets from Excel Export:
[CODE]
Count := Xl.Worksheets().Count
if (Count > 1)
{
    Loop, %Count%
    {
      if (A_Index = 1)
        Continue
      ws := "Sheet" . A_Index
      Xl.Worksheets(ws).Delete()
    }
}
[/CODE]

To print excel and keep (somewhat) hidden:
Xl := ComObjCreate("Excel.Application")
Xl.Workbooks.Add                            ;add a new workbook
Xl.Range("A1").Value := "Some Text"
XL.Application.Dialogs(8).Show
Xl.Visible := False
XL.ActiveWorkbook.Saved := True
XL.ActiveWorkbook.Close()

  
