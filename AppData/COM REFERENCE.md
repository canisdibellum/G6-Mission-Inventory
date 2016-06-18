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
```
Xl.Range("A1").Value := helloworld
```



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
