Build_Menus()
{
Menu, OHIEdit, Add, Undo Last Inventory Change`tCTRL+ALT+Z, UndoChange
  Menu, AllFile, Add, Backup to Archive`tCTRL+F12, Backup
  Menu, OHIEdit, Add, Select All`tCTRL+A, OHSelAll
  Menu, ATIMenu, Add, &File, :AllFile
  Menu, AllMenu, Add, &Edit, :OHIEdit
  Menu, OHIEdit, Disable, Select All`tCTRL+A, 
  return
}
