BuildMenus()
{
  Menu, AllFile, Add, Import Inventory from Excel Spreadsheet, ImportXLSX
  Menu, AllFile, Add, Backup to Archive`tCTRL+F12, Backup
  Menu, AllFile, Add, Restore Archive...`tCTRL+F11, Restore
  Menu, AllFile, Add, Clear Inventory, ClearInventory
  Menu, AllFile, Add, 
  Menu, AllFile, Add, Preferences...`tCTRL+F5, Prefs
  Menu, AllFile, Add, Print Labels...`tCTRL+SHIFT+P, PrintLbls
  Menu, AllFile, Add, Print...`tCTRL+P, Print
  Menu, AllFile, Add, Exit

  Menu, AllTools, Add, Export to Excel…`tCTRL+T, ExpToXLGUI
  Menu, AllTools, Add, Generate MAC List...`tCTRL+M, ExpMAC
  Menu, AllTools, Add, Generate Hand Receipt...`tCTRL+H, Full2062
  Menu, AllTools, Add, 
  Menu, AllTools, Add, Verify Items for Issue..., Verify
  Menu, AllTools, Add, Deadline Items..., Deadline
  Menu, AllTools, Add, 
  Menu, AllTools, Add, Generate Number Reports...`tCTRL+F7, RunReports

  Menu, ATIEdit, Add, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, UndoChange
  Menu, ATIEdit, Add, 
  Menu, ATIEdit, Add, Undo Last Queue Addition`tCTRL+Z, undo
  Menu, ATIEdit, Add, Clear Queue, clear

  Menu, ATIAdd, Add, Register Item Code...`tCTRL+F2, RegisterItem
  Menu, ATIAdd, Add, 
  Menu, ATIAdd, Add, Add Item Without Serial Number`tCTRL+F3, NoSNAddGui
  Menu, ATIAdd, Add, Add Issuable Location`tCTRL+F4, AddIL

  Menu, FIEdit, Add, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, UndoChange
  Menu, FIEdit, Add, 
  Menu, FIEdit, Add, Make Changes, MakeChanges
  Menu, FIEdit, Add, Apply Changes, ApplyChanges
  Menu, FIEdit, Add, Cancel Making Changes`tEsc, CancelChanges

  Menu, OHIEdit, Add, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, UndoChange
  Menu, OHIEdit, Add, 
  Menu, OHIEdit, Add, Select All`tCTRL+A, OHSelAll
  Menu, OHIEdit, Add, Select None, OHSelN

  Menu, TIEdit, Add, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, UndoChange
  Menu, TIEdit, Add, 
  Menu, TIEdit, Add, Select All`tCTRL+A, TISelAll
  Menu, TIEdit, Add, Clear Scanned Serials and Selections, TIClear

  Menu, HelpMenu, Add, Guide`tCTRL+F1, Guide
  Menu, HelpMenu, Add, 
  Menu, HelpMenu, Add, Changelog
  Menu, HelpMenu, Add, 
  Menu, HelpMenu, Add, About

  Menu, ATIMenu, Add, &File, :AllFile
  Menu, ATIMenu, Add, &Tools, :AllTools
  Menu, ATIMenu, Add, &Edit, :ATIEdit
  Menu, ATIMenu, Add, &Add…, :ATIAdd
  Menu, ATIMenu, Add, &Help, :HelpMenu
  Menu, FIMenu, Add, &File, :AllFile
  Menu, FIMenu, Add, &Tools, :AllTools
  Menu, FIMenu, Add, &Edit, :FIEdit
  Menu, FIMenu, Add, &Help, :HelpMenu
  Menu, OHIMenu, Add, &File, :AllFile
  Menu, OHIMenu, Add, &Tools, :AllTools
  Menu, OHIMenu, Add, &Edit, :OHIEdit
  Menu, OHIMenu, Add, &Help, :HelpMenu
  Menu, TIMenu, Add, &File, :AllFile
  Menu, TIMenu, Add, &Tools, :AllTools
  Menu, TIMenu, Add, &Edit, :TIEdit
  Menu, TIMenu, Add, &Help, :HelpMenu
  
  ;~ Disable Items until they are enabled:
  Menu, FIEdit, Disable, Make Changes
  Menu, FIEdit, Disable, Apply Changes
  
  Menu, ATIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, FIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, OHIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, TIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
 

  ;~ Disable Non-Working Menu Items:
  Menu, AllFile, Disable, Import Inventory from Excel Spreadsheet
  Menu, AllFile, Disable, Backup to Archive`tCTRL+F12, 
  Menu, AllFile, Disable, Restore Archive...`tCTRL+F11, 
  Menu, AllFile, Disable, Clear Inventory
  Menu, AllFile, Disable, Preferences...`tCTRL+F5, 
  Menu, AllFile, Disable, Print Labels...`tCTRL+SHIFT+P, 
  Menu, AllFile, Disable, Print...`tCTRL+P, 
  Menu, AllTools, Disable, Deadline Items...
  Menu, AllTools, Disable, Generate Number Reports...`tCTRL+F7, 
  Menu, ATIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, FIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, OHIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, OHIEdit, Disable, Select All`tCTRL+A, 
  Menu, OHIEdit, Disable, Select None
  Menu, TIEdit, Disable, Undo Last Inventory Change`tCTRL+ALT+SHIFT+Z, 
  Menu, TIEdit, Disable, Select All`tCTRL+A, 
  Menu, HelpMenu, Disable, Guide`tCTRL+F1, 
  Menu, HelpMenu, Disable, Changelog
  Menu, HelpMenu, Disable, About
  
  Return
}