# G6-Mission-Inventory
Inventory Program for G6, GLTD made in Autohotkey with Attachment to FoxIt PDF Reader and Excel  
    
# Version: 2.2.1.0  
    
## Version Breakdown:  
    [Major UI or functionality change] . [Feature-Add] . [Feature-Fix/Bug-Fix] . [Code Cleanup]  

## ScanGUI Auto-Change Data:    
	L-[Location]    
	I-[Item]-[Model]    
	IL-[Location Type]-[Name/Label]    
    
    
## Archived Filename Explanation:    
	Archived filenames look like this:    
	Inventory.20160603.234505.Startup.csv    
	Inventory.20160603.234506.Changed.csv    
	  
	Heres what it means:    
	The numbers are a time stamp that is exactly when backup was made, it breaks down as below    
	  
	Inventory.	2016	06		03.		23		45		05	  
	Inventory.	2016	06		03.		23		45		06	  
				Year	Month	Day		Hour	Minute	Second  
	  
	After that is Why it was automatically backed up:  
	.Startup.csv  
	.Changed.csv  
	  
	If this archive folder ever contains more than 100 backups, the oldest files will automatically be purged to keep the total file count at 100.  
	  
	If you want to change that value (keep more or less):  
	 - Open the info.ini in the folder above (\G6 Mission Inventory\AppData\info.ini)  
	 - Change the number value under [Archived Records to Keep]  
	  
	One small note on changing the value: the Inventory.csv used for demonstration had 840 Lines and was 51KB, when the folder had 100 of them, the folder's size was 5.08 MB.   
	Use that information as you will.  
    
    
    
    
# To Do:  
    
 [ ] To do  
 [/] Started  
 [x] Done  
 [?] Not Sure, test  
  X  Not Going to do, see [Note] at end  
 [>] in progress  

##PRIORITY: (To be done ASAP)
	 [ ] parse through ideas file from OneCloud, add in to To-Do lists as appropriate
	 [ ] Make sure everything important made it to this Priority List
	 [ ] Fix Major AInventory (Add to Inventory) Issues:
		 [ ] Some labels are still showing dashes!  
		 [ ] Fix label Auto-Numbering  
		 [ ] Queue is not considered when looking for duplicate serials (SSList) [Add to as part of scan procedure?]  
			[ ] build ScanCheck Lists: (Serials, Items, Locations) 
				[ ] Build list of all Serials, Items, and Locations in Inventory when BuildInventory run [I think it already does] 
				[ ] As soon as scan button is hit, Build Serial, Item, and Location lists from Current Queue [use or update and then use already existing functions to facilitate 
				[ ] append them to already existing lists,   
				[ ] look into building lists as arrays and what happens when you try to retrieve the key for a value that doesn't exist,   (checking array for dups: https://autohotkey.com/board/topic/68613-add-variables-to-array-and-check-for-duplicates/)
				[ ] When scanning, add to lists  
				[ ] When exiting scan routine, sort, U Location and Item Lists, and update comboboxes  
		 [ ] Queue is not considered when looking for highest Label Number  
		 [ ] Keeps trying to re-add IL scans instead of just setting the location  
		 [ ] Adding quantified items to queue does not update the comboboxes in the "Add Item Without Serial Number" GUI  
	 [ ] In prompt, blank Model if it says "Select Item First"  
	 [ ] "Model" manual combobox doesn't populate  
 	 [ ] when Scan InputBox terminates, check for MAC, if none found make sure MAC Column is hidden   

		 
		 
	 [ ] Determine which items need to be researched  
	 [ ] Anything else that does not need to be researched  
	 [ ] Bugfixes  
	 [ ] Feature Fixes  
	 [ ] Make 100% Stable 
	 [ ] Add Features
    
## Bug List: (Something acts like it isn't supposed to)  
	Add to Inventory: (ScanGui)  
	 [ ] "Add IL" GUI Controls too short for contained text  
	 [ ] Casing for "Item" needs to be handled
	
	Inventory:  
	 [ ] "Export to Excel" not looking for MAC Addresses  
	 [ ] Still tries to generate Full 2062 when nothing issued  
	 [ ] BuildInventory needs to LV_Delete OHInventory before adding  
 	 [ ] bug: Single-Cell Edits for bottom 2 Rows are wonky  
	 [ ] Figure out how to auto-size filter DDL's to fit contents    
	 [ ] Make Changes needs full inventory before enabling changes    
		[ ] See if that means FReset & BuildInventory or one or the other    
	 [ ] Edit multiple needs to runthrough and autohdr whatever columns got changed  
	
	On-Hand Inventory:  
	 [ ] When IL is checked, previously checked items in that location have QTY set to 0  
	 [ ] Issue Splashscreen delayed, make come up faster  
	 [ ] Issue Excel export doesn't include MAC Addresses  
		Verify Issue:  
		 [ ] when remove 1st Serial, it leaves white space at top  
		 [ ] Needs to handle scanning same serial twice  
		 [ ] Needs to handle scanning IL tag as serial  
		 [ ] When all clear some sort of message would be satisfying  
		 [ ] Put listbox like TIInventory that lists already scanned items  

	Turn-In Inventory:   
		 [ ] Reset # of scanned Items label  
	 

## Feature Fix: (Something works but it's kluge-y, I can implement better, or is missing certain functionality [ex. doesn't handle Zero or Negative Input])  
	 [ ] Clarify Error Message Buttons on ScanGui Serial Error  
	 [ ] Make sure Menu hides if not on Tab 1 and shows if on Tab 1  
	 [/] Make anything that triggers Filtration run FReset First, then [Guicontrol, text] the appropriate filter with the appropriate value then gosub, Filtration  
	 [ ] Find and fix all instances where main gui isn't activated when a 2nd GUI window closes    
	 [ ] Get Rid of Blank space at the top of CustomerSelect DDL  
	 [ ] Get Rid all R and H options on DDLs and ComboBoxes (don't forget to check Functions.ahk)  
    
## Feature-Add:  
	 [ ] make ComboBoxes instead be drop-downs with [Add...] option at bottom of list so they have to look through the list before typing something in.  
	 [ ] Make Undo Hotkey ^z, make return if tab <> 1 (See if it can be useful on other tabs)  
	 [ ] Figure out better colored buttons  
	     [ ] Check other gui styles, see if button color change possible  
	         [ ] See if WinX style gui available  
	 [ ] Look into ifexist C:\Program Files\Microsoft Office 15\root\office15\excel.exe (or detect version and if >2010), don't delete Sheets 2 & 3  
	 [ ] When scanning in Add to Inventory ScanGUI (and possibly Turn-In) make Splashtext with autochange (item labels [i-<>-<>]) and small tips and tricks (ex. type "undo" to undo and "clear" to clear)  
	     [ ] make typing "Undo" and "Clear" execute those functions  
	 [ ] Run reports Menu Item....Pop up Msg Box with all stats and option to export to text file.
	 [ ] Make "Set Special Status" Menu Item for Full Inventory [Currently possible with edit multitiple]  
		[ ] for things like "Deadline" or "In Use"   
	 [ ] Add "Duplicate Row" to MakeChanges Context Menu  
	 [ ] Add ability to Copy Cell to OHInventory   
    
## General To Do:  
	 [ ] Test all controls make sure everything is Nominal
	 [ ] Make sure all Inputs handle blank, 0, and Negative Entries  
	 [ ] Clear Out old stuff and set HRNum to 0  
	 [ ] Write Up Documentation, to include commenting, especially Functions  
	 [ ] Continue planning and then implement Menus  
	 [ ] Comb for unused/old Variables and Subroutines  
	 [?] During make changes if you have something selected and then search for something, it adds found row to selection...can't decide if thats a bug or a feature  
   
    
   
______________________________________________________________________________  
    
# Done: 
    ----------------------2.2.1.0----------------------  
	 [x] New Menu subroutines already added to bottom of script  
	 [x] Test if Capital Letters in Hotkeys matter (doesn't matter)  
	 [x] Test if F-keys work without CTRL (They Don't)  
	 [x] Test if hotkeys work if menu not showing (They Do)  
	 [x] Test building folder in function (Works just fine)  
	 [x] Test disabling	(needs to be ENTIRE display name including hotkey)  
			- Syntax: Menu, [MenuName], Disable, [Entire DisplayName including hotkeys]  
	 [x] Test disabling entire menu ("&" matters)  
	 [x] Test if F-keys work with CTRL (They do)  
	 [x] Get working Proof of concept, display:  
		 [x] Build all menus in one #Included function  
		 [x] Changing using Array  
		 [x] Using and disabling Keyboard Accelerators  
		 [x] Disabling within build function  

    ----------------------2.1.1.0----------------------  
	Priority:	
	 [x] Rework CustomerSelect DDL to fill TIInventory Independently of Full Inventory  
		[x] Make sure MAC Column doesn't show if empty  
	 [x] Move CustomerSelect DDL down and add text above "Choose Customer"  
		[x] Handle all instances of Script looking for "Choose Customer" as selection  
	 [x] Move Scanned Serials ListBox down and add text above "Serials Scanned"  
		[x] Handle all instances of Script looking for "Serials Scanned" as selection  
	 [x] QI Turn-in needs to run BuildInventory (and possibly FReset) before making changes!!  (wiped out my inventory because I had a filter set)  
	On-Hand Inventory:  
	 [x] OHInventory was duplicating if it showed last item in csv  
		Verify Issue:  
		 [x] Needs sounds to indicate good scan  
	Turn-In Inventory:   
	 [x] Make CustomerSelect disable Controls first
	 [x] at the end have TINonSN store ScannedSerials in OldScanned then
	 [x] run customerselect 
		 [x] GuiControl, , Scanned, |%OldScanned%
		 [x] if OldScanned = ""; Don't Repopulate Check Marks
		 [x] Repopulate Check Marks:
	 [x] For Turning-In Quantified items: instead of running through the csv twice and then looking up each qty before building gui to turn in, store qty in array while building TIInventory then just call the value for the gui   
	 [x] TI Quantified Items doesn't remove them from the listview  
    ----------------------2.0.1.0----------------------   
	 Major Functionality Change:   
	 [x] Re-Do Build Inventory, use StringSplit.....make sure Quantified Items (if Serial is blank), don't get added to CBIList [2.0.1.0]  
	 BugFixes/FeatureFixes:   
	 [x] Bug: Scanning Registered Item number attempts to add Serial  
		 [x] Make sure scanned data is analyzed in correct order  
	 [x] Bug: Choose Customer DDL for Turn-In Broken  
	 [x] Bug: ComboBoxes broken completely    
	 [x] bug: not storing StoredSerials [I think I fixed it]    
	 [x] Bug: When you scan registered Number, it keeps the value onto the next scan [Can't Reproduce at home....attempt with scanner]  
	 [x] Location/Model on scangui not updating [Seems to be working fine now]  
	 [x] Don't add Quantified Items to Item ComboBox!  
	 [x] Fix: ScanGUIAdd Functionality completely broken  
	 [x] Test all controls make sure everything is Nominal
    ----------------------1.4.2.0----------------------   
	 [x] bug: not clearing scanned serials from turn-in [1.4.2.0] 
	 [x] Edit Individual cells [1.4.2.0] 
	     [x] Made Gui, made it popup....get better positioning...find col width, find upper left corner of cell [1.4.2.0]    
    ----------------------1.3.1.0----------------------   
	 [x] When issuing, If user checks IL, check all items in location [1.3.1.0]	 
	    [x] RowNumber := 0, RowNumber := Search(RowNumber...) Search for location and check[1.3.1.0]    
	 [x] Make Changes popup if other tab selected [1.3.1.0]	 
    ----------------------1.2.1.0----------------------  
	 [x] Figure way to edit multiple lines at once in "Make Changes" [1.2.1.0]  
    ----------------------1.1.1.0----------------------  
	 [x] Fix having to right-click twice to edit  [1.1.1.0]  
	 [x] Make spacebar check multi-selection *****IMPORTANT******  [1.1.1.0]  
    ----------------------1.0.1.0----------------------   
     [x] bug: Export MAC exports everything, not just items w/ MAC [1.0.1.0]    
     [x] Bug: Continues to Hide MAC stuff even if MAC Present in Inventory [1.0.1.0]  
    	 [x] Maybe because it says "VoIP Phone" instead of just VoIP? {Was looking at an item list IListC that didn't exist} [1.0.1.0]  
     [x] Bug: Still Exports MAC Address if nothing issued [1.0.1.0]      
     [x] Bug: Turn-In for Quantified Items populates field with "QTY" instead of quantity. [1.0.1.0]      
    ----------------------Pre-1.0.0.0----------------------  
	 [x] Disable controls that don't work  
     [x] Edit BuildInventory:  
       [x] get rid of anything that messes w/ filters  
       [x] gosub, FReset    [BuildInventory still has to build other LV's and dropdowns(/combo boxes)]  
         
     [x] give option for Full2062 to have multiple customers  
     [x] add button to gui for multiple customers then gosub to BuildInventory then gosub to gui for multiple customers and carry on from there.  
       [x] if customer list has more than one, InputBox and ask "Who is this being issued to?" for the 2062.  
        
     [x] Add line above filters.  
     [x] Add text (ex: Location, Label Set, Item, Model...)  
     [x] Make filters have alphabetical, unique list of contents of Category  
     [x] Make Label Names [–] then number (if that doesn't work find some other ASCII character that doesn't return as a hyphen)  
        
     X add Change sounds to  Edit Menu, for tabs 1 and 4 [Vetoed by Saint Martin]  
     X Item (STOCK NUMBER) [Veto...too vague]  
     X Make Label Correspond to Item NOT location(?)	[veto, too complicated]  
     X Expand ComboBox Dropdowns when clicked	[veto, too complicated]  
     [x] Bug: Issue broken....changes all Serial Numbers to 1  
     [x] Bug: Turn-In for power supplies cuts off first letter, screws everything up  
     [x] Fix Filters  
     [x] Bug: SplashText Disappears during ExportPDF  
     [x] Add Verify Issue function (menu select item or something.....that means keeping records! have like popup asking HR Number and have the Serial List Stored and ready to pop up in another scan gui of some method)  
     [x] date-time stamp Archived files   
     [x] Archive when starting up and when changing  
     [x] if more than 100 archived, delete oldest archive  
     [x] Make splashtext uniform  
     [x] SplashText on Immediate on Add to Inventory  
     [x] Cancel Changes needs to revert to non-Make Changes  
     [x] Make Changes Splashtext....Have Digital Patience  
     [x] Make Generate Complete DA 2062 work [Forced filtration is ready, just need exacts]  
     [x] Find Army Star to use for icon  
     [x] Change Menu and Gui Upper Left icon(s)  
     [x] Decide on Program Name  
     [x] Make Compiled .exe in outer folder to run script  
     [x] Put Hand Receipts folder in Parent Directory and change references  
     [x] Fix that only one Iteration of Serialized issues  
     [x] Make Issue Export Excel to Show Labels  
     [x] Add "Please Wait" SplashText When Issuing  
     [x] Add "Please Wait" SplashText When Turning In  
     [x] Add Item to StockNR[?] on next page  
     [x] Script Importing into FoxIt  
     [x] Handle other PDFs already being open in FoxIt  
     [x] After Turn-In, Rebuild CustomerSelect and Choose "1"  
     [x] Make Export delete Sheet2 and Sheet3  
     [x] Add All Inventory and test Multi-Page Issue  
     [x] Make CustomerGui Function  
     [x] If there is a VoIP Item, then Export should also include the MAC Column  
     [x] If there is a Quantified Item, then Export should also include QTY  
     [x] Tweak GenerateXML to be able to be used by other subs  
     [x] Turn GenerateXML into a function  
     [x] Make Export MAC Sheet work  
     [x] Handle "Add to Inventory" with empty grid  
     [x] Make Add "Issuable Location":  
    	 [x] Button  
    	 [x] GUI[Current (with CustGui) won't work, needs to be able to define Location (label?)  
    	 [x] subroutine  
     [x] Add Sounds to Turn-In ScanGUI  
