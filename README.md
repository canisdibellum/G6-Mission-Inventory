# G6-Mission-Inventory
Inventory Program for G6, GLTD made in Autohotkey with Attachment to FoxIt PDF Reader and Excel  
    
Version: 1.0.1.0  
    
Version Breakdown:  
    Major UI or functionality change . Feature-Add . Feature-Fix/Bug-Fix . Code Cleanup  

ScanGUI Auto-Change Data:    
	L-[Location]    
	I-[Item]-[Model]    
	IL-[Location Type]-[Name/Label]    
    
    
Archived Filename Explanation:    
	Archived filenames look like this:    
	Inventory.20160603.234505.Startup.csv    
	Inventory.20160603.234506.Changed.csv    
	  
	Heres what it means:    
	The numbers are a time stamp that is exactly when backup was made, it breaks down as below    
	  
	Inventory.	2016	06	03.	23	45	05	  
	Inventory.	2016	06	03.	23	45	06	  
			Year	Month	Day	Hour	Minute	Second  
	  
	After that is Why it was automatically backed up:  
	.Startup.csv  
	.Changed.csv  
	  
	If this archive folder ever contains more than 100 backups, the oldest files will automatically be purged to keep the total file count at 100.  
	  
	If you want to change that value (keep more or less):  
	 - Open the info.ini in the folder above (\G6 Mission Inventory\AppData\info.ini)  
	 - Change the number value under [Archived Records to Keep]  
	  
	One small note on changing the value: the Inventory.csv used for demonstration had 840 Lines and was 51KB, when the folder had 100 of them, the folder's size was 5.08 MB.   
	Use that information as you will.  
    
    
    
    
To Do:  
    
 - To do  
 / Started  
 x Done  
 ? Not Sure, test  
 * Not Going to do, see [Note] at end  
 > in progress  
    
    
BugFixes: (Something acts like it isn't supposed to)  
	 - Bug: Scanning Registered Item number attempts to add Serial
		 - Make sure scanned data is analyzed in correct order
	 - Bug: Choose Customer DDL for Turn-In Broken
	 - Bug: ComboBoxes broken completely    
	 - bug: not clearing scanned serials from turn-in    
	 ? bug: not storing StoredSerials [I think I fixed it]    
	 ? Bug: When you scan registered Number, it keeps the value onto the next scan [Can't Reproduce at home....attempt with scanner]  
	 - bug: not looking at ScanGUI for most recent Label number  
	 - bug: During Verify Issue, when remove 1st Serial, it leaves white space at top  
    
Feature Fix: (Something works but it's kluge-y, I can implement better, or is missing certain functionality [ex. doesn't handle Zero or Negative Input])  
	 - Re-Do Build Inventory, use StringSplit.....make sure Quantified Items (if Serial is blank), don't get added to CBIList  
	 - Don't add Quantified Items to Item ComboBox!  
	 - Clarify Error Message Buttons on ScanGui Serial Error  
	 - Handle scanning "IL-" codes in Verify Issue  
	 ? Location/Model on scangui not updating [Seems to be working fine now]  
	 - Make Undo and Clear boxes refresh the SSList  
	 - Make sure Menu hides if not on Tab 1 and shows if on Tab 1  
	 / Make anything that triggers Filtration run FReset First, then [Guicontrol, text] the appropriate filter with the appropriate value then gosub, Filtration  
	 - Figure out how to auto-size filter DDL's to fit contents  
    
Feature-Add:  
     - Make spacebar check multi-selection *****IMPORTANT******
	 - Make Changes popup if other tab selected  
	 - make ComboBoxes instead be drop-downs with [Add...] option at bottom of list so they have to look through the list before typing something in.  
	 - Figure way to edit multiple lines at once in "Make Changes"  
	 - Make Undo Hotkey ^z, make return if tab <> 1 (See if it can be useful on other tabs)  
	 - Figure out better colored buttons  
	     - Check other gui styles, see if button color change possible  
	         - See if WinX style gui available  
	 - Look into ifexist C:\Program Files\Microsoft Office 15\root\office15\excel.exe (or detect version and if >2010), don't delete Sheets 2 & 3  
	 - When issuing, If user checks IL, check all items in location	[If too complicated, been directed to veto]  
	    - RowNumber := 0, RowNumber := Search(RowNumber...) Search for location and check  
	 - When scanning in Add to Inventory ScanGUI (and possibly Turn-In) make Splashtext with autochange (item labels [i-<>-<>]) and small tips and tricks (ex. type "undo" to undo and "clear" to clear  
	     - make typing "Undo" and "Clear" execute those functions  
    
Still To Do:  
	 / Test all controls make sure everything is Nominal [ScanGUI seems to be fine after private testing, real testing commences tomorrow]  
	 - Make sure all Inputs handle blank, 0, and Negative Entries  
	 - Clear Out old crap and set HRNum to 0  
	 - Write Up Documentation, to include commenting, especially Functions  
	 / Disable controls that don't work [I THINK I got em all]  
	 - Continue planning and then implement Menus  
   
    
   
______________________________________________________________________________  
    
Done: 

----------------------1.0.1.0----------------------  
 x bug: Export MAC exports everything, not just items w/ MAC [1.0.1.0]    
 x Bug: Continues to Hide MAC stuff even if MAC Present in Inventory [1.0.1.0]  
	 x Maybe because it says "VoIP Phone" instead of just VoIP? {Was looking at an item list IListC that didn't exist} [1.0.1.0]  
 x Bug: Still Exports MAC Address if nothing issued [1.0.1.0]      
 x Bug: Turn-In for Quantified Items populates field with "QTY" instead of quantity. [1.0.1.0]      
----------------------Pre-1.0.0.0----------------------  
 x Edit BuildInventory:  
   x get rid of anything that messes w/ filters  
   x gosub, FReset    [BuildInventory still has to build other LV's and dropdowns(/combo boxes)]  
     
 x give option for Full2062 to have multiple customers  
 x add button to gui for multiple customers then gosub to BuildInventory then gosub to gui for multiple customers and carry on from there.  
   x if customer list has more than one, InputBox and ask "Who is this being issued to?" for the 2062.  
    
 x Add line above filters.  
 x Add text (ex: Location, Label Set, Item, Model...)  
 x Make filters have alphabetical, unique list of contents of Category  
 x Make Label Names [–] then number (if that doesn't work find some other ASCII character that doesn't return as a hyphen)  
    
 * add Change sounds to  Edit Menu, for tabs 1 and 4 [Vetoed by Saint Martin]  
 * Item (STOCK NUMBER) [Veto...too vague]  
 * Make Label Correspond to Item NOT location(?)	[veto, too complicated]  
 * Expand ComboBox Dropdowns when clicked	[veto, too complicated]  
 x Bug: Issue broken....changes all Serial Numbers to 1  
 x Bug: Turn-In for power supplies cuts off first letter, screws everything up  
 x Fix Filters  
 x Bug: SplashText Disappears during ExportPDF  
 x Add Verify Issue function (menu select item or something.....that means keeping records! have like popup asking HR Number and have the Serial List Stored and ready to pop up in another scan gui of some method)  
 x date-time stamp Archived files   
 x Archive when starting up and when changing  
 x if more than 100 archived, delete oldest archive  
 x Make splashtext uniform  
 x SplashText on Immediate on Add to Inventory  
 x Cancel Changes needs to revert to non-Make Changes  
 x Make Changes Splashtext....Have Digital Patience  
 x Make Generate Complete DA 2062 work [Forced filtration is ready, just need exacts]  
 x Find Army Star to use for icon  
 x Change Menu and Gui Upper Left icon(s)  
 x Decide on Program Name  
 x Make Compiled .exe in outer folder to run script  
 x Put Hand Receipts folder in Parent Directory and change references  
 x Fix that only one Iteration of Serialized issues  
 x Make Issue Export Excel to Show Labels  
 x Add "Please Wait" SplashText When Issuing  
 x Add "Please Wait" SplashText When Turning In  
 x Add Item to StockNR[?] on next page  
 x Script Importing into FoxIt  
 x Handle other PDFs already being open in FoxIt  
 x After Turn-In, Rebuild CustomerSelect and Choose "1"  
 x Make Export delete Sheet2 and Sheet3  
 x Add All Inventory and test Multi-Page Issue  
 x Make CustomerGui Function  
 x If there is a VoIP Item, then Export should also include the MAC Column  
 x If there is a Quantified Item, then Export should also include QTY  
 x Tweak GenerateXML to be able to be used by other subs  
 x Turn GenerateXML into a function  
 x Make Export MAC Sheet work  
 x Handle "Add to Inventory" with empty grid  
 x Make Add "Issuable Location":  
	 x Button  
	 x GUI[Current (with CustGui) won't work, needs to be able to define Location (label?)  
	 x subroutine  
 x Add Sounds to Turn-In ScanGUI  
