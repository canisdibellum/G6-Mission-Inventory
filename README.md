# G6-Mission-Inventory
Inventory Program for G6, GLTD made in Autohotkey with Attachment to FoxIt PDF Reader and Excel

Version: 1.0.0.0

 - To do
 / Started
 x Done
 ? Not Sure, test
 * Not Going to do, see [Note] at end
 > in progress

To Do:

BugFixes: (Something acts like it isn't supposed to)
 - Bug: ComboBoxes broken completely
 - Bug: Still Exports MAC Address if nothing issued
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

Feature-Add:
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

Still To Do:
 / Test all controls make sure everything is Nominal [ScanGUI seems to be fine after private testing, real testing commences tomorrow]
 - Make sure all Inputs handle blank, 0, and Negative Entries
 - Clear Out old crap and set HRNum to 0
 - Write Up Documentation, to include commenting, especially Functions
 / Disable controls that don't work [I THINK I got em all]
 - Continue planning and then implement Menus
 

 
----------------------------------------------------------

Done:
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
