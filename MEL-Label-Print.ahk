;MEL-Location-Print.ahk	An AutoHotKey script for use by Circ to print 
;			a set of MeL return labels using a Word template.
;
; 			Mod 00, 06/11/2021, new script to auto-complete
;				and print MeL labels.
;			Mod 01, 06/23/2021, user requested tweaks
;			Mod 02, 06/28/2021, user requested tweaks, part 2
;				add a notes field and modify max branch to 24.
;			Mod 03, 07/07/2021, typo fixed in the branch table display, 
;				and added bolding to the branch name for printing.
;			Mod 04, 08/30/2021, some blank lines removed to shorten. 
;
;			Always Trim Spaces and Tabs.
;			Match Anywhere in Window Title.
;			Create CTRL-X CTRL-x as hotkeys to stop script.
;
AUTOTRIM, ON
SETTITLEMATCHMODE, 2
#WinActivateForce
;			Reminder on what the Send key modifiers are:
;			   ^   before char is   control-char
;			   !   before char is   alt-char
;			   +   before char is   shift-char
;
;			Name the Table Arrays used and initialize varables.
LocTable 		:= []
DupLocs		:= []
LocationCode	:= ZZ000
LocName		:= ""
; 			Read the MEL-Locations.csv file by row.
;			This .csv file was created by extracting the library
;			locations from the MeL locations web page and
;			editing into a .csv file.
;
;			For each row (location), parse each field into a
;			cell of the 2D array LocTable
;				each row is a MeL location
;				each with 4 fields; 	1st 3 are location ID
;					     	4th is location name
;
;			Converts each field into all upper case.
;			The last row number # will be held in LocationIndex.
;
;			A_Index is auto defined, # times the loop is executed.
;			Each Loop, Read will bring the next data record into
;			varable A_LoopReadLine.
;
;			Loop Parse has the designation of CSV, so it 
;			understands its working with a .CSV data and it
;			breaks the record into fields one at a time, placing
;			the content of each field in A_LoopField.
;
;			Each field is stored into the table matrix LocTable.
;	
UrlDownloadToFile, https://portal.cantonpl.org/docs/_layouts/15/DocIdRedir.aspx?ID=CANTONPL-34-32396, MEL-Locations.csv		
Loop, Read, MEL-Locations.csv
{
    LocationIndex := A_Index
;
    Loop, Parse, A_LoopReadLine, CSV
{
	StringUpper, LocationText, A_LoopField
	LocTable%LocationIndex%_%A_Index% := LocationText
}
}
;			Display select rows of table for debugging.
;			Note how the cell addressing is done.
;MSGBOX,1,,
;(
;The table length is:      %LocationIndex%
;Row 1:	%LocTable1_1%	%LocTable1_2%	%LocTable1_3%	%LocTable1_4%
;Row 5:	%LocTable5_1%	%LocTable5_2%	%LocTable5_3%	%LocTable5_4%
;Row 9:	%LocTable9_1%	%LocTable9_2%	%LocTable9_3%	%LocTable9_4%
;)
;=====================	Open Word using the template file in the same
;			directory as this script.  Originally the template was
;			from: "https://portal.cantonpl.org/docs/Documents/Circulation Services/Forms/ILL Forms and Stickers/MeL sticky label.dotx"
;
Run "MeL sticky label.dotx", Max
Sleep, 300
WinWaitActive, - Word,,10
Sleep, 500
;
Reason := "Timeout occurred waiting for Word to open template file."
If (ErrorLevel <> 0)
	GoTo, AppExit
;===================== Set the Default Printer and Get the User's Initials.
;
RunWait, rundll32 printui.dll PrintUIEntry /y /n "EPSON TM-T88IV ReStick"
Sleep, 250
;			Request the user to enter their initials.
;
;			The GUI boxes in this script are like subroutines.
;			The return at the end of their setup generates a
;			wait until a button action occurs.
;
;			Each GUI box has some instructions (Add, Text),
;			a data entry field (Add, Edit) which also defines
;			the size of the field and the variable into which to
;			place the data keyed in by the user.  Along with
;			2 buttons (Done and Cancel) and where to GoTo
;			in the script when pressed.
;
;			An [Enter] pressed during data entry will trigger
;			the default button, which is Done.
UsersInitials := "CPL"
;
GUI, New,,Enter Initials
GUI, Add, Text,T1, Type in your initials for use on these MeL Labels:
GUI, Add, Edit, w30 h21 vUsersInitials
GUI, Add, Button, Default gDONEpressed0, Done
GUI, Add, Button, gCANCELpressed0, Cancel
GUI, Show, Center
Return
;			The user pressed cancel.
;			Any time the script is sent to the exit app routine, a
;			reason is set so we can determine why it ended.
CANCELpressed0:
GUI, Destroy
Reason := "Request for Initials, user clicked cancel."
GoTo, AppExit
;			User indicated data entry was completed,
;			so capture the data and close the window.
;			In this case the data is converted to all uppercase.
DONEpressed0:
GUI, Submit
GUI, Destroy
StringUpper, UsersInitials, UsersInitials
Sleep, 200
;
;===================== For each label.
;			Request user to enter the MeL Location Code
;			from the printed Mel request form.  If same as
;			last location, the user can click Again to print 
;			another label with the same location information.
GetLocCode:
;
LocationCode	:= ZZ999
;
GUI, New,,Enter Location Code 
GUI, Add, Text, T1, Type in the MeL Location Code for this Label,
GUI, Add, Text, T1, ............or click AGAIN to reprint the last Label:
GUI, Add, Edit, w47 h21 vLocationCode
GUI, Add, Button, gItemTypeEntry, Again
GUI, Add, Button, Default gDONEpressed1, Done
GUI, Add, Button, gCANCELpressed1, Cancel
GUI, Show, Center
Return
;			The user pressed cancel.
CANCELpressed1:
GUI, Destroy
Reason := "Request for MeL Location Code, user clicked cancel."
GoTo, AppExit
;			Data entry complete, capture data and close window.
DONEpressed1:
GUI, Submit
GUI, Destroy
Sleep, 200
;			Ensure data entered is all uppercase and
;			5 characters long.
;
StringUpper, LocationText, LocationCode
FieldLength := StrLen(LocationText)
;
If (FieldLength <> 5)
{
      MSGBOX,1,,The Location Code entered was not 5 characters long, press OK to try again.
      GoTo, GetLocCode
}
;===================== Initialize the duplicates table (24 max).
;			Some of the MeL location codes are duplicated due
;			to branch locations.  The user will need to tell us
;			which one to use.  So the DupLocs matrix table will
;			hold each of the branch names for this location code.
;
Loop 24
{
DupLocs%A_Index% := "                                        "
}
LibLocation := ""
DupCount := 0
;
;===================== Find the location code which was entered by the 
;			user in the LocTable that we read in earlier.
Loop
{
LocationCode := LocTable%A_Index%_2
CodeLength := StrLen(LocationCode)
; 
If (LocationCode  =  LocationText)
{
	DupCount := DupCount + 1
	LocName := LocTable%A_Index%_4 "                                        "
	DupLocs%DupCount% := SubStr(LocName,1,40)
	GoTo, NextLocRecord
}
;			Once all of the matches for a location code are
;			found, or all codes checked, the loop can end.
;
If (DupCount > 0) or (A_Index > LocationIndex)
	Break
;
NextLocRecord:
}
;===================== The matching process is complete.
;			If DupCount is 0, no match was found, so
;			have the user type in the location name.
;			If DupCount is 1, no dups were found, so
;			no need to show the dups table.
;			If DupCount is more than 1, need to show the
;			dups table so user can pick location to use.
;
;			At the end, LibLocation will contain the library
;			to be printed on the label.
If (DupCount < 1)
	GoTo, UserTypesLocation
;
LibLocation := DupLocs1
SelectedLocation := 0
;
If (DupCount = 1)
	GoTo, ItemTypeEntry
;
;===================== Display the branch names and have user pick.
GetLocDup:
;
GUI, New,,Select Location
GUI, Add, Text,T1, Type in the Number of the Correct Branch for %LocationText%:
GUI, Add, Text,T1, (1) %DupLocs1%			(13) %DupLocs13%
GUI, Add, Text,T1, (2) %DupLocs2%			(14) %DupLocs14%
GUI, Add, Text,T1, (3) %DupLocs3%			(15) %DupLocs15%
GUI, Add, Text,T1, (4) %DupLocs4%			(16) %DupLocs16%	
GUI, Add, Text,T1, (5) %DupLocs5%			(17) %DupLocs17%
GUI, Add, Text,T1, (6) %DupLocs6%			(18) %DupLocs18%
GUI, Add, Text,T1, (7) %DupLocs7%			(19) %DupLocs19%
GUI, Add, Text,T1, (8) %DupLocs8%			(20) %DupLocs20%
GUI, Add, Text,T1, (9) %DupLocs9%			(21) %DupLocs21%
GUI, Add, Text,T1, (10) %DupLocs10%			(22) %DupLocs22%
GUI, Add, Text,T1, (11) %DupLocs11%			(23) %DupLocs23%
GUI, Add, Text,T1, (12) %DupLocs12%			(24) %DupLocs24%
GUI, Add, Edit, w21 h21 vSelectedLocation
GUI, Add, Button, Default gDONEpressed2, Done
GUI, Add, Button, gCANCELpressed2, Cancel
GUI, Show, Center
Return
;			The user pressed cancel.
CANCELpressed2:
GUI, Destroy
Reason := "Request for MeL Branch, user clicked cancel."
GoTo, AppExit
;			Data entry complete, capture data and close window.
DONEpressed2:
GUI, Submit
GUI, Destroy
Sleep, 200
;
If (SelectedLocation > DupCount) or (SelectedLocation < 1)
{
      MSGBOX,1,,The Branch option entered was not a valid choice, press OK to try again.
      GoTo, GetLocDup
}
LibLocation := DupLocs%SelectedLocation%
GoTo, ItemTypeEntry
;
;====================== Have user type in the library location's name.
UserTypesLocation:
;
LibLocation := ""
;
GUI, New,,No Match Found - Type in the Location's Name
GUI, Add, Text,T1, Since no match was found, type in the Name of the Library's Location:
GUI, Add, Edit, w450 h21 vLibLocation
GUI, Add, Button, Default gDONEpressed3, Done
GUI, Add, Button, gCANCELpressed3, Cancel
GUI, Show, Center
Return
;			The user pressed cancel.
CANCELpressed3:
GUI, Destroy
Reason := "Request for entry of Location Name, user clicked cancel."
GoTo, AppExit
;			Data entry complete, capture data and close window.
DONEpressed3:
GUI, Submit
GUI, Destroy
Sleep, 200
;
;===================== Get the Item Type/Count.  Regular books aren't the
;			type that have an item count, but CDs and DVDs do.
ItemTypeEntry:
;
ItemCount := 0
ItemNote   := ""
;
GUI, New,,Multiple Items?
GUI, Add, Text,T1, Press enter if no entries are required, or:
GUI, Add, Text,T1, ..........Key in the number of discs/items.
GUI, Add, Text,T1, ..........To add a note, Tab to the notes field & type.
GUI, Add, Text,T1, ..........Then press Enter.
GUI, Add, Text,T1, ===================================
GUI, Add, Text,T1, Number of items:
GUI, Add, Edit, w45 h21 vItemCount
GUI, Add, Text,T1, Note:
GUI, Add, Edit, w200 h21 vItemNote
GUI, Add, Button, Default gDONEpressed4, Done
GUI, Add, Button, gCANCELpressed4, Cancel
GUI, Show, Center
Return
;			The user pressed cancel.
CANCELpressed4:
GUI, Destroy
Reason := "Request for number of items, the user clicked cancel."
GoTo, AppExit
;			Data entry complete, capture data and close window.
DONEpressed4:
GUI, Submit
GUI, Destroy
Sleep, 100
;
;===================== Edit the label template starting with Initials.
EditForm:
;			Activate Word Window. 		
WinActivate, - Word
Sleep, 150
;			Type the user's initials and the number of items on Line 1.
Send %UsersInitials%{Tab}
Sleep, 100
;
If (ItemCount > 0)
	Send Item Count is: %ItemCount%
;
;			If there is a note, type on Line 2.
Sleep, 100
;
If (ItemNote <> "")
{
	Send {Enter}
	Sleep, 100
	Send {Tab}%ItemNote%
	Sleep, 250
}
;			Move down the page to the line with "Borrowing Library:" 
Send {Down 2}
Sleep, 50
;			Move to the left margin.
Send {Home}
Sleep, 100
;			Move beyond the "Borrowing Library:" label.
Send {Right 5}
Sleep, 30
Send {Right 5}
Sleep, 30
Send {Right 5}
Sleep, 30
Send {Right 4}
Sleep, 30
;			Type the selected library location's name, but highlighted:
;				[Ctrl-B] to turn on then off bold
;				[Ctrl-]] to increase font one size, we do 4 times.
;				[Ctrl-[] to decrease font one size, we set it back.
Send ^b%LocationText%   ^]
Sleep, 25
Send ^]^]
Sleep, 25
Send ^]%LibLocation%^b^[
Sleep, 25
Send ^[^[
Sleep, 25
Send ^[{Del}
Sleep, 350
;
;===================== Print the label to the default printer.
Send ^p{Enter}
Sleep, 3500
;===================== Get Ready for Next Label by Undo-ing changes.
;			Undo is control-z.
Loop 48
{
	Send ^z
	Sleep, 5
}
;			Next Label
GoTo, GetLocCode
;
;===================== End Script
AppExit:
;			The script is ending, so close Word and then
;			display the reason for ending.
;
;			Close Word without saving. Alt-F4 to close, and
;			then choose the do not save option.
WinActivate, - Word
Sleep, 300
Send !{F4}
Sleep, 100
Send n
Sleep, 100
;
MSGBOX,1,,%Reason%    Script is Ending.
ExitApp

