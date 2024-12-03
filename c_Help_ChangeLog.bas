Sub ChangeLog()
Application.ScreenUpdating = False
If Evaluate("ISREF(Merlin_ChangeLog!A1)") = True Then
    Application.DisplayAlerts = False
    Sheets("Merlin_ChangeLog").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = "Merlin_ChangeLog"
Else
    Sheets.Add.Name = "Merlin_ChangeLog"
End If

'Blank
'Cells(r, 1) = ""
'Cells(r, 2) = ""
'Cells(r, 3) = ""
'r = r + 1
r = 1

Cells(r, 1) = "Menu Item"
Cells(r, 2) = "Keyboard Shortcut"
Cells(r, 3) = "Description"

Cells(r, 1).Font.Bold = True
Cells(r, 2).Font.Bold = True
Cells(r, 3).Font.Bold = True

r = r + 1
Cells(r, 1) = "Merlin"
Cells(r, 2) = ""
Cells(r, 3) = "Merlin was originally created by Kyle Whitmire, but has had many contributors over the years with a special thanks to Daniel Whiteman who contributed much better VBA code than Kyle.  Much of the code is not my own, but was 'stolen' from the web. I've attempted to leave credit to the original author in the comments where large chunks of code were used."

r = r + 1
Cells(r, 1) = "Update Merlin"
Cells(r, 2) = ""
Cells(r, 3) = "Checks for new version of Merlin online (https://merlinaddin.xyz) and updates local copy accordingly."

r = r + 1
Cells(r, 1) = "Merlin Support"
Cells(r, 2) = ""
Cells(r, 3) = "Google Groups forum to ask questions, submit suggestions, or to upload code for bug fix or enhancements. URL: https://groups.google.com/forum/#!forum/merlin-add-in."

r = r + 1
Cells(r, 1).Font.Bold = True
Cells(r, 1) = "Formatting"
Cells(r, 2) = ""
Cells(r, 3) = "Self-explanatory, but generally alters visual formatting of cells for consistency and allows use of keyboard shortcuts."

r = r + 1
Cells(r, 1) = "     Yellow"
Cells(r, 2) = "Ctrl+Shft+Y"
Cells(r, 3) = "Formats cell background color yellow."

r = r + 1
Cells(r, 1) = "     Green"
Cells(r, 2) = "Ctrl+Shft+G"
Cells(r, 3) = "Formats cell background color green."

r = r + 1
Cells(r, 1) = "     Blue"
Cells(r, 2) = "Ctrl+Shft+B"
Cells(r, 3) = "Formats cell background color blue."

r = r + 1
Cells(r, 1) = "     Red"
Cells(r, 2) = "Ctrl+Shft+R"
Cells(r, 3) = "Formats cell background color red.  Also formats text color white."

r = r + 1
Cells(r, 1) = "     Post-It Note Yellow"
Cells(r, 2) = "Ctrl+Shft+N"
Cells(r, 3) = "Formats cell background color ""post-it note"" yellow."

r = r + 1
Cells(r, 1) = "     Clear Formatting"
Cells(r, 2) = "Ctrl+Shft+C"
Cells(r, 3) = "Clear ALL formatting on selected cells."

r = r + 1
Cells(r, 1) = "     Paste Special - Formatting"
Cells(r, 2) = "Ctrl+Shft+P"
Cells(r, 3) = "Paste formatting on selected cells (must have copied or Ctrl+C first)."

r = r + 1
Cells(r, 1) = "     Paste Special - Values"
Cells(r, 2) = "Ctrl+Shft+S"
Cells(r, 3) = "Paste values in selected cells (must have copied or Ctrl+C first)."

r = r + 1
Cells(r, 1) = "     Paste Special - Formulas"
Cells(r, 2) = "Ctrl+Shft+F"
Cells(r, 3) = "Paste Formulas in selected cells (must have copied or Ctrl+C first)."

r = r + 1
Cells(r, 1) = "     Page Setup"
Cells(r, 2) = ""
Cells(r, 3) = "Modified print settings to narrow margins and add dynamic date/time footer "

r = r + 1
Cells(r, 1) = "     Color Columns"
Cells(r, 2) = ""
Cells(r, 3) = "Formats all columns to the left of selected column to have the same background color as the selected column."


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

r = r + 1
Cells(r, 1).Font.Bold = True
Cells(r, 1) = "Number Formatting"
Cells(r, 2) = ""
Cells(r, 3) = ""

r = r + 1
Cells(r, 1) = "     ""Number Format"" through ""Percent Format no Red"""
Cells(r, 2) = ""
Cells(r, 3) = "Self-explanatory, but alters visual formatting of numbers for consistency and allows use of keyboard shortcuts. Does not alter the actual number/value in any way.  Anything specifying ""no Red"" leaves negative numbers with default color.  Others format negative numbers Red."

r = r + 1
Cells(r, 1) = "     Basis Point Format"
Cells(r, 2) = ""
Cells(r, 3) = "Replaces value or formula with formula that multiplies by 10,000 and adds ""bps"" to formatting. (ex. 0.05% will be multiplied by 10,000 and read ""5 bps”)."

r = r + 1
Cells(r, 1) = "     Ordinal Number Format"
Cells(r, 2) = ""
Cells(r, 3) = "Adds ordinal indicator (1st, 2nd, etc.) via number formatting based on value in cell."

r = r + 1
Cells(r, 1) = "     If Error then 0"
Cells(r, 2) = "Ctrl+Shft+E"
Cells(r, 3) = "Wraps value/formula in IFERROR(x,0) for error correction.  Can be used on contiguous ranges even if formulas are all different."

r = r + 1
Cells(r, 1) = "     Round"
Cells(r, 2) = ""
Cells(r, 3) = "Wraps formula in ROUND(x,y).  Prompted for y decimals for rounding.  Can be used on contiguous ranges even if formulas are all different."

r = r + 1
Cells(r, 1) = "     Increase Decimal"
Cells(r, 2) = ""
Cells(r, 3) = "Adds a decimal place to number formatting."

r = r + 1
Cells(r, 1) = "     Decrease Decimal"
Cells(r, 2) = ""
Cells(r, 3) = "Removes a decimal place from number formatting."


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

r = r + 1
Cells(r, 1).Font.Bold = True
Cells(r, 1) = "Workbook Efficiency"
Cells(r, 2) = ""
Cells(r, 3) = ""

r = r + 1
Cells(r, 1) = "     Two Range List Builder"
Cells(r, 2) = ""
Cells(r, 3) = "Loops through two ranges of text to build two columns containing each possible combination of the items in each range."

r = r + 1
Cells(r, 1) = "     Three Range List Builder"
Cells(r, 2) = ""
Cells(r, 3) = "Same as above, but for three unique ranges."

r = r + 1
Cells(r, 1) = "     Evaluate as Formula/Number"
Cells(r, 2) = ""
Cells(r, 3) = "After pasting values, sometimes formulas/numbers get stuck as Text.  This converts to General and evalutates as formulas/numbers again."

r = r + 1
Cells(r, 1) = "     Convert Text to Formula"
Cells(r, 2) = ""
Cells(r, 3) = "Wraps any text/number value in a cell with ="""""".  Skips cells with existing formula."

r = r + 1
Cells(r, 1) = "     Copy/Paste Exact Formulas"
Cells(r, 2) = ""
Cells(r, 3) = "Creates an exact copy of formulas in a different place.  Used when absolute reference is not desirable/practical, but formulas need to be duplicated.  (Example: Cell A1 contains formula of ""=A2"".  Normally a copy/paste to B1 would result in formula of ""=B2"".  Using the macro would result in formula of ""=A2""... an exact copy.)"

r = r + 1
Cells(r, 1) = "     Crack Internal Passwords"
Cells(r, 2) = ""
Cells(r, 3) = "Uses brute force to breaks password protection in protected Worksheets and VBA Modules in the Active Workbook."

r = r + 1
Cells(r, 1) = "     Find Errors in Formulas"
Cells(r, 2) = ""
Cells(r, 3) = "Creates a hyperlinked list of all formulas on current Worksheet with Errors in them for easier cleanup."

r = r + 1
Cells(r, 1) = "     Manage Hidden Objects"
Cells(r, 2) = ""
Cells(r, 3) = "Lists all hidden objects in Workbook including worksheets, shapes (charts, controls, etc.), and named ranges.  Allows user to selectively unhide each one by placing a ""Y"" in column C and rerunning the macro."

r = r + 1
Cells(r, 1) = "     List External Links"
Cells(r, 2) = ""
Cells(r, 3) = "Creates a hyperlinked list of all external links in Workbook for easier cleanup/management."

r = r + 1
Cells(r, 1) = "     Create Workbook Table of Contents"
Cells(r, 2) = ""
Cells(r, 3) = "Creates a hyperlinked list of all Worksheets in the Active Workbook."

r = r + 1
Cells(r, 1) = "     Count Worksheets"
Cells(r, 2) = ""
Cells(r, 3) = "Counts number of Worksheets in Workbook whether Hidden or Visible.  This is for you Sven!"

r = r + 1
Cells(r, 1) = "     Unhide/Rehide Worksheets"
Cells(r, 2) = ""
Cells(r, 3) = "Unhides any hidden Worksheets.  When re-run, the macro re-hides the same worksheets. Thanks Ryan!"

r = r + 1
Cells(r, 1) = "     Worksheet Selector"
Cells(r, 2) = "Ctrl+Shft+W"
Cells(r, 3) = "Shows a pop-up of Worksheets that user can click to navigate to selected Worksheet.  Thanks Maddy and Ross!"

r = r + 1
Cells(r, 1) = "     Unhide/Rehide Worksheets"
Cells(r, 2) = "Ctrl+Shft+U"
Cells(r, 3) = "Acts as Toggle. Run once to unhide any hidden Worksheets.  Run again to re-hide the same Worksheets."

r = r + 1
Cells(r, 1) = "     Size of Worksheets"
Cells(r, 2) = ""
Cells(r, 3) = "Attempts to give a directional guide of the relative sizes of each Worksheet.  Creates a new Workbook out of each Worksheet and lists the size of each."

r = r + 1
Cells(r, 1) = "     Auto-Group Hidden Rows/Columns"
Cells(r, 2) = ""
Cells(r, 3) = "Finds any hidden rows/columns that aren't already grouped and groups them."

r = r + 1
Cells(r, 1) = "     Disable AutoRecover"
Cells(r, 2) = ""
Cells(r, 3) = "Disables AutoRecover on current Workbook.  Useful if workbook is very large or has long recalc time."

r = r + 1
Cells(r, 1) = "     Export to Delimited File"
Cells(r, 2) = ""
Cells(r, 3) = "Allows for user-input controls on exporting the Active Sheet to a delimited file.  Various options presented during subroutine."

r = r + 1
Cells(r, 1) = "     Calc Timers"
Cells(r, 2) = ""
Cells(r, 3) = "Calculates multiple iterations and averages calculation time to help with troubleshooting long recalc times by isolating the section causing trouble."

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

r = r + 1
r = r + 1
Cells(r, 1) = "Percent Variance"
Cells(r, 2) = "Ctrl+Shft+V"
Cells(r, 3) = "Inserts the formula to calculate percent variance between the two columns to the left of the current cell."

r = r + 1
Cells(r, 1) = "Highlight Selection"
Cells(r, 2) = "Ctrl+Shft+H"
Cells(r, 3) = "Creates a highlight around the current Selection. Run again to toggle highlights off. Used when 'presenting' a document to other people and you want to call attention to a particular cell."

r = r + 1
Cells(r, 1) = "GoTo Precedent"
Cells(r, 2) = "Ctrl+Shft+X"
Cells(r, 3) = "Works similar to Trace Precedent function in Excel, but presents results with more detail.  If only one precedent, will auto-jump to that cell.  If multiple precedents, user inputs which jump to take.  With INDEX/MATCH and VLOOKUP formulas, the jump takes you to the cell where the value returned from the formula originated."

r = r + 1
Cells(r, 1) = "Go Back from Precedent"
Cells(r, 2) = "Ctrl+Shft+Z"
Cells(r, 3) = "After executing GoTo Precendent function above, this takes you back to the cell you initially jumped from."

r = r + 1
Cells(r, 1) = "Scale All Charts on Sheet"
Cells(r, 2) = ""
Cells(r, 3) = "Scales all Primary and Secondary axes of charts on the active sheet.  Does not play well with Stacked Column. If chart name contains the text 'NoScale' it will be skipped.  If chart name contains 'Split' the data points on the Primary Axis will be contained in top half of chart and data points in Secondary Axis will be contained in bottom half."


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

r = r + 1
r = r + 1
Cells(r, 1).Font.Bold = True
Cells(r, 1) = "Active Chart Functions"
Cells(r, 2) = ""
Cells(r, 3) = "This menu is only visible when a chart is selected"

r = r + 1
Cells(r, 1) = "     Rename Chart"
Cells(r, 2) = ""
Cells(r, 3) = "Dialog Box popup to rename Active Chart."

r = r + 1
Cells(r, 1) = "     Transparent Chart"
Cells(r, 2) = ""
Cells(r, 3) = "Removes background fill and border for Active Chart.  Looks nicer in PowerPoint."

r = r + 1
Cells(r, 1) = "     Center Chart Title"
Cells(r, 2) = ""
Cells(r, 3) = "Yup, just like it sounds."

r = r + 1
Cells(r, 1) = "     Scale Active Chart"
Cells(r, 2) = ""
Cells(r, 3) = "Same as Scale All Charts, but only on Active Chart."

r = r + 1
Cells(r, 1) = "     Scale All Charts on Sheet"
Cells(r, 2) = ""
Cells(r, 3) = "Scales all Primary and Secondary axes of charts on the active sheet.  Does not play well with Stacked Column. If chart name contains the text 'NoScale' it will be skipped.  If chart name contains 'Split' the data points on the Primary Axis will be contained in top half of chart and data points in Secondary Axis will be contained in bottom half."

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

r = r + 1
r = r + 1
Cells(r, 1).Font.Bold = True
Cells(r, 1) = "Under Development (use at own risk)"
Cells(r, 2) = ""
Cells(r, 3) = "No documentation provided and no guarantee that these work."

r = r + 1
Cells(r, 1) = "User Defined Formulas"
Cells(r, 2) = ""
Cells(r, 3) = ""

Cells(r, 1).Font.Bold = True
Cells(r, 2).Font.Bold = True
Cells(r, 3).Font.Bold = True

r = r + 1
Cells(r, 1) = "{=List_Unique}"
Cells(r, 2) = ""
Cells(r, 3) = "Lists only the unique entries in a range.  Must be entered as a CSE Array."

r = r + 1
Cells(r, 1) = "StockQuote"
Cells(r, 2) = ""
Cells(r, 3) = "Returns High, Low, Open, Close, Volume data for a given Ticker symbol"


'*************************************************************************************************************************
'*************************************************************************************************************************
'*************************************************************************************************************************
'*********************************** END HELP - BEGIN CHANGE LOG *********************************************************
'*************************************************************************************************************************
'*************************************************************************************************************************
'*************************************************************************************************************************

r = r + 2
r = r + 2
Cells(r, 1) = "Change Item"
Cells(r, 2) = "Date"
Cells(r, 3) = "Description"

Cells(r, 1).Font.Bold = True
Cells(r, 2).Font.Bold = True
Cells(r, 3).Font.Bold = True

r = r + 1
Cells(r, 1) = "bps formatting enhancement"
Cells(r, 2) = "08.02.23"
Cells(r, 3) = "Enhancement to allow selection range instead of just one cell for bps_format."

r = r + 1
Cells(r, 1) = "Clipboard Empty code fix"
Cells(r, 2) = "08.03.23"
Cells(r, 3) = "Fixed bug where nothing is in the clipboard.  Used for paste_special, paste_formulas, and paste_formatting."

r = r + 1
Cells(r, 1) = "Trace Precedent code fix"
Cells(r, 2) = "09.09.19"
Cells(r, 3) = "Fixed bug in opening external files."

r = r + 1
Cells(r, 1) = "Unhide/Rehide Worksheets"
Cells(r, 2) = "03.03.19"
Cells(r, 3) = "New feature added."

r = r + 1
Cells(r, 1) = "Convert Text to Formula"
Cells(r, 2) = "02.19.19"
Cells(r, 3) = "New feature added."

r = r + 1
Cells(r, 1) = "Evaluate as Formula/Number"
Cells(r, 2) = "02.19.19"
Cells(r, 3) = "Same as Convert Text to Formula from 02.14.19, but clarified name/function after adding new Convert function."
r = r + 1
Cells(r, 1) = "Removed Essbase sub-menu"
Cells(r, 2) = "02.19.19"
Cells(r, 3) = "Moved features to Workbook Efficiency sub-menu."
r = r + 1
Cells(r, 1) = "New Number Formatting Functions and sub-menu"
Cells(r, 2) = "02.19.19"
Cells(r, 3) = "Added various K & M formatting options and created new Number Formatting sub-menu."

r = r + 1
Cells(r, 1) = "Ordinal Format"
Cells(r, 2) = "02.14.19"
Cells(r, 3) = "New feature added."

r = r + 1
Cells(r, 1) = "Round"
Cells(r, 2) = "02.14.19"
Cells(r, 3) = "New feature added."

r = r + 1
Cells(r, 1) = "Convert Text to Formula"
Cells(r, 2) = "02.14.19"
Cells(r, 3) = "New feature added. Thanks Daniel."

r = r + 1
Cells(r, 1) = "View All Worksheets"
Cells(r, 2) = "11.16.18"
Cells(r, 3) = "New feature added. Thanks Maddy and Ross."

r = r + 1
Cells(r, 1) = "Update Merlin"
Cells(r, 2) = "05.24.18"
Cells(r, 3) = "Update Merlin function working again both inside and outside AT&T's firewall."

r = r + 1
Cells(r, 1) = "Merlin Support"
Cells(r, 2) = "05.24.18"
Cells(r, 3) = "Added link to Merlin Add-In Google Group."

r = r + 1
Cells(r, 1) = "Percent Variance"
Cells(r, 2) = "04.27.18"
Cells(r, 3) = "Inserts the formula to calculate percent variance between the two columns to the left of the current cell."

r = r + 1
Cells(r, 1) = "Added Column Coloring"
Cells(r, 2) = "03.15.18"
Cells(r, 3) = "Applies colors by row from selected column to all columns to the left of that column."

r = r + 1
Cells(r, 1) = "Added Clear formatting"
Cells(r, 2) = "09.08.17"
Cells(r, 3) = "Clears formatting from current range (Ctrl + Shft + C)"

r = r + 1
Cells(r, 1) = "Added Paste formatting"
Cells(r, 2) = "09.08.17"
Cells(r, 3) = "Pastes format from copied range to selected range (same as Alt+E+S+F) (Ctrl + Shft + P)"

r = r + 1
Cells(r, 1) = "Tweaked Highlight function"
Cells(r, 2) = "07.13.17"
Cells(r, 3) = "Added multiple Selection highlight function and modified for highlight Selection instead of ActiveCell"

r = r + 1
Cells(r, 1) = "Added $M formatting"
Cells(r, 2) = "06.29.17"
Cells(r, 3) = "Added $M format that applies 0;;M format to display $1,111,111 as $1M (Ctrl + Shft + M)"

r = r + 1
Cells(r, 1) = "Added Increase Decimal"
Cells(r, 2) = "06.29.17"
Cells(r, 3) = "Added decimal increase routine that allows shortcut key (Ctrl + Shft + I)"

r = r + 1
Cells(r, 1) = "Added Decrease Decimal"
Cells(r, 2) = "06.29.17"
Cells(r, 3) = "Added decimal decrease routine that allows shortcut key (Ctrl + Shft + D)"

r = r + 1
Cells(r, 1) = "Error correction on GoTo_Precedent/INDEXTRACE/VLOOKUPTRACE"
Cells(r, 2) = "03.30.17"
Cells(r, 3) = "Fixed to work on named ranges including ranges in external workbooks either opened or closed"

r = r + 1
Cells(r, 1) = "Error correction on IFERROR"
Cells(r, 2) = "03.14.17"
Cells(r, 3) = "Modified IFERROR Subroutine to ignore cells with text or blank."

r = r + 1
Cells(r, 1) = "Added 'Manage Hidden Objects' Sub"
Cells(r, 2) = "03.14.17"
Cells(r, 3) = "Lists all hidden objects in workbook including worksheets, shapes (charts, controls, etc.), and named ranges.  Allows user to selectively unhide each one."

r = r + 1
Cells(r, 1) = "Fixed INDEXTRACE and VLOOKUPTRACE to work across workbooks"
Cells(r, 2) = "03.14.17"
Cells(r, 3) = "Jump code now opens a linked workbook, and all jumps (WS, WB, etc.) all function."

r = r + 1
Cells(r, 1) = "64-Bit Compatibility"
Cells(r, 2) = "03.10.17"
Cells(r, 3) = "Added ptrSafe code to Subs & Functions for Long variable to fix 64-Bit compatibility."
'Long needs to be LongPtr.  Declare Function needs to be Declare PtrSafe Function

r = r + 1
Cells(r, 1) = "New Menu"
Cells(r, 2) = "11.10.16"
Cells(r, 3) = "New Merlin menu created with sub-menu hierarchy."


Worksheets("Merlin_ChangeLog").Activate

Worksheets("Merlin_ChangeLog").Columns("A:A").ColumnWidth = "40"
Worksheets("Merlin_ChangeLog").Range("A:A").WrapText = True
Worksheets("Merlin_ChangeLog").Columns("B:B").EntireColumn.AutoFit
Worksheets("Merlin_ChangeLog").Columns("C:C").ColumnWidth = "100"
Worksheets("Merlin_ChangeLog").Range("C:C").WrapText = True
Worksheets("Merlin_ChangeLog").Range("A:C").VerticalAlignment = xlTop
Application.ScreenUpdating = True
End Sub
