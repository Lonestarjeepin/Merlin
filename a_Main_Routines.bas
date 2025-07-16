'Shortcut Keys are in "ThisWorkbook" under left menu VBAProject (Default.xlam) Excel Objects,
'then in dropdowns, "Workbook" "Open"

'Declare Function must be Declare ptrSafe Function for 64-bit compatibility
'Long variable type must be LongPtr for 64-bit compatibility

'-----------------------------------------------------------------------------------------------------------
'Just some intersting code that I use a lot so saving it here
'
'For "onClick"
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    If Not Intersect(ActiveCell, Range("A1")) Is Nothing Then
'       Call xxxxx
'    End If
'End Sub


'For "onChange"
'Private Sub Worksheet_Change(ByVal Target As Range)
'    If Not Intersect(Range("E8"), Target) Is Nothing Then
'     Call xxxxx
'    End If
'End Sub

'**** For Calculation Timers ****
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As LongPtr
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As LongPtr


'***** For Scale Chart *****
Public Type scaleAxisScale
  ' Calculated Axis Scale Parameters
  dMin As Double
  dMax As Double
  dMajor As Double
  dMinor As Double
End Type


Public FindFormulaToggle As Integer
'Public Const vbDoubleQuote As String = Chr(34) '"""" 'represents 1 double quote (")
Public colHiddenWS As New Collection

' Declare the API function for checking clipboard formats
' Declaration used for error checking in paste routines like paste_special, paste_formulas, etc. to accomodate ClipboardHasContent Function
#If VBA7 Then
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
#Else
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
#End If








Function ClipboardHasContent() As Boolean
    Dim i As Long
    
    ' Loop through possible clipboard formats (1 to 20 should cover most common formats)
    For i = 1 To 20
        ' Check if the current format is available in the clipboard
        If IsClipboardFormatAvailable(i) <> 0 Then
            ClipboardHasContent = True
            Exit Function
        End If
    Next i
    
    ' If no active clipboard element is found, return False
    ClipboardHasContent = False
End Function

    
'**********************************************************************************************

Sub Green()
' 07.09.10 - Kyle Whitmire

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65280
    End With
End Sub
Sub Red()
' 07.09.10 - Kyle Whitmire
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        '.TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
    End With
End Sub
Sub Blue()
' 07.09.10 - Kyle Whitmire
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16763904
    End With
End Sub
Sub Yellow()
' 07.09.10 - Kyle Whitmire

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
End Sub
Sub post_it_note()
' 07.09.10 - Kyle Whitmire

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
    End With
    Selection.Locked = False
    Selection.FormulaHidden = False
End Sub
Sub number_format()
' 07.14.09 - Kyle Whitmire
    Selection.NumberFormat = "#,##0_);[Red](#,##0);""-""_)"
End Sub
Sub number_nored_format()
' 07.14.09 - Kyle Whitmire
    Selection.NumberFormat = "#,##0_);(#,##0);""-""_)"
End Sub
Sub million_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_(#,##0,,""M""_);[Red]_((#,##0,,""M"");_(""-""_)"
End Sub
Sub million_nored_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_(#,##0,,""M""_);_((#,##0,,""M"");_(""-""_)"
End Sub
Sub thousand_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_(#,##0,""K""_);[Red]_((#,##0,""K"");_(""-""_)"
End Sub
Sub thousand_nored_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_(#,##0,""K""_);_((#,##0,""K"");_(""-""_)"
End Sub
Sub dollar_format()
' 07.14.09 - Kyle Whitmire (aka Sharon's dollar format in honor of our dear friend Sharon Smith)
    Selection.NumberFormat = "_($* #,##0_);[Red]_($* (#,##0);_($* ""-""_)"
End Sub
Sub dollar_nored_format()
' 07.14.09 - Kyle Whitmire
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_)"
End Sub
Sub dollar_million_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_($* #,##0,,""M""_);[Red]_($* (#,##0,,""M"");_($* ""-""_)"
End Sub
Sub dollar_million_nored_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_($* #,##0,,""M""_);_($* (#,##0,,""M"");_($* ""-""_)"
End Sub
Sub dollar_thousand_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_($* #,##0,""K""_);[Red]_($* (#,##0,""K"");_($* ""-""_)"
End Sub
Sub dollar_thousand_nored_format()
' 06.29.17 - Kyle Whitmire
    Selection.NumberFormat = "_($* #,##0,""K""_);_($* (#,##0,""K"");_($* ""-""_)"
End Sub
Sub percent_format()
' 07.14.09 - Kyle Whitmire
    Selection.NumberFormat = "#,##0.0%_);[Red](#,##0.0%);""-""_)"
End Sub
Sub percent_nored_format()
' 07.14.09 - Kyle Whitmire
    Selection.NumberFormat = "#,##0.0%_);(#,##0.0%);""-""_)"
End Sub
Sub bps_format()
' 07.14.09 - Kyle Whitmire

    Dim cell As Range

    ' Loop through each cell in the selected range
    For Each cell In Selection

        ' Check if the cell contains a formula
        If cell.HasFormula Then
            If InStr(cell.Formula, "*10000") Or InStr(cell.Formula, "*10^4") Then
            Else
                cell.Formula = Replace(cell.Formula, "=", "=(") & ")*10000"
            End If
        Else
            cell.Value = "=(" & cell.Value & ")*10000"
        End If

        ' Apply the number format to the cell
        cell.NumberFormat = "#,##0"" bps""_);[Red](#,##0"" bps"");""-""_)"

    Next cell

End Sub
Sub increase_decimal()
' 06.29.17 - Kyle Whitmire
    Application.CommandBars.FindControl(ID:=398).Execute
End Sub
Sub decrease_decimal()
' 06.29.17 - Kyle Whitmire
    Application.CommandBars.FindControl(ID:=399).Execute
End Sub
Sub paste_special()
    ' Keyboard Shortcut: Ctrl+Shift+S
    
    ' Check if the clipboard has content
    If Not ClipboardHasContent() Then
        MsgBox "There is no content in the clipboard. Please copy something first.", vbInformation, "No Clipboard Content"
    Else
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If

End Sub
Sub paste_formulas()
' Keyboard Shortcut: Ctrl+Shift+F
    
    ' Check if the clipboard has content
    If Not ClipboardHasContent() Then
        MsgBox "There is no content in the clipboard. Please copy something first.", vbInformation, "No Clipboard Content"
    Else
    
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
End Sub
Sub clear_formatting()
' Keyboard Shortcut: Ctrl+Shift+C
    Selection.ClearFormats
End Sub
Sub paste_formatting()
' Keyboard Shortcut: Ctrl+Shift+P
    
    ' Check if the clipboard has content
    If Not ClipboardHasContent() Then
        MsgBox "There is no content in the clipboard. Please copy something first.", vbInformation, "No Clipboard Content"
    Else
    
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    End If
End Sub
Sub Page_Setup()
    With ActiveSheet.PageSetup
        .RightFooter = "&D &T"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With
End Sub
Sub LoadEssbase()
    On Error Resume Next
        AddIns("Essexcln").Installed = True
    AddIns("Oracle Essbase OLAP Server DLL (non-Unicode)").Installed = True
    AddIns("Oracle Essbase Query Designer AddIn").Installed = True
    AddIns("Oracle Hyperion Smart View for Office, Fusion Edition").Installed = _
        True
    AddIns("In2Hyp Essbase Ribbon").Installed = True
End Sub

Sub Variance_Percent()
' 01.15.2014 - Daniel Whiteman
    'Help/(Hurt) 1-x or Inc/(Dec) x-1
    varScenType = InputBox("Would you like to express as Help/(Hurt) (option 1) or Inc/(Dec) (option 2)?")
    If varScenType = 1 Then
        Selection.FormulaR1C1 = "=1-RC[-2]/RC[-1]"
    ElseIf varScenType = 2 Then
        Selection.FormulaR1C1 = "=RC[-2]/RC[-1]-1"
    End If
    Call percent_nored_format
End Sub
Sub IFERROR()
' 01.15.2014 - Daniel Whiteman
Application.ScreenUpdating = False
    Dim cell As Range
    
    For Each cell In Selection
        If cell.HasFormula = True Then
            cell.Formula = "=iferror(" & Right(cell.Formula, Len(cell.Formula) - 1) & ",0)"
        End If
    Next
Application.ScreenUpdating = True
End Sub

Sub RenameChart()

Dim varChartName As String

varChartName = InputBox(Prompt:="Input New Chart Name.")
ActiveChart.Parent.Name = varChartName

End Sub

Sub transparent_chart()
    With ActiveChart
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
    End With
End Sub

Sub center_chart_title()
  Application.ScreenUpdating = False
    ActiveChart.ChartTitle.Left = ActiveChart.ChartArea.Width
    ActiveChart.ChartTitle.Left = ActiveChart.ChartTitle.Left / 2
  Application.ScreenUpdating = True
End Sub

Sub AddIns_Off()

    AddIns("PTS Cluster Stack Utility").Installed = False
    AddIns("PTS Waterfall Plotter").Installed = False
    AddIns("XY Chart Labeler 7.0").Installed = False
    AddIns("Solver Add-in").Installed = False

End Sub

Sub AddIns_On()

    AddIns("PTS Cluster Stack Utility").Installed = True
    AddIns("PTS Waterfall Plotter").Installed = True
    AddIns("XY Chart Labeler 7.0").Installed = True
    AddIns("Solver Add-in").Installed = True

End Sub


Sub Two_Range_List_Builder()

'Created by Kyle Whitmire - 07.19.11

    Application.Calculation = xlManual
    
    Dim rng_Range1 As Range
    Set rng_Range1 = Application.InputBox(Prompt:="Select 1st Range (e.g. Locations).", Type:=8)
    
    Dim rng_Range2 As Range
    Set rng_Range2 = Application.InputBox(Prompt:="Select 2nd Range (e.g. Departments/Accounts).", Type:=8)

    Dim rng_Paste_Cell As Range
    
    Set rng_Paste_Cell = Application.InputBox(Prompt:="Select Cell to Paste Output", Type:=8)
    
    With rng_Paste_Cell
        .Parent.Parent.Activate
        .Parent.Activate
        .Select
    End With
    
Application.ScreenUpdating = False
    For Each x In rng_Range1
        For Each y In rng_Range2
            ActiveCell.Value = x
            ActiveCell.Offset(0, 1).Value = y
            ActiveCell.Offset(1, 0).Select
        Next
    Next
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
    rng_Paste_Cell.Select
End Sub
Sub Three_Range_List_Builder()

'Created by Kyle Whitmire - 07.19.11
    Application.Calculation = xlManual
    
    Dim rng_Range1 As Range
    Set rng_Range1 = Application.InputBox(Prompt:="Select 1st Range (e.g. Locations).", Type:=8)
    
    Dim rng_Range2 As Range
    Set rng_Range2 = Application.InputBox(Prompt:="Select 2nd Range (e.g. Departments/Accounts).", Type:=8)
    
    Dim rng_Range3 As Range
    Set rng_Range3 = Application.InputBox(Prompt:="Select 3rd Range (e.g. Departments/Accounts).", Type:=8)

    Dim rng_Paste_Cell As Range
    
    Set rng_Paste_Cell = Application.InputBox(Prompt:="Select Cell to Paste Output", Type:=8)
    rng_Paste_Cell.Select
    
Application.ScreenUpdating = False
    For Each x In rng_Range1
        For Each y In rng_Range2
            For Each Z In rng_Range3
                ActiveCell.Value = x
                ActiveCell.Offset(0, 1).Value = y
                ActiveCell.Offset(0, 2).Value = Z
                ActiveCell.Offset(1, 0).Select
            Next
        Next
    Next
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
    rng_Paste_Cell.Select
End Sub

Sub auto_group()

Application.ScreenUpdating = False

    lastCol = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.Count
    lastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    
    varCol = 1
    varRow = 1
    
    ActiveWindow.DisplayOutline = True

    Do While varRow <= lastRow
    If Rows(varRow & ":" & varRow).Hidden = True Then
        If Rows(varRow & ":" & varRow).OutlineLevel = 1 Then
          Rows(varRow & ":" & varRow).Group
        End If
        Rows(varRow & ":" & varRow).EntireRow.Hidden = False
    End If
    varRow = varRow + 1
    Loop
    
    Do While varCol <= lastCol
    If Columns(varCol).Columns.Hidden = True Then
        If Columns(varCol).OutlineLevel = 1 Then
          Columns(varCol).Columns.Group
        End If
        Columns(varCol).EntireColumn.Hidden = False
    End If
    varCol = varCol + 1
    Cells(1, varCol).Select
    Loop

Application.ScreenUpdating = True

End Sub

Sub Count_Worksheets()

    varCount = ActiveWorkbook.Worksheets.Count
    varhidden = 0
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
        varhidden = varhidden + 1
        End If
    Next
    
    MsgBox "There are " & varCount & " worksheets in this report. " & varhidden & " are hidden."

End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExportToTextFile
' This exports a sheet or range to a text file, using a
' user-defined separator character.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExportToTextFile(FName As String, Sep As String, AppendData As Boolean, SkipBlank As Boolean)


Dim WholeLine As String
Dim FNum As Integer
Dim RowNdx As LongPtr
Dim ColNdx As Integer
Dim StartRow As LongPtr
Dim EndRow As LongPtr
Dim StartCol As Integer
Dim EndCol As Integer
Dim CellValue As String


Application.ScreenUpdating = False
On Error GoTo EndMacro:
FNum = FreeFile


With ActiveSheet.UsedRange
StartRow = Application.InputBox("Input 1st Row Number to Export.", Type:=1) '.Cells(1).Row
StartCol = Application.InputBox("Input 1st Column *Number* to Export.", Type:=1) '.Cells(1).Column
EndRow = .Cells(.Cells.Count).Row
EndCol = .Cells(.Cells.Count).Column
End With

If AppendData = True Then
Open FName For Append Access Write As #FNum
Else
Open FName For Output Access Write As #FNum
End If

For RowNdx = StartRow To EndRow
WholeLine = ""
For ColNdx = StartCol To EndCol
CellValue = Cells(RowNdx, ColNdx).Value
WholeLine = WholeLine & CellValue & Sep
Next ColNdx
WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))

If SkipBlank = True Then
        Print #FNum, WholeLine
    Else
    If WorksheetFunction.CountA(ActiveSheet.Rows(RowNdx)) <> 0 Then
        Print #FNum, WholeLine
    End If
End If

Next RowNdx

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True
Close #FNum

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExportToDelimited
' This prompts the user for the FileName, separator, blanks, and start row/column
' then calls the ExportToTextFile procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ExportToDelimited()
Dim fileName As Variant
Dim Sep As String
Dim AppendDataInput As LongPtr
Dim AppendDataBool As Boolean
Dim varSkipBlank As LongPtr
Dim boolSkipBlank As Boolean
fileName = Application.GetSaveAsFilename(InitialFileName:=vbNullString, filefilter:="Text Files (*.txt),*.txt")
If fileName = False Then
''''''''''''''''''''''''''
' user cancelled, get out
''''''''''''''''''''''''''
Exit Sub
End If
Sep = Application.InputBox("Enter a separator character. Type ""Tab"" for Tab", Type:=2)  ''Sep = "|"
If Sep = "Tab" Then
    Sep = vbTab
End If
If Sep = vbNullString Then
''''''''''''''''''''''''''
' user cancelled, get out
''''''''''''''''''''''''''
Exit Sub
End If

AppendDataInput = MsgBox("Do you want to Append *Yes* or Overwrite *No* the file?", vbYesNo)
If AppendDataInput = vbYes Then
    AppendDataBool = True
    Else
    AppendDataBool = False
End If

varSkipBlank = MsgBox("Do you want to export blank rows?" & vbCr & "Yes = Export; No = Skip", vbYesNo)
If varSkipBlank = vbYes Then
    boolSkipBlank = True
    Else
    boolSkipBlank = False
End If

Debug.Print "FileName: " & fileName, "Separator: " & Sep
ExportToTextFile FName:=CStr(fileName), Sep:=CStr(Sep), AppendData:=AppendDataBool, SkipBlank:=boolSkipBlank
End Sub

Sub MRNETOPSUpload()
Application.ScreenUpdating = False

Dim wb As Workbook
Dim myFileName As String
Dim filePath As String
filePath = ActiveWorkbook.path
Dim FileNm As String
'FileNm = ActiveWorkbook.Name
FileNm = Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 1))

ActiveSheet.Copy
Range("A1:XFD1048576").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("A1").Select
ActiveSheet.Name = "MRNETOPS Upload"

'Can't upload blank rows - this loop looks for the year in the "Channel Code" column and if blank, deletes the row
IRow = 1
varColumn = Application.Match("Channel Code", Range(IRow & ":" & IRow), 0)
varUsedRows = ActiveSheet.UsedRange.Rows.Count
IRow = varUsedRows
Do While IRow <> 0
    If Cells(IRow, varColumn) = "" Then
        Range(IRow & ":" & IRow).EntireRow.Delete
    End If
    IRow = IRow - 1
Loop

''Find all zero-length text strings and clear contents
'For Each x In ActiveSheet.UsedRange.Cells
'    If x.Value = vbNullString Then
'        x.ClearContents
'    End If
'Next


myFileName = "Upload_" & FileNm & "_" & Format(Now(), "mmddyy")
ChDir filePath
Application.Dialogs(xlDialogSaveAs).Show myFileName, 51

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Reference for other FileTypes to use in xlDialogSaveAs command
'50 = .xlsb
'51 = .xlsx
'52 = .xlsm
'xlWorkbook = .xls
'xlTextWindows = .txt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = True
End Sub

Sub DisableAutoRecover()

    ' Check to see if the feature is enabled, if not, disable it.
    If ActiveWorkbook.EnableAutoRecover = True Then
        ActiveWorkbook.EnableAutoRecover = False
        'Application.StatusBar = "AutoRecover Disabled"
        MsgBox "The AutoRecover feature has been disabled."
    Else
        MsgBox "The AutoRecover feature is already disabled."
        'Application.StatusBar = "AutoRecover Disabled"
    End If

End Sub


Sub TableOfContents()
Dim ws As Worksheet

On Error GoTo ErrorCatch
varGoBack_ws = ActiveSheet.Name
varGoBack_rng = ActiveCell.Address
Sheets.Add.Name = "Table of Contents"
Range("A1").Value = "Worksheet Name"
Range("A1").Font.Bold = True

Range("A2").Select
For Each ws In Worksheets
    ActiveCell.Value = ws.Name
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
    Selection.Offset(1, 0).Select
Next
ActiveSheet.Range("A:A").Columns.EntireColumn.AutoFit
Range("A2").Select

For Each ws In Worksheets
    ws.Activate
    Sheets("Table of Contents").Move Before:=ActiveSheet
    Exit Sub
Next

ErrorCatch:

Application.DisplayAlerts = False
ActiveSheet.Delete
MsgBox "Unable to create a table of contents. You may already have one."

End Sub

'Subroutine to find and display all formulas with errors.
'Prompts user to scan the active worksheet or the entire workbook.
Sub Find_Formula_Errors()

    ' --- Variable Declarations ---
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetsToScan As Collection
    Dim resultsSht As Worksheet
    Dim errorRng As Range
    Dim cell As Range
    Dim userChoice As VbMsgBoxResult
    Dim resultsRow As Long
    Const RESULTS_SHEET_NAME As String = "Formula Errors Report"

    ' --- Initialization and Setup ---
    Set wb = ActiveWorkbook
    
    ' Exit if no workbook is open to scan.
    If wb Is Nothing Then
        MsgBox "There is no active workbook to scan.", vbExclamation, "Action Canceled"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    ' --- Get User Input for Scan Scope ---
    userChoice = MsgBox("Scan the entire workbook?" & vbCrLf & vbCrLf & _
                        "• Yes = Scan all worksheets" & vbCrLf & _
                        "• No = Scan only the active sheet" & vbCrLf & _
                        "• Cancel = Exit", _
                        vbYesNoCancel + vbQuestion, "Select Scan Scope")

    If userChoice = vbCancel Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- Create a collection of worksheets to process ---
    Set sheetsToScan = New Collection
    If userChoice = vbYes Then 'Scan entire workbook
        For Each ws In wb.Worksheets
            'Avoid scanning a previous report sheet
            If ws.Name <> RESULTS_SHEET_NAME Then
                sheetsToScan.Add ws
            End If
        Next ws
    Else 'Scan active sheet only
        sheetsToScan.Add wb.ActiveSheet
    End If

    ' --- Prepare the Results Sheet ---
    ' Delete old report sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(RESULTS_SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Add a new sheet for the results and set up headers
    Set resultsSht = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    resultsSht.Name = RESULTS_SHEET_NAME
    resultsRow = 1
    
    With resultsSht.Range("A1:D1")
        .Value = Array("Worksheet", "Cell Address", "Formula", "Error Value")
        .Font.Bold = True
    End With

    ' --- Main Processing Loop ---
    For Each ws In sheetsToScan
        ' Find all cells with formula errors in the current sheet
        Set errorRng = Nothing 'Reset range for each sheet
        On Error Resume Next
        Set errorRng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0

        If Not errorRng Is Nothing Then
            For Each cell In errorRng
                resultsRow = resultsRow + 1
                
                ' Column A: Worksheet Name
                resultsSht.Cells(resultsRow, "A").Value = ws.Name
                
                ' Column B: Cell Address (with Hyperlink)
                resultsSht.Hyperlinks.Add Anchor:=resultsSht.Cells(resultsRow, "B"), _
                                           Address:="", _
                                           SubAddress:="'" & ws.Name & "'!" & cell.Address, _
                                           TextToDisplay:=cell.Address(External:=False)
                
                ' Column C: The formula as text
                resultsSht.Cells(resultsRow, "C").Value = "'" & cell.Formula
                
                ' Column D: The resulting error value
                ' *** FIX: Use .Text property to get the visible error string ***
                resultsSht.Cells(resultsRow, "D").Value = cell.Text
            Next cell
        End If
    Next ws

    ' --- Finalization ---
    If resultsRow > 1 Then 'If any errors were found
        resultsSht.Columns("A:D").AutoFit
        resultsSht.Activate
    Else 'No errors found, so clean up
        Application.DisplayAlerts = False
        resultsSht.Delete
        Application.DisplayAlerts = True
        MsgBox "No formula errors were found in the selected scope.", vbInformation, "Scan Complete"
    End If

    Application.ScreenUpdating = True

End Sub

Sub CenterAcrossColumns()
    With Selection
        .UnMerge
        .HorizontalAlignment = xlCenterAcrossSelection
        .MergeCells = False
    End With
End Sub

Sub boldselection()
    Application.ScreenUpdating = False
    txt = ActiveCell.Value
    x = Len(txt)
    y = 1
    Z = True
    Do
        ActiveCell.Characters(Start:=y, Length:=1).Font.Bold = Z
        If Mid(txt, y, 1) = "-" Then
            Z = False
        End If
        If Mid(txt, y, 1) = Chr(10) Then
            If Mid(txt, y + 1, 4) = "   •" Then
                Z = False
            Else
                Z = True
            End If
        End If
        y = y + 1
    Loop Until y = x
    Application.ScreenUpdating = True
End Sub

Public Function NextText(Textrange As Range, OmitRange As Range) As String
Dim str_text As String
Dim cell As Range
Dim cell2 As Range

'X = OmitRange.Count
x = -1
y = -1
    For Each cell In Textrange
        If Left(cell.Value, 1) = "-" Then
            If InStr(1, NextText, "-") = 0 Then
                NextText = NextText & " " & cell.Value
            Else
                NextText = NextText & ", " & Replace(cell.Value, "- ", "")
            End If
        End If
        If Not Left(cell.Value, 1) = "-" And cell.Value <> "" Then
            If y = 0 Then Exit Function
        End If
        'Check to see if duplicated
            'If we have text then...
            If x = 1 Then
                'Reset Y
                y = 0
                For Each cell2 In OmitRange
                    'If any OmitRange cell is the same, y = -1
                    If cell2.Value = NextText Then
                        y = -1
                        x = -1
                    End If
                Next
            End If
        'Normal Routine
            If cell.Value <> "" And x <> 1 And Left(cell.Value, 1) <> "-" Then
                NextText = cell.Value
                x = 1
            End If
    Next
    If y = -1 Then NextText = ""
End Function
Public Function OneText(Textrange As Range) As String
Dim str_text As String
Dim cell As Range
Dim cell2 As Range
    'combines a range of text
    
    For Each cell In Textrange
        Z = Len(cell.Value)
        If cell.Value <> "" Then
            If Len(OneText) = 0 Then
                OneText = cell.Value
            Else
                OneText = OneText & "" & Chr(10) & "" & cell.Value
            End If
        End If

    Next
End Function
Public Function Concatenate_Text(Textrange As Range, TopCnt As Integer) As String
Dim str_text As String
Dim cell As Range
str_text = ""
    For Each cell In Textrange
        If Not str_text = "" Then
            str_text = str_text & ", " & cell.Value
        Else
            str_text = str_text & cell.Value
        End If
        TopCnt = TopCnt - 1
        If TopCnt = 0 Then Exit For
    Next
Concatenate_Text = Trim(str_text)
End Function


Option Explicit


Sub ListLinks()
    
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Dim wsResults As Worksheet
    Dim wsToScan As Worksheet
    Dim sheetsToScan As Collection
    Dim linkSources As Variant
    Dim source As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim fileName As String
    Dim resultsRow As Long
    Dim scopeChoice As VbMsgBoxResult
    
    Set wb = ActiveWorkbook
    
    'Check if any external links exist before proceeding
    linkSources = wb.linkSources(xlExcelLinks)
    If IsEmpty(linkSources) Then
        MsgBox "This workbook does not contain any external links.", vbInformation, "No Links Found"
        GoTo CleanUp
    End If
    
    'Ask user to define the scope: entire workbook or active sheet only
    scopeChoice = MsgBox("Do you want to scan the entire workbook?" & vbCrLf & vbCrLf & _
                         "   •  Click 'Yes' to scan all worksheets." & vbCrLf & _
                         "   •  Click 'No' to scan only the active sheet.", _
                         vbYesNoCancel, "Select Scan Scope")

    Set sheetsToScan = New Collection
    Select Case scopeChoice
        Case vbYes
            For Each wsToScan In wb.Worksheets
                sheetsToScan.Add wsToScan
            Next wsToScan
        Case vbNo
            sheetsToScan.Add wb.ActiveSheet
        Case vbCancel
            GoTo CleanUp
    End Select
    
    Set wsResults = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsResults.Name = "LinkList_" & Format(Now, "HH-mm-ss")
    
    With wsResults.Range("A1:D1")
        .Value = Array("Link FilePath", "Link FileName", "Reference Cell", "Ref Cell Formula")
        .Font.Bold = True
    End With
    
    resultsRow = 2 'Start writing results from the second row
    
    'Loop through each external link source found in the workbook
    For Each source In linkSources
        fileName = GetFileName(CStr(source)) 'Extract just the filename from the full path
        
        'Loop through the collection of sheets designated by the user
        For Each wsToScan In sheetsToScan
            'Ignore the results sheet to prevent self-referencing
            If wsToScan.Name <> wsResults.Name Then
                
                'Isolate only cells with formulas
                Set searchRange = Nothing
                On Error Resume Next
                Set searchRange = wsToScan.Cells.SpecialCells(xlCellTypeFormulas)
                On Error GoTo 0
                
                If Not searchRange Is Nothing Then
                    'Find the first cell containing the link's filename
                    Set foundCell = searchRange.Find(What:=fileName, LookIn:=xlFormulas, LookAt:=xlPart)
                    
                    If Not foundCell Is Nothing Then
                        firstAddress = foundCell.Address
                        
                        'Loop through all found cells until it circles back to the first one
                        Do
                            'Write link details to the results sheet
                            wsResults.Cells(resultsRow, 1).Value = source
                            wsResults.Cells(resultsRow, 2).Value = fileName
                            
                            'Add a hyperlink back to the source cell for easy navigation
                            wsResults.Hyperlinks.Add Anchor:=wsResults.Cells(resultsRow, 3), _
                                                     Address:="", _
                                                     SubAddress:="'" & wsToScan.Name & "'!" & foundCell.Address, _
                                                     TextToDisplay:=wsToScan.Name & "!" & foundCell.Address
                                                     
                            wsResults.Cells(resultsRow, 4).Value = "'" & foundCell.Formula
                            
                            resultsRow = resultsRow + 1
                            
                            'Find the next occurrence
                            Set foundCell = searchRange.FindNext(foundCell)
                            
                        Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
                    End If
                End If
            End If
        Next wsToScan
    Next source
    
    'AutoFit columns for readability if any results were found
    If resultsRow > 2 Then
        wsResults.Columns("A:D").AutoFit
    Else
        'If no links were found within the specified scope, add a note
        wsResults.Range("A2").Value = "No external links were found in the selected scope."
    End If
    
    wsResults.Activate

'Re-enable screen updating and exit the sub
CleanUp:
    Application.ScreenUpdating = True
    
End Sub

Private Function GetFileName(ByVal fullPath As String) As String
    
    Dim pos As Integer
    
    'Find the position of the last backslash or forward slash
    pos = InStrRev(fullPath, "\")
    If pos = 0 Then
        pos = InStrRev(fullPath, "/")
    End If
    
    'Extract the substring after the last slash
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        'If no slash is found, the path itself is the filename
        GetFileName = fullPath
    End If
    
End Function



Sub WorksheetSizes()
    Dim wks As Worksheet
    Dim c As Range
    Dim sFullFile As String
    Dim sReport As String
    Dim sWBName As String

    sReport = "Size Report"
    sWBName = "Erase Me.xls"
    sFullFile = ActiveWorkbook.path & _
      Application.PathSeparator & sWBName

    ' Add new worksheet to record sizes
    On Error Resume Next
    Set wks = Worksheets(sReport)
    If wks Is Nothing Then
        With ActiveWorkbook.Worksheets.Add(Before:=Worksheets(1))
            .Name = sReport
            .Range("A1").Value = "Worksheet Name"
            .Range("B1").Value = "Approximate Size"
        End With
    End If
    On Error GoTo 0
    With ActiveWorkbook.Worksheets(sReport)
        .Select
        .Range("A1").CurrentRegion.Offset(1, 0).ClearContents
        Set c = .Range("A2")
    End With

    Application.ScreenUpdating = False
    ' Loop through worksheets
    For Each wks In ActiveWorkbook.Worksheets
        If wks.Name <> sReport Then
            If wks.Visible = xlSheetHidden Then
                hidden_flag = 1
                wks.Visible = xlSheetVisible
            End If
            wks.Copy
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFullFile
            ActiveWorkbook.Close SaveChanges:=False
            Application.DisplayAlerts = True
            c.Offset(0, 0).Value = wks.Name
            c.Offset(0, 1).Value = FileLen(sFullFile)
            Set c = c.Offset(1, 0)
            Kill sFullFile
            If hidden_flag = 1 Then
                hidden_flag = 0
                wks.Visible = xlSheetHidden
            End If
        End If
    Next wks
    Application.ScreenUpdating = True
End Sub

Sub CopyExactFormulas()

    'If Application.CutCopyMode = False Then

    Dim rng_CopyRange As Range
    Set rng_CopyRange = Application.InputBox(Prompt:="Select Range you wish to copy.", Type:=8)
    
    'Else
    
    
       
    Dim rng_PasteCell As Range
    Set rng_PasteCell = Application.InputBox(Prompt:="Select Cell to Paste Output", Type:=8)
        
    var_RangeRows = rng_CopyRange.Rows.Count
    var_RangeCols = rng_CopyRange.Columns.Count
        
    Dim rng_PasteRange As Range
    Set rng_PasteRange = Range(rng_PasteCell, Cells(rng_PasteCell.Row + var_RangeRows - 1, rng_PasteCell.Column + var_RangeCols - 1))

Application.ScreenUpdating = False
    
    rng_CopyRange.Replace What:="=", Replacement:="#", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    rng_CopyRange.Copy
    rng_PasteRange.Select
    ActiveSheet.Paste
    
    rng_PasteRange.Replace What:="#", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    rng_CopyRange.Replace What:="#", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    rng_PasteCell.Select
    Application.CutCopyMode = False
    'End If
Application.ScreenUpdating = True
End Sub

Sub HighlightSelection()
'KW - to highlight/de-highlight the active cell for use during presentations to call attention to that cell

'test to see if highlight already exists
    On Error Resume Next
    Set varCell = ActiveSheet.Shapes("Cell Highlight")
    Set varVertical = ActiveSheet.Shapes("Vertical Highlight")
    Set varHorizontal = ActiveSheet.Shapes("Horizontal Highlight")
    On Error GoTo 0
'If any part of highlight exists, remove all highlights
    varCount = 0
    If IsEmpty(varCell) Then varCount = varCount Else varCount = varCount + 1
    If IsEmpty(varVertical) Then varCount = varCount Else varCount = varCount + 1
    If IsEmpty(varHorizontal) Then varCount = varCount Else varCount = varCount + 1
    
    If varCount <> 0 Then
        On Error Resume Next
'Loop through all Cell Highlight objects if multiple cells were selected in previous run
        For Each shp In ActiveSheet.Shapes
            If shp.Name = "Cell Highlight" Then
                ActiveSheet.Shapes("Cell Highlight").Delete
            End If
        Next shp
        ActiveSheet.Shapes("Vertical Highlight").Delete
        ActiveSheet.Shapes("Horizontal Highlight").Delete
        On Error GoTo 0
    Else

'If multiple cells are selected, only highlight each cell else do vertical and horizontal highlights also
If InStr(Selection.Address, ",") <> 0 Then
    varRangeLong = Selection.Address
    arrRange = Split(varRangeLong, ",")
    For Each hRange In arrRange
        varCellLeft = Range(hRange).Left
        varCellWidth = Range(hRange).Width
        varCellTop = Range(hRange).Top
        varCellHeight = Range(hRange).Height
        varCellBottom = varCellTop + varCellHeight
        varCellRight = varCellLeft + varCellWidth
    
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, varCellLeft, varCellTop, varCellWidth, varCellHeight)
            .Select
            .Name = "Cell Highlight"
            .Fill.Visible = msoFalse
            .Line.Visible = msoTrue
            .Line.Weight = 5
            .Line.ForeColor.RGB = RGB(64, 255, 64)
        End With
    
    Next hRange

Else
'Only one cell selected
'Grab relative position and dimensions of active cell
varCellLeft = Selection.Left
varCellWidth = Selection.Width
varCellTop = Selection.Top
varCellHeight = Selection.Height
varCellBottom = varCellTop + varCellHeight
varCellRight = varCellLeft + varCellWidth

'Create Cell highlight if not already present
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, varCellLeft, varCellTop, varCellWidth, varCellHeight)
        .Select
        .Name = "Cell Highlight"
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.Weight = 5
        .Line.ForeColor.RGB = RGB(64, 255, 64)
    End With
    
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, varCellLeft, 0, varCellWidth, varCellBottom)
        .Select
        .Name = "Vertical Highlight"
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(64, 255, 64)
    End With
    
    With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, varCellTop, varCellRight, varCellHeight)
        .Select
        .Name = "Horizontal Highlight"
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(64, 255, 64)
    End With
    End If
End If
    ActiveCell.Select
End Sub

Sub PeekaBoo()
'I'd like to thank Alex G. for the inspiration and three fingers of Oban Little Bay for the motivation.
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

If Evaluate("ISREF(PeekaBoo!A1)") = False Then
    Sheets.Add.Name = "PeekaBoo"
Else
varNamedRangeSection = "N"
'Run Unhide code where column C = Y
For r = 1 To Worksheets("PeekaBoo").UsedRange.Rows.Count
    'determine if PeekaBoo code is down to the Named Range section (handles WS and Shapes section differently)
    If Cells(r, 1) = "Range Scope" Then
        varNamedRangeSection = "Y"
    End If
    
    If Cells(r, 3) = "Y" Or Cells(r, 3) = "y" Then
        If varNamedRangeSection = "N" Then
            If Cells(r, 2) = "" Then
                'evaluate if worksheet exists; if false then it is a shape according to PeekaBoo logic
                If Evaluate("ISREF('" & Replace(Cells(r, 1).Value, "'", "") & "'!A1)") = True Then
                Worksheets(Cells(r, 1).Value).Visible = True
                End If
            Else
                Worksheets(Cells(r, 1).Value).Shapes(Cells(r, 2).Value).Visible = True
            End If
        Else
            strNamedRange = Cells(r, 2).Value
            'for some reason if the named range is scoped to a WS, the first single quote doesn't get picked up - add it back
            If InStr(strNamedRange, "'!") <> 0 And Left(strNamedRange, 1) <> "'" Then
                strNamedRange = "'" & strNamedRange
            End If
                ActiveWorkbook.Names(strNamedRange).Visible = True
        End If
    End If
Next

    Sheets("PeekaBoo").Select
    Sheets("PeekaBoo").Cells.Clear
End If


Cells(1, 1) = "To Unhide any of the objects, type 'Y' in column C and re-run macro.  This WS will be rebuilt."
Range("A1:C1").MergeCells = True

r = 3
Cells(r, 1) = "Worksheet"
Cells(r, 2) = "Shape Name"
Cells(r, 3) = "Unhide"
Cells(r, 1).Font.Bold = True
Cells(r, 2).Font.Bold = True
Cells(r, 3).Font.Bold = True
r = 4

'Find Hidden Shapes
For Each ws In ActiveWorkbook.Worksheets
    For Each s In ws.Shapes
        If s.Visible <> -1 Then
            Cells(r, 1) = ws.Name
            Cells(r, 2) = s.Name
            r = r + 1
            countShapes = 1
        End If
    Next s
Next ws
If countShapes <> 1 Then
    Cells(r, 1) = "No hidden Shapes found."
End If

'Find Hidden and VeryHidden Worksheets
r = r + 1
Cells(r, 1) = "Worksheets"
Cells(r, 3) = "Unhide"
Cells(r, 1).Font.Bold = True
Cells(r, 3).Font.Bold = True
r = r + 1

For Each ws In ActiveWorkbook.Worksheets
    If ws.Visible <> -1 Then
        Cells(r, 1) = ws.Name
        r = r + 1
        countWS = 1
    End If
Next ws
If countWS <> 1 Then
    Cells(r, 1) = "No hidden Worksheets found."
End If

'Find Hidden Named Ranges
r = r + 1
Cells(r, 1) = "Range Scope"
Cells(r, 2) = "Named Range"
Cells(r, 3) = "Unhide"
Cells(r, 1).Font.Bold = True
Cells(r, 2).Font.Bold = True
Cells(r, 3).Font.Bold = True
r = r + 1

For Each N In ActiveWorkbook.Names
    If N.Visible <> -1 Then
        If InStr(N.Name, "'!") <> 0 Then
            Cells(r, 1) = Replace(Replace(Left(N.Name, InStr(N.Name, "!")), "'", ""), "!", "")
        Else
            Cells(r, 1) = "Global"
        End If
        'Cells(r, 1) = n.Name
        Cells(r, 2) = N.Name
        r = r + 1
        countNames = 1
    End If
Next N
'Next ws
If countNames <> 1 Then
    Cells(r, 1) = "No hidden Named Ranges found."
End If


'fit columns to size
Worksheets("PeekaBoo").Range("A:B").Columns.EntireColumn.AutoFit
Worksheets("PeekaBoo").Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'Applies color from one column to all cells to the left of that column
Sub ColorColumn()
Application.ScreenUpdating = False

rngColorColumnLetter = InputBox("Which Column (e.g. AY) contains the colors you want to apply?")
rngColorColumnNumber = Range(rngColorColumnLetter & 1).Column
rngColorNonNumeric = MsgBox("Do you want to apply color from all cells in " & rngColorColumnLetter & " (Yes) or only where values in " & rngColorColumnLetter & " are NUMERIC (No)?", vbYesNo)
rngColorRows = MsgBox("Do you want to apply color to all active rows (Yes) or select certain rows (No)?", vbYesNo)

If rngColorRows = vbNo Then
    Set rngColorRows = Application.InputBox("Select range to apply colors:", Type:=8)
    rngMin = rngColorRows.Rows(1).Row
    rngMax = rngMin + rngColorRows.Rows.Count - 1

Else
    rngMin = 1
    rngMax = rngMin + ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1
End If
    

For r = rngMin To rngMax
    If rngColorNonNumeric = vbYes Then
        Range(Cells(r, 1).Address, Cells(r, rngColorColumnNumber - 1).Address).Interior.Color = Cells(r, rngColorColumnNumber).DisplayFormat.Interior.Color
    ElseIf Not IsEmpty(Cells(r, rngColorColumnNumber)) Then
        If IsNumeric(Cells(r, rngColorColumnNumber).Value) Then
            Range(Cells(r, 1).Address, Cells(r, rngColorColumnNumber - 1).Address).Interior.Color = Cells(r, rngColorColumnNumber).DisplayFormat.Interior.Color
        End If
    End If
Next
Application.ScreenUpdating = True

End Sub

Sub Merlin_Support()
'opens Merlin Add-In forum in default browser
    ActiveWorkbook.FollowHyperlink "https://github.com/Lonestarjeepin/Merlin"

End Sub

Function List_Unique(rng As Range)
'From www.get-digital-help.com
'Must be entered as a CSE Array

Dim cell As Range, temp() As String, i As Single, iRows As Integer
ReDim temp(0)
For Each cell In rng
  For i = LBound(temp) To UBound(temp)
    If temp(i) = cell Then
       i = i + 1
       Exit For
    End If
  Next i
  i = i - 1
  If temp(i) <> cell Then
    temp(UBound(temp)) = cell
    ReDim Preserve temp(UBound(temp) + 1)
  End If
Next cell
iRows = Range(Application.Caller.Address).Rows.Count
If iRows < UBound(temp) Then
  temp(iRows - 1) = "More values.."
Else
  For i = UBound(temp) To iRows
    ReDim Preserve temp(UBound(temp) + 1)
    temp(UBound(temp)) = ""
  Next i
End If
List_Unique = Application.Transpose(temp)
End Function

Sub View_All_WorkSheets()
'11.26.18 Contributed by Maddy M.

Application.CommandBars("Workbook tabs").ShowPopup

End Sub

Sub EvaluateAsFormula()
'02.14.19 Contributed by Daniel W.

    Application.ScreenUpdating = False
    Dim cell As Range
    
    For Each cell In Selection
        If cell.HasFormula = False Then
            cell.NumberFormat = "General"
            cell.Formula = cell.Value
        End If
    Next
    Application.ScreenUpdating = True
End Sub

Sub ConvertToFormula()
' 01.15.2014 - Kyle W.
Application.ScreenUpdating = False
    Dim cell As Range
    
    For Each cell In Selection
        If cell.HasFormula = False Then
            cell.Formula = "=""" & cell.Formula & """"
        End If
    Next
Application.ScreenUpdating = True
End Sub

Sub Round()
' 01.15.2014 - Kyle W.
Application.ScreenUpdating = False
    Dim cell As Range
    
    varRound = InputBox("How many decimal places do you want to round to?")
    
    For Each cell In Selection
        If cell.HasFormula = True Then
            cell.Formula = "=round(" & Right(cell.Formula, Len(cell.Formula) - 1) & "," & varRound & ")"
        End If
    Next
Application.ScreenUpdating = True
End Sub

Sub Ordinal_Format()
'Adds ordinal indicator to the end of number via formatting
Application.ScreenUpdating = False
Dim cell As Range

For Each cell In Selection
    varNumber = CLngPtr(cell)
    Select Case Right(varNumber, 1)
        Case 1
        cell.NumberFormat = "#,##0""st""_);(#,##0""st"");"" - ""_)"
        '"#,##0""st""_);(#,##0""st"");"" - ""_)"
        Case 2
        cell.NumberFormat = "#,##0""nd""_);(#,##0""nd"");"" - ""_)"
        Case 3
        cell.NumberFormat = "#,##0""rd""_);(#,##0""rd"");"" - ""_)"
        Case Else
        cell.NumberFormat = "#,##0""th""_);(#,##0""th"");"" - ""_)"
    End Select

    varNumber = CLngPtr(cell)
    Select Case Right(varNumber, 2)
        Case 11, 12, 13
        cell.NumberFormat = "#,##0""th""_);(#,##0""th"");"" - ""_)"
    End Select
Next
Application.ScreenUpdating = True
End Sub

Sub Unhide_Rehide_WS()
'03.01.19 - Base code and inspiration (unhide) from Ryan C., but re-hide from Kyle W.
'Gets a list of hidden WSs and unhides them.  When re-run, the macro re-hides the same WSs

'check if array is blank - if not, then proceed with unhide

Application.ScreenUpdating = False

'Check to see if collection is empty (not holding any hidden ws names)
If colHiddenWS.Count = 0 Then
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        'unhides any hidden ws and adds to collection
        If ws.Visible = xlSheetHidden Then
            ws.Visible = xlSheetVisible
            colHiddenWS.Add ws.Name
        End If
    Next ws
Else
'if collection is NOT empty, then proceed with rehide
    'hide previously unhidden WS
    For i = 1 To colHiddenWS.Count
        wsName = colHiddenWS(i)
        Worksheets(wsName).Visible = xlSheetHidden
    Next i
    'reset array to blank
    Set colHiddenWS = New Collection
End If

Application.ScreenUpdating = True
End Sub


'*********************************
'******Calculation Timers*********
'*********************************


'
Function MicroTimer() As Double
'

' Returns seconds.
'
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0

' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency

' Get ticks.
    getTickCount cyTicks1

' Seconds
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function



Sub RangeTimer()
    DoCalcTimer 1
End Sub
Sub SheetTimer()
    DoCalcTimer 2
End Sub
Sub RecalcTimer()
    DoCalcTimer 3
End Sub
Sub FullcalcTimer()
    DoCalcTimer 4
End Sub

Sub DoCalcTimer(jMethod As LongPtr)
    Dim dTime As Double
    Dim dOvhd As Double
    Dim oRng As Range
    Dim oCell As Range
    Dim oArrRange As Range
    Dim sCalcType As String
    Dim lCalcSave As XlCalculation
    Dim bIterSave As Boolean
    '
    On Error GoTo Errhandl

' Initialize
    dTime = MicroTimer

    ' Save calculation settings.
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    Select Case jMethod
    Case 1

        ' Switch off iteration.

        If Application.Iteration <> False Then
            Application.Iteration = False
        End If
        
        ' Max is used range.

        If Selection.Count > 1000 Then
            Set oRng = Intersect(Selection, Selection.Parent.UsedRange)
        Else
            Set oRng = Selection
        End If

        ' Include array cells outside selection.

        For Each oCell In oRng
            If oCell.HasArray Then
                If oArrRange Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                End If
                If Intersect(oCell, oArrRange) Is Nothing Then
                    Set oArrRange = oCell.CurrentArray
                    Set oRng = Union(oRng, oArrRange)
                End If
            End If
        Next oCell

        sCalcType = "Calculate " & CStr(oRng.Count) & _
            " Cell(s) in Selected Range (Avg of 5 Iterations): "
    Case 2
        sCalcType = "Recalculate Sheet " & ActiveSheet.Name & " (Avg of 5 Iterations): "
    Case 3
        sCalcType = "Recalculate open workbooks (Avg of 5 Iterations): "
    Case 4
        sCalcType = "Full Calculate open workbooks (Avg of 5 Iterations): "
    End Select

' Get start time.
    dTime = MicroTimer
    Select Case jMethod
    Case 1
        If Val(Application.VERSION) >= 12 Then
            oRng.CalculateRowMajorOrder
            oRng.CalculateRowMajorOrder
            oRng.CalculateRowMajorOrder
            oRng.CalculateRowMajorOrder
            oRng.CalculateRowMajorOrder
        Else
            oRng.Calculate
            oRng.Calculate
            oRng.Calculate
            oRng.Calculate
            oRng.Calculate
        End If
    Case 2
        ActiveSheet.Calculate
        ActiveSheet.Calculate
        ActiveSheet.Calculate
        ActiveSheet.Calculate
        ActiveSheet.Calculate
    Case 3
        Application.Calculate
        Application.Calculate
        Application.Calculate
        Application.Calculate
        Application.Calculate
    Case 4
        Application.CalculateFull
        Application.CalculateFull
        Application.CalculateFull
        Application.CalculateFull
        Application.CalculateFull
    End Select

' Calc duration.
    dTime = MicroTimer - dTime

    On Error GoTo 0

    dTime = Application.WorksheetFunction.Round(dTime, 5) / 5
    MsgBox sCalcType & " " & CStr(dTime) & " Seconds", _
        vbOKOnly + vbInformation, "CalcTimer"

Finish:

    ' Restore calculation settings.
    If Application.Calculation <> lCalcSave Then
         Application.Calculation = lCalcSave
    End If
    If Application.Iteration <> bIterSave Then
         Application.Calculation = bIterSave
    End If
    Exit Sub
Errhandl:
    On Error GoTo 0
    MsgBox "Unable to Calculate " & sCalcType, _
        vbOKOnly + vbCritical, "CalcTimer"
    GoTo Finish
End Sub
'************************************* End Calculation Timers *********************************************


Sub AllInternalPasswords()
        ' Breaks worksheet and workbook structure passwords. Bob McCormick
        '  probably originator of base code algorithm modified for coverage
        '  of workbook structure / windows passwords and for multiple passwords
        '
        ' Norman Harker and JE McGimpsey 27-Dec-2002 (Version 1.1)
        ' Modified 2003-Apr-04 by JEM: All msgs to constants, and
        '   eliminate one Exit Sub (Version 1.1.1)
        ' Reveals hashed passwords NOT original passwords
        Const DBLSPACE As String = vbNewLine & vbNewLine
        Const AUTHORS As String = DBLSPACE & vbNewLine & _
                "Adapted from Bob McCormick base code by" & _
                "Norman Harker and JE McGimpsey"
        Const HEADER As String = "AllInternalPasswords User Message"
        Const VERSION As String = DBLSPACE & "Version 1.1.1 2003-Apr-04"
        Const REPBACK As String = DBLSPACE & "Please report failure " & _
                "to the microsoft.public.excel.programming newsgroup."
        Const ALLCLEAR As String = DBLSPACE & "The workbook should " & _
                "now be free of all password protection, so make sure you:" & _
                DBLSPACE & "SAVE IT NOW!" & DBLSPACE & "and also" & _
                DBLSPACE & "BACKUP!, BACKUP!!, BACKUP!!!" & _
                DBLSPACE & "Also, remember that the password was " & _
                "put there for a reason. Don't stuff up crucial formulas " & _
                "or data." & DBLSPACE & "Access and use of some data " & _
                "may be an offense. If in doubt, don't."
        Const MSGNOPWORDS1 As String = "There were no passwords on " & _
                "sheets, or workbook structure or windows." & AUTHORS & VERSION
        Const MSGNOPWORDS2 As String = "There was no protection to " & _
                "workbook structure or windows." & DBLSPACE & _
                "Proceeding to unprotect sheets." & AUTHORS & VERSION
        Const MSGTAKETIME As String = "After pressing OK button this " & _
                "will take some time." & DBLSPACE & "Amount of time " & _
                "depends on how many different passwords, the " & _
                "passwords, and your computer's specification." & DBLSPACE & _
                "Just be patient! Make me a coffee!" & AUTHORS & VERSION
        Const MSGPWORDFOUND1 As String = "You had a Worksheet " & _
                "Structure or Windows Password set." & DBLSPACE & _
                "The password found was: " & DBLSPACE & "$$" & DBLSPACE & _
                "Note it down for potential future use in other workbooks by " & _
                "the same person who set this password." & DBLSPACE & _
                "Now to check and clear other passwords." & AUTHORS & VERSION
        Const MSGPWORDFOUND2 As String = "You had a Worksheet " & _
                "password set." & DBLSPACE & "The password found was: " & _
                DBLSPACE & "$$" & DBLSPACE & "Note it down for potential " & _
                "future use in other workbooks by same person who " & _
                "set this password." & DBLSPACE & "Now to check and clear " & _
                "other passwords." & AUTHORS & VERSION
        Const MSGONLYONE As String = "Only structure / windows " & _
                 "protected with the password that was just found." & _
                 ALLCLEAR & AUTHORS & VERSION & REPBACK
        Dim w1 As Worksheet, w2 As Worksheet
        Dim i As Integer, j As Integer, k As Integer, l As Integer
        Dim m As Integer, N As Integer, i1 As Integer, i2 As Integer
        Dim i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
        Dim PWord1 As String
        Dim ShTag As Boolean, WinTag As Boolean
        
        Application.ScreenUpdating = False
        With ActiveWorkbook
            WinTag = .ProtectStructure Or .ProtectWindows
        End With
        ShTag = False
        For Each w1 In Worksheets
                ShTag = ShTag Or w1.ProtectContents
        Next w1
        If Not ShTag And Not WinTag Then
            MsgBox MSGNOPWORDS1, vbInformation, HEADER
            Exit Sub
        End If
        MsgBox MSGTAKETIME, vbInformation, HEADER
        If Not WinTag Then
            MsgBox MSGNOPWORDS2, vbInformation, HEADER
        Else
          On Error Resume Next
          Do      'dummy do loop
            For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
            For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
            For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
            For i5 = 65 To 66: For i6 = 65 To 66: For N = 32 To 126
            With ActiveWorkbook
              .Unprotect Chr(i) & Chr(j) & Chr(k) & _
                 Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                 Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)
              If .ProtectStructure = False And _
              .ProtectWindows = False Then
                  PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
                    Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                    Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)
                  MsgBox Application.Substitute(MSGPWORDFOUND1, _
                        "$$", PWord1), vbInformation, HEADER
                  Exit Do  'Bypass all for...nexts
              End If
            End With
            Next: Next: Next: Next: Next: Next
            Next: Next: Next: Next: Next: Next
          Loop Until True
          On Error GoTo 0
        End If
        If WinTag And Not ShTag Then
          MsgBox MSGONLYONE, vbInformation, HEADER
          Exit Sub
        End If
        On Error Resume Next
        For Each w1 In Worksheets
          'Attempt clearance with PWord1
          w1.Unprotect PWord1
        Next w1
        On Error GoTo 0
        ShTag = False
        For Each w1 In Worksheets
          'Checks for all clear ShTag triggered to 1 if not.
          ShTag = ShTag Or w1.ProtectContents
        Next w1
        If ShTag Then
            For Each w1 In Worksheets
              With w1
                If .ProtectContents Then
                  On Error Resume Next
                  Do      'Dummy do loop
                    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
                    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
                    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
                    For i5 = 65 To 66: For i6 = 65 To 66: For N = 32 To 126
                    .Unprotect Chr(i) & Chr(j) & Chr(k) & _
                      Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                      Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)
                    If Not .ProtectContents Then
                      PWord1 = Chr(i) & Chr(j) & Chr(k) & Chr(l) & _
                        Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                        Chr(i4) & Chr(i5) & Chr(i6) & Chr(N)
                      MsgBox Application.Substitute(MSGPWORDFOUND2, _
                            "$$", PWord1), vbInformation, HEADER
                      'leverage finding Pword by trying on other sheets
                      For Each w2 In Worksheets
                        w2.Unprotect PWord1
                      Next w2
                      Exit Do  'Bypass all for...nexts
                    End If
                    Next: Next: Next: Next: Next: Next
                    Next: Next: Next: Next: Next: Next
                  Loop Until True
                  On Error GoTo 0
                End If
              End With
            Next w1
        End If
        MsgBox ALLCLEAR & AUTHORS & VERSION & REPBACK, vbInformation, HEADER
    End Sub
    
'**************************************** End Crack Passwords ********************************************





'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'***********************************          Scale Chart Macro Start       *******************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'Original Code by Jon Peltier, modified by Kyle Whitmire as annotated KW
'**********************************************************************************************************************
'Option Explicit



Public Function fnAxisScale(ByVal dMin As Double, ByVal dMax As Double) As scaleAxisScale
  ' Calculates tidy settings for the chart axes
  Dim dPower As Double, dScale As Double, dSmall As Double, dTemp As Double

  'Check if the max and min are the same
  If dMax = dMin Then
    dTemp = dMax
    dMax = dMax * 1.01
    dMin = dMin * 0.99
  End If

  'Check if dMax is bigger than dMin - swap them if not
  If dMax < dMin Then
    dTemp = dMax
    dMax = dMin
    dMin = dTemp
  End If

  'Make dMax a little bigger and dMin a little smaller (by 1% of their difference)
  If dMax > 0 Then
    dMax = dMax + (dMax - dMin) * 0.01
  ElseIf dMax < 0 Then
    dMax = WorksheetFunction.Min(dMax + (dMax - dMin) * 0.01, 0)
  Else
    dMax = 0
  End If
  If dMin > 0 Then
    dMin = WorksheetFunction.Max(dMin - (dMax - dMin) * 0.01, 0)
  ElseIf dMin < 0 Then
    dMin = dMin - (dMax - dMin) * 0.01
  Else
    dMin = 0
  End If

  'What if they are both 0?
  If (dMax = 0) And (dMin = 0) Then dMax = 1

  'This bit rounds the maximum and minimum values to reasonable values
  'to chart.  If not done, the axis numbers will look very silly
  'Find the range of values covered
  dPower = Log(dMax - dMin) / Log(10)
  dScale = 10 ^ (dPower - Int(dPower))

  'Find the scaling factor
  Select Case dScale
    Case 0 To 2.5
      dScale = 0.2
      dSmall = 0.1
    Case 2.5 To 5
      dScale = 0.5
      dSmall = 0.1
    Case 5 To 7.5
      dScale = 1
      dSmall = 0.2
    Case Else
      dScale = 2
      dSmall = 0.5
  End Select

  'Calculate the scaling factor (major & minor unit)
  dScale = dScale * 10 ^ Int(dPower)
  dSmall = dSmall * 10 ^ Int(dPower)

  'Round the axis values to the nearest scaling factor
  fnAxisScale.dMin = dScale * Int(dMin / dScale)
  fnAxisScale.dMax = dScale * (Int(dMax / dScale) + 1)
  'KW - replaced dScale with "hardcode" to show 5 tickmarks on each vertical axis
  fnAxisScale.dMajor = (fnAxisScale.dMax - fnAxisScale.dMin) / 5 'dScale
  fnAxisScale.dMinor = dSmall

End Function

Public Function udfAxisScale(ByVal dMin As Double, ByVal dMax As Double) As Variant
  ' Worksheet interface to fnAxisScale
  ' Returns a horizontal array to the worksheet
  Dim scaleMyScale As scaleAxisScale
  Dim scaleOutput As Variant

  scaleMyScale = fnAxisScale(dMin, dMax)

  ReDim scaleOutput(1 To 4)
  scaleOutput(1) = scaleMyScale.dMin
  scaleOutput(2) = scaleMyScale.dMax
  scaleOutput(3) = scaleMyScale.dMajor
  scaleOutput(4) = scaleMyScale.dMinor

  udfAxisScale = scaleOutput
End Function

Sub ScaleActiveSheetCharts()
  Dim cht As ChartObject
  Application.ScreenUpdating = False
  For Each cht In ActiveSheet.ChartObjects
    ScaleChartAxes cht.Chart
  Next
  Application.ScreenUpdating = True
End Sub
Sub ScaleActiveChart()
  If Not ActiveChart Is Nothing Then
    ScaleChartAxes ActiveChart
  End If
End Sub

Sub ScaleChartAxes(cht As Chart)
  
  Dim AxisScaleP As scaleAxisScale
  Dim AxisScaleS As scaleAxisScale
  Dim dSMin As Double, dSMax As Double, dPMin As Double, dPMax As Double
  Dim vSValues As Variant, vPValues As Variant
  Dim iSeries As LongPtr, iPoint As LongPtr
  Dim srs As Series

  'KW - set min and max to non-zero values that are sure to be overwritten - was having issues with values between -1 and 1.
  dSMin = 99999999999#
  dSMax = -99999999999#
  dPMin = 99999999999#
  dPMax = -99999999999#
  TwoAxisToggle = 0
  
  With cht
    
    ' loop through all series and all points to find min and max
    For iSeries = 1 To .SeriesCollection.Count
      Set srs = .SeriesCollection(iSeries)
        'find Primary Axis values
        If .SeriesCollection(iSeries).AxisGroup = xlPrimary Then
          vPValues = srs.values

            If iSeries = 1 Then
                dPMin = vPValues(1)
                dPMax = vPValues(1)
            End If

            'KW - Empty qualifier added to ignore "#N/A" points on graph
            For iPoint = 1 To srs.Points.Count
                If dPMin > vPValues(iPoint) Then If vPValues(iPoint) <> Empty Then dPMin = vPValues(iPoint)
                If dPMax < vPValues(iPoint) Then If vPValues(iPoint) <> Empty Then dPMax = vPValues(iPoint)
            Next
        End If
        'KW - find Secondary axis values
        If .SeriesCollection(iSeries).AxisGroup = xlSecondary Then
          vSValues = srs.values

            If iSeries = 1 Then
                dSMin = vSValues(1)
                dSMax = vSValues(1)
            End If

            For iPoint = 1 To srs.Points.Count
                If dSMin > vSValues(iPoint) Then
                    If vSValues(iPoint) <> Empty Then
                        dSMin = vSValues(iPoint)
                    End If
                End If
                If dSMax < vSValues(iPoint) Then
                    If vSValues(iPoint) <> Empty Then
                        dSMax = vSValues(iPoint)
                    End If
                End If
            Next
            TwoAxisToggle = 1
        End If
    Next

    'KW - If the text "NoScale" is present anywhere in chartobject name, scaling won't be applied.  Allows way to ignore certain charts when executing sub for all chartobjects
    If InStr(cht.Name, "NoScale") = 0 Then
        'DW - if the text "Split" appears in the chart name, make sure that the lines do not intersect
        If InStr(cht.Name, "Split") <> 0 Then
            If TwoAxisToggle = 1 Then
                dPMin = dPMax - 2 * (dPMax - dPMin)
                dSMax = dSMin + 2 * (dSMax - dSMin)
            End If
        End If
        'DW - build in slightly more white
        dPMax = (dPMax - dPMin) * 0.2 + dPMax
        dSMax = (dSMax - dSMin) * 0.2 + dSMax
        If dPMin > 0 Then
            dPMin = dPMin - (dPMax - dPMin) * 0.2
            If dPMin < 0 Then dPMin = 0
        End If
        If dSMin > 0 Then
            dSMin = dSMin - (dSMax - dSMin) * 0.2
            If dSMin < 0 Then dSMin = 0
        End If
        
        ' compute axis scales
        AxisScaleP = fnAxisScale(dPMin, dPMax)
        AxisScaleS = fnAxisScale(dSMin, dSMax)
        
        
        ' apply Primary axis scale
        With .Axes(xlValue, xlPrimary)
          If .MinimumScale > AxisScaleP.dMax Then
            .MaximumScale = AxisScaleP.dMax
            .MinimumScale = AxisScaleP.dMin
          Else
            .MinimumScale = AxisScaleP.dMin
            .MaximumScale = AxisScaleP.dMax
          End If
          .MajorUnit = AxisScaleP.dMajor
        End With
        ' apply Secondary axis Scale
        If cht.HasAxis(xlValue, xlSecondary) = True Then
            With .Axes(xlValue, xlSecondary)
                If .MinimumScale > AxisScaleS.dMax Then
                    .MaximumScale = AxisScaleS.dMax
                    .MinimumScale = AxisScaleS.dMin
                Else
                    .MinimumScale = AxisScaleS.dMin
                    .MaximumScale = AxisScaleS.dMax
                End If
                    .MajorUnit = AxisScaleS.dMajor
            End With
        End If
    End If

'KW - center chart title if present
    If cht.HasTitle = True Then
        cht.ChartTitle.Left = cht.ChartArea.Width
        cht.ChartTitle.Left = cht.ChartTitle.Left / 2
    End If

  End With
End Sub

'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'***********************************          Scale Chart Macro End       *********************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************

Sub HighlightContributingCells()
    Dim formulaCell As Range
    Dim formulaText As String
    Dim sumRange As Range
    Dim criteriaRange As Range
    Dim criteriaRanges() As Range
    Dim criteriaOperators() As String
    Dim criteriaValues() As String
    Dim cell As Range
    Dim formulaParts() As String
    Dim i As Long, j As Long, k As Long
    Dim sumRangeAddress As String
    Dim criteriaRangeAddress As String
    Dim criteria As String
    Dim operatorPos As Long
    Dim comparisonOperator As String
    Dim criteriaValue As String
    Dim validCriteria As Boolean
    Dim isSUMIF As Boolean
    Dim isCOUNTIF As Boolean
    
    ' Get user input for the formula cell
    On Error Resume Next
    Set formulaCell = Application.InputBox("Select the cell containing the SUMIF/SUMIFS/COUNTIF/COUNTIFS formula:", Type:=8)
    On Error GoTo 0
    If formulaCell Is Nothing Then Exit Sub
    
    ' Extract formula text
    formulaText = formulaCell.Formula
    
    ' Debug: Print formula text
    Debug.Print "Formula Text: " & formulaText
    
    ' Check if the formula is SUMIF, SUMIFS, COUNTIF, or COUNTIFS
    isSUMIF = InStr(formulaText, "SUMIF(") > 0 And InStr(formulaText, "SUMIFS(") = 0
    isCOUNTIF = InStr(formulaText, "COUNTIF(") > 0 And InStr(formulaText, "COUNTIFS(") = 0
    
    If isSUMIF Or isCOUNTIF Then
        ' Handle SUMIF or COUNTIF
        formulaParts = Split(Mid(formulaText, InStr(1, formulaText, "(") + 1, Len(formulaText) - InStr(1, formulaText, "(") - 1), ",")
        
        ' Extract criteria range (first argument)
        criteriaRangeAddress = Trim(formulaParts(0))
        Set criteriaRange = Range(criteriaRangeAddress)
        ' Debug: Print criteria range address
        Debug.Print "Criteria Range: " & criteriaRange.Address
        
        ' Extract criteria (second argument)
        criteria = Trim(formulaParts(1))
        criteria = Replace(criteria, """", "")
        
        ' Extract sum range (third argument, optional for SUMIF)
        If isSUMIF And UBound(formulaParts) >= 2 Then
            sumRangeAddress = Trim(formulaParts(2))
            Set sumRange = Range(sumRangeAddress)
        Else
            Set sumRange = criteriaRange
        End If
        ' Debug: Print sum range address
        Debug.Print "Sum Range: " & sumRange.Address
        
        ' Identify and extract the comparison operator and criteria value
        If Left(criteria, 2) = "<>" Or Left(criteria, 2) = "<=" Or Left(criteria, 2) = ">=" Then
            comparisonOperator = Left(criteria, 2)
            criteriaValue = Mid(criteria, 3)
        ElseIf Left(criteria, 1) = "=" Or Left(criteria, 1) = "<" Or Left(criteria, 1) = ">" Then
            comparisonOperator = Left(criteria, 1)
            criteriaValue = Mid(criteria, 2)
        Else
            comparisonOperator = "="
            criteriaValue = criteria
        End If
        
        ' Debug: Print comparison operator and criteria value
        Debug.Print "Comparison Operator: " & comparisonOperator
        Debug.Print "Criteria Value: " & criteriaValue
        
        ' Evaluate and highlight cells
        For i = 1 To criteriaRange.Cells.Count
            ' Compare values, considering both numeric and string comparisons
            Select Case comparisonOperator
                Case "="
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value = Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value = criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
                Case "<"
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value < Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value < criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
                Case ">"
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value > Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value > criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
                Case "<="
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value <= Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value <= criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
                Case ">="
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value >= Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value >= criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
                Case "<>"
                    If IsNumeric(criteriaValue) Then
                        If criteriaRange.Cells(i).Value <> Val(criteriaValue) Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    Else
                        If criteriaRange.Cells(i).Value <> criteriaValue Then
                            criteriaRange.Rows(i).Interior.Color = vbGreen
                        End If
                    End If
            End Select
        Next i
    ElseIf InStr(formulaText, "SUMIFS(") > 0 Or InStr(formulaText, "COUNTIFS(") > 0 Then
        ' Handle SUMIFS or COUNTIFS
        formulaParts = Split(Mid(formulaText, InStr(1, formulaText, "(") + 1, Len(formulaText) - InStr(1, formulaText, "(") - 1), ",")
        
        ' Extract sum range (first argument for SUMIFS, not used for COUNTIFS)
        If InStr(formulaText, "SUMIFS(") > 0 Then
            sumRangeAddress = Trim(formulaParts(0))
            Set sumRange = Range(sumRangeAddress)
            ' Debug: Print sum range address
            Debug.Print "Sum Range: " & sumRange.Address
        Else
            Set sumRange = Nothing
        End If
        
        ' Initialize arrays
        ReDim criteriaRanges((UBound(formulaParts) - 1) \ 2)
        ReDim criteriaOperators((UBound(formulaParts) - 1) \ 2)
        ReDim criteriaValues((UBound(formulaParts) - 1) \ 2)
        
        ' Extract criteria ranges and criteria values
        For i = IIf(InStr(formulaText, "SUMIFS(") > 0, 1, 0) To UBound(formulaParts) Step 2
            criteriaRangeAddress = Trim(formulaParts(i))
            Set criteriaRanges((i - IIf(InStr(formulaText, "SUMIFS(") > 0, 1, 0)) \ 2) = Range(criteriaRangeAddress)
            ' Debug: Print criteria range address
            Debug.Print "Criteria Range: " & criteriaRanges((i - IIf(InStr(formulaText, "SUMIFS(") > 0, 1, 0)) \ 2).Address
            
            criteria = Trim(formulaParts(i + 1))
            criteria = Replace(criteria, """", "")
            
            ' Identify and extract the comparison operator and criteria value
            If Left(criteria, 2) = "<>" Or Left(criteria, 2) = "<=" Or Left(criteria, 2) = ">=" Then
                comparisonOperator = Left(criteria, 2)
                criteriaValue = Mid(criteria, 3)
            ElseIf Left(criteria, 1) = "=" Or Left(criteria, 1) = "<" Or Left(criteria, 1) = ">" Then
                comparisonOperator = Left(criteria, 1)
                criteriaValue = Mid(criteria, 2)
            Else
                comparisonOperator = "="
                criteriaValue = criteria
            End If
            
            criteriaOperators((i - IIf(InStr(formulaText, "SUMIFS(") > 0, 1, 0)) \ 2) = comparisonOperator
            criteriaValues((i - IIf(InStr(formulaText, "SUMIFS(") > 0, 1, 0)) \ 2) = criteriaValue
            
            ' Debug: Print comparison operator and criteria value
            Debug.Print "Comparison Operator: " & comparisonOperator
            Debug.Print "Criteria Value: " & criteriaValue
        Next i
        
        ' Evaluate and highlight cells
        For i = 1 To criteriaRanges(0).Cells.Count
            validCriteria = True
            For j = 0 To UBound(criteriaRanges)
                Set criteriaRange = criteriaRanges(j)
                comparisonOperator = criteriaOperators(j)
                criteriaValue = criteriaValues(j)
                
                ' Compare values, considering both numeric and string comparisons
                Select Case comparisonOperator
                    Case "="
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value = Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value = criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                    Case "<"
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value < Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value < criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                    Case ">"
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value > Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value > criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                    Case "<="
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value <= Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value <= criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                    Case ">="
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value >= Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value >= criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                    Case "<>"
                        If IsNumeric(criteriaValue) Then
                            If Not criteriaRange.Cells(i).Value <> Val(criteriaValue) Then
                                validCriteria = False
                            End If
                        Else
                            If Not criteriaRange.Cells(i).Value <> criteriaValue Then
                                validCriteria = False
                            End If
                        End If
                End Select
                If Not validCriteria Then Exit For
            Next j
            If validCriteria Then
                ' Highlight all columns in the evaluated row
                For k = 0 To UBound(criteriaRanges)
                    criteriaRanges(k).Rows(i).Interior.Color = vbGreen
                Next k
                If Not sumRange Is Nothing Then
                    sumRange.Rows(i).Interior.Color = vbGreen
                End If
            End If
        Next i
    Else
        MsgBox "The formula is not a SUMIF, SUMIFS, COUNTIF, or COUNTIFS formula. This code currently only supports SUMIF, SUMIFS, COUNTIF, and COUNTIFS."
    End If
    ' --- End of adaptation section ---
End Sub
