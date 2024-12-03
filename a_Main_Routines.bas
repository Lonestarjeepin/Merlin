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


Public varGoBack_ws As String
Public varGoBack_rng As String
Public varGoback_WB As String
Public varF1_Delay As String
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

'**********************************************************************************************
'GetKeyState function inserted for use in INDEXTRACE and VLOOKUPTRACE to detect presence of Shift Key Press before opening target WB if not already open
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Function ShiftPressed() As Boolean
'Returns True if shift key is pressed - used by Trace Precedent sub where link is to external workbook
    ShiftPressed = GetKeyState(SHIFT_KEY) < 0
End Function




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

Sub GoToPrecedents()
     ' written by Bill Manville
     ' With edits from PaulS and Kyle Whitmire
     ' this procedure finds the cells which are the direct precedents of the active cell
    Dim rLast As Range, iLinkNum As Integer, iArrowNum As Integer
    Dim stMsg As String
    Dim bNewArrow As Boolean
    Dim varGoTo As String
    Dim ws As Worksheet
    Dim cntHidden As Variant
    Dim cntPrecedents As Variant
    Dim i As Variant
    Dim element As Variant
    
    Application.ScreenUpdating = False
    On Error GoTo errcatch
    OtherSheetFlag = 0
    Current_WB = ActiveWorkbook.Name
    
lbl_PostOpenFile:
'KW ***********************************************
'Determine if there is a named range in the formula and if the named points to an external WB then lbl_OpenFile will open that WB
                namerng = ""
                For Each namerng In ActiveWorkbook.Names
                    If InStr(CStr(ActiveCell.Formula), namerng.Name) > 0 Then
                        strFormula = Replace(ActiveCell.Formula, namerng.Name, namerng.RefersTo)
                        strNamedRangeAddress = namerng.RefersTo

                    End If
                Next namerng
'***********************************************
    
    
'KW ***********************************************
    'Open destination workbook if not already open
lbl_OpenFile:


'Open destination file if external link is found in formula
    If InStr(ActiveCell.Formula, ":\") <> 0 Then
        tempIArray = Mid(ActiveCell.Formula, InStr(ActiveCell.Formula, "'") + 1, InStr(ActiveCell.Formula, "]") + 1)
        Do While ShiftPressed()
            DoEvents
        Loop
        
        homeWB = ActiveWorkbook.Name
        Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
        refWB = ActiveWorkbook.Name
        Workbooks(homeWB).Activate
        GoTo lbl_PostOpenFile
    End If
    
'Open destination file if external link is found in named range
    If InStr(strNamedRangeAddress, ":\") <> 0 Then
    tempIArray = Mid(strNamedRangeAddress, InStr(strNamedRangeAddress, "'") + 1, InStr(strNamedRangeAddress, "]") + 1)
    Do While ShiftPressed()
        DoEvents
    Loop

    homeWB = ActiveWorkbook.Name
    Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
    refWB = ActiveWorkbook.Name
    Workbooks(homeWB).Activate
    GoTo lbl_PostOpenFile
    End If
'***********************************************
    
    If InStr(Selection.Formula, "[") > 0 Then
        DestinationWB = Mid(Selection.Formula, InStr(Selection.Formula, "[") + 1, InStr(Selection.Formula, "]") - InStr(Selection.Formula, "[") - 1)
        OtherSheetFlag = 1
        Workbooks(DestinationWB).Activate
        Else
        'Create Flag if there is a sheet reference
        If InStr(Selection.Formula, "!") > 0 Then
            OtherSheetFlag = 1
        End If
    End If
        
    If OtherSheetFlag = 1 Then
        cntHidden = 0
        cntPrecedents = 0
        '''Count Hidden Worksheets
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
                cntHidden = cntHidden + 1
            End If
        Next
    '------------------------------------
        ReDim varwsarray(cntHidden)
    
        i = 1
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
                varwsarray(i) = ws.Name
                ws.Visible = True
                i = i + 1
            End If
        Next
'------------------------------------
    End If
    Workbooks(Current_WB).Activate
    
    ActiveCell.ShowPrecedents
    Set rLast = ActiveCell
    iArrowNum = 1
    iLinkNum = 1
    bNewArrow = True
    Do
        Do
            Application.GoTo rLast
            On Error Resume Next
            ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=iArrowNum, LinkNumber:=iLinkNum
            If Err.Number > 0 Then Exit Do
            On Error GoTo 0
            If rLast.Address(External:=True) = ActiveCell.Address(External:=True) Then Exit Do
            bNewArrow = False
            If rLast.Worksheet.Parent.Name = ActiveCell.Worksheet.Parent.Name Then
                    stMsg = stMsg & vbNewLine & "Arrow: " & iArrowNum & " - Link: " & iLinkNum & " - " & "'" & Selection.Parent.Name & "'!" & Selection.Address
                    cntPrecedents = cntPrecedents + 1
            Else
                stMsg = stMsg & vbNewLine & "Arrow: " & iArrowNum & " - Link: " & iLinkNum & " - " & Selection.Address(External:=True)
                cntPrecedents = cntPrecedents + 1
            End If

            
            
            iLinkNum = iLinkNum + 1 ' try another  link
        Loop
        If bNewArrow Then Exit Do
        iLinkNum = 1
        bNewArrow = True
        iArrowNum = iArrowNum + 1 'try another arrow
    Loop
    
    Application.GoTo rLast
    
    
    
    'KW - Sets persistent variable for GoBack feature - jump back to originating cell after following a precedent jump
    'Must put "Public varGoBack_ws and varGoBack_rng As String" at top of procedure to establish persistent variables across subroutines
    varGoBack_ws = ActiveSheet.Name
    varGoBack_rng = ActiveCell.Address
    varGoback_WB = ActiveWorkbook.Name
    
    'Auto-jump feature if only one precedent is found.  Added by Kyle Whitmire
    If stMsg <> "" Then
        'MsgBox "Precedents are" & stMsg & vbNewLine & vbNewLine & "test"
        If cntPrecedents = 1 Then
            varGoTo = "1-1"
            Else
            varGoTo = InputBox(stMsg & vbNewLine & vbNewLine & "To GoTo Precedent cell, input as x-x (Arrow-Link).")
        End If
        If varGoTo <> "" Then
            varArray = Split(varGoTo, "-")
            varArrow = varArray(0)
            varLink = varArray(1)
        
            ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=varArrow, LinkNumber:=varLink
        
            Application.ScreenUpdating = True
            ActiveCell.Activate
            Application.ScreenUpdating = False
        End If
        Else
        MsgBox ("No Precedents found in this Cell.")
    End If
    rLast.Parent.ClearArrows
    
    '''Rehide Array of Hidden Worksheets
    Destination = ActiveSheet.Name
    
   On Error Resume Next
   If OtherSheetFlag = 1 Then
        For i = 1 To cntHidden
            If Not varwsarray(i) = Destination Then
                Sheets(varwsarray(i)).Visible = False
            End If
            Sheets(Destination).Activate
        Next
    End If
    
    If Not varwsarray(cntHidden) = Destination Then
        Sheets(varwsarray(cntHidden)).Visible = False
    End If
    Application.ScreenUpdating = True
    Exit Sub
errcatch:
    Application.ScreenUpdating = True
    MsgBox "Oops, there was a problem"
End Sub

Sub GoBack()
    If Not IsEmpty(varGoback_WB) Then
        Windows(varGoback_WB).Activate
    End If
    If Not IsEmpty(varGoBack_ws) Then
        Worksheets(varGoBack_ws).Select
        Range(varGoBack_rng).Select
    End If

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
Dim Filename As Variant
Dim Sep As String
Dim AppendDataInput As LongPtr
Dim AppendDataBool As Boolean
Dim varSkipBlank As LongPtr
Dim boolSkipBlank As Boolean
Filename = Application.GetSaveAsFilename(InitialFileName:=vbNullString, filefilter:="Text Files (*.txt),*.txt")
If Filename = False Then
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

Debug.Print "FileName: " & Filename, "Separator: " & Sep
ExportToTextFile FName:=CStr(Filename), Sep:=CStr(Sep), AppendData:=AppendDataBool, SkipBlank:=boolSkipBlank
End Sub

Sub MRNETOPSUpload()
Application.ScreenUpdating = False

Dim wb As Workbook
Dim myFileName As String
Dim FilePath As String
FilePath = ActiveWorkbook.Path
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
ChDir FilePath
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

Function ISNamedRange(strRangeName As String) As Boolean
'Checks to see if a string is a valid named range in the workbook
    Dim rngExists As String
     
    On Error Resume Next
    rngExists = Names(strRangeName).Name
    If Len(rngExists) = 0 Then
    ISNamedRange = False
    Else
    ISNamedRange = True
    End If
    On Error GoTo 0
     
End Function


Sub IndexTrace()
    Dim FormulaStr As String
    Dim PO As Integer
    Dim PC As Integer
    Dim C1 As Integer
    Dim C2 As Integer
    Dim IArray As String
    Dim IArrayJump As String
    Dim IRow As Integer
    Dim IColumn As Integer
    Dim Destination As String
    Dim Match1 As String
    Dim Match2 As String
    Dim Searchfor As String
    Dim Searchformula As String
    Dim SearchformulaCnt As Integer
    Dim POcnt As Integer
    Dim ic As Integer
    Dim userchoice As Integer
    
'    On Error GoTo errcatch
        
    'Searchfor is the formula being searched. At this time it can only be index - Do not change
    Searchfor = "INDEX"
    FormulaStr = ActiveCell.Formula

    'This loop will count how many index formulas are in the selected formula
    Indexcnt = 0
    For x = 1 To Len(ActiveCell.Formula)
        If Mid(ActiveCell.Formula, x, 5) = "INDEX" Then
            Indexcnt = Indexcnt + 1
        End If
    Next
    VLookup = 0
    For x = 1 To Len(ActiveCell.Formula)
        If Mid(ActiveCell.Formula, x, 7) = "VLOOKUP" Then
            VLookup = VLookup + 1
        End If
    Next
    'This case handles which index will be used
    Select Case Indexcnt
        Case Is < 1
            If VLookup > 0 Then
                VlookupTrace
                Exit Sub
            Else
                GoToPrecedents
                Exit Sub
            End If
        Case 1
            userchoice = 1
        Case Is > 1
            userchoice = InputBox(Prompt:="There are " & Indexcnt & " index formulas. Which one would you like? Please choose a whole number between 1 and " & Indexcnt & ".", Title:="Index Trace")
            userchoice = Int(userchoice)
'            Debug.Print userchoice > 0 And userchoice < (Indexcnt + 1)
            If Not userchoice > 0 And userchoice < (Indexcnt + 1) Then
                MsgBox "You must chose a whole number between 1 and " & Indexcnt & ". Please run the macro again."
                Exit Sub
            End If
        
    End Select
    
    'This finds what character our index starts on
    ic = 0
    x = 0
    Do
        x = x + 1
        If x > Len(ActiveCell.Formula) Then
            Exit Do
        Else
            If Mid(ActiveCell.Formula, x, 5) = "INDEX" Then
                ic = ic + 1
                If ic = userchoice Then
                startingchar = x
                    Exit Do
                End If
            End If
        End If
        
    Loop Until ic = userchoice
    
're-builds the formula arrays once external WB has been opened
lbl_PostOpenFile:
    
    
    'Separates the index formula being used and breaks down the formula to determine row and column
    FormulaStr = ActiveCell.Formula
    Formula = Searchfor & Mid(FormulaStr, InStr(InStr(startingchar, FormulaStr, Searchfor), FormulaStr, "("), Len(FormulaStr) - InStr(FormulaStr, Searchfor) + 1)
    FormulaCnt = Len(Formula)
    
    POcnt = 0
    ccnt = 0
    For x = 1 To FormulaCnt
        Select Case Mid(Formula, x, 1)
            Case Is = "("
                POcnt = POcnt + 1
                If POcnt = 1 Then
                    PO = x
                End If
            Case Is = ")"
                If POcnt > 0 Then
                    POcnt = POcnt - 1
                    If POcnt = 0 Then
                        PC = x
                        Formula = Mid(Formula, 1, x)
                        Exit For
                    End If
                End If
            Case Is = ","
                If POcnt = 1 Then
                    ccnt = ccnt + 1
                    Select Case ccnt
                        Case Is = 1
                            C1 = x
                        Case Is = 2
                            C2 = x
                    End Select
                End If
                        
        End Select
    Next
    
    IArray = Mid(Formula, PO + 1, C1 - PO - 1)
    IArrayJump = Replace(Mid(Formula, PO + 1, C1 - PO - 1), "'", "")
    IRowForm = Mid(Formula, C1 + 1, C2 - C1 - 1)
    IColform = Mid(Formula, C2 + 1, PC - C2 - 1)
    
    Dim strIArray As Variant
    
    '***********************************************
    'Open destination workbook if not already open
lbl_OpenFile:
    
    'Evaluate formula to determine if external file
    If InStr(ActiveCell.Formula, ":\") <> 0 Then
    tempIArray = Mid(ActiveCell.Formula, InStr(ActiveCell.Formula, "'") + 1, InStr(ActiveCell.Formula, "]") + 1)
    Do While ShiftPressed()
        DoEvents
    Loop
    
    homeWB = ActiveWorkbook.Name
    Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
    refWB = ActiveWorkbook.Name
    Workbooks(homeWB).Activate
    GoTo lbl_PostOpenFile
    End If
    
    'Evaluate Named Range to determine if external file
    
    If InStr(IArray, ":\") <> 0 Then
    tempIArray = Mid(IArray, InStr(IArray, "'") + 1, InStr(IArray, "]") + 1)
    Do While ShiftPressed()
        DoEvents
    Loop
    
    homeWB = ActiveWorkbook.Name
    Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
    refWB = ActiveWorkbook.Name
    Workbooks(homeWB).Activate
    GoTo lbl_PostOpenFile
    End If
    '***********************************************

'Test if any of the INDEX/MATCH variables are named ranges and convert to cell address if they are
    
    If ISNamedRange(IArray) = True Then
        IArray = Names(IArray).RefersTo
        If InStr(IArray, ":\") <> 0 Then
        GoTo lbl_OpenFile
        End If
    End If
    If ISNamedRange(IArrayJump) = True Then
        IArrayJump = Names(IArrayJump).RefersTo
        If InStr(IArrayJump, ":\") <> 0 Then
        GoTo lbl_OpenFile
        End If
    End If
    

    
'Not sure I need any of the kwIArray conversion since I am caring for named ranges above
    Dim kwIArray As Range
    Set kwIArray = Range(IArray)

    If InStr(IArray, "!") = 0 Then
        RefSheet = ActiveSheet.Name
    Else
        If InStr(IArray, "]") = 0 Then
            refWB = ActiveWorkbook.Name
            IArray = "'" & kwIArray.Parent.Name & "'!" & kwIArray.Address(External:=False)
            RefSheet = Replace(Left(IArray, InStr(IArray, "!") - 1), "'", "")
        Else
            refWB = Replace(Replace(Replace(Mid(IArrayJump, 2, InStr(IArrayJump, "]") - 2), "[", ""), "'", ""), "=", "")
            IArray = "'[" & refWB & "]" & kwIArray.Parent.Name & "'!" & kwIArray.Address(External:=False)
            RefSheet = Replace(Mid(IArray, InStr(IArray, "]") + 1, InStr(IArray, "!") - InStr(IArray, "]") - 1), "'", "")
        End If
    End If
    
    
    'Debug.Print refWB
    
    If Not InStr(IRowForm, "ROW()") = 0 Then
        ro = ActiveCell.Row
        IRowForm = Replace(IRowForm, "ROW()", ro)
    End If
    
    If Not IsNumeric(IRowForm) Then
        IRow = Evaluate(IRowForm)
    Else
        IRow = IRowForm
    End If
    IRow = IRow + Range(IArray).Row - 1

    If Not InStr(IColform, "COLUMN()") = 0 Then
        co = ActiveCell.Column
        IColform = Replace(IColform, "COLUMN()", co)
    End If
    
    If Not IsNumeric(IColform) Then
        IColumn = Application.Evaluate(IColform)
    Else
        IColumn = IColform
    End If
    IColumn = IColumn + Range(IArray).Column - 1
        
    Destination = Cells(IRow, IColumn).Address
    
    varGoBack_ws = ActiveSheet.Name
    varGoBack_rng = ActiveCell.Address
    varGoback_WB = ActiveWorkbook.Name

    If Not InStr(IArray, "]") = 0 Then
        Windows(refWB).Activate
    End If
    
    If Sheets(RefSheet).Visible = False Then
        Sheets(RefSheet).Visible = True
    End If
    
    Sheets(RefSheet).Select
    Range(Destination).Select
    Exit Sub
    
errcatch:
    Application.ScreenUpdating = True
    MsgBox "Oops, there was a problem"
End Sub
Sub VlookupTrace()
    Dim FormulaStr As String
    Dim PO As Integer
    Dim PC As Integer
    Dim C1 As Integer
    Dim C2 As Integer
    Dim IArray As String
    Dim IArrayJump As String
    Dim IRow As Integer
    Dim IColumn As Integer
    Dim Destination As String
    Dim Match1 As String
    Dim Match2 As String
    Dim Searchfor As String
    Dim Searchformula As String
    Dim SearchformulaCnt As Integer
    Dim POcnt As Integer
    Dim ic As Integer
    Dim userchoice As Integer
'    On Error GoTo errcatch
        
    'Searchfor is the formula being searched. At this time it can only be VLookup - Do not change
    Searchfor = "VLOOKUP"
    FormulaStr = ActiveCell.Formula

    'This loop will count how many VLookup formulas are in the selected formula
    Indexcnt = 0
    For x = 1 To Len(ActiveCell.Formula)
        If Mid(ActiveCell.Formula, x, 7) = "VLOOKUP" Then
            VLookup = VLookup + 1
        End If
    Next

    'This case handles which VLookup will be used
    Select Case VLookup
        Case Is < 1
            GoToPrecedents
        Case 1
            userchoice = 1
        Case Is > 1
            userchoice = InputBox(Prompt:="There are " & VLookup & " Vlookup formulas. Which one would you like? Please choose a whole number between 1 and " & VLookup & ".", Title:="Vlookup Trace")
            userchoice = Int(userchoice)
'            Debug.Print userchoice > 0 And userchoice < (Indexcnt + 1)
            If Not userchoice > 0 And userchoice < (VLookup + 1) Then
                MsgBox "You must chose a whole number between 1 and " & VLookup & ". Please run the macro again."
                Exit Sub
            End If
    End Select
    
    'This finds what character our VLookup starts on
    ic = 0
    x = 0
    Do
        x = x + 1
        If x > Len(ActiveCell.Formula) Then
            Exit Do
        Else
            If Mid(ActiveCell.Formula, x, 7) = "VLOOKUP" Then
                ic = ic + 1
                If ic = userchoice Then
                startingchar = x
                    Exit Do
                End If
            End If
        End If
        
    Loop Until ic = userchoice
    
're-builds the formula arrays once external WB has been opened
lbl_PostOpenFile:
    
    'Separates the index formula being used and breaks down the formula to determine row and column
    
    FormulaStr = ActiveCell.Formula
    Formula = Searchfor & Mid(FormulaStr, InStr(InStr(startingchar, FormulaStr, Searchfor), FormulaStr, "("), Len(FormulaStr) - InStr(FormulaStr, Searchfor) + 1)
    FormulaCnt = Len(Formula)
       
    POcnt = 0
    ccnt = 0
    For x = 1 To FormulaCnt
        Select Case Mid(Formula, x, 1)
            Case Is = "("
                POcnt = POcnt + 1
                If POcnt = 1 Then
                    PO = x
                End If
            Case Is = ")"
                If POcnt > 0 Then
                    POcnt = POcnt - 1
                    If POcnt = 0 Then
                        PC = x
                        Formula = Mid(Formula, 1, x)
                        Exit For
                    End If
                End If
            Case Is = ","
                If POcnt = 1 Then
                    ccnt = ccnt + 1
                    Select Case ccnt
                        Case Is = 1
                            C1 = x
                        Case Is = 2
                            C2 = x
                        Case Is = 3
                            C3 = x
                    End Select
                End If
                        
        End Select
    Next
    
    IRowForm = Mid(Formula, PO + 1, C1 - PO - 1)
    IArrayJump = Replace(Mid(Formula, PO + 1, C1 - PO - 1), "'", "")
    IArray = ""
    refWB = ""
    RefSheet = ""
    IArray = Mid(Formula, C1 + 1, C2 - C1 - 1)
    IColform = Mid(Formula, C2 + 1, C3 - C2 - 1)
    
    '***********************************************
    'Open destination workbook if not already open
lbl_OpenFile:
    
    'Evaluate formula to determine if external file
    If InStr(ActiveCell.Formula, ":\") <> 0 Then
    tempIArray = Mid(ActiveCell.Formula, InStr(ActiveCell.Formula, "'") + 1, InStr(ActiveCell.Formula, "]") + 1)
    Do While ShiftPressed()
        DoEvents
    Loop
    
    homeWB = ActiveWorkbook.Name
    Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
    refWB = ActiveWorkbook.Name
    Workbooks(homeWB).Activate
    GoTo lbl_PostOpenFile
    End If
    
    'Evaluate Named Range to determine if external file
    
    If InStr(IArray, ":\") <> 0 Then
    tempIArray = Mid(IArray, InStr(IArray, "'") + 1, InStr(IArray, "]") + 1)
    Do While ShiftPressed()
        DoEvents
    Loop
    
    homeWB = ActiveWorkbook.Name
    Workbooks.Open Replace(Replace(Replace((Left(tempIArray, InStr(tempIArray, "]"))), "[", ""), "]", ""), "'", "")
    refWB = ActiveWorkbook.Name
    Workbooks(homeWB).Activate
    GoTo lbl_PostOpenFile
    End If
    '***********************************************

    
'Test if any of the VLOOKUP variables are named ranges and convert to cell address if they are
    
    If ISNamedRange(IArray) = True Then
        IArray = Names(IArray).RefersTo
        If InStr(IArray, ":\") <> 0 Then
        GoTo lbl_OpenFile
        End If
    End If

    If ISNamedRange(IArrayJump) = True Then
        Debug.Print ISNamedRange(IArrayJump)
        IArrayJump = Names(IArrayJump).RefersTo
        If InStr(IArrayJump, ":\") <> 0 Then
        GoTo lbl_OpenFile
        End If
    End If
    
    
    
    'Dim kwIArray As Range
    Set kwIArray = Range(IArray)
    
    IArrayCol = Range(IArray).Columns(1).Address
    
'    If InStr(IArray, "!") = 0 Then     'Same WB, Same WS
'        RefSheet = ActiveSheet.Name
'        IArray = "'" & kwIArray.Parent.Name & "'!" & kwIArray.Address(External:=False)
'        IRow = WorksheetFunction.Match(Evaluate(iRowForm), Worksheets(RefSheet).Range(IArrayCol), 0)
'    Else
        If InStr(IArray, "]") = 0 Then 'Same WB, Different WS
            refWB = ActiveWorkbook.Name
            IArray = "'" & kwIArray.Parent.Name & "'!" & kwIArray.Address(External:=False)
            RefSheet = Replace(Left(IArray, InStr(IArray, "!") - 1), "'", "")
            IRow = WorksheetFunction.Match(Evaluate(IRowForm), Worksheets(RefSheet).Range(IArrayCol), 0)
        Else                           'Different WB, Different WS
            refWB = Replace(Replace(Mid(IArray, 3, InStr(IArray, "]") - 3), "[", ""), "'", "")
            tempIArray = Replace(IArray, "'", "")
            IArray = "'[" & refWB & "]" & kwIArray.Parent.Name & "'!" & kwIArray.Address(External:=False)
            RefSheet = Replace(Mid(IArray, (InStr(IArray, "]") + 1), (InStr(tempIArray, "!") - (InStr(IArray, "]")))), "'", "")    'Replace(Left(IArray, InStr(IArray, "!") - 1), "'", "")
            IRow = WorksheetFunction.Match(Evaluate(IRowForm), Workbooks(refWB).Worksheets(RefSheet).Range(IArrayCol), 0)
        End If
'    End If

    If Not IsNumeric(IColform) Then
        IColumn = Application.Evaluate(IColform)
    Else
        IColumn = IColform
    End If
    IColumn = IColumn + Range(IArray).Column - 1
    IRow = IRow + Range(IArray).Row - 1
    Destination = Cells(IRow, IColumn).Address
    
    varGoBack_ws = ActiveSheet.Name
    varGoBack_rng = ActiveCell.Address
    varGoback_WB = ActiveWorkbook.Name

    If Not InStr(IArray, "]") = 0 Then
        Windows(refWB).Activate
    End If
    
    If Sheets(RefSheet).Visible = False Then
        Sheets(RefSheet).Visible = True
    End If
    
    Sheets(RefSheet).Select
    Range(Destination).Select
    Application.ScreenUpdating = True
    Exit Sub


errcatch:
    Application.ScreenUpdating = True
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

'Subroutine to find and display (in new sheet) all formulas with ERRORs in active sheet
Sub form_errors()
Dim s1 As String
s1 = ActiveSheet.Name
Dim rng, cell As Range
Set rng = Sheets(s1).UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
Application.DisplayAlerts = False
Select Case rng Is Nothing
    Case False
        On Error Resume Next
        Sheets(s1 & "_Invalid").Delete
        On Error GoTo 0
        Sheets.Add After:=Sheets(s1)
        ActiveSheet.Name = s1 & "_Invalid"
        Sheets(s1 & "_Invalid").Select
        
        Cells(1, 1) = "Cell Address"
        Cells(1, 2) = "Formula"
        Cells(1, 3) = "Formula Result"
        
        varRow = 1
        For Each cell In rng
        varRow = varRow + 1
        Debug.Print Rows.Count
            Sheets(s1 & "_Invalid").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = cell.Address
            ActiveSheet.Hyperlinks.Add Anchor:=Cells(varRow, 1), Address:="", SubAddress:= _
                            "'" & s1 & "'!" & cell.Address, TextToDisplay:=cell.Address
            
            
            Sheets(s1 & "_Invalid").Cells(Rows.Count, 2).End(xlUp).Offset(1, 0) = Replace(cell.Formula, "=", "'")
            Sheets(s1 & "_Invalid").Cells(Rows.Count, 3).End(xlUp).Offset(1, 0) = cell.Value
        Next cell
        
End Select

ActiveSheet.Range("A:C").Columns.EntireColumn.AutoFit
Application.DisplayAlerts = True
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
            If Mid(txt, y + 1, 4) = "   " Then
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


Sub ListLinks()
Application.ScreenUpdating = False
Application.EnableCancelKey = xlDisabled
    
    If IsEmpty(ActiveWorkbook.LinkSources) Then
        msg = "This workbook does not contain any links!"
        MsgBox msg
        Exit Sub
    End If
    
    Sheets.Add.Name = "LinkList_" & Replace(Time, ":", ".")
    Range("A1").Select
    ActiveCell.Value = "Link FilePath"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Link FileName"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Reference Cell"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Ref Cell Formula"
    ActiveCell.Offset(0, 1).Select
    Range("A2").Select

    varRow = 2
    For Each lSource In ActiveWorkbook.LinkSources
        Cells(varRow, 1).Value = lSource
        'Repeated to be used to display only filename later in code
        Cells(varRow, 2).Value = lSource
        Range("B:B").Replace "*\", "", xlPart
        varFileName = Cells(varRow, 2).Value
        
        For Each sh In ActiveWorkbook.Sheets
        
            Set rng1 = Nothing
            Set rng2 = Nothing

            On Error Resume Next
            Set rng1 = sh.Cells.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            
            If Not rng1 Is Nothing Then
                 'look for *.xls
                With rng1
                    Set rng2 = .Find("*" & Replace(varFileName, "'", "''") & "*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
                    If Not rng2 Is Nothing Then
                        FirstAddress = sh.Name & "!" & rng2.Address
                         'repeat till code loops back to first formula cell containing lSource
                        Do
                           Cells(varRow, 3).Value = sh.Name & "!" & rng2.Address
                            ActiveSheet.Hyperlinks.Add Anchor:=Cells(varRow, 3), Address:="", SubAddress:= _
                            "'" & sh.Name & "'!" & rng2.Address, TextToDisplay:=sh.Name & "!" & rng2.Address
    
                           Cells(varRow, 4).Value = "'" & rng2.Formula
                           varRow = varRow + 1
                           Set rng2 = .FindNext(rng2)
                        Loop Until sh.Name & "!" & rng2.Address = FirstAddress
                    End If
                End With
            End If
        
        Next
        varRow = varRow + 1
    Next
Columns("A:D").AutoFit

Application.ScreenUpdating = True
End Sub



Sub WorksheetSizes()
    Dim wks As Worksheet
    Dim c As Range
    Dim sFullFile As String
    Dim sReport As String
    Dim sWBName As String

    sReport = "Size Report"
    sWBName = "Erase Me.xls"
    sFullFile = ActiveWorkbook.Path & _
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


Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As LongPtr
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As LongPtr
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
    Dim lCalcSave As LongPtr
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

    dTime = Round(dTime, 5) / 5
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

Public Type scaleAxisScale
  ' Calculated Axis Scale Parameters
  dMin As Double
  dMax As Double
  dMajor As Double
  dMinor As Double
End Type

Function fnAxisScale(ByVal dMin As Double, ByVal dMax As Double) As scaleAxisScale
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
          vPValues = srs.Values

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
          vSValues = srs.Values

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
