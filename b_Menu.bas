
'*****************************************************************
'Code for Add-In Menu Dropdown
'*****************************************************************
Sub AddMenus()

Dim cMenu1 As CommandBarControl

Dim cbMainMenuBar As CommandBar
Dim cbChartMenuBar As CommandBar

Dim iViewMenu As Integer

Dim cbcCustomMenu As CommandBarControl
    'Dim subFormattingMenu As CommandBarControl
Dim cbcCustomChartMenu As CommandBarControl



    '(1)Delete any existing one.We must use On Error Resume next in case it does not exist.

    On Error Resume Next
    
    Application.CommandBars("Worksheet Menu Bar").Controls("&Merlin").Delete

    '(2)Set a CommandBar variable to Worksheet menu bar

    Set cbMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
 
    '(3)Return the Index number of the View menu. We can then use this to place a custom menu before.

    iViewMenu = cbMainMenuBar.Controls("View").Index

    '(4)Add a Control to the "Worksheet Menu Bar" before View

    'Set a CommandBarControl variable to it

    Set cbcCustomMenu = cbMainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=iViewMenu)

    '(5)Give the control a caption

    cbcCustomMenu.Caption = "&Merlin"


    '(6)Working with our new Control, add a sub control and give it a Caption and tell it which macro to run (OnAction).
            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Merlin ChangeLog and Help"
                .OnAction = "ChangeLog"
            End With

            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Update Merlin"
                .OnAction = "ManualUpdate"
            End With

            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Merlin Support"
                .OnAction = "Merlin_Support"
            End With

            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "-------------------------"
                .OnAction = ""
            End With
            
    'Add a sub-menu that will contain other menu items

        Set subFormattingMenu = cbcCustomMenu.Controls.Add(Type:=msoControlPopup) ', Before:=iViewMenu)
        subFormattingMenu.Caption = "Formatting"
    
    'Menu items contained in sub-menu
        With subFormattingMenu
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Yellow (Ctrl + Shft Y)"
                .OnAction = "yellow"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Green (Ctrl + Shft G)"
                .OnAction = "green"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Blue (Ctrl + Shft B)"
                .OnAction = "blue"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Red (Ctrl + Shft R)"
                .OnAction = "red"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Post-It Note Yellow (Ctrl + Shft N)"
                .OnAction = "post_it_note"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Clear Formatting (Ctrl + Shft C)"
                .OnAction = "clear_formatting"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Paste Special - Formatting (Ctrl + Shft P)"
                .OnAction = "paste_formatting"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Paste Special - Values (Ctrl + Shft S)"
                .OnAction = "paste_special"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Paste Special - Formulas (Ctrl + Shft F)"
                .OnAction = "paste_formulas"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Page Setup - Narrow w/ Date/Time Footer"
                .OnAction = "Page_Setup"
            End With
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Color Columns"
                .OnAction = "ColorColumn"
            End With
        End With  'end Formatting sub-menu
        
        
        
        'Add a sub-menu that will contain other menu items

            Set subNumberFormattingMenu = cbcCustomMenu.Controls.Add(Type:=msoControlPopup) ', Before:=iViewMenu)
            subNumberFormattingMenu.Caption = "Number Formatting"
    
        'Menu items contained in sub-menu
            With subNumberFormattingMenu
        
        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Number Format"
                .OnAction = "number_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Number Format No Red (Ctrl + Shft 1)"
                .OnAction = "number_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Million Format"
                .OnAction = "million_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Million Format No Red"
                .OnAction = "million_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Thousand Format"
                .OnAction = "thousand_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Thousand Format No Red"
                .OnAction = "thousand_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Sharon's Dollar Format"
                .OnAction = "dollar_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Dollar Format No Red (Ctrl + Shft 4)"
                .OnAction = "dollar_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Dollar Million Format"
                .OnAction = "dollar_million_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Dollar Million Format No Red (Ctrl + Shft M)"
                .OnAction = "dollar_million_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Dollar Thousand Format"
                .OnAction = "dollar_thousand_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Dollar Thousand Format No Red (Ctrl + Shft K)"
                .OnAction = "dollar_thousand_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Percent Format"
                .OnAction = "percent_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Percent Format No Red (Ctrl + Shft 5)"
                .OnAction = "Percent_nored_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Basis Point Format"
                .OnAction = "bps_format"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Ordinal Number Format"
                .OnAction = "Ordinal_Format"
            End With
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "If Error then 0 (Ctrl + Shft E)"
                .OnAction = "Iferror"
            End With
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Round"
                .OnAction = "Round"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Increase Decimal (Ctrl + Shft I)"
                .OnAction = "increase_decimal"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Decrease Decimal (Ctrl + Shft D)"
                .OnAction = "decrease_decimal"
            End With
            
            End With  'end Number Formatting sub-menu
            
        
    'Add a sub-menu that will contain other menu items

        Set subEfficiencyMenu = cbcCustomMenu.Controls.Add(Type:=msoControlPopup)
        subEfficiencyMenu.Caption = "Workbook Efficiency"

    'Menu items contained in sub-menu
        With subEfficiencyMenu
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Two Range List Builder"
                .OnAction = "Two_Range_List_Builder"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Three Range List Builder"
                .OnAction = "Three_Range_List_Builder"
            End With
                                  
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Evaluate as Formula/Number"
                .OnAction = "EvaluateAsFormula"
            End With
               
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Convert Text to Formula"
                .OnAction = "ConvertToFormula"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Copy/Paste Exact Formulas"
                .OnAction = "CopyExactFormulas"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Crack Internal Passwords"
                .OnAction = "AllInternalPasswords"
            End With

            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Highlight Contributing Cells (SUMIFS/COUNTIFS)"
                .OnAction = "HighlightContributingCells"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Find Errors in Formulas"
                .OnAction = "Find_Formula_Errors"
            End With
            '''''''''''''''''''''''''''''''''''''''''''''
            'NEED TO UPDATE FOR WS VS. WB
            '''''''''''''''''''''''''''''''''''''''''''''
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Manage Hidden Objects"
                .OnAction = "PeekaBoo"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "List External Links"
                .OnAction = "ListLinks"
            End With

            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Create Workbook Table of Contents"
                .OnAction = "TableOfContents"
            End With
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Count Worksheets"
                .OnAction = "Count_Worksheets"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Worksheet Selector (Ctrl + Shft W)"
                .OnAction = "View_All_Worksheets"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Unhide / Rehide Worksheets (Ctrl + Shft U)"
                .OnAction = "Unhide_Rehide_WS"
            End With
                        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Size of Worksheets"
                .OnAction = "WorksheetSizes"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Auto-Group Hidden Rows/Cols"
                .OnAction = "auto_group"
            End With
                                    
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Disable AutoRecover"
                .OnAction = "DisableAutoRecover"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Export to Delimited File"
                .OnAction = "ExportToDelimited"
            End With

            With .Controls.Add(Type:=msoControlButton)
                .Caption = "-------------------------"
                .OnAction = ""
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Calc Timer - Range"
                .OnAction = "RangeTimer"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Calc Timer - Sheet"
                .OnAction = "SheetTimer"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Calc Timer - Recalc"
                .OnAction = "RecalcTimer"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Calc Timer - Full Calc"
                .OnAction = "FullcalcTimer"
            End With
            
        End With 'end sub-menu
        
        

        
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Percent Variance (Ctrl + Shft V)"
                .OnAction = "Variance_Percent"
            End With
            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Highight Selection (Ctrl + Shft H)"
                .OnAction = "HighlightSelection"
            End With
            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Trace Precedents (Ctrl + Shft X)"
                .OnAction = "TracePrecedents"
            End With
            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Go Back from Precedent (Ctrl + Shft Z)"
                .OnAction = "GoBack"
            End With
            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "-------------------------"
                .OnAction = ""
            End With
                        
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Scale All Charts on Sheet"
                .OnAction = "ScaleActiveSheetCharts"
            End With

            
            With cbcCustomMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "-------------------------"
                .OnAction = ""
            End With
            
            
'    'Add a sub-menu that will contain other menu items
'
'        Set subDevMenu = cbcCustomMenu.Controls.Add(Type:=msoControlPopup)
'        subDevMenu.Caption = "Under Development (use at own risk)"

    'Menu items contained in sub-menu
'        With subDevMenu

'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = "Find Numbers in SUM"
'                .OnAction = "FindNumSub"
'            End With
'
'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = "List_Unique"
'                .OnAction = ""
'            End With
            
'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = "Toggle Smart View (Ctrl + Shft V)"
'                .OnAction = "toggle_SmartView"
'            End With
'        End With 'end sub-menu

On Error GoTo 0





    '(1)Delete any existing one.We must use On Error Resume next in case it does not exist.

    On Error Resume Next


    Application.CommandBars("Chart Menu Bar").Controls("&Merlin").Delete

    '(2)Set a CommandBar variable to Worksheet menu bar
    Set cbChartMenuBar = Application.CommandBars("Chart Menu Bar")
   

    '(3)Return the Index number of the View menu. We can then use this to place a custom menu before.
    iViewMenu = cbChartMenuBar.Controls("View").Index

    '(4)Add a Control to the "Worksheet Menu Bar" before View

    'Set a CommandBarControl variable to it
    Set cbcCustomChartMenu = cbChartMenuBar.Controls.Add(Type:=msoControlPopup, Before:=iViewMenu)

    '(5)Give the control a caption

    cbcCustomChartMenu.Caption = "&Merlin"

'''''''''''''''''
            With cbcCustomChartMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Rename Chart"
                .OnAction = "RenameChart"
            End With
            
            With cbcCustomChartMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Transparent Chart (Ctrl + Shft T)"
                .OnAction = "transparent_chart"
            End With
            
            With cbcCustomChartMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Center Chart Title"
                .OnAction = "center_chart_title"
            End With
                        
            With cbcCustomChartMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Scale Active Chart"
                .OnAction = "ScaleActiveChart"
            End With
            
            With cbcCustomChartMenu.Controls.Add(Type:=msoControlButton)
                .Caption = "Scale All Charts on Sheet"
                .OnAction = "ScaleActiveSheetCharts"
            End With
''''''''''''''''''
On Error GoTo 0
End Sub



Sub DeleteMenu()

    

On Error Resume Next

Application.CommandBars("Worksheet Menu Bar").Controls("&Merlin").Delete
Application.CommandBars("Chart Menu Bar").Controls("&Merlin").Delete
Application.CommandBars("Worksheet Menu Bar").Controls("&Merlin").Delete
Application.CommandBars("Chart Menu Bar").Controls("&Merlin").Delete
On Error GoTo 0

End Sub
