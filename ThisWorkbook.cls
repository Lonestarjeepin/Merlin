Private Sub Workbook_Open()

DeleteMenu

AddMenus

Application.OnKey "^+b", "Blue"
Application.OnKey "^+g", "Green"
Application.OnKey "^+r", "Red"
Application.OnKey "^+y", "Yellow"
Application.OnKey "^+n", "Post_it_note"
Application.OnKey "^+s", "paste_special"
Application.OnKey "^+f", "paste_formulas"
Application.OnKey "^+1", "number_nored_format"
Application.OnKey "^+4", "dollar_nored_format"
Application.OnKey "^+M", "dollar_million_nored_format"
Application.OnKey "^+K", "dollar_thousand_nored_format"
Application.OnKey "^+5", "percent_nored_format"
Application.OnKey "^+I", "increase_decimal"
Application.OnKey "^+D", "decrease_decimal"
Application.OnKey "^+H", "HighlightSelection"
Application.OnKey "^+T", "transparent_chart"
Application.OnKey "^+C", "clear_formatting"
Application.OnKey "^+P", "paste_formatting"
Application.OnKey "^+x", "TracePrecedents"
Application.OnKey "^+z", "GoBack"
Application.OnKey "^+e", "Iferror"
Application.OnKey "^+q", "CenterAcrossColumns"
Application.OnKey "^+v", "Variance_Percent"
Application.OnKey "^+W", "View_All_WorkSheets"
Application.OnKey "^+U", "Unhide_Rehide_WS"

        
End Sub


Private Sub Workbook_AddinInstall()

Run "AddMenus"

End Sub



Private Sub Workbook_AddinUninstall()

Run "DeleteMenu"

On Error Resume Next

End Sub
