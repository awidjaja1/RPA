Attribute VB_Name = "Module4"
Sub Team_Register()
Attribute Team_Register.VB_Description = "Team Register updated for CSV export 6/5/24"
Attribute Team_Register.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Team_Register Macro
' Team Register updated for CSV export 6/5/24
'

'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.DELETE Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Invoice #"
    Columns("D:D").Select
    Selection.DELETE Shift:=xlToLeft
    Selection.DELETE Shift:=xlToLeft
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "EE ID"
    Columns("F:F").Select
    Selection.DELETE Shift:=xlToLeft
    Selection.Style = "Comma"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Columns("G:G").Select
    Selection.DELETE Shift:=xlToLeft
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Invoice Date"
    Columns("D:D").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    Selection.DELETE Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Selection.ColumnWidth = 16.67
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("A1:I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A3").Select
    Columns("A:A").ColumnWidth = 46.22
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("H:H").ColumnWidth = 18.78
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("G1:G183"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=0
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "Pmts "
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "Applied "
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "Balance "
    Range("N2:N4").Select
    Selection.Style = "Comma"
    Range("M2:N4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("M4:N4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("N4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-2]C-R[-1]C"
    Range("N5").Select
End Sub
