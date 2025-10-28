Attribute VB_Name = "Module9"
Sub Oracle_Payments()
Attribute Oracle_Payments.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Oracle_Payments Macro
'
' Keyboard Shortcut: Ctrl+i
'
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$I$28").AutoFilter Field:=2, Criteria1:= _
        "=Credit Memo", Operator:=xlOr, Criteria2:="=Invoice"
    Range("B2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.Range("$A$1:$I$9").AutoFilter Field:=2
    ActiveWorkbook.Worksheets("Exported").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Exported").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("C1:C9"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Exported").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*-1"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E16"), Type:=xlFillDefault
    Range("E2:E16").Select
    Columns("D:F").Select
    Selection.Style = "Comma"
    Range("G15").Select
    Columns("G:G").ColumnWidth = 7.89
    Columns("F:F").ColumnWidth = 6.89
    Columns("H:H").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Reversed", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 8420607
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("F9").Select
    Columns("F:F").EntireColumn.AutoFit
    Range("A1").Select
End Sub
