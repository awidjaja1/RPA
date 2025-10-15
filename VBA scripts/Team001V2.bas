Attribute VB_Name = "Module7"
Sub Team_001V2()
Attribute Team_001.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Team_001 Macro
'
'   Set Variable to count rows
    Dim r As Long
    r = (Range("B1") * 1) + 1
'   Rearrange columns to usable order

    Application.ScreenUpdating = False
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("F:F,H:H").Select
    Range("H1").Activate
    'ActiveWindow.ScrollColumn = 2
    'ActiveWindow.ScrollColumn = 3
    'ActiveWindow.ScrollColumn = 4
    Range("F:F,H:H,L:L").Select
    Range("L1").Activate
'    ActiveWindow.ScrollColumn = 5
'    ActiveWindow.ScrollColumn = 6
'    ActiveWindow.ScrollColumn = 7
'    ActiveWindow.ScrollColumn = 8
'    ActiveWindow.ScrollColumn = 9
'    ActiveWindow.ScrollColumn = 10
'    ActiveWindow.ScrollColumn = 11
'    ActiveWindow.ScrollColumn = 12
    Range("F:F,H:H,L:L,S:S").Select
    Range("S1").Activate
    Selection.Delete Shift:=xlToLeft
'    ActiveWindow.ScrollColumn = 13
'    ActiveWindow.ScrollColumn = 12
'    ActiveWindow.ScrollColumn = 11
'    ActiveWindow.ScrollColumn = 10
'    ActiveWindow.ScrollColumn = 9
'    ActiveWindow.ScrollColumn = 8
'    ActiveWindow.ScrollColumn = 7
'    ActiveWindow.ScrollColumn = 6
'    ActiveWindow.ScrollColumn = 5
'    ActiveWindow.ScrollColumn = 4
'    ActiveWindow.ScrollColumn = 3
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 1
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    'Add EE status column
    'r = Range("A" & Rows.Count).End(xlUp).Row 'Test
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Empl Status"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "A"
    Selection.AutoFill Destination:=Range("D2:D" & r)
    Range("D2:D" & r).Select
    
    'Adjust Column Width
    Columns("A:A").ColumnWidth = 9.29
    Columns("B:B").ColumnWidth = 14.57
    Columns("C:C").ColumnWidth = 7.71
    Columns("D:D").ColumnWidth = 3.57
    Columns("E:E").ColumnWidth = 3.86
    Columns("F:F").ColumnWidth = 6.29
    Columns("G:G").ColumnWidth = 11.14
    Columns("H:H").ColumnWidth = 15.43
    Columns("I:I").ColumnWidth = 8.57
    Columns("J:J").ColumnWidth = 8.71
    Columns("K:K").ColumnWidth = 15.57
    Columns("L:L").ColumnWidth = 9.57
    Columns("M:M").ColumnWidth = 7.57
    Columns("N:N").ColumnWidth = 6.71
    Columns("O:O").ColumnWidth = 6.57
    Columns("P:P").ColumnWidth = 7
    Columns("Q:Q").ColumnWidth = 9.14
    Columns("R:R").ColumnWidth = 9
    Columns("S:S").ColumnWidth = 3
    
    'Conditional Formatting: if L = 0
    Columns("L:L").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if M&O = 0
    Range("M:M,O:O").Select
    Range("O1").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if N = 0
    Columns("N:N").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if P is not empty
    Columns("P:P").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(P1))>0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if Q is anything but "Confirmed"
    Columns("Q:Q").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Confirmed", _
        TextOperator:=xlDoesNotContain
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 16764108
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if R is anything of Advice
    Columns("R:R").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Advice", _
        TextOperator:=xlDoesNotContain
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    'Conditional Formatting: if S anything but "N"
    Columns("S:S").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="N", _
        TextOperator:=xlDoesNotContain
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("G1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'   Filter for Before Tax (K)& Medical (G)
    ActiveSheet.Range("$A$1:$Y$1440").AutoFilter Field:=8, Criteria1:="Medical"
    ActiveSheet.Range("$A$1:$Y$1440").AutoFilter Field:=11, Criteria1:= _
        "Before-Tax"
        
'   ???
    'Rows("1:1").Select
    'With ActiveWindow
        '.SplitColumn = 0
        '.SplitRow = 1
    'End With
    
    'Insert and name new sheets
    'ActiveWindow.FreezePanes = True
    Sheets("sheet1").Select
    Sheets("sheet1").Name = "001"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Summary"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Register"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Payments"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "Original_Reg"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "Notes"
    Sheets("001").Select
    Sheets("001").Move Before:=Sheets(5)
    Range("A1").Select
    
    'Include Styles for Book(c).xlsx
        'make sure WKBK is open prior to running macro
    ActiveWorkbook.Styles.Merge Workbook:=Workbooks("Book(c).xlsx")
End Sub
