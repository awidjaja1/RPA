Attribute VB_Name = "Module3"
Sub PrepPvTbl()
Attribute PrepPvTbl.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Test_PrepPvTbl Macro V.1.3
'Version 1.2
'   -Updated to not reference PV_Summary Macro
'   -Will now clear rows that have "total" in column A
'   -Looking into Adding Left boarder on column with years
'Version 1.3
'   -Will now auto fill in After tax for the first "Voluntary *" Benefits & Nontaxable for first Plan type after "Basic Disability"
' Keyboard Shortcut: Ctrl+Shift+P
'
Dim TRCount As Long
    
    With ActiveSheet.PivotTables(1).PivotFields( _
        "Deduction Classification")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(1).PivotFields("Plan Type")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables(1).PivotFields("Benefit Plan")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables(1).AddDataField ActiveSheet.PivotTables( _
        1).PivotFields("Current Deduction"), "Sum of Current Deduction", _
        xlSum
    On Error Resume Next ' update 1.1
    ActiveSheet.PivotTables(1).PivotFields( _
        "Months (Paycheck Issue Date)").Orientation = xlHidden
    On Error Resume Next 'update 1.1
    ActiveSheet.PivotTables(1).PivotFields( _
        "Quarters (Paycheck Issue Date)").Orientation = xlHidden
    Range("D5").Select
    ActiveSheet.PivotTables(1).PivotFields( _
        "Years (Paycheck Issue Date)").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    Range("D4").Select
    ActiveSheet.PivotTables(1).PivotFields( _
        "Years (Paycheck Issue Date)").ShowDetail = True
    Range("B12").Select
    ActiveSheet.PivotTables(1).PivotFields("Plan Type").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    With ActiveSheet.PivotTables(1).PivotFields("Plan Type")
        On Error Resume Next
        .PivotItems("403(b)").Visible = False
        On Error Resume Next
        .PivotItems("403(b) ROTH").Visible = False
        On Error Resume Next
        .PivotItems("457(b)").Visible = False
        On Error Resume Next
        .PivotItems("457(b) ROTH").Visible = False
        On Error Resume Next
        .PivotItems("DCP Contribution").Visible = False
        On Error Resume Next
        .PivotItems("Flex Spending - Dependent Care").Visible = False
        On Error Resume Next
        .PivotItems("Flex Spending - Health").Visible = False
        On Error Resume Next
        .PivotItems("General Deduction").Visible = False
        On Error Resume Next
        .PivotItems("Identity Theft Protection").Visible = False
        On Error Resume Next
        .PivotItems("UC Retirement Plan").Visible = False
        .PivotItems("UC Retirement Plan 2016").Visible = False
        On Error Resume Next
        .PivotItems("Employee Assistance Program").Visible = False
        On Error Resume Next
        .PivotItems("DC Supplement to UCRP").Visible = False
        On Error Resume Next
        .PivotItems("DC Choice in Lieu of UCRP").Visible = False
        On Error Resume Next
        .PivotItems("Health Savings Account").Visible = False
        On Error Resume Next 'Updated 6/12/25 Per Denise 1.1
        .PivotItems("Summer Salary DC Plan").Visible = False
    End With
    
    'Get to the last collumn/row of usable data
     Application.Goto Reference:="R100C1"
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(rowOffset:=-2, columnOffset:=-1).Activate.Select
    Range(Selection, Cells(1)).Select
    
    'Application.Run "PERSONAL.XLSB!PV_Summary" ''Added this Macro below(1.2)
    
    Selection.Copy
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Rows("3:5").Select
    Selection.Style = "Header(c)"
    Range("A6").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Style = "Recon"
    Range("D6").Select
    ActiveWindow.FreezePanes = True
   
    
    ' Loop to remove subtotals in Column A (1.2)
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).Row
    Dim i As Long
    For i = TRCount To 6 Step -1
        If InStr(1, Cells(i, 1).Value, "Total", vbTextCompare) > 0 Then
            Range("A" & i).EntireRow.Clear
        End If
    Next i
    
    'Insert left boarders to on columns with years
    Dim j As Integer
    Dim yrs As Integer
    
    TRCount = Cells(Sheets(1).Rows.Count, 2).End(xlUp).Row - 3
    yrs = Worksheets(1).Range("4:4").Cells.SpecialCells(xlCellTypeConstants).Count
    
    Range("A4").Select
    If yrs > 0 Then
        For j = 1 To yrs
            Selection.End(xlToRight).Select
            Selection.Resize(TRCount).Select
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next j
    End If
    
    
    'Add "After-Tax" Lable  to first "Voluntary * Disability" in Column A (V1.3)
    TRCount = Cells(Sheets(1).Rows.Count, 2).End(xlUp).Row
    For k = 6 To TRCount
        If InStr(1, Cells(k, 2).Value, "Voluntary", vbTextCompare) > 0 Then
            Range("A" & k).FormulaR1C1 = "After-Tax"
            Exit For
        End If
    Next k
    
    
    ' Add "Nontaxable Benefit" to line following "Basic Disability" Plan (v1.3)
    ' Does not check
    Dim X As Boolean
    
    X = False
    
    For n = 6 To TRCount
        If Cells(n, 2).Value <> "" And (X = True) Then
            Range("A" & n).FormulaR1C1 = "Nontaxable Benefit"
            Exit For
        ElseIf InStr(1, Cells(n, 2).Value, "Basic Disability", vbTextCompare) > 0 Then
                X = True
        End If
    Next n
            
    
     'Run next Macro
    Application.Run "PERSONAL.XLSB!Key_Checklist"
    ActiveWindow.Zoom = 85
    Columns("D:IV").Select
    Columns("D:IV").EntireColumn.AutoFit
    
 End Sub
