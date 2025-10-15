Attribute VB_Name = "Module1"
Sub AutoBWDates()

'V 0.8 Add BWDates if missing in Row 5
'   -Start on D5
'   -Bug:For every column it add, it does not check another column in the end
'       Fix: Run from back to front "C to 0 Step -1"
'   -Bug:If a ben holiday is before a missing date, does not add the date
'       Fix: Added ElseIF after inictial check to check for missing date after Ben Hol
'   -Bug:Updated to run from back to front, will only add one missing date, so will have to be run multiple times to get all missing dates
'       -Fix: Add i=i+2 at the end of process will check the new date if there is a missing date after it

Dim c As Integer

c = Worksheets(1).Range("5:5").Cells.SpecialCells(xlCellTypeConstants).Count - 5



For i = c To 1 Step -1
    CurrentCell = Cells(5, 4 + i)
    NextCell = Cells(5, 5 + i)
    
    'checks to see if there is more that 14 days bwtween the current
    'and the next date and makes space and adds date via BWDate Macro
    
    'If Application.WorksheetFunction.EoMonth(CurrentCell + 5, 0) - _
        Application.WorksheetFunction.EoMonth(NextCell + 5, 0) = 0 _
        And Not NextCell - CurrentCell < 16 Then ' Check if Monthly
        
    If NextCell - CurrentCell > 16 And DatePart("d", CurrentCell + 14) < 29 Then 'If the second part is true >28 then the date is a missing Ben holiday
        Columns(5 + i).Insert Shift:=xlToRight
        Cells(5, 5 + i).Select
        Application.Run "PERSONAL.XLSB!BWDate"
        i = i + 2 'Will allow to check if there is a missing date after date that was just added
       
       
    'If the missing date is after a ben holiday this should add the missing date
    ElseIf NextCell - CurrentCell > 40 Then
       Columns(5 + i).Insert Shift:=xlToRight
       Cells(5, 5 + i).Select
       ' copy of BWDate macro but add 28 days instead of 14
       ActiveCell.FormulaR1C1 = "=RC[-1]+28"
       ActiveCell.Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
       Application.CutCopyMode = False
       Selection.Style = "Accent2"
       i = i + 2 'Will allow to check if there is a missing date after date that was just added
    
    'To catch and mark off cycles
    'ElseIf NextCell - CurrentCell < 13 And 12 < CurrentCell - Cells(5, 3 + i) < 16 Then
        'Cells(5, 4 + i).Style = "Accent5"
    
    End If
Next i

Range("D4").Select 'v0.801

End Sub

Sub Key_Checklist()
'
' Key_Checklist Macro
'Version 1.01 Updated to include Waived, Partial Arrears/Wash & Partial EE paid
' V1.2 by VAM 8/6/25
' -Updated so the bottom starts in based on "i"
' -"i" can be definded in one of two ways (only one should be active, the other commented out)
'       1)Formual that looks at last cell in column C
'       2)variable that can be updated Manually (default 34)
'V1.21 by VAM 8/20/25
'-Added "Union Change - No Charge" to Key
'-Set EE ID in Cell B1 as text
'
    Application.ScreenUpdating = False
'
    'Top Left of sheet
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "EE#"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "RFL"
    
    ' Saves Cell B1 as text v1.21
    Range("B1").Select
    ActiveCell.Formula = "='001'!A2" 'Pulled EE ID from first row of data from 001
    Range("B1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("A3").Select  'Commented out as well as next line. Used with the Save_RFL macro when active
    ActiveCell.FormulaR1C1 = _
        "=R[-2]C[1]&""_""&TEXT(TODAY(),""mm.dd.yyyy"")&""_RFL"""
    Range("A1:A2").Select
    Selection.Font.Bold = True
    
    
    
    
    'Bottom Left of sheet
    'Formula to find the last cell to help define where checklist should start v1.2
    ' To manually set, comment out and change value of the i below to set row
    i = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row + 6 'Last num minus 1 is total number of blank rows below last Benefit Plan v1.2
    
    'i = 34 ' Select Row Manually
        
    
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Waived"
    Selection.Style = "Recon2"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Arrears/Wash"
    Selection.Style = "Accent1"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Partial Arrears/Wash"
    'Selection.Style = "Bad"
    Selection.Style = "PartialArrears"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "EE owes"
    Selection.Style = "Bad"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = " Partial EE paid"
    Selection.Style = "Bad"
    Selection.Style = "Partial Oracle"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "EE paid"
    Selection.Style = "Neutral"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "UC owes"
    Selection.Style = "Good"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Refund "
    Selection.Style = "Accent6"
    
    'v1.21
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Union Change - No Charge"
    Selection.Style = "60% - Accent5"
    
    i = i + 2 ' will leave a blank space after refund
    j = i ' Used to select section to be bold on line 141 in v 1.2
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Case#"
    
    i = i + 3 ' will leave two blank spaces after Case#
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "OTD"
    
    i = i + 3 ' will leave two blank spaces after OTD
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Register"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Payments"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Adjustments in Oracle"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Summary"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Payment Calc"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Sticky Note "
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Create Case"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Add to OTD (Monthly or Bi-Weekly)"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Zero out Oracle "
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Update RFL"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Update Metrics"
    
    i = i + 1
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = "Save to Recon Folder "
    
    'i = i + 1
    
    Columns("A:A").ColumnWidth = 16.5 'Updated to fit "Partial Arrears/Wash", originaly 13.11 v1.2
    
    
    Range("A" & j, "A" & i).Select
    Selection.Font.Bold = True
    
    Application.ScreenUpdating = True
    
    Range("A1").Select
    
End Sub

Sub Oracle_Payments()
Attribute Oracle_Payments.VB_ProcData.VB_Invoke_Func = "I\n14"
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


Sub PrepPvTbl()
Attribute PrepPvTbl.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Test_PrepPvTbl Macro V.1.4
'Version 1.2 VAM
'   -Updated to not reference PV_Summary Macro
'   -Will now clear rows that have "total" in column A
'   -Will now add Left boarder on column with years
'Version 1.3 VAM 7/16/23
'   -Will now auto fill in After tax for the first "Voluntary *" Benefits & Nontaxable for first Plan type after "Basic Disability"
'   -Will now move Disabilitys down to the bottom of the rows
'Version 1.4
'   -Added a message box at the end asking to run AutoBWDates Macro


'   Keyboard Shortcut: Ctrl+Shift+P
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
        On Error Resume Next
        .PivotItems("UC Retirement Plan 2016").Visible = False
        On Error Resume Next
        .PivotItems("Employee Assistance Program").Visible = False
        On Error Resume Next
        .PivotItems("DC Supplement to UCRP").Visible = False
        On Error Resume Next
        .PivotItems("DC Choice in Lieu of UCRP").Visible = False
        On Error Resume Next
        .PivotItems("Health Savings Account").Visible = False
        On Error Resume Next 'Updated 6/12/25 Per Denise 1.1.1
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
    ActiveSheet.Previous.Select 'can be update to move to sheet 1
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Rows("3:3").Select
    Selection.ClearContents
    Rows("3:5").Select
    Selection.Style = "Header(c)"
    Range("A6").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Style = "Recon"
    Range("D6").Select
    ActiveWindow.FreezePanes = True
        
   
    
    ' Loop to remove subtotals in Column A (v1.2)
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row
    Dim i As Long
    For i = TRCount To 6 Step -1
        If InStr(1, Cells(i, 1).Value, "Total", vbTextCompare) > 0 Then
            Range("A" & i).EntireRow.Clear
        End If
    Next i
    
    'Insert left boarders to on columns with years (v1.2)
    Dim j As Integer
    Dim yrs As Integer
    
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row - 3
    yrs = Worksheets(1).Range("4:4").Cells.SpecialCells(xlCellTypeConstants).Count
    
    Range("A4").Select
    If yrs > 1 Then '(v1.3.1) update to run if there is more than one year
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
    
    
    'Add "After-Tax" Lable  to first "Voluntary * Disability" in Column A (v1.3)
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row
    For k = 6 To TRCount
        If InStr(1, Cells(k, 2).Value, "Voluntary", vbTextCompare) > 0 Then
            Range("A" & k).FormulaR1C1 = "After-Tax"
            Exit For
        End If
    Next k
    
    
    ' Add "Nontaxable Benefit" to line following "Basic Disability" Plan (v1.3)
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
    
    ' Move Voluntary Disability Rows to the bottom (v1.3)
        '(v1.3.1) added "or" to if formula logic, else would not move Voluntary benefits if it is the only After-tax benefit
    
    If WorksheetFunction.CountIf(Range("A6:A" & TRCount), "After-Tax") = 2 Or _
        WorksheetFunction.CountIf(Range("B6:B" & TRCount), "Voluntary*") > 0 Then
        
        Range("A6").Select
        If WorksheetFunction.CountIf(Range("A6:A" & TRCount), "After-Tax") = 2 Then
            '(v1.3.1) If there are 2 "After-Tax" will move to the second, else stay on the first
            Selection.End(xlDown).Select
        End If
        Range(Selection, Selection.End(xlDown)).Select
        r = Selection.Rows.Count - 2
        ActiveCell.Resize(r).EntireRow.Cut
        Rows(TRCount + 2).Select
        Selection.Insert Shift:=xlDown
    End If
    
        '(v1.3.1) If only after tax benefits is disability, will leave row 6 blank. Code will check if its blank and delete row
    If WorksheetFunction.CountA(Range("6:6")) = 0 Then
        Range("A6").EntireRow.Delete
    End If
    
    
    'Move Basic Disability row/s to the bottom (v1.3)
    Dim row As Long
    
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row '(v1.3.1)
    
    If WorksheetFunction.CountIf(Range("A6:A" & TRCount), "Nontaxable Benefit") = 2 Then
        row = Application.WorksheetFunction.Match("Nontaxable Benefit", Range("A1:A" & TRCount), 0)
        Range("A" & row).Select
        Range(Selection, Selection.End(xlDown)).Select
        r = Selection.Rows.Count - 1
        ActiveCell.Resize(r).Select
        Selection.EntireRow.Select
        Selection.Cut
        Rows(TRCount + 2).Select  '(v1.3.1)
        Selection.Insert Shift:=xlDown
    End If
    
            
    
     'Run next Macro
    Application.Run "PERSONAL.XLSB!Key_Checklist"
    'ActiveWindow.Zoom = 85
    Columns("D:IV").EntireColumn.AutoFit
          
    'Create message box asking to run AutoBWDates Macro v1.4
   Response = MsgBox("Fill in missing BW dates?", vbYesNo, "Biweekly?")
   If Response = vbYes Then
        Application.Run "PERSONAL.XLSB!AutoBWDates"
   End If
    
    Columns("D:IV").EntireColumn.AutoFit 'v1.4.1
    
 End Sub

Sub RASC_Summary()
'
' RASC_Summary Macro V1.1
' V0.99 7/7/25
' v 1.1 7/18/25 VAM
'   -Remove zero amount line
'   -Format for printing
'   -Updated worksheet referance to ws
'v 1.1.1 8/7/25 VAM
'   -If there is Positive amount CM, will lead to replacing several Invoices with blank

   'Set Variable to count rows
    Dim TRCount As Long
    Dim PTCount As Long
    Dim Name As String
    Dim ws As Worksheet
    
    
    TRCount = Cells(Sheets(1).Rows.Count, 1).End(xlUp).row
    PTCount = Cells(Sheets(1).Rows.Count, 9).End(xlUp).row
    Name = Range("F2") 'Get EE name
    Set ws = ActiveSheet
    
    'Application.ScreenUpdating = False

    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Invoice #"
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "EE ID"
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Style = "Comma"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
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
    Selection.Delete Shift:=xlToLeft
   
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("A1:I1").Select
    Selection.Interior.Color = RGB(191, 191, 191)
    Selection.Font.Bold = True
    'Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    'Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    'Selection.Borders(xlEdgeLeft).LineStyle = xlNone
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
    'Selection.Borders(xlEdgeRight).LineStyle = xlNone
    'Selection.Borders(xlInsideVertical).LineStyle = xlNone
    'Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A3").Select
    Columns("A:A").ColumnWidth = 46.22
    Columns("H:H").ColumnWidth = 18.78
    Range("A1").Select
    ws.AutoFilter.Sort. _
        SortFields.Clear
    ws.AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("G1:G" & TRCount), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=0
    
    ' Loop to remove zero values (v1.1)
    TRCount = Cells(Sheets(1).Rows.Count, 1).End(xlUp).row
    Dim j As Long
    For j = TRCount To 2 Step -1
        If Cells(j, 7).Value = 0 Then
            Range("C" & j).EntireRow.Delete
        End If
    Next j
    
    ' Added: if there is Positive CM, Subotal will replace several invoices with blanks (v1.1.1)
    ws.AutoFilter.Sort. _
        SortFields.Clear
    ws.AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("I1:I" & TRCount), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Subtotal
    PTCount = Cells(Sheets(1).Rows.Count, 9).End(xlUp).row
    
    If PTCount > 1 Then 'If there are no Previous transactions PTCount = 1 and this step would not be nessesary
        Range("I2:I" & PTCount).Select
        Selection.Copy
        Range("C2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If

    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=R[2]C[3]"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 1).Formula = Name
    Columns("E:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Payment Received"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Amount Due"
    Range("G3").Select
    Range("F3").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Rows("2:2").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ws.AutoFilter.Sort. _
        SortFields.Clear
    ws.AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("B2"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ws.AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=9
    Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5, 6, 7), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
        
    ' Loop to format Subtotal Rows
    TRCount = Cells(Sheets(1).Rows.Count, 2).End(xlUp).row
    Dim i As Long
    For i = 2 To TRCount
        If InStr(1, Cells(i, 2).Value, "Total", vbTextCompare) > 0 Then
            Range("B" & i).Clear 'Clear Default total
            'Format subtotal row
            Range("A" & i & ":G" & i).Select
            Selection.Font.Bold = True
            Selection.Interior.Color = RGB(217, 217, 217)
            Selection.Style = "Comma"
            Range("D" & i).Select
            ActiveCell.FormulaR1C1 = "Total"
            Selection.HorizontalAlignment = xlRight
            Range("G" & i).Formula2R1C1 = "= RC[-2]-RC[-1]"
        End If
    Next i
    
    'Format Balance due summary
    Range("F" & TRCount + 1 & ":G" & TRCount + 1).Select
        Selection.Interior.Color = RGB(217, 217, 217)
        Selection.Font.Bold = True
        Selection.HorizontalAlignment = xlRight
        Selection.Style = "Comma"
    Range("F" & TRCount + 1).Formula2R1C1 = "Balance Due"
    Cells(TRCount + 1, 7).Formula = "=SUBTOTAL(9,G3:G" & TRCount & ")"
    
    'Format Header
    Range("A1:G2").Select
    Selection.Interior.Color = RGB(217, 217, 217)
    Selection.Font.Bold = True
    
    'Outside border
    Range("A1:G" & TRCount + 1).BorderAround _
        ColorIndex:=1, Weight:=xlThick
    
    'Format Coverage Month
    Range("C:C").NumberFormat = "MMMM YYYY"
    
    'Format Invoice#,Invoice Date, Coverage Month
    Range("A:C").Select
        Selection.Font.Bold = True
        Selection.VerticalAlignment = xlCenter
    
    'Remove Defalult total
    Range("A" & TRCount & ":G" & TRCount).Select
        Selection.ClearContents
        Selection.Cells.Interior.Pattern = xlNone
        
    'Format for Printing (v1.1)
    
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintTitleRows = "$1:$2"
        .CenterFooter = "Page &P of &N"
        .Zoom = False   ' Set to false or FitToPage* will not work
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Set to false or would auto set to 1 and made to fit in one page
    End With
     
        
    
End Sub

Sub Team_001()
Attribute Team_001.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Team_001 Macro
'   Updated to not rely on
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
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
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
        "Nontaxable Benefit"
        

    
    
    'Insert and name new sheets
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

Sub Team_Register()
Attribute Team_Register.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Team_Register Macro
' Team Register updated for CSV export 6/5/24
' Version2.0 7/2/25 by VAM
'   Set Variable to count rows
' Version 2.2 7/10/25 by VAM
'   Remove zero value from data
    Dim TRCount As Long
    Dim PTCount As Long
    TRCount = Cells(Sheets(1).Rows.Count, 1).End(xlUp).row
    PTCount = Cells(Sheets(1).Rows.Count, 9).End(xlUp).row
    
    Application.ScreenUpdating = False

    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Invoice #"
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "EE ID"
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Style = "Comma"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
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
    Selection.Delete Shift:=xlToLeft
    Selection.ColumnWidth = 16.67
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
    Columns("H:H").ColumnWidth = 18.78
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
    
     ' Loop to remove zero values V2.1
    TRCount = Cells(Sheets(1).Rows.Count, 7).End(xlUp).row
    Dim j As Long
    For j = TRCount To 2 Step -1
        If Cells(j, 7).Value = 0 Then
            Range("C" & j).EntireRow.Delete
        End If
    Next j
    
    ' Subtotal V2.0
    PTCount = Cells(Sheets(1).Rows.Count, 9).End(xlUp).row
    
    If PTCount > 1 Then 'If there are no Previous transaction number PTCount = 1 and this step would not be nessesary
        Range("I2:I" & PTCount).Select
        Selection.Copy
        Range("C2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("C1:C" & TRCount), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("C1:G" & TRCount).Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    
     ' Loop to format Subtotal Rows V2.0
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row
    Dim i As Long
    For i = 2 To TRCount
        If InStr(1, Cells(i, 3).Value, "Total", vbTextCompare) > 0 Then
            Range("C" & i & ":G" & i).Select
            Selection.Font.Bold = True
            Selection.Interior.Color = RGB(228, 158, 221)
            Selection.Style = "Comma"
            If i = TRCount Then
                Range("C" & i & ":G" & i).Select
                Selection.Interior.Color = RGB(192, 230, 245)
            End If
        End If
    Next i
    
    'Set up Pmt & Applied calc V2.0
    TRCount = Cells(Sheets(1).Rows.Count, 3).End(xlUp).row
    
    Range("F" & TRCount + 2).Select
    ActiveCell.FormulaR1C1 = "Pmts "
    Range("F" & TRCount + 3).Select
    ActiveCell.FormulaR1C1 = "Applied "
    Range("F" & TRCount + 4).Select
    ActiveCell.FormulaR1C1 = "Balance "
    Range("G" & TRCount + 2 & ":G" & TRCount + 4).Select
    Selection.Style = "Comma"
    Range("F" & TRCount + 2 & ":G" & TRCount + 4).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("F" & TRCount + 2 & ":G" & TRCount + 4).Select
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
    Range("G" & TRCount + 2).Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G" & TRCount + 3).Select
    ActiveCell.FormulaR1C1 = "=R[-3]C"
    Range("G" & TRCount + 4).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-2]C-R[-1]C"
      
    
End Sub

Sub AutoWaivedDisability()
'
' AutoWaivedDisability Macro
' v 1.0 by VAM
'
' Keyboard Shortcut: Ctrl+Shift+Q
'

    Dim Sel As Range
    
    Set Sel = Selection
        
    'Set border around waived disability
    Sel.BorderAround _
        ColorIndex:=1, Weight:=xlThin
    
    For Each cell In Sel
        If Not IsEmpty(cell) And IsNumeric(cell) Then 'IsNumeric returns blank cells as true, so added "Not IsEmpty()" to void empty cells
            cell.Style = "Accent6"
        ElseIf IsEmpty(cell) And cell.Interior.Color = RGB(217, 217, 217) Then ' Check to see if cell is empty and Recon grey
            cell.Style = "Recon2"
        End If
    Next cell
       
End Sub

Sub BWDate()
Attribute BWDate.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' BWDate Macro
'

'
    ActiveCell.FormulaR1C1 = "=RC[-1]+14"
    'ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Style = "Accent2"
End Sub

Sub PartialArrearsWas()
Attribute PartialArrearsWas.VB_ProcData.VB_Invoke_Func = "B\n14"
'
' PartialArrearsWas Macro
'
' Keyboard Shortcut: Ctrl+Shift+B
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub

Sub Save_RFL()
Attribute Save_RFL.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' Save_RFL Macro
'
' Make sure to update "Folder" prior to uses
'
' Keyboard Shortcut: Ctrl+Shift+S
    Dim Name As String
    Dim Folder As String
    
    Folder = "C:\Users\vmedina\Box\#PII-UCPC-FN-BB\#Account Reconciliation\"
    
    Worksheets("Summary").Range("A3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
        
    Name = Range("A3")
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        Folder & Name & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

