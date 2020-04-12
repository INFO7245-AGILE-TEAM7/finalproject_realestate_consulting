Attribute VB_Name = "Module3"
Sub FinancialReportMacro()
Attribute FinancialReportMacro.VB_ProcData.VB_Invoke_Func = " \n14"
' Macro1 Macro
'
    Range("F1").Value = "OT"
    Range("G1").Value = "Total"
    Range("F2:F" & Cells(Rows.count, "A").End(xlUp).Row).Formula = "=D2*0.1338"
    Range("G2:G" & Cells(Rows.count, "A").End(xlUp).Row).Formula = "=SUM(D2:F2)"
    
    totalSum = Application.WorksheetFunction.Sum(Range("G2:G" & Cells(Rows.count, "A").End(xlUp).Row))
    Range("J1").Value = "Amount to Investor"
    Range("J2").Value = totalSum * 0.75
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.EntireColumn.AutoFit
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
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
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveCell.Range("A1:G1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    
    ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.EntireColumn.AutoFit
    ActiveCell.Offset(0, 9).Range("A1:A2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.NumberFormat = "$#,##0.00"
    ActiveCell.Offset(0, -6).Columns("A:D").EntireColumn.Select
    Selection.Style = "Currency"
End Sub

