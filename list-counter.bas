Attribute VB_Name = "Module1"
Sub List_count_clean_relative()
Attribute List_count_clean_relative.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' List_count_clean_relative Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Range(Selection, Selection.End(xlDown)).Select
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
    Selection.Style = "Check Cell"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 7960846
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "COUNT"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(-1, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    ActiveCell.Columns("A:B").EntireColumn.RemoveDuplicates Columns:=1, Header:= _
        xlYes
    ActiveCell.Columns("A:B").EntireColumn.Select
    ActiveCell.Columns("A:B").EntireColumn.EntireColumn.AutoFit
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 Key:=ActiveCell _
        .Offset(0, 1).Range("A1:A4863"), SortOn:=xlSortOnValues, Order:=xlDescending _
        , DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 Key:=ActiveCell _
        .Offset(0, 1).Range("A1:A4863"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Selection
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
