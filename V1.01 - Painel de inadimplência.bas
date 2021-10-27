Attribute VB_Name = "Módulo1"
Sub Geral()

    Call Data_Ult_Venda
    Call Base_Inicial_Etapa_01
    Call TD
    Call Base_Inicial_Etapa_02
    Call Base_Geral
    Call Base_Ação_Crítica
    
    Sheets("MACROS").Select
    Range("B7").Select

End Sub

Sub Data_Ult_Venda()

    Application.ScreenUpdating = False

    Sheets("DATA ÚLT. VENDA").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B3").Select
    Sheets("BD - DATAS").Select
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("C3").Select
    ActiveSheet.Range("$B$3:$C$100000").AutoFilter Field:=2, Criteria1:="<>-", _
        Operator:=xlAnd
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("DATA ÚLT. VENDA").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B3").Select
    Sheets("BD - DATAS").Select
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Selection.AutoFilter
    Range("E4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("DATA ÚLT. VENDA").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("BD - DATAS").Select
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("DATA ÚLT. VENDA").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Columns("C:C").Select
    Range("C3").Activate
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4), TrailingMinusNumbers:=True
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("C2").Select
    ActiveWorkbook.Worksheets("DATA ÚLT. VENDA").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("DATA ÚLT. VENDA").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C2:C300000"), SortOn:=xlSortOnValues, Order:=xlDescending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA ÚLT. VENDA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B3").Select
    ActiveSheet.Range("$B$2:$C$300000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("B3").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("B3").Select
    Columns("B:C").Select
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
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
    Range("B3").Select

    Application.ScreenUpdating = True

End Sub


Sub Base_Inicial_Etapa_01()

Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE INICIAL").Range("C4").Value)
    final = Abs(Worksheets("BASE INICIAL").Range("B4").Value)
 
    Do While atual > final
        Sheets("BASE INICIAL").Select
        Range("B6").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B6").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE INICIAL").Range("C4").Value)
        final = Abs(Worksheets("BASE INICIAL").Range("B4").Value)
    Loop

    Sheets("BASE INICIAL").Select
    Range("B5").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C4").Value > 0 Then
        linhaf = linhai - Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C4").Value < 0 Then
        linhaf = linhai + Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B6").Select
    Sheets("BD - INADIMPLÊNCIA").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BASE INICIAL").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select
    Range("AW6:BP6").Select
    Selection.Copy
    Range("AW7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

Sub TD()

Application.ScreenUpdating = False

    Sheets("TD").Select
    ActiveWorkbook.RefreshAll
    Range("F7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("F2:I2").Select
    Selection.Copy
    Range("D5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(-1, 2).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TD").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TD").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "H5"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TD").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("TD").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TD").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "G5"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TD").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("I2").Select
    Selection.Copy
    Range("I6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("I7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

Sub Base_Inicial_Etapa_02()

Application.ScreenUpdating = False

    Sheets("BASE INICIAL").Select
    Range("BQ2").Select
    Selection.Copy
    Range("BQ6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select
    ActiveWorkbook.RefreshAll
    
Application.ScreenUpdating = True

End Sub

Sub Base_Geral()

Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE GERAL").Range("C1").Value)
    final = Abs(Worksheets("BASE GERAL").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE GERAL").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE GERAL").Range("C1").Value)
        final = Abs(Worksheets("BASE GERAL").Range("B1").Value)
    Loop

    Sheets("BASE GERAL").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE INICIAL").Select
    Range("BP5").Select
    ActiveSheet.Range("$B$5:$BQ$10000").AutoFilter Field:=67, Criteria1:="=1", _
        Operator:=xlAnd
    Range("AX5:BO5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE GERAL").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Columns.AutoFit
    Range("J3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("J3:J10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("I3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("I3:I10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("L3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("L3:L10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("M3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("M3:M10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("N3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N3:N10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("O3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("O3:O10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("BASE INICIAL").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B6").Select
    Sheets("BASE GERAL").Select
    Range("B4").Select

Application.ScreenUpdating = True

End Sub

Sub Base_Ação_Crítica()

Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE - AÇÃO CRÍTICA").Range("C1").Value)
    final = Abs(Worksheets("BASE - AÇÃO CRÍTICA").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE - AÇÃO CRÍTICA").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE - AÇÃO CRÍTICA").Range("C1").Value)
        final = Abs(Worksheets("BASE - AÇÃO CRÍTICA").Range("B1").Value)
    Loop

    Sheets("BASE - AÇÃO CRÍTICA").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE INICIAL").Select
    Range("BQ5").Select
    ActiveSheet.Range("$B$5:$BQ$10000").AutoFilter Field:=68, Criteria1:="=1", _
        Operator:=xlAnd
    Range("AX5:BO5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE - AÇÃO CRÍTICA").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Columns.AutoFit
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BASE INICIAL").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B6").Select
    Sheets("BASE - AÇÃO CRÍTICA").Select
    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("E3:E10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("R3:R10000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("M3:M10000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE - AÇÃO CRÍTICA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B4").Select
    ActiveWorkbook.RefreshAll

Application.ScreenUpdating = True

End Sub

Sub Arquivo_Envio()

Application.ScreenUpdating = False

    ActiveWorkbook.Save
    ChDir _
        ActiveWorkbook.Path
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\Gestão de Inadimplência - MS - Dados de 01.03.20 a " & Worksheets("MACROS").Range("C14").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets(Array("MACROS", "BD - INADIMPLÊNCIA", "BD - CI - M0", "BD - DATAS", "DATA ÚLT. VENDA", _
        "BASE INICIAL", "TD")).Select
    Sheets("MACROS").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=3
    Sheets(Array("MACROS", "BD - INADIMPLÊNCIA", "BD - CI - M0", "BD - DATAS", "DATA ÚLT. VENDA", _
        "BASE INICIAL", "TD", "GRÁFICOS")).Select
    Sheets("GRÁFICOS").Activate
    ActiveWindow.SelectedSheets.Delete
    ActiveWindow.DisplayHeadings = False
    Sheets("TOP 199").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE - AÇÃO CRÍTICA").Select
    Selection.End(xlUp).Select
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE GERAL").Select
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWorkbook.Save

Application.ScreenUpdating = True

End Sub


