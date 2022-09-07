Function LastRow(SheetName As String) As Long
    Dim Last As RowCol
    Last = LastRowCol(SheetName)
    LastRow = Last.Row
End Function

Function LastCol(SheetName As String) As Long
    Dim Last As RowCol
    Last = LastRowCol(SheetName)
    LastCol = Last.Col
End Function

Function LastRowCol(SheetName As String) As RowCol

    'get last row/col variables
    Dim MaxColCheck As Long, LastRow As Long, LastCol As Long, i As Long, j As Long, Blank As Long
    MaxColCheck = 10

    'get last row/col
    LastRow = 1
    LastCol = 1
    i = 1
    j = 1
    Blank = 0

    Do While True
        'check if number of contigious blank columns has been reached
        If Blank >= MaxColCheck Then
            Exit Do
        End If
        'find last row is column j
        i = Sheets(SheetName).Cells(Sheets(SheetName).Rows.Count, j).End(xlUp).Row
        'if cell is an error, set cell to blank and find last row again
        If IsError(Sheets(SheetName).Cells(i, j)) Then
            Sheets(SheetName).Cells(i, j) = ""
        Else
            'if last row cell is not an error check if row is one and cell is blank, if yes then is a blank column
            If i = 1 And Sheets(SheetName).Cells(i, j) = "" Then
                Blank = Blank + 1
            Else
                'if not a blank column then check if row is greater than current max row
                Blank = 0
                LastCol = j
                If i > LastRow Then
                    LastRow = i
                End If
            End If
            'go to next column
            j = j + 1
        End If
    Loop

    'return max row and col is sheet as data type Row Col
    Dim Last As RowCol
    Last.Row = LastRow
    Last.Col = LastCol

    LastRowCol = Last

End Function

Function CreateCleanSheet(SheetName As String)

    Dim datasht As Boolean
    datasht = False
    For Each Sht In ActiveWorkbook.Worksheets
        If Sht.Name = SheetName Then
            datasht = True
            Sheets(SheetName).Select
            If ActiveSheet.AutoFilterMode Then
                Selection.AutoFilter
            End If
            Cells.Select
            Selection.ClearContents
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Next Sht
    If datasht = False Then
        ActiveWorkbook.Worksheets.Add
        ActiveSheet.Name = SheetName
    End If
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
    If ActiveWindow.FreezePanes = True Then
        ActiveWindow.FreezePanes = False
    End If
    Range("A1").Select
    ActiveWindow.ScrollColumn = Range("A1").Column
    ActiveWindow.ScrollRow = Range("A1").Row
        

End Function

#Const EarlyBound = False
Function WriteUTF8(sText As String, sFile As String) As Boolean
  ' Returns True if sText saved successfully as UTF-8 in sFile

    On Error GoTo Oops

    #If EarlyBound Then
        ' Requires a reference to Microsoft ActiveX Data Objects
        With New ADODB.stream
    #Else
        ' No reference required
        Const adTypeText As Long = 2
        Const adSaveCreateOverWrite As Long = 2
        With CreateObject("ADODB.Stream")
    #End If
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText sText
        .SaveToFile Filename:=sFile, Options:=adSaveCreateOverWrite
        WriteUTF8 = True
        End With
        Exit Function

    Oops:
    MsgBox Err.Description
  
End Function

Function AppsOff()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Function

Function AppsOn()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Function

Function CalcOn()
    Application.Calculation = xlCalculationAutomatic
End Function

Function CalcOff()
    Application.Calculation = xlCalculationManual
End Function

Function RemoveshtError(SheetName As String)

Dim Last As RowCol

    Last = LastRowCol(SheetName)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If IsError(Sheets(SheetName).Cells(i, j)) Then
                Sheets(SheetName).Cells(i, j) = ""
            End If
        Next
    Next

End Function

Sub RemoveSheets(Keep As Variant)

    'remove all sheets in workbook that are not in array Keep

    For Each Sht In ActiveWorkbook.Worksheets
        Dim DeleteSht As Boolean
            DeleteSht = True
        For Each ShtKeep In Keep
            If Sht.Name = ShtKeep Then
                DeleteSht = False
                Exit For
            End If
        Next
        If DeleteSht = True Then
            Sheets(Sht.Name).Delete
        End If
    Next

End Sub

Function AsciiRestrict(Line As String) As String

    Dim Chars As String
    Chars = ""
    For j = 1 To Len(Line)
        If Asc(Mid(Line, j, 1)) > 31 And Asc(Mid(Line, j, 1)) < 127 Then
            Chars = Chars & Chr(Asc(Mid(Line, j, 1)))
        End If
    Next

    AsciiRestrict = Chars

End Function

Sub LastColBorder(ShtName As String)

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    With Sheets(ShtName).Columns(Last.Col)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With

End Sub

Sub RemoveBorders(ShtName As String)

    With Sheets(ShtName).Cells
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    End Sub

    Sub AddBorderAtCol(ShtName As String, Col As Long)

    With Sheets(ShtName).Columns(Col)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThick
        End With
    End With
    
End Sub

Sub RmFilterSplitFreeze(ShtName As String)

    current = ActiveSheet.Name
    Sheets(ShtName).Activate
    If ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter
    End If
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
    If ActiveWindow.FreezePanes = True Then
        ActiveWindow.FreezePanes = False
    End If
    Sheets(current).Activate

End Sub

Sub AddFilterSplitFreeze(ShtName As String)

    RmFilterSplitFreeze ShtName
    current = ActiveSheet.Name
    Sheets(ShtName).Activate

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    Range(Cells(1, 1), Cells(1, Last.Col)).AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Sheets(current).Activate

End Sub

Sub RangeGet(ShtName As String, Optional RowTrim As Integer = 0, Optional ColTrim As Integer = 0)

    Dim Last As RowCol
    Last = LastRowCol(ShtName)
            
    Dim RangeSet As String
    RangeSet = "=" & ShtName & "!R1C1:R" & Last.Row - RowTrim & "C" & Last.Col - ColTrim

    Dim Found As Boolean
    Found = False
    For Each Name In ActiveWorkbook.Names
        If Name.Name = ShtName & "Data" Then
            Found = True
        End If
    Next
    If Found = True Then
        With ActiveWorkbook.Names(ShtName & "Data")
            .Name = ShtName & "Data"
            .RefersToR1C1 = RangeSet
        End With
    Else
        ActiveWorkbook.Names.Add Name:=ShtName & "Data", RefersToR1C1:=RangeSet
    End If

End Sub

Sub UpdatePivots()
    ActiveWorkbook.RefreshAll
End Sub

Sub ClearRange(ShtName As String, StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)

    current = ActiveSheet.Name
    Sheets(ShtName).Activate
    Worksheets(ShtName).Range(Cells(StartRow, StartCol), Cells(EndRow, EndCol)).Clear
    Sheets(current).Activate

End Sub

Sub PasteSpecial(ShtName As String)

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If IsError(Sheets(ShtName).Cells(i, j)) Then
                Sheets(ShtName).Cells(i, j) = ""
            Else
                Sheets(ShtName).Cells(i, j) = Sheets(ShtName).Cells(i, j)
            End If
        Next
    Next

End Sub

Sub PrintColData(Headers As Variant, Formulas As Variant, ShtName As String, FirstCol As Long, LastRow As Long)

    'Check for row offset
    j = FirstCol
    For Each a In Headers
        Sheets(ShtName).Cells(1, j).FormulaR1C1 = a
        j = j + 1
    Next

    'print out column data
    j = FirstCol
    For Each a In Formulas
        For i = 2 To LastRow
            Sheets(ShtName).Cells(i, j).FormulaR1C1 = a
        Next
        j = j + 1
    Next

End Sub

Function FindInSheet(ShtName As String, Value As String) As RowCol

    'Find row and col of value in a sheet

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If Sheets(ShtName).Cells(i, j) = Value Then
                FindInSheet.Row = i
                FindInSheet.Col = j
                Exit Function
            End If
        Next
    Next
    FindInSheet.Row = 0
    FindInSheet.Col = 0

End Function

Function FindInSheetRow(ShtName As String, Value As String) As Integer

    'Find row and col of value in a sheet

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If Sheets(ShtName).Cells(i, j) = Value Then
                FindInSheetRow = i
                Exit Function
            End If
        Next
    Next
    FindInSheetRow = 0

End Function

Function FindInSheetCol(ShtName As String, Value As String) As Integer

    'Find row and col of value in a sheet

    Dim Last As RowCol
    Last = LastRowCol(ShtName)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If Sheets(ShtName).Cells(i, j) = Value Then
                FindInSheetCol = j
                Exit Function
            End If
        Next
    Next
    FindInSheetCol = 0

End Function

Function FindValRow(ShtName As String, Value As String, Last As RowCol) As Integer

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If Sheets(ShtName).Cells(i, j) = Value Then
                FindValRow = i
                Exit Function
            End If
        Next
    Next
    FindValRow = 0

End Function

Function FindValCol(ShtName As String, Value As String, Last As RowCol) As Integer

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            If Sheets(ShtName).Cells(i, j) = Value Then
                FindValCol = j
                Exit Function
            End If
        Next
    Next
    FindValCol = 0

End Function

Function FindChartType(No As Long) As String

    Dim ChartNo As String
    ChartNo = "" & No & ""

    Dim ChartTypes As Object
    Set ChartTypes = CreateObject("Scripting.Dictionary")

    ChartTypes("-4098") = "xl3DArea"
    ChartTypes("78") = "xl3DAreaStacked"
    ChartTypes("79") = "xl3DAreaStacked100"
    ChartTypes("60") = "xl3DBarClustered"
    ChartTypes("61") = "xl3DBarStacked"
    ChartTypes("62") = "xl3DBarStacked100"
    ChartTypes("-4100") = "xl3DColumn"
    ChartTypes("54") = "xl3DColumnClustered"
    ChartTypes("55") = "xl3DColumnStacked"
    ChartTypes("56") = "xl3DColumnStacked100"
    ChartTypes("-4101") = "xl3DLine"
    ChartTypes("-4102") = "xl3DPie"
    ChartTypes("70") = "xl3DPieExploded"
    ChartTypes("1") = "xlArea"
    ChartTypes("76") = "xlAreaStacked"
    ChartTypes("77") = "xlAreaStacked100"
    ChartTypes("57") = "xlBarClustered"
    ChartTypes("71") = "xlBarOfPie"
    ChartTypes("58") = "xlBarStacked"
    ChartTypes("59") = "xlBarStacked100"
    ChartTypes("15") = "xlBubble"
    ChartTypes("87") = "xlBubble3DEffect"
    ChartTypes("51") = "xlColumnClustered"
    ChartTypes("52") = "xlColumnStacked"
    ChartTypes("53") = "xlColumnStacked100"
    ChartTypes("102") = "xlConeBarClustered"
    ChartTypes("103") = "xlConeBarStacked"
    ChartTypes("104") = "xlConeBarStacked100"
    ChartTypes("105") = "xlConeCol"
    ChartTypes("99") = "xlConeColClustered"
    ChartTypes("100") = "xlConeColStacked"
    ChartTypes("101") = "xlConeColStacked100"
    ChartTypes("95") = "xlCylinderBarClustered"
    ChartTypes("96") = "xlCylinderBarStacked"
    ChartTypes("97") = "xlCylinderBarStacked100"
    ChartTypes("98") = "xlCylinderCol"
    ChartTypes("92") = "xlCylinderColClustered"
    ChartTypes("93") = "xlCylinderColStacked"
    ChartTypes("94") = "xlCylinderColStacked100"
    ChartTypes("-4120") = "xlDoughnut"
    ChartTypes("80") = "xlDoughnutExploded"
    ChartTypes("4") = "xlLine"
    ChartTypes("65") = "xlLineMarkers"
    ChartTypes("66") = "xlLineMarkersStacked"
    ChartTypes("67") = "xlLineMarkersStacked100"
    ChartTypes("63") = "xlLineStacked"
    ChartTypes("64") = "xlLineStacked100"
    ChartTypes("5") = "xlPie"
    ChartTypes("69") = "xlPieExploded"
    ChartTypes("68") = "xlPieOfPie"
    ChartTypes("109") = "xlPyramidBarClustered"
    ChartTypes("110") = "xlPyramidBarStacked"
    ChartTypes("111") = "xlPyramidBarStacked100"
    ChartTypes("112") = "xlPyramidCol"
    ChartTypes("106") = "xlPyramidColClustered"
    ChartTypes("107") = "xlPyramidColStacked"
    ChartTypes("108") = "xlPyramidColStacked100"
    ChartTypes("-4151") = "xlRadar"
    ChartTypes("82") = "xlRadarFilled"
    ChartTypes("81") = "xlRadarMarkers"
    ChartTypes("88") = "xlStockHLC"
    ChartTypes("89") = "xlStockOHLC"
    ChartTypes("90") = "xlStockVHLC"
    ChartTypes("91") = "xlStockVOHLC"
    ChartTypes("83") = "xlSurface"
    ChartTypes("85") = "xlSurfaceTopView"
    ChartTypes("86") = "xlSurfaceTopViewWireframe"
    ChartTypes("84") = "xlSurfaceWireframe"
    ChartTypes("-4169") = "xlXYScatter"
    ChartTypes("74") = "xlXYScatterLines"
    ChartTypes("75") = "xlXYScatterLinesNoMarkers"
    ChartTypes("72") = "xlXYScatterSmooth"
    ChartTypes("73") = "xlXYScatterSmoothNoMarkers"

    FindChartType = ChartTypes(ChartNo)

Set ChartTypes = Nothing

End Function

Function XORD(DataIn As String) As String

    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer

    Dim CodeKey As String
    CodeKey = "Hash"

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))

        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr

    XORD = strDataOut

End Function


Function XORE(DataIn As String) As String

    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer

    Dim CodeKey As String
    CodeKey = "Hash"

    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))

        temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring

        strDataOut = strDataOut + tempstring
    Next lonDataPtr

    XORE = strDataOut

End Function

Function SF(i As Integer) As String

    Dim Selections As Variant
    ReDim Selections(1 To 6)
    Selections(1) = "#,##0"
    Selections(2) = "[$-F400]h:mm:ss AM/PM"
    Selections(3) = "0.00%"
    Selections(4) = "m/d/yyyy"
    Selections(5) = "0.00"
    Selections(6) = "dd/mm/yyyy hh:mm:ss"

    SF = Selections(i)

End Function

Sub FormatCols(ShtName As String, Col As Integer, Formatting As String, Optional HozAlign As String = "")

    With Sheets(ShtName).Columns(Col)
        
        .NumberFormat = Formatting
        
        If LCase(HozAlign) = "left" Then
            .HorizontalAlignment = xlLeft
        ElseIf LCase(HozAlign) = "right" Then
            .HorizontalAlignment = xlRight
        ElseIf LCase(HozAlign) = "center" Or LCase(Alightment) = "middle" Then
            .HorizontalAlignment = xlCenter
        Else
            .HorizontalAlignment = xlLeft
        End If
        
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    
    End With

End Sub
