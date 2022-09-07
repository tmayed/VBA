Function ReIndex1dArray(array1 As Variant) As Variant

    Dim array2 As Variant
    ReDim array2(LBound(array1) - LBound(array1) + 1 To UBound(array1) - LBound(array1) + 1)
    For i = LBound(array1) To UBound(array1)
        array2(i - LBound(array1) + 1) = array1(i)
    Next

    ReIndex1dArray = array2

End Function

Function ReIndex2dArray(array1 As Variant) As Variant

    Dim array2 As Variant
    ReDim array2(LBound(array1, 1) - LBound(array1, 1) + 1 To UBound(array1, 1) - LBound(array1, 1) + 1, LBound(array1, 2) - LBound(array1, 2) + 1 To UBound(array1, 2) - LBound(array1, 2) + 1)
    For i = LBound(array1, 1) To UBound(array1, 1)
        For j = LBound(array1, 2) To UBound(array1, 2)
            array2(i - LBound(array1, 1) + 1, j - LBound(array1, 2) + 1) = array1(i, j)
        Next
    Next

    ReIndex2dArray = array2

End Function

Function Array2dToCSVString(sArray As Variant) As String

    Dim sText As String
    sText = ""
    For i = LBound(sArray, 1) To UBound(sArray, 1)
        For j = LBound(sArray, 2) To UBound(sArray, 2)
            Text = sArray(i, j)
            Text = Replace(Text, Chr(10), "")
            sText = sText & Replace(Text, ",", ";") & ","
        Next
        sText = Left(sText, Len(sText) - 1) & Chr(10)
    Next

    Array2dToCSVString = sText

End Function

Function Print1dArray(array1 As Variant, SheetName As String)

    CreateCleanSheet (SheetName)

    Dim ReIndexArray As Variant
    ReIndexArray = ReIndex1dArray(array1)

    For i = LBound(ReIndexArray) To UBound(ReIndexArray)
        Sheets(SheetName).Cells(i, 1) = ReIndexArray(i)
    Next

End Function

Function Print2dArray(array1 As Variant, SheetName As String)

    On Error Resume Next
    CreateCleanSheet (SheetName)

    Dim array2 As Variant
    array2 = ReIndex2dArray(array1)

    For i = LBound(array2, 1) To UBound(array2, 1)
        For j = LBound(array2, 2) To UBound(array2, 2)
            Sheets(SheetName).Cells(i, j) = array2(i, j)
        Next
    Next

End Function

Function Print1dArrayP(array1 As Variant, SheetName As String, Row As Integer, Col As Integer)

    Dim ReIndexArray As Variant
    ReIndexArray = ReIndex1dArray(array1)

    For i = LBound(ReIndexArray) To UBound(ReIndexArray)
        Sheets(SheetName).Cells(i + Row - 1, Col) = ReIndexArray(i)
    Next

End Function

Function Print2dArrayP(array1 As Variant, SheetName As String, Row As Integer, Col As Integer)

    Dim array2 As Variant
    array2 = ReIndex2dArray(array1)

    For i = LBound(array2, 1) To UBound(array2, 1)
        For j = LBound(array2, 2) To UBound(array2, 2)
            Sheets(SheetName).Cells(i + Row - 1, j + Col - 1) = array2(i, j)
        Next
    Next

End Function

Function CSV2Array(Text As String) As Variant

    'find number of commas seperated entries in string
    Dim Entries As Integer
    Entries = 1
    For i = 1 To Len(Text)
        If Mid(Text, i, 1) = "," Then
            Entries = Entries + 1
        End If
    Next

    'create array to hold entires
    Dim holder As Variant
    ReDim holder(1 To Entries)

    'set loop variables
    Dim SP As Integer, EP As Integer, Entry As Integer, Found As Boolean
    SP = 1
    EP = 1
    Entry = 1
    Found = False

    'loop through text charcter by character and fill array with entries
    For i = 1 To Len(Text)
        If Entry = Entries Then
            EP = Len(Text)
            i = EP
            Found = True
        ElseIf Mid(Text, i, 1) = "," Then
            EP = i - 1
            Found = True
        End If
        If Found Then
            holder(Entry) = Mid(Text, SP, EP - SP + 1)
            SP = i + 1
            EP = i + 1
            Found = False
            Entry = Entry + 1
        End If
    Next

    CSV2Array = holder

End Function

Function RangeTo1dArray(rng As Variant) As Variant

    'find size of range
    i = 0
    For Each element In rng
        i = i + 1
    Next

    'create array of same size as range
    Dim Hold As Variant
    ReDim Hold(1 To i)

    'populate array
    i = 0
    For Each element In rng
        i = i + 1
        Hold(i) = element
    Next

    RangeTo1dArray = Hold

    End Function


    Function Merge1dArrays(arr1 As Variant, arr2 As Variant) As Variant

    'merges together two 1d arrays into a single array

    arr1 = ReIndex1dArray(arr1)
    arr2 = ReIndex1dArray(arr2)

    Total = UBound(arr1) + UBound(arr2)

    Dim Merge As Variant
    ReDim Merge(1 To Total)

    i = 0
    For Each element In arr1
        i = i + 1
        Merge(i) = element
    Next
    For Each element In arr2
        i = i + 1
        Merge(i) = element
    Next

    Merge1dArrays = Merge

    End Function

Function Merge2dArrays(arr1 As Variant, arr2 As Variant) As Variant

    'merges together two 1d arrays into a single array

    arr1 = ReIndex2dArray(arr1)
    arr2 = ReIndex2dArray(arr2)

    Dim Merge As Variant

    If UBound(arr1, 2) = UBound(arr2, 2) Then

        Row = UBound(arr1, 1) + UBound(arr2, 1)
        Col = UBound(arr1, 2)
        
        ReDim Merge(1 To Row, 1 To Col)
        
        a = 0
        For i = 1 To UBound(arr1, 1)
            a = a + 1
            b = 0
            For j = 1 To UBound(arr1, 2)
                b = b + 1
                Merge(a, b) = arr1(i, j)
            Next
        Next
        For i = 1 To UBound(arr2, 1)
            a = a + 1
            b = 0
            For j = 1 To UBound(arr2, 2)
                b = b + 1
                Merge(a, b) = arr2(i, j)
            Next
        Next

    End If

    Merge2dArrays = Merge

End Function

Function SheetTo2dArray(Sht As String) As Variant

    'Turns Sheet into 2d array

    RemoveshtError (Sht)

    Dim Last As RowCol
    Last = LastRowCol(Sht)

    Dim Hold As Variant
    ReDim Hold(1 To Last.Row, 1 To Last.Col)

    For i = 1 To Last.Row
        For j = 1 To Last.Col
            Hold(i, j) = Sheets(Sht).Cells(i, j)
        Next
    Next

    SheetTo2dArray = Hold

End Function

Sub Array2dToCSVFile(Arr As Variant, OutputFile As String)

    'create output file
    Set OutFO = CreateObject("Scripting.FileSystemObject")
    Set OutF = OutFO.CreateTextFile(OutputFile, True)

    'create input file object
    Set InFO = CreateObject("Scripting.FileSystemObject")

    'iterate over files to combine
    For i = 1 To UBound(Arr, 1)
        
        Dim Line As String
        Line = ""
        For j = 1 To UBound(Arr, 2)
            Line = Line & Replace(Arr(i, j), ",", ";") & ","
        Next
        Line = Trim(Left(Line, Len(Line) - 1))
        Line = AsciiRestrict(Line)
                
        OutF.WriteLine (Line)
        
    'next file
    Next

    'close output file
    OutF.Close

End Sub
