Sub RestorePlatePositions()
    '  プレート位置情報を読み込んで、各プレートを再配置する
    Dim platePosition As Worksheet, plateWs As Worksheet, dataWs As Worksheet
    Dim shp As Shape
    Dim i As Long, lastRow As Long
    Dim atdCol As Long, ovtCol As Long
    Dim currentDate As String
    Dim lastNameDict As Object
    Dim fullNameDict As Object
    Dim employeeCode As String
    Dim positionData As Variant
    Dim leftPos As Single, topPos As Single
    Dim name As String, lastName As String, firstName As String
    Dim nameToDisplay As String
    
    ' シートの設定
    Set platePosition = ThisWorkbook.Sheets("配置記録")
    Set plateWs = ThisWorkbook.Sheets("配置")
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    
    ' 最終行を取得
    lastRow = platePosition.Cells(platePosition.Rows.Count, "A").End(xlUp).Row
    
    ' 現在の日付から列を特定
    currentDate = Format(Date, "YYYYMMDD")
    atdCol = platePosition.Cells(1, platePosition.Columns.Count).End(xlToLeft).Column
    ovtCol = atdCol + 1
    
    ' 社員名の重複Dictionaryの初期化
    Set lastNameDict = CreateObject("Scripting.Dictionary")
    Set fullNameDict = CreateObject("Scripting.Dictionary")
    
    ' 配置シートの全図形を削除
    plateWs.Shapes.SelectAll
    On Error Resume Next
    Selection.Delete
    On Error GoTo 0
    
    ' 名字の重複回数をカウント
    For i = 2 To lastRow
        employeeCode = platePosition.Cells(i, 1).Value
        name = GetEmployeeName(dataWs, employeeCode)
        lastName = Split(name, " ")(0)
        firstName = Split(name, " ")(1)
        If lastNameDict.Exists(lastName) Then
            lastNameDict(lastName) = lastNameDict(lastName) + 1
        Else
            lastNameDict.Add lastName, 1
        End If
        fullNameDict(employeeCode) = name
    Next i
    
    ' 出勤プレートと残業プレートを再配置
    For i = 2 To lastRow
        employeeCode = platePosition.Cells(i, 1).Value
        name = fullNameDict(employeeCode)
        lastName = Split(name, " ")(0)
        firstName = Split(name, " ")(1)
        
        If lastNameDict(lastName) > 1 Then
            nameToDisplay = lastName & " " & Left(firstName, 1)
        Else
            nameToDisplay = lastName
        End If
        
        ' 出勤プレート
        positionData = platePosition.Cells(i, atdCol).Value
        If positionData <> "" Then
            leftPos = Split(positionData, ",")(0)
            topPos = Split(positionData, ",")(1)
            Call CreatePlate(plateWs, leftPos, topPos, "atd" & employeeCode, RGB(255, 0, 0), nameToDisplay)
        End If
        
        ' 残業プレート
        positionData = platePosition.Cells(i, ovtCol).Value
        If positionData <> "" Then
            leftPos = Split(positionData, ",")(0)
            topPos = Split(positionData, ",")(1)
            Call CreatePlate(plateWs, leftPos, topPos, "ovt" & employeeCode, RGB(0, 204, 255), nameToDisplay)
        End If
    Next i

    ' 辞書オブジェクトを解放　メモリーリークを防ぐ
    Set lastNameDict = Nothing
    Set fullNameDict = Nothing
End Sub

Sub CreatePlate(ws As Worksheet, x As Single, y As Single, shapeName As String, fillColor As Long, displayName As String)
    Dim shp As Shape
    Const SHAPE_WIDTH As Single = 57 ' 2cm
    Const SHAPE_HEIGHT As Single = 22.8 ' 0.8cm
    
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, SHAPE_WIDTH, SHAPE_HEIGHT)
    With shp
        .Name = shapeName
        .Fill.ForeColor.RGB = fillColor
        .TextFrame2.TextRange.Text = displayName
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Function GetEmployeeName(dataWs As Worksheet, employeeCode As String) As String
    Dim i As Long, lastRow As Long
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If dataWs.Cells(i, 1).Value = employeeCode Then
            GetEmployeeName = dataWs.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    GetEmployeeName = ""
End Function
