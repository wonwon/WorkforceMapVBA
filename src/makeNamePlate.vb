Sub makePlateNames()
    ' 社員名簿からデータを読み込み、出勤、残業配置プレートを作成するコード
    Dim employeeCode As String, name As String, lastName As String, firstName As String
    Dim shp As shape, blueShape As shape
    Dim i As Long, lastRow As Long, totalRows As Long, columnOffset As Single, rowOffset As Single
    Dim dataWs As Worksheet, plateWs As Worksheet
    Dim lastNameDict As Object
    Dim sheetName As String
    Dim totalHeight As Single

    Const MAX_COLUMNS As Integer = 10
    Const SHAPE_WIDTH As Single = 57 ' 2cm, Excelの既定では1cm = 28.35pt
    Const SHAPE_HEIGHT As Single = 22.8 ' 0.8cm
    Const GAP As Single = 5
    Const BLUE_OFFSET As Single = 10
    
    ' シートの設定
    ' 配置シートを生成する
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    sheetName = "配置"

    ' 配置シートの存在を確認、なければ作成
    On Error Resume Next
    Set plateWs = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If plateWs Is Nothing Then
        Set plateWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        plateWs.name = sheetName
    End If

    ' 社員データの最終行を取得
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row

    ' 社員名の重複Dictionaryの初期化
    Set lastNameDict = CreateObject("Scripting.Dictionary")
    
    ' 配置シートに図形がある場合すべて消す
    plateWs.Shapes.SelectAll
    Selection.Delete

    ' 名字の重複回数をカウント
    For i = 2 To lastRow
        lastName = Split(dataWs.Cells(i, 2).Value, " ")(0)
        If lastNameDict.Exists(lastName) Then
            lastNameDict(lastName) = lastNameDict(lastName) + 1
        Else
            lastNameDict.Add lastName, 1
        End If
    Next i

    ' 赤色のプレートを作成
    maxRedBottom = 0
    
    For i = 2 To lastRow
        employeeCode = dataWs.Cells(i, 1).Value
        name = dataWs.Cells(i, 2).Value
        lastName = Split(name, " ")(0)
        If lastNameDict(lastName) > 1 Then
            firstName = Split(name, " ")(1)
        Else
            firstName = ""
        End If
        
        columnOffset = ((i - 2) Mod MAX_COLUMNS) * (SHAPE_WIDTH + GAP)
        rowOffset = ((i - 2) \ MAX_COLUMNS) * (SHAPE_HEIGHT + GAP)
        
        ' プレート作成処理
        Call CreatePlate(plateWs, columnOffset, rowOffset, SHAPE_WIDTH, SHAPE_HEIGHT, lastName, firstName, "atd" & employeeCode, RGB(255, 0, 0))
        
                    
        ' 赤色のプレートの位置情報を更新
        If rowOffset + SHAPE_HEIGHT > maxRedBottom Then
            maxRedBottom = rowOffset + SHAPE_HEIGHT
        End If
    Next i
    
    ' 水色のプレートを作成
    For i = 2 To lastRow
        employeeCode = dataWs.Cells(i, 1).Value
        name = dataWs.Cells(i, 2).Value
        lastName = Split(name, " ")(0)
        If lastNameDict(lastName) > 1 Then
            firstName = Split(name, " ")(1)
        Else
            firstName = ""
        End If

        columnOffset = ((i - 2) Mod MAX_COLUMNS) * (SHAPE_WIDTH + GAP)
        rowOffset = maxRedBottom + BLUE_OFFSET + ((i - 2) \ MAX_COLUMNS) * (SHAPE_HEIGHT + GAP)
        
        ' プレート作成処理
        Call CreatePlate(plateWs, columnOffset, rowOffset, SHAPE_WIDTH, SHAPE_HEIGHT, lastName, firstName, "ovt" & employeeCode, RGB(0, 204, 255))
    Next i

    ' 辞書オブジェクトを解放　メモリーリークを防ぐ
    Set lastNameDict = Nothing
End Sub

Sub CreatePlate(ws As Worksheet, x As Single, y As Single, width As Single, height As Single, lastName As String, firstName As String, shapeName As String, fillColor As String)

   Dim shp As shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, width, height)
    With shp
        .name = shapeName
        .Fill.ForeColor.RGB = fillColor
        .TextFrame2.TextRange.Text = lastName & IIf(firstName <> "", Left(firstName, 1), "")
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        If firstName <> "" Then
            .TextFrame2.TextRange.Characters(Len(lastName) + 1, 1).Font.Size = 10
        End If
    End With
End Sub
