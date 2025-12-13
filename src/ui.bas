Option Explicit

' シート名定数
Private Const SHEET_PLATE As String = "配置"
Private Const SHEET_DATA As String = "社員データ"
Private Const SHEET_POSITION As String = "配置記録"

' 図形定数
Private Const SHAPE_WIDTH As Single = 57
Private Const SHAPE_HEIGHT As Single = 22.8
Private Const GAP As Single = 5
Private Const BLUE_OFFSET As Single = 10
Private Const MAX_COLUMNS As Integer = 10

' 色定数
Private Const COLOR_RED As Long = 255          ' RGB(255, 0, 0)
Private Const COLOR_YELLOW As Long = 65535     ' RGB(255, 255, 0)
Private Const COLOR_BLUE_LIGHT As Long = 16764108 ' RGB(0, 204, 255)
Private Const COLOR_BLUE_DARK As Long = 16711680  ' RGB(0, 0, 255)

Sub InitializeBadges()
    ' 社員データから名札を初期生成（再生成） - 高速化
    Call log.StartTimer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim dataWs As Worksheet, plateWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim lastNameDict As Object
    Dim employeeCode As String, name As String, lastName As String, firstName As String
    Dim columnOffset As Single, rowOffset As Single
    Dim maxRedBottom As Single
    Dim arrData As Variant
    
    Set dataWs = ThisWorkbook.Sheets(SHEET_DATA)
    Set plateWs = GetOrCreateSheet(SHEET_PLATE)
    
    ' 既存図形削除
    plateWs.Shapes.SelectAll
    Selection.Delete
    
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then
        arrData = dataWs.Range("A2:B" & lastRow).Value ' A: Code, B: Name
    End If
    
    Set lastNameDict = CreateObject("Scripting.Dictionary")
    
    ' 重複チェック（配列処理）
    If IsArray(arrData) Then
        For i = 1 To UBound(arrData, 1)
            name = arrData(i, 2)
            lastName = Split(name, " ")(0)
            If lastNameDict.Exists(lastName) Then
                lastNameDict(lastName) = lastNameDict(lastName) + 1
            Else
                lastNameDict.Add lastName, 1
            End If
        Next i
    
        ' 赤色（退勤/未出勤）プレート作成
        maxRedBottom = 0
        For i = 1 To UBound(arrData, 1)
            employeeCode = arrData(i, 1)
            name = arrData(i, 2)
            lastName = Split(name, " ")(0)
            If lastNameDict(lastName) > 1 Then
                firstName = Split(name, " ")(1)
            Else
                firstName = ""
            End If
            
            columnOffset = ((i - 1) Mod MAX_COLUMNS) * (SHAPE_WIDTH + GAP)
            rowOffset = ((i - 1) \ MAX_COLUMNS) * (SHAPE_HEIGHT + GAP)
            
            Call CreatePlate(plateWs, columnOffset, rowOffset, "atd" & employeeCode, COLOR_RED, lastName, firstName)
            
            If rowOffset + SHAPE_HEIGHT > maxRedBottom Then maxRedBottom = rowOffset + SHAPE_HEIGHT
        Next i
        
        ' 青色（残業用）プレート作成
        For i = 1 To UBound(arrData, 1)
            employeeCode = arrData(i, 1)
            name = arrData(i, 2)
            lastName = Split(name, " ")(0)
            If lastNameDict(lastName) > 1 Then
                firstName = Split(name, " ")(1)
            Else
                firstName = ""
            End If
            
            columnOffset = ((i - 1) Mod MAX_COLUMNS) * (SHAPE_WIDTH + GAP)
            rowOffset = maxRedBottom + BLUE_OFFSET + ((i - 1) \ MAX_COLUMNS) * (SHAPE_HEIGHT + GAP)
            
            Call CreatePlate(plateWs, columnOffset, rowOffset, "ovt" & employeeCode, COLOR_BLUE_LIGHT, lastName, firstName)
        Next i
    End If
    
    Set lastNameDict = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call log.WriteLog("INFO", "名札初期化完了 (" & log.StopTimer() & ")", "ui", "InitializeBadges")
End Sub

Private Sub CreatePlate(ws As Worksheet, x As Single, y As Single, shapeName As String, fillColor As Long, lastName As String, Optional firstName As String = "")
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, SHAPE_WIDTH, SHAPE_HEIGHT)
    With shp
        .Name = shapeName
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

Sub UpdateBadgeColor(employeeCode As String, status As String)
    ' ステータスに応じて名札の色を変更
    ' status: "CHECKIN" -> Yellow, "CHECKOUT" -> Red
    Dim plateWs As Worksheet
    Dim shp As Shape
    Dim targetName As String
    
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    targetName = "atd" & employeeCode
    
    On Error Resume Next
    Set shp = plateWs.Shapes(targetName)
    On Error GoTo 0
    
    If shp Is Nothing Then
        Call log.WriteLog("WARNING", "名札が見つかりません: " & targetName, "ui", "UpdateBadgeColor")
        Exit Sub
    End If
    
    If status = "CHECKIN" Then
        shp.Fill.ForeColor.RGB = COLOR_YELLOW
    ElseIf status = "CHECKOUT" Then
        shp.Fill.ForeColor.RGB = COLOR_RED
    End If
    
    Call log.WriteLog("INFO", "名札色変更: " & employeeCode & " -> " & status)
End Sub

Sub UpdateOvertimeBorder(employeeCode As String, isAllowed As Boolean)
    ' 残業可否に応じて枠線を変更
    Dim plateWs As Worksheet
    Dim shp As Shape
    Dim targetName As String
    
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    targetName = "atd" & employeeCode ' 残業表示は通常名札（出勤名札）に行う仕様のようです
    
    On Error Resume Next
    Set shp = plateWs.Shapes(targetName)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    If isAllowed Then
        shp.Line.ForeColor.RGB = COLOR_BLUE_DARK
        shp.Line.Weight = 2.25
    Else
        shp.Line.Visible = msoFalse ' または標準に戻す
        ' 標準に戻す場合は以下
        ' shp.Line.ForeColor.RGB = RGB(0,0,0)
        ' shp.Line.Weight = 0.75
    End If
End Sub

Sub SavePositions()
    ' 全プレートの位置を保存（配列処理による高速化）
    Call log.StartTimer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim posWs As Worksheet, plateWs As Worksheet
    Dim shp As Shape
    Dim lastRow As Long, targetCol As Long
    Dim employeeCode As String
    Dim dateStr As String
    Dim dictPositions As Object
    Dim keys As Variant
    Dim arrEmpCodes As Variant
    Dim arrOut As Variant
    Dim i As Long
    
    Set posWs = ThisWorkbook.Sheets(SHEET_POSITION)
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    Set dictPositions = CreateObject("Scripting.Dictionary")
    
    dateStr = Format(Date, "YYYYMMDD")
    targetCol = posWs.Cells(1, posWs.Columns.Count).End(xlToLeft).Column + 1
    posWs.Cells(1, targetCol).Value = "出勤_" & dateStr
    posWs.Cells(1, targetCol + 1).Value = "残業_" & dateStr
    
    ' 図形情報をメモリに収集（オブジェクト操作自体は重いが、書き込みをまとめる）
    For Each shp In plateWs.Shapes
        If InStr(shp.Name, "atd") > 0 Then
            dictPositions("atd_" & Mid(shp.Name, 4)) = shp.Left & "," & shp.Top
        ElseIf InStr(shp.Name, "ovt") > 0 Then
            dictPositions("ovt_" & Mid(shp.Name, 4)) = shp.Left & "," & shp.Top
        End If
    Next shp
    
    ' 既存の社員コードリストを読み込み
    lastRow = posWs.Cells(posWs.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        arrEmpCodes = posWs.Range("A2:A" & lastRow).Value
        ' 書き込み用配列を準備
        ReDim arrOut(1 To UBound(arrEmpCodes, 1), 1 To 2) ' 1: atd, 2: ovt
        
        For i = 1 To UBound(arrEmpCodes, 1)
            employeeCode = CStr(arrEmpCodes(i, 1))
            If dictPositions.Exists("atd_" & employeeCode) Then
                arrOut(i, 1) = dictPositions("atd_" & employeeCode)
            End If
            If dictPositions.Exists("ovt_" & employeeCode) Then
                arrOut(i, 2) = dictPositions("ovt_" & employeeCode)
            End If
        Next i
        
        ' 一括書き込み
        posWs.Cells(2, targetCol).Resize(UBound(arrOut, 1), 2).Value = arrOut
    Else
        Call log.WriteLog("WARNING", "配置記録時に社員コードが見つかりません", "ui", "SavePositions")
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call log.WriteLog("INFO", "位置情報保存完了 (" & log.StopTimer() & ")", "ui", "SavePositions")
End Sub

Private Sub WritePosToSheet(ws As Worksheet, empCode As String, col As Long, x As Single, y As Single)
    ' 廃止（SavePositionsで一括処理するため）
End Sub

Sub RestorePositions()
    ' 最新の位置情報を復元
    Dim posWs As Worksheet, plateWs As Worksheet
    Dim shp As Shape
    Dim lastRow As Long, targetCol As Long, i As Long
    Dim dateStr As String
    Dim empCode As String
    Dim posData As String
    Dim xy() As String
    Dim atdCol As Long, ovtCol As Long
    
    Set posWs = ThisWorkbook.Sheets(SHEET_POSITION)
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    lastRow = posWs.Cells(posWs.Rows.Count, 1).End(xlUp).Row
    
    ' 最新の日付の列を探す（簡易的に最終列付近を見る）
    ' 実運用では日付で検索すべきだが、ここでは最終列と仮定
    atdCol = posWs.Cells(1, posWs.Columns.Count).End(xlToLeft).Column - 1
    If atdCol < 2 Then atdCol = 2 ' 最低限の列
    ovtCol = atdCol + 1
    
    Call log.WriteLog("INFO", "位置復元開始: Col " & atdCol, "ui", "RestorePositions")
    
    For i = 2 To lastRow
        empCode = posWs.Cells(i, 1).Value
        
        ' 出勤プレート
        posData = posWs.Cells(i, atdCol).Value
        If posData <> "" And InStr(posData, ",") > 0 Then
            xy = Split(posData, ",")
            On Error Resume Next
            Set shp = plateWs.Shapes("atd" & empCode)
            On Error GoTo 0
            If Not shp Is Nothing Then
                shp.Left = Val(xy(0))
                shp.Top = Val(xy(1))
            End If
        End If
        
        ' 残業プレート
        posData = posWs.Cells(i, ovtCol).Value
        If posData <> "" And InStr(posData, ",") > 0 Then
            xy = Split(posData, ",")
            On Error Resume Next
            Set shp = plateWs.Shapes("ovt" & empCode)
            On Error GoTo 0
            If Not shp Is Nothing Then
                shp.Left = Val(xy(0))
                shp.Top = Val(xy(1))
            End If
        End If
    Next i
    
    Call log.WriteLog("INFO", "位置情報を復元しました", "ui", "RestorePositions")
End Sub

Sub ToggleDutyMark(employeeCode As String, isVisible As Boolean)
    ' 唱和当番マーク（オレンジ色の円）を表示/非表示
    Dim plateWs As Worksheet
    Dim shpBadge As Shape
    Dim shpMark As Shape
    Dim markName As String
    Dim badgeName As String
    
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    badgeName = "atd" & employeeCode
    markName = "duty_" & employeeCode
    
    ' 親となる名札を探す
    On Error Resume Next
    Set shpBadge = plateWs.Shapes(badgeName)
    Set shpMark = plateWs.Shapes(markName)
    On Error GoTo 0
    
    If shpBadge Is Nothing Then Exit Sub ' 名札がなければ何もしない
    
    If isVisible Then
        ' 表示（なければ作成、あれば表示）
        If shpMark Is Nothing Then
            ' 名札の左横に作成 (10pt x 10pt)
            Set shpMark = plateWs.Shapes.AddShape(msoShapeOval, shpBadge.Left - 12, shpBadge.Top + (shpBadge.Height - 10) / 2, 10, 10)
            shpMark.Name = markName
            shpMark.Fill.ForeColor.RGB = RGB(255, 165, 0) ' Orange
            shpMark.Line.Visible = msoFalse
        Else
            shpMark.Visible = msoTrue
            ' 位置補正（名札が動いている可能性があるため）
            shpMark.Left = shpBadge.Left - 12
            shpMark.Top = shpBadge.Top + (shpBadge.Height - 10) / 2
        End If
    Else
        ' 非表示（削除または非表示）
        If Not shpMark Is Nothing Then
            shpMark.Delete ' すっきりさせるため削除
        End If
    End If
End Sub

Sub ResetAllDutyMarks()
    ' 全ての当番マークを削除（日付変更時など）
    Dim plateWs As Worksheet
    Dim shp As Shape
    
    Set plateWs = ThisWorkbook.Sheets(SHEET_PLATE)
    For Each shp In plateWs.Shapes
        If Left(shp.Name, 5) = "duty_" Then
            shp.Delete
        End If
    Next shp
End Sub

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
    On Error GoTo 0
End Function
