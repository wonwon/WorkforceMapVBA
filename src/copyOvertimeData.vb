Option Explicit

Sub CopyOvertimeData()
    ' 社員データシートに入力された残業時間を本データシートにコピーするコード
    Dim dataWs As Worksheet, mainDataWs As Worksheet
    Dim employeeCode As String, employeeName As String, overtimeHours As String
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim todayDate As String
    Dim dateCol As Range
    Dim findEmployee As Range
    
    ' シートの設定
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    Set mainDataWs = ThisWorkbook.Sheets("本データ")
    
    ' 画面更新をオフ（処理速度向上）
    Application.ScreenUpdating = False
    
    ' 今日の日付を取得（シートとフォーマットを統一）
    todayDate = Format(Date, "yyyy/mm/dd")
    
    ' 本データシートで今日の日付の列を検索（Findを使用）
    Set dateCol = mainDataWs.Rows(1).Find(what:=todayDate, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' 今日の日付の列が見つからない場合、メッセージを表示して終了
    If dateCol Is Nothing Then
        MsgBox "今日の日付の列が見つかりません。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 社員データシートの最終行を取得
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    
    ' 残業時間を本データシートにコピー
    For i = 2 To lastRow
        employeeName = dataWs.Cells(i, 2).Value
        overtimeHours = dataWs.Cells(i, 3).Value
        
        ' 空白行をスキップ
        If employeeName = "" Or overtimeHours = "" Then GoTo NextIteration
        
        ' 本データシートで社員名の行を検索（Findを使用）
        Set findEmployee = mainDataWs.Range("A:A").Find(what:=employeeName, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' 社員名が見つかった場合、残業時間をコピー
        If Not findEmployee Is Nothing Then
            mainDataWs.Cells(findEmployee.Row, dateCol.Column).Value = overtimeHours
        End If
        
NextIteration:
    Next i
    
    ' 画面更新をオン
    Application.ScreenUpdating = True

    ' オブジェクト変数の解放
    Set dataWs = Nothing
    Set mainDataWs = Nothing
    Set dateCol = Nothing
    Set findEmployee = Nothing
End Sub