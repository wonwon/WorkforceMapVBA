Sub CopyOvertimeData()
     ' 社員データシートに入力された残業時間を本データシートにコピーするコード
    Dim dataWs As Worksheet, mainDataWs As Worksheet
    Dim employeeCode As String, employeeName As String, overtimeHours As String
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim todayDate As String
    Dim dateCol As Long
    
    ' シートの設定
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    Set mainDataWs = ThisWorkbook.Sheets("本データ")
    
    ' 今日の日付を取得
    todayDate = Format(Date, "yyyy/mm/dd")
    
    ' 本データシートで今日の日付の列を探す
    lastCol = mainDataWs.Cells(1, mainDataWs.Columns.Count).End(xlToLeft).Column
    For j = 1 To lastCol
        If mainDataWs.Cells(1, j).Value = todayDate Then
            dateCol = j
            Exit For
        End If
    Next j
    
    ' 今日の日付の列が見つからない場合、メッセージを表示して終了
    If dateCol = 0 Then
        MsgBox "今日の日付の列が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 社員データシートの最終行を取得
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    
    ' 残業時間を本データシートにコピー
    For i = 2 To lastRow
        employeeCode = dataWs.Cells(i, 1).Value
        employeeName = dataWs.Cells(i, 2).Value
        overtimeHours = dataWs.Cells(i, 3).Value
        
        ' 本データシートで社員名の行を探す
        lastRow = mainDataWs.Cells(mainDataWs.Rows.Count, 1).End(xlUp).Row
        For j = 2 To lastRow
            If mainDataWs.Cells(j, 1).Value = employeeName Then
                mainDataWs.Cells(j, dateCol).Value = overtimeHours
                Exit For
            End If
        Next j
    Next i
End Sub