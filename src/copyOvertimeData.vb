Option Explicit

Sub CopyOvertimeData()
    ' 社員データシートに入力された残業時間を本データシートにコピーするコード（メモリ最適化版）
    
    Dim dataWs As Worksheet, mainDataWs As Worksheet
    Dim lastRowData As Long, lastRowMain As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim todayDate As String
    Dim dateCol As Long
    Dim employeeList As Variant, overtimeList As Variant
    Dim mainEmployeeList As Variant

    ' シートの設定
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    Set mainDataWs = ThisWorkbook.Sheets("本データ")
    
    ' 画面更新をオフ（処理速度向上）
    Application.ScreenUpdating = False
    
    ' 今日の日付を取得（シートのフォーマットに統一）
    todayDate = Format(Date, "yyyy/mm/dd")
    
    ' 本データシートで今日の日付の列を検索（ループ処理）
    lastCol = mainDataWs.Cells(1, mainDataWs.Columns.Count).End(xlToLeft).Column
    dateCol = 0
    
    For j = 1 To lastCol
        If mainDataWs.Cells(1, j).Value = todayDate Then
            dateCol = j
            Exit For
        End If
    Next j

    ' 今日の日付の列が見つからない場合、メッセージを表示して終了
    If dateCol = 0 Then
        MsgBox "今日の日付の列が見つかりません。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 社員データシートの最終行を取得
    lastRowData = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    
    ' 本データシートの最終行を取得
    lastRowMain = mainDataWs.Cells(mainDataWs.Rows.Count, "A").End(xlUp).Row
    
    ' 社員データと残業時間を配列に格納（メモリ最適化）
    employeeList = dataWs.Range("B2:B" & lastRowData).Value
    overtimeList = dataWs.Range("C2:C" & lastRowData).Value
    
    ' 本データシートの社員リストを配列に格納（検索用）
    mainEmployeeList = mainDataWs.Range("A2:A" & lastRowMain).Value

    ' 残業時間を本データシートにコピー
    For i = 1 To UBound(employeeList, 1) ' 配列は1-based
        If employeeList(i, 1) <> "" And overtimeList(i, 1) <> "" Then
            ' 本データシートで社員名を検索（ループだが、配列なのでメモリ負担を軽減）
            For j = 1 To UBound(mainEmployeeList, 1)
                If mainEmployeeList(j, 1) = employeeList(i, 1) Then
                    ' 本データシートの該当セルに残業時間を記入
                    mainDataWs.Cells(j + 1, dateCol).Value = overtimeList(i, 1)
                    Exit For ' 早期終了でループ回数を削減
                End If
            Next j
        End If
    Next i
    
    ' 画面更新をオン
    Application.ScreenUpdating = True

    ' オブジェクト変数の解放（メモリ節約）
    Set dataWs = Nothing
    Set mainDataWs = Nothing
End Sub