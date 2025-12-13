Sub BackupAttendanceData()
    ' 勤怠データをバックアップシートとCSVに保存
    Dim wsMain As Worksheet, wsBackup As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim backupRow As Long
    Dim backupPath As String, backupFileName As String
    Dim startTime As Double

    startTime = Timer
    On Error GoTo ErrorHandler

    ' シートの設定
    Set wsMain = ThisWorkbook.Sheets("本データ")

    ' バックアップシート作成
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets("バックアップ")
    On Error GoTo 0
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsMain)
        wsBackup.Name = "バックアップ"
        WriteLog "INFO", "バックアップシートを作成", "BackupAttendanceData"
    End If

    ' 最新のデータ範囲を取得
    lastRow = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column

    ' バックアップ先の行を取得
    backupRow = wsBackup.Cells(wsBackup.Rows.Count, 1).End(xlUp).Row + 1

    ' データをコピー
    wsMain.Range(wsMain.Cells(1, 1), wsMain.Cells(lastRow, lastCol)).Copy
    wsBackup.Cells(backupRow, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    WriteLog "INFO", "バックアップ成功", "BackupAttendanceData"

    ' CSVファイルに保存
    backupPath = ThisWorkbook.Path & "\backup\"
    If Dir(backupPath, vbDirectory) = "" Then MkDir backupPath

    backupFileName = backupPath & "backup_" & Format(Date, "yyyymmdd") & ".csv"
    SaveRangeAsCSV wsBackup, backupFileName
    WriteLog "INFO", "CSV バックアップ成功: " & backupFileName, "BackupAttendanceData"

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "バックアップ完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒", "BackupAttendanceData"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", Err.Description, "BackupAttendanceData", Erl
    Resume Next
End Sub