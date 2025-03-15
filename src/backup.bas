Option Explicit

Sub BackupAttendanceData()
    ' 勤怠データをバックアップシートとCSVに保存する
    Dim wsMain As Worksheet, wsBackup As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim backupRow As Long
    Dim backupPath As String
    Dim backupFileName As String
    Dim startTime As Double

    ' パフォーマンス計測開始
    startTime = Timer
    On Error GoTo ErrorHandler

    ' シートの設定
    Set wsMain = ThisWorkbook.Sheets("本データ")

    ' バックアップシートの存在確認、なければ作成
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets("バックアップ")
    On Error GoTo 0
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsMain)
        wsBackup.Name = "バックアップ"
        WriteLog "INFO", "バックアップシートを作成"
    End If

    ' 最新のデータ範囲を取得
    lastRow = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column

    ' バックアップ先の行を取得
    backupRow = wsBackup.Cells(wsBackup.Rows.Count, "A").End(xlUp).Row + 1

    ' データをコピー
    wsMain.Range(wsMain.Cells(1, 1), wsMain.Cells(lastRow, lastCol)).Copy
    wsBackup.Cells(backupRow, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    WriteLog "INFO", "勤怠データをバックアップシートにコピー"

    ' CSVファイルに保存
    backupPath = ThisWorkbook.Path & "\backup\"
    If Dir(backupPath, vbDirectory) = "" Then MkDir backupPath

    backupFileName = backupPath & "backup_" & Format(Date, "yyyymmdd") & ".csv"
    SaveRangeAsCSV wsBackup, backupFileName
    WriteLog "INFO", "バックアップデータをCSVに保存: " & backupFileName

    ' 30日以上前のデータを削除
    DeleteOldBackups backupPath

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "データバックアップ完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "バックアップ処理でエラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
