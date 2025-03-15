Private Sub Workbook_Open()
    ' Workbook を開いたときに実行される処理（ログ機能付き）
    
    On Error Resume Next ' エラーが発生しても処理を継続
    WriteLog "INFO", "Workbook_Open が実行されました"

    ' タスクスケジュールを設定
    Application.OnTime TimeValue("19:00:00"), "CopyOvertimeData"
    WriteLog "INFO", "CopyOvertimeData のスケジュール設定: 19:00"

    ' バックアップ処理を設定
    Application.OnTime TimeValue("23:59:00"), "BackupAttendanceData"
    WriteLog "INFO", "BackupAttendanceData のスケジュール設定: 23:59"

    On Error GoTo 0 ' エラー処理を元に戻す
End Sub
