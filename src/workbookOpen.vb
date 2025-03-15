Private Sub Workbook_Open()
    ' 毎日19:00にCopyOvertimeDataサブルーチンを実行するコード（ログ機能付き）
    
    On Error Resume Next ' エラーが発生しても処理を継続
    WriteLog "INFO", "Workbook_Open が実行されました"

    ' タスクスケジュールを設定
    Application.OnTime TimeValue("19:00:00"), "CopyOvertimeData"

    ' 設定完了ログ
    WriteLog "INFO", "CopyOvertimeData のスケジュール設定: 19:00"
    
    On Error GoTo 0 ' エラー処理を元に戻す
End Sub
