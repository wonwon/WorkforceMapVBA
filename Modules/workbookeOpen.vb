Private Sub Workbook_Open()
    ' 毎日19:00にCopyOvertimeDataサブルーチンを実行するコード
    Application.OnTime TimeValue("19:00:00"), "CopyOvertimeData"
End Sub