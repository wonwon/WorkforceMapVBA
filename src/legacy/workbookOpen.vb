Private Sub Workbook_Open()
    ' 起動時のスケジュール設定
    On Error GoTo ErrorHandler

    WriteLog "INFO", "Workbook_Open が実行されました", "Workbook_Open"

    Application.OnTime TimeValue("19:00:00"), "CopyOvertimeData"
    WriteLog "INFO", "CopyOvertimeData のスケジュール設定: 19:00", "Workbook_Open"

    Application.OnTime TimeValue("23:59:00"), "BackupAttendanceData"
    WriteLog "INFO", "BackupAttendanceData のスケジュール設定: 23:59", "Workbook_Open"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", Err.Description, "Workbook_Open", Erl
    Resume Next
End Sub