Sub WriteLog(logType As String, message As String, Optional moduleName As String = "", Optional lineNumber As String = "")
    ' ログを記録する（エラー発生時のモジュール名・行番号も記録）

    Dim logWs As Worksheet
    Dim lastRow As Long
    Dim logMessage As String
    Dim logTime As String

    ' シートの設定（ログ専用シート）
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("ログ")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add
        logWs.Name = "ログ"
        logWs.Cells(1, 1).Value = "日時"
        logWs.Cells(1, 2).Value = "タイプ"
        logWs.Cells(1, 3).Value = "モジュール"
        logWs.Cells(1, 4).Value = "行番号"
        logWs.Cells(1, 5).Value = "メッセージ"
    End If
    On Error GoTo 0

    ' ログ記録
    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    logTime = Format(Now, "yyyy/mm/dd HH:MM:SS")

    ' エラー時にモジュール名と行番号を記録
    If logType = "ERROR" And moduleName <> "" And lineNumber <> "" Then
        logMessage = "[エラー] " & message & " (モジュール: " & moduleName & ", 行: " & lineNumber & ")"
    Else
        logMessage = message
    End If

    logWs.Cells(lastRow, 1).Value = logTime
    logWs.Cells(lastRow, 2).Value = logType
    logWs.Cells(lastRow, 3).Value = moduleName
    logWs.Cells(lastRow, 4).Value = lineNumber
    logWs.Cells(lastRow, 5).Value = logMessage
End Sub