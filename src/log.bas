' === Performance Timer ===
Private startTime As Double

Sub StartTimer()
    startTime = Timer
End Sub

Function StopTimer() As String
    Dim elapsed As Double
    elapsed = Timer - startTime
    StopTimer = Format(elapsed, "0.000") & " sec"
End Function

Sub WriteLog(logType As String, message As String, Optional moduleName As String = "", Optional lineNumber As String = "")
    ' ログを記録する（エラー発生時のモジュール名・行番号も記録）
    ' テキストファイルのみに記録する（Excelシートへの書き込みは遅いため廃止推奨だが、要件により残す場合は配列化が必要。今回は高速化優先でファイルのみにするか、バッファリングする）
    ' ここではファイル書き込みをメインとし、シート書き込みは削除して高速化する（またはオプションにする）
    
    Dim logMessage As String
    Dim logTime As String
    Dim logFileName As String
    Dim fullLogPath As String
    Dim fileNum As Integer
    
    logTime = Format(Now, "yyyy/mm/dd HH:mm:ss")
    
    ' メッセージの整形
    If logType = "ERROR" And moduleName <> "" And lineNumber <> "" Then
        logMessage = "[" & logTime & "] [" & logType & "] " & message & " (Module: " & moduleName & ", Line: " & lineNumber & ")"
    Else
        logMessage = "[" & logTime & "] [" & logType & "] " & message
    End If
    
    ' 1. テキストファイルへの出力
    logFileName = "log_" & Format(Date, "yyyymmdd") & ".txt"
    fullLogPath = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "logs" & Application.PathSeparator & logFileName
    
    ' フォルダ作成（存在確認のオーバーヘッドを減らすためエラーハンドリングで対応）
    On Error Resume Next
    fileNum = FreeFile
    Open fullLogPath For Append As #fileNum
    If Err.Number <> 0 Then
        ' フォルダがない場合のみ作成試行
        MkDir ThisWorkbook.Path & Application.PathSeparator & "data"
        MkDir ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "logs"
        Err.Clear
        Open fullLogPath For Append As #fileNum
    End If
    On Error GoTo 0
    
    Print #fileNum, logMessage
    Close #fileNum
    
End Sub


Sub WriteErrorLog(message As String, moduleName As String, lineNumber As String)
    ' エラーログ専用（error_log.txtにも記録）
    Call WriteLog("ERROR", message, moduleName, lineNumber)
    
    Dim fileNum As Integer
    Dim fullLogPath As String
    
    fullLogPath = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "logs" & Application.PathSeparator & "error_log.txt"
    
    fileNum = FreeFile
    Open fullLogPath For Append As #fileNum
    Print #fileNum, "[" & Format(Now, "yyyy/mm/dd HH:mm:ss") & "] [ERROR] " & message & " (" & moduleName & ":" & lineNumber & ")"
    Close #fileNum
End Sub