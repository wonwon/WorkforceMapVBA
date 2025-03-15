Option Explicit

Sub WriteLog(logType As String, message As String)
    ' ログを記録する共通関数（フォルダ構成に対応）
    Dim logFile As Object
    Dim logFolder As String, logFilePath As String
    Dim fs As Object

    ' ログフォルダのパス
    logFolder = ThisWorkbook.Path & "\data\logs\"
    
    ' フォルダが存在しない場合は作成
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(logFolder) Then
        fs.CreateFolder logFolder
    End If
    
    ' ログファイル（本日の日付ごとに新規作成）
    logFilePath = logFolder & "log_" & Format(Date, "yyyymmdd") & ".txt"
    
    ' ログを追記モードで開く
    Set logFile = fs.OpenTextFile(logFilePath, 8, True)
    logFile.WriteLine Format(Now, "yyyy/mm/dd HH:MM:SS") & " [" & logType & "] " & message
    logFile.Close
    
    ' エラーログ（特定のエラーのみ保存）
    If logType = "ERROR" Then
        logFilePath = logFolder & "error_log.txt"
        Set logFile = fs.OpenTextFile(logFilePath, 8, True)
        logFile.WriteLine Format(Now, "yyyy/mm/dd HH:MM:SS") & " [ERROR] " & message
        logFile.Close
    End If
    
    ' オブジェクト解放
    Set logFile = Nothing
    Set fs = Nothing
End Sub