Option Explicit

' 定数
Private Const SHEET_ATTENDANCE As String = "勤怠データ"
Private Const SHEET_EMPLOYEE As String = "社員データ"

Sub RecordAttendance(employeeCode As String, status As String)
    ' 勤怠データを記録する
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = GetSheet(SHEET_ATTENDANCE)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = employeeCode
    ws.Cells(lastRow, 2).Value = Date
    ws.Cells(lastRow, 3).Value = Format(Time, "HH:MM:SS")
    ws.Cells(lastRow, 4).Value = status
    
    Call log.WriteLog("INFO", "勤怠記録: " & employeeCode & " - " & status)
End Sub

Function GetEmployeeName(employeeCode As String) As String
    ' 社員コードから名前を取得
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = GetSheet(SHEET_EMPLOYEE)
    
    ' A列（社員コード）を検索
    Set cell = ws.Columns(1).Find(What:=employeeCode, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cell Is Nothing Then
        GetEmployeeName = cell.Offset(0, 1).Value ' B列が名前と仮定
    Else
        GetEmployeeName = ""
    End If
End Function

Sub UpdateOvertimeStatus(employeeCode As String, isAllowed As Boolean)
    ' 残業可否を更新
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = GetSheet(SHEET_EMPLOYEE)
    Set cell = ws.Columns(1).Find(What:=employeeCode, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cell Is Nothing Then
        If isAllowed Then
            cell.Offset(0, 2).Value = "◯" ' C列と仮定
        Else
            cell.Offset(0, 2).Value = ""
        End If
        Call log.WriteLog("INFO", "残業ステータス更新: " & employeeCode & " -> " & isAllowed)
    Else
        Call log.WriteLog("ERROR", "社員が見つかりません: " & employeeCode, "db", "UpdateOvertimeStatus")
    End If
End Sub

Private Function GetSheet(sheetName As String) As Worksheet
    ' シート取得（なければエラーログ）
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    If GetSheet Is Nothing Then
        Call log.WriteLog("ERROR", "シートが見つかりません: " & sheetName, "db", "GetSheet")
        ' 必要なら作成するロジックをここに追加
    End If

' === Backup & Data Management ===

Sub BackupData()
    ' データをバックアップ
    Dim wsMain As Worksheet, wsBackup As Worksheet
    Dim lastRow As Long, lastCol As Long, backupRow As Long
    Dim backupPath As String, backupFileName As String
    
    On Error Resume Next
    Set wsMain = ThisWorkbook.Sheets(SHEET_ATTENDANCE)
    Set wsBackup = ThisWorkbook.Sheets("バックアップ")
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add
        wsBackup.Name = "バックアップ"
    End If
    On Error GoTo 0
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
    lastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column
    backupRow = wsBackup.Cells(wsBackup.Rows.Count, 1).End(xlUp).Row + 1
    
    If lastRow > 1 Then
        wsMain.Range(wsMain.Cells(2, 1), wsMain.Cells(lastRow, lastCol)).Copy
        wsBackup.Cells(backupRow, 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    
    ' CSV出力
    backupPath = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "backup" & Application.PathSeparator
    If Dir(backupPath, vbDirectory) = "" Then MkDir backupPath
    
    backupFileName = backupPath & "backup_" & Format(Date, "yyyymmdd_HHMMSS") & ".csv"
    Call SaveRangeAsCSV(wsBackup, backupFileName)
    
    Call log.WriteLog("INFO", "バックアップ完了: " & backupFileName, "db", "BackupData")
End Sub

Private Sub SaveRangeAsCSV(ws As Worksheet, filePath As String)
    Dim fileNum As Integer
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim strBuf As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    For r = 1 To lastRow
        strBuf = ""
        For c = 1 To lastCol
            strBuf = strBuf & ws.Cells(r, c).Value & ","
        Next c
        If Len(strBuf) > 0 Then strBuf = Left(strBuf, Len(strBuf) - 1)
        Print #fileNum, strBuf
    Next r
    
    Close #fileNum
End Sub

Sub UploadDailyData()
    ' 日次データのアップロード（配列処理による高速化）
    Call log.StartTimer
    
    Dim wsEmp As Worksheet, wsMain As Worksheet
    Dim lastRowEmp As Long, lastColMain As Long
    Dim targetCol As Long
    Dim todayStr As String, empCode As String, overtime As Variant
    Dim arrEmp As Variant
    Dim arrMainHeader As Variant
    Dim arrMainKeys As Variant
    Dim r As Long, c As Long
    Dim outputData As Variant ' 書き込み用配列
    Dim empRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsEmp = ThisWorkbook.Sheets(SHEET_EMPLOYEE)
    Set wsMain = GetOrCreateSheet("本データ")
    
    todayStr = Format(Date, "yyyy/mm/dd")
    
    ' ヘッダー読み込み（配列）
    lastColMain = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column
    If lastColMain < 1 Then lastColMain = 1
    arrMainHeader = wsMain.Range(wsMain.Cells(1, 1), wsMain.Cells(1, lastColMain)).Value
    
    ' 今日の列を探す
    targetCol = 0
    For c = 1 To UBound(arrMainHeader, 2)
        If Format(arrMainHeader(1, c), "yyyy/mm/dd") = todayStr Then
            targetCol = c
            Exit For
        End If
    Next c
    
    If targetCol = 0 Then
        targetCol = lastColMain + 1
        wsMain.Cells(1, targetCol).Value = todayStr
    End If
    
    ' 社員データ読み込み（配列）
    lastRowEmp = wsEmp.Cells(wsEmp.Rows.Count, 1).End(xlUp).Row
    If lastRowEmp < 2 Then GoTo Cleanup
    arrEmp = wsEmp.Range("A2:D" & lastRowEmp).Value ' A: Code, D: Overtime
    
    ' 本データの社員コード一覧読み込み（更新用マッピング作成）
    Dim lastRowMain As Long
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
    If lastRowMain > 1 Then
        arrMainKeys = wsMain.Range("A2:A" & lastRowMain).Value
    End If
    
    ' ディクショナリで本データの行番号を管理
    Dim dictMain As Object
    Set dictMain = CreateObject("Scripting.Dictionary")
    If IsArray(arrMainKeys) Then
        For r = 1 To UBound(arrMainKeys, 1)
            If Not IsEmpty(arrMainKeys(r, 1)) Then
                dictMain(CStr(arrMainKeys(r, 1))) = r + 1 ' 行番号（+1はヘッダー分）
            End If
        Next r
    End If
    
    ' 書き込みデータの準備（都度Writeは遅いので、本データに行を追加するか既存行を更新するか判断）
    ' ここでは実装を簡単にするため、逐次書き込みを避けるが、
    ' 行が追加される可能性があるため、Rangeオブジェクトを操作する
    
    For r = 1 To UBound(arrEmp, 1)
        empCode = CStr(arrEmp(r, 1))
        overtime = arrEmp(r, 4)
        
        If empCode <> "" And overtime <> "" Then
            If dictMain.Exists(empCode) Then
                ' 既存行更新
                wsMain.Cells(dictMain(empCode), targetCol).Value = overtime
            Else
                ' 新規行追加
                lastRowMain = lastRowMain + 1
                wsMain.Cells(lastRowMain, 1).Value = empCode
                wsMain.Cells(lastRowMain, targetCol).Value = overtime
                dictMain(empCode) = lastRowMain
            End If
        End If
    Next r
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call log.WriteLog("INFO", "日次データアップロード完了 (" & log.StopTimer() & ")", "db", "UploadDailyData")
End Sub

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = sheetName
        If sheetName = "本データ" Then GetOrCreateSheet.Cells(1, 1).Value = "社員コード"
    End If
    On Error GoTo 0
End Function

Sub UpdateOvertimeHours(employeeCode As String, hours As Double)
    ' 残業時間を記録（単発処理だが高速化のため画面停止）
    Call log.StartTimer
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim cell As Range
    Set ws = GetSheet(SHEET_EMPLOYEE)
    Set cell = ws.Columns(1).Find(What:=employeeCode, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cell Is Nothing Then
        cell.Offset(0, 3).Value = hours ' D列と仮定
        Call log.WriteLog("INFO", "残業時間記録: " & employeeCode & " -> " & hours & "h (" & log.StopTimer() & ")", "db", "UpdateOvertimeHours")
    Else
        Call log.WriteLog("ERROR", "社員が見つかりません: " & employeeCode, "db", "UpdateOvertimeHours")
    End If
    
    Application.ScreenUpdating = True
End Sub

Function IsChantingDuty(employeeCode As String) As Boolean
    ' 唱和当番判定
    ' シート「唱和当番」のA列(日付)とB列(社員コード)をチェック
    ' 高速化のため配列処理
    
    Dim ws As Worksheet
    Dim arrData As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim todayStr As Variant ' Date型比較または文字列比較
    Dim targetDate As Date
    
    IsChantingDuty = False
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("唱和当番")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function ' シートがなければ判定なし
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    arrData = ws.Range("A2:B" & lastRow).Value
    targetDate = Date
    
    If IsArray(arrData) Then
        For i = 1 To UBound(arrData, 1)
            If IsDate(arrData(i, 1)) Then
                If arrData(i, 1) = targetDate And CStr(arrData(i, 2)) = employeeCode Then
                    IsChantingDuty = True
                    Exit For
                End If
            End If
        Next i
    End If
End Function

