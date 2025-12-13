Option Explicit

Sub Auto_Open()
    ' ブックを開いたときの処理
    Call log.WriteLog("INFO", "システム起動", "main", "Auto_Open")
    
    ' 前回の起動日を確認し、日付が変わっていればデータをアップロードしてバックアップ
    Dim lastDate As String
    lastDate = GetSetting("WorkforceMap", "System", "LastLaunchDate", "")
    
    If lastDate <> Format(Date, "YYYYMMDD") Then
        Call log.WriteLog("INFO", "日付変更を検知: " & lastDate & " -> " & Format(Date, "YYYYMMDD"), "main", "Auto_Open")
        Call db.BackupData
        Call db.UploadDailyData
        SaveSetting "WorkforceMap", "System", "LastLaunchDate", Format(Date, "YYYYMMDD")
    End If
End Sub


' === V1.1 New Flows ===

Sub InputCheckIn()
    ' 出勤フォーム（シミュレーション）
    Dim code As String
    code = InputBox("【出勤】社員コードを入力/スキャンしてください", "出勤打刻")
    If code = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim name As String
    name = db.GetEmployeeName(code)
    If name = "" Then
        MsgBox "社員コードが登録されていません", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 出勤処理
    Call CheckIn(code)
    
    ' 唱和当番チェック
    If db.IsChantingDuty(code) Then
        Call ui.ToggleDutyMark(code, True)
        Application.ScreenUpdating = True
        MsgBox "本日は【唱和当番】です。" & vbCrLf & "朝礼で唱和をお願いします。", vbInformation, "当番通知"
    Else
        Application.ScreenUpdating = True
        MsgBox name & " さん、おはようございます。（出勤）", vbInformation
    End If
    
End Sub

Sub InputCheckOut()
    ' 退勤フォーム（シミュレーション）
    Dim code As String
    Dim overtimeHours As String
    
    code = InputBox("【退勤】社員コードを入力/スキャンしてください", "退勤打刻")
    If code = "" Then Exit Sub
    
    ' 残業時間入力
    ' 簡易的にInputBoxで "0", "0.5", "1.0" 等を入力させる
    ' 本来はフォームでボタン選択だが、ここではロジックのみ実装
    overtimeHours = InputBox("本日の残業時間（h）を入力してください" & vbCrLf & "（例: 0, 0.5, 1.0, 1.5 ...）", "残業時間", "0")
    If overtimeHours = "" Then Exit Sub ' キャンセル
    
    Application.ScreenUpdating = False
    
    Dim name As String
    name = db.GetEmployeeName(code)
    
    ' 退勤処理
    Call CheckOut(code)
    
    ' 残業記録
    If IsNumeric(overtimeHours) Then
        Dim val As Double
        val = CDbl(overtimeHours)
        If val > 0 Then
            Call db.UpdateOvertimeHours(code, val)
            ' 残業がある場合は自動的に残業可（青枠）にする運用も考えられるが
            ' V1.1要件では「残業可フラグは従来通り」とのことなので、ここでは記録のみ
            ' ただし、残業時間>0なら青枠をつける方が自然かもしれない。
            ' ここでは記録のみとする。
        End If
    End If
    
    Application.ScreenUpdating = True
    MsgBox name & " さん、お疲れ様でした。（退勤）", vbInformation
End Sub

' Deprecated: V1.0 Toggle Flow
' Sub InputBarcode() ... End Sub 
' Sub ProcessScan() ... End Sub


Sub CheckIn(employeeCode As String)
    Call db.RecordAttendance(employeeCode, "出勤")
    Call ui.UpdateBadgeColor(employeeCode, "CHECKIN")
End Sub

Sub CheckOut(employeeCode As String)
    Call db.RecordAttendance(employeeCode, "退勤")
    Call ui.UpdateBadgeColor(employeeCode, "CHECKOUT")
    ' 退勤時に残業枠もクリアする場合
    ' Call ui.UpdateOvertimeBorder(employeeCode, False)
End Sub

Sub ToggleOvertime()
    ' 残業申請（選択中の社員、または入力）
    Dim employeeCode As String
    employeeCode = InputBox("残業申請する社員コードを入力してください", "残業申請")
    If employeeCode = "" Then Exit Sub
    
    ' 現在の状態を確認してトグル（ここでは常に許可にする例）
    Dim response As VbMsgBoxResult
    response = MsgBox("残業を許可しますか？（「いいえ」で解除）", vbYesNoCancel + vbQuestion, "残業許可")
    
    If response = vbCancel Then Exit Sub
    
    If response = vbYes Then
        Call db.UpdateOvertimeStatus(employeeCode, True)
        Call ui.UpdateOvertimeBorder(employeeCode, True)
    Else
        Call db.UpdateOvertimeStatus(employeeCode, False)
        Call ui.UpdateOvertimeBorder(employeeCode, False)
    End If
End Sub

Sub InputOvertimeHours()
    ' 残業時間の手入力
    Dim employeeCode As String
    Dim hours As String
    
    employeeCode = InputBox("残業時間を記録する社員コードを入力してください", "残業時間入力")
    If employeeCode = "" Then Exit Sub
    
    hours = InputBox("残業時間（h）を入力してください", "残業時間入力")
    If IsNumeric(hours) Then
        Call db.UpdateOvertimeHours(employeeCode, CDbl(hours))
        MsgBox "記録しました。", vbInformation
    Else
        MsgBox "数値で入力してください。", vbExclamation
    End If
End Sub

Sub InitializeSystem()
    ' システムの初期セットアップ（名札生成など）
    If MsgBox("名札を初期化・再生成しますか？配置はリセットされます。", vbYesNo + vbQuestion) = vbYes Then
        Call ui.InitializeBadges
    End If
End Sub

Sub SaveCurrentArrangement()
    Call ui.SavePositions
    MsgBox "現在の配置を保存しました。", vbInformation
End Sub

Sub LoadArrangement()
    Call ui.RestorePositions
    MsgBox "配置を復元しました。", vbInformation
End Sub
