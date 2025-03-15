Option Explicit

Sub RestorePlatePositions()
    ' プレート位置情報を読み込んで、各プレートを再配置する（ログ機能付き）
    Dim platePosition As Worksheet, plateWs As Worksheet, dataWs As Worksheet
    Dim shp As Shape
    Dim i As Long, lastRow As Long
    Dim atdCol As Long, ovtCol As Long
    Dim currentDate As String
    Dim lastNameDict As Object
    Dim fullNameDict As Object
    Dim employeeCode As String
    Dim positionData As Variant
    Dim leftPos As Single, topPos As Single
    Dim name As String, lastName As String, firstName As String
    Dim nameToDisplay As String
    Dim startTime As Double

    ' パフォーマンス計測開始
    startTime = Timer
    On Error GoTo ErrorHandler

    ' シートの設定
    Set platePosition = ThisWorkbook.Sheets("配置記録")
    Set plateWs = ThisWorkbook.Sheets("配置")
    Set dataWs = ThisWorkbook.Sheets("社員データ")

    ' 最終行を取得
    lastRow = platePosition.Cells(platePosition.Rows.Count, "A").End(xlUp).Row
    WriteLog "INFO", "配置記録シートの最終行: " & lastRow

    ' 現在の日付から列を特定
    currentDate = Format(Date, "YYYYMMDD")
    atdCol = platePosition.Cells(1, platePosition.Columns.Count).End(xlToLeft).Column
    ovtCol = atdCol + 1
    WriteLog "INFO", "対象列 - 出勤: " & atdCol & " / 残業: " & ovtCol

    ' 社員名の重複Dictionaryの初期化
    Set lastNameDict = CreateObject("Scripting.Dictionary")
    Set fullNameDict = CreateObject("Scripting.Dictionary")

    ' 配置シートの全図形を削除
    plateWs.Cells.Clear
    WriteLog "INFO", "配置シートの既存プレートを削除"

    ' 名字の重複回数をカウント
    For i = 2 To lastRow
        employeeCode = platePosition.Cells(i, 1).Value
        name = GetEmployeeName(dataWs, employeeCode)
        
        ' 名前が取得できない場合の処理
        If name = "" Then
            WriteLog "WARNING", "社員コード " & employeeCode & " に対応する名前が見つかりません"
            GoTo NextIteration
        End If
        
        ' 名前を適切に分割
        If InStr(name, " ") > 0 Then
            lastName = Split(name, " ")(0)
            firstName = Split(name, " ")(1)
        Else
            lastName = name
            firstName = ""
        End If

        ' 辞書へ追加
        If lastNameDict.Exists(lastName) Then
            lastNameDict(lastName) = lastNameDict(lastName) + 1
        Else
            lastNameDict.Add lastName, 1
        End If
        fullNameDict(employeeCode) = name

NextIteration:
    Next i

    ' 出勤プレートと残業プレートを再配置
    For i = 2 To lastRow
        employeeCode = platePosition.Cells(i, 1).Value
        name = fullNameDict(employeeCode)

        If name = "" Then GoTo NextLoop

        If lastNameDict(Split(name, " ")(0)) > 1 Then
            nameToDisplay = Split(name, " ")(0) & " " & Left(Split(name, " ")(1), 1)
        Else
            nameToDisplay = Split(name, " ")(0)
        End If

        ' 出勤プレート
        positionData = platePosition.Cells(i, atdCol).Value
        If positionData <> "" Then
            leftPos = Val(Split(positionData, ",")(0))
            topPos = Val(Split(positionData, ",")(1))
            Call CreatePlate(plateWs, leftPos, topPos, "atd" & employeeCode, RGB(255, 0, 0), nameToDisplay)
            WriteLog "INFO", "出勤プレート復元: " & employeeCode & " - (" & leftPos & "," & topPos & ")"
        End If

        ' 残業プレート
        positionData = platePosition.Cells(i, ovtCol).Value
        If positionData <> "" Then
            leftPos = Val(Split(positionData, ",")(0))
            topPos = Val(Split(positionData, ",")(1))
            Call CreatePlate(plateWs, leftPos, topPos, "ovt" & employeeCode, RGB(0, 204, 255), nameToDisplay)
            WriteLog "INFO", "残業プレート復元: " & employeeCode & " - (" & leftPos & "," & topPos & ")"
        End If

NextLoop:
    Next i

    ' 辞書オブジェクトを解放
    Set lastNameDict = Nothing
    Set fullNameDict = Nothing

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "プレート再配置完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "エラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
