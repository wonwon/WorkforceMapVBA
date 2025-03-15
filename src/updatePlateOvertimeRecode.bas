Option Explicit

Sub UpdatePlateAndOvertimeRecorde()
    ' 2Dコードリーダーで社員コードを読み込み、プレート色変更
    ' 社員データの残業時間に入力する（ログ機能付き）
    
    Dim employeeCode As String
    Dim plateWs As Worksheet, dataWs As Worksheet
    Dim shp As Shape
    Dim plateName As String
    Dim foundPlate As Boolean
    Dim lastRow As Long, i As Long
    Dim startTime As Double

    ' パフォーマンス計測開始
    startTime = Timer
    On Error GoTo ErrorHandler

    ' シートの設定
    Set plateWs = ThisWorkbook.Sheets("配置")
    Set dataWs = ThisWorkbook.Sheets("社員データ")

    ' 2Dコードリーダーで読み込んだ社員コード（仮: 入力ボックスで受け取る）
    employeeCode = InputBox("社員コードを入力してください:", "2Dコード入力")

    ' 入力がない場合は処理を終了
    If employeeCode = "" Then
        WriteLog "WARNING", "社員コードが入力されませんでした"
        Exit Sub
    End If

    ' 配置シートから該当するプレートを探す
    foundPlate = False
    For Each shp In plateWs.Shapes
        If shp.Name = "atd" & employeeCode Then
            plateName = shp.Name
            foundPlate = True
            Exit For
        End If
    Next shp

    ' 該当するプレートが見つかった場合、色を赤色に変更する
    If foundPlate Then
        plateWs.Shapes(plateName).Fill.ForeColor.RGB = RGB(255, 0, 0) ' 赤色に変更
        WriteLog "INFO", "プレート色変更: " & plateName & " → 赤色"
    Else
        MsgBox "該当するプレートが見つかりませんでした。", vbExclamation
        WriteLog "WARNING", "プレートが見つかりませんでした: " & employeeCode
        Exit Sub
    End If

    ' 社員データシートから該当する社員の行を探す
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If dataWs.Cells(i, 1).Value = employeeCode Then
            ' 残業時間を入力
            dataWs.Cells(i, 4).Value = 2 ' 残業時間を仮入力（後で変更可能）
            WriteLog "INFO", "残業時間を入力: " & employeeCode & " → 2時間"
            Exit For
        End If
    Next i

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "プレート更新完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "エラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
