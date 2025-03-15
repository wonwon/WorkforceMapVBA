Option Explicit

Sub UpdatePlateAndOvertimeStatus()
    ' 出勤時に2Dコードリーダーで読み込んだ社員コードの社員のプレートを赤色から黄色へ変更する
    ' 残業可否の◯のついた社員は、プレートの枠を太く青色に変更する（ログ機能付き）
    
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

    ' 該当するプレートが見つかった場合、色を黄色に変更する
    If foundPlate Then
        With plateWs.Shapes(plateName)
            .Fill.ForeColor.RGB = RGB(255, 255, 0) ' 黄色に変更
        End With
        WriteLog "INFO", "プレート色変更: " & plateName & " → 黄色"
    Else
        MsgBox "該当するプレートが見つかりませんでした。", vbExclamation
        WriteLog "WARNING", "プレートが見つかりませんでした: " & employeeCode
        Exit Sub
    End If

    ' 社員データシートから該当する社員の行を探す
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If dataWs.Cells(i, 1).Value = employeeCode Then
            ' 残業可否の部分に◯印を入力する
            dataWs.Cells(i, 3).Value = "◯"
            WriteLog "INFO", "社員データ更新: " & employeeCode & " → 残業可"

            ' 残業可否が◯の社員のプレートの枠を太い青色に変更する
            If dataWs.Cells(i, 3).Value = "◯" Then
                With plateWs.Shapes(plateName)
                    .Line.ForeColor.RGB = RGB(0, 0, 255) ' 青色に変更
                    .Line.Weight = 2.25 ' 枠を太くする
                End With
                WriteLog "INFO", "プレート枠変更: " & plateName & " → 青色 (残業可)"
            End If
            Exit For
        End If
    Next i

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "プレート状態更新完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "エラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
