Option Explicit

Sub RecordPlatePositions()
    ' 配置シートに配置されている図形の位置情報を配置記録シートに記録するコード（ログ機能付き）
    ' 出勤プレートは、出勤_記録日、残業プレートは、残業_記録日の列名を追加
    ' 記録する時に、過去の記録をクリアしてから、新たに記録する

    Dim platePosition As Worksheet, plateWs As Worksheet
    Dim shp As Shape
    Dim i As Long, lastRow As Long
    Dim platePositionColumn As Long
    Dim currentDate As String
    Dim employeeCode As String
    Dim startTime As Double

    ' パフォーマンス計測開始
    startTime = Timer
    On Error GoTo ErrorHandler

    ' シートの設定
    Set platePosition = ThisWorkbook.Sheets("配置記録")
    Set plateWs = ThisWorkbook.Sheets("配置")

    ' 最終行を取得
    lastRow = platePosition.Cells(platePosition.Rows.Count, "A").End(xlUp).Row

    ' 既存の記録を削除
    If lastRow > 1 Then
        platePosition.Range("B:Z").ClearContents
        WriteLog "INFO", "過去の位置記録をクリア"
    End If

    ' 現在の日付を取得してカラム名に設定
    currentDate = Format(Date, "YYYYMMDD")
    platePositionColumn = platePosition.Cells(1, platePosition.Columns.Count).End(xlToLeft).Column + 1
    platePosition.Cells(1, platePositionColumn).Value = "出勤_" & currentDate
    platePosition.Cells(1, platePositionColumn + 1).Value = "残業_" & currentDate

    WriteLog "INFO", "位置記録のカラム名を設定: 出勤_" & currentDate & " / 残業_" & currentDate

    ' 配置シートの全図形の位置情報を配置記録シートに記録
    For Each shp In plateWs.Shapes
        If InStr(shp.Name, "atd") > 0 Then ' 出勤プレートの場合
            employeeCode = Mid(shp.Name, InStr(shp.Name, "atd") + 3)
            For i = 2 To lastRow
                If platePosition.Cells(i, 1).Value = employeeCode Then
                    platePosition.Cells(i, platePositionColumn).Value = shp.Left & "," & shp.Top
                    WriteLog "INFO", "出勤プレート位置記録: " & employeeCode & " - " & shp.Left & "," & shp.Top
                    Exit For
                End If
            Next i
        ElseIf InStr(shp.Name, "ovt") > 0 Then ' 残業プレートの場合
            employeeCode = Mid(shp.Name, InStr(shp.Name, "ovt") + 3)
            For i = 2 To lastRow
                If platePosition.Cells(i, 1).Value = employeeCode Then
                    platePosition.Cells(i, platePositionColumn + 1).Value = shp.Left & "," & shp.Top
                    WriteLog "INFO", "残業プレート位置記録: " & employeeCode & " - " & shp.Left & "," & shp.Top
                    Exit For
                End If
            Next i
        End If
    Next shp

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "プレート位置記録完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "エラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
