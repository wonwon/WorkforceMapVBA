Option Explicit

Sub RearrangeOVTPlates()
    ' 残業時間プレートを任意の位置に再配列する機能（ログ機能付き）
    
    Dim plateWs As Worksheet
    Dim shp As Shape
    Dim ovtPlates As Collection
    Dim newTop As Single, newLeft As Single
    Dim i As Integer
    Dim startTime As Double

    ' パフォーマンス計測開始
    startTime = Timer
    On Error GoTo ErrorHandler

    ' 配置シートの設定
    Set plateWs = ThisWorkbook.Sheets("配置")
    
    ' ovtプレートのコレクションを作成
    Set ovtPlates = New Collection
    For Each shp In plateWs.Shapes
        If InStr(shp.Name, "ovt") > 0 Then
            ovtPlates.Add shp
        End If
    Next shp

    ' ログ: 残業プレートの数を記録
    WriteLog "INFO", "再配置対象の残業プレート数: " & ovtPlates.Count

    ' 新しい配置位置を設定
    newTop = 400
    newLeft = 300

    ' ovt残業時間プレートを新しい位置に移動して再配置
    For i = 1 To ovtPlates.Count
        Set shp = ovtPlates(i)
        shp.Top = newTop + Int((i - 1) / 10) * shp.Height ' 上からの位置
        shp.Left = newLeft + ((i - 1) Mod 10) * shp.Width ' 右にの位置
        
        ' 各プレートの移動後の座標をログに記録
        WriteLog "INFO", "残業プレート再配置: " & shp.Name & " -> Top: " & shp.Top & ", Left: " & shp.Left
    Next i

    ' パフォーマンスログ
    WriteLog "PERFORMANCE", "残業プレート再配置完了 - 処理時間: " & Format(Timer - startTime, "0.00") & " 秒"

    Exit Sub

ErrorHandler:
    WriteLog "ERROR", "エラー発生: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub
