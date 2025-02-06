Sub UpdatePlateAndOvertimeRecorde()
    ' 2Dコードリーダーで社員コードを読み込み、プレート色変更
    ' 社員データの残業時間に入力する
    Dim employeeCode As String
    Dim plateWs As Worksheet, dataWs As Worksheet
    Dim shp As Shape
    Dim plateName As String
    Dim foundPlate As Boolean
    Dim lastRow As Long, i As Long
    
    ' シートの設定
    Set plateWs = ThisWorkbook.Sheets("配置")
    Set dataWs = ThisWorkbook.Sheets("社員データ")
    
    ' 2Dコードリーダーで読み込んだ社員コード
    employeeCode = ' ここに2Dコードリーダーで読み込んだ社員コードを設定
    
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
    Else
        MsgBox "該当するプレートが見つかりませんでした。"
    End If
    
    ' 社員データシートから該当する社員の行を探す
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If dataWs.Cells(i, 1).Value = employeeCode Then
            ' 残業時間を入力
            dataWs.Cells(i, 4).Value = 2 ' 残業時間を入力（仮入力2時間）
            Exit For
        End If
    Next i
End Sub
