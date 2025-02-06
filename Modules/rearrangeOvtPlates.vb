Sub RearrangeOVTPlates()
    ' 残業時間プレートを任意の位置に再配列する機能
    ' 再配列ボタンをコントローラーに加える
    Dim plateWs As Worksheet
    Dim shp As Shape
    Dim ovtPlates As Collection
    Dim newTop As Single, newLeft As Single
    Dim i As Integer, j As Integer
    
    ' 配置シートの設定
    Set plateWs = ThisWorkbook.Sheets("配置")
    
    ' ovtプレートのコレクションを作成
    Set ovtPlates = New Collection
    For Each shp In plateWs.Shapes
        If InStr(shp.Name, "ovt") > 0 Then
            ovtPlates.Add shp
        End If
    Next shp
    
    ' 新しい配置位置を設定
    newTop = 400
    newLeft = 300
    
    ' ovt残業時間プレートを新しい位置に移動して再配置
    For i = 1 To ovtPlates.Count
        Set shp = ovtPlates(i)
        shp.Top = newTop + Int((i - 1) / 10) * shp.Height ' 上からの位置
        shp.Left = newLeft + ((i - 1) Mod 10) * shp.Width ' 右にの位置
    Next i
End Sub