Sub allShapeDelete()
    ' 配置シートに削除ボタン（コントローラー）を作成し、すべての図形を消す機能
    Dim shp As shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next shp
End Sub
