Sub AllShapesDelete()
    ' 指定シートのすべてのシェイプを削除する
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' 対象シートを指定（必要に応じて "シート名" を変更）
    Set ws = ActiveSheet ' または Worksheets("シート名")

    ' エラー時のスキップ（存在しないシェイプがある場合）
    On Error Resume Next
    
    ' すべてのシェイプを削除
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' エラーハンドリングを解除
    On Error GoTo 0
    
    ' オブジェクト変数の解放
    Set ws = Nothing
    Set shp = Nothing
End Sub
