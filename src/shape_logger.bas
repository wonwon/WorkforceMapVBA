Sub RecordUnnamedTextShapes()
    Dim ws As Worksheet, logWs As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim tempName As String

    Set ws = ThisWorkbook.Sheets("配置")

    ' ログ用シートの準備
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("Shape記録")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
        logWs.Name = "Shape記録"
    Else
        logWs.Cells.Clear
    End If
    On Error GoTo 0

    ' ヘッダー
    logWs.Range("A1:F1").Value = Array("ShapeName", "Left", "Top", "Width", "Height", "Text")

    i = 2

    For Each shp In ws.Shapes
        If shp.Name Like "Rectangle*" Then ' デフォルト名のものだけ対象
            If shp.Type = msoAutoShape And shp.TextFrame2.HasText Then
                tempName = Trim(shp.TextFrame2.TextRange.Text)
                If tempName <> "" Then
                    With logWs
                        .Cells(i, 1).Value = tempName ' オブジェクト名用に使用
                        .Cells(i, 2).Value = shp.Left
                        .Cells(i, 3).Value = shp.Top
                        .Cells(i, 4).Value = shp.Width
                        .Cells(i, 5).Value = shp.Height
                        .Cells(i, 6).Value = tempName
                    End With
                    i = i + 1
                End If
            End If
        End If
    Next shp

    MsgBox "名前なし＋テキスト付きシェイプの記録が完了しました。", vbInformation
End Sub