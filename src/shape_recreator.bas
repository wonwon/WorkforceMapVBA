Sub RecreateNamedShapesFromLog()
    Dim ws As Worksheet, logWs As Worksheet
    Dim i As Long, lastRow As Long
    Dim shp As Shape
    Dim nameText As String

    Set logWs = ThisWorkbook.Sheets("Shape記録")

    ' 新しいシート準備
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("再配置")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "再配置"
    Else
        ws.Cells.Clear
        ws.Shapes.SelectAll
        Selection.Delete
    End If
    On Error GoTo 0

    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        nameText = logWs.Cells(i, 1).Value
        If nameText <> "" Then
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
                logWs.Cells(i, 2).Value, _
                logWs.Cells(i, 3).Value, _
                logWs.Cells(i, 4).Value, _
                logWs.Cells(i, 5).Value)
            With shp
                .Name = nameText
                .TextFrame2.TextRange.Text = nameText
                .TextFrame2.TextRange.Font.Size = 12
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        End If
    Next i

    MsgBox "再配置完了：テキストを名前に設定したシェイプを生成しました。", vbInformation
End Sub