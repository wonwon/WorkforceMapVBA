Sub SaveRangeAsCSV(ws As Worksheet, filePath As String)
    Dim fso As Object, stream As Object
    Dim row As Long, col As Long, lastRow As Long, lastCol As Long
    Dim line As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(filePath, True)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For row = 1 To lastRow
        line = ""
        For col = 1 To lastCol
            line = line & ws.Cells(row, col).Value & ","
        Next col
        line = Left(line, Len(line) - 1)
        stream.WriteLine line
    Next row

    stream.Close
    Set fso = Nothing
    Set stream = Nothing
End Sub
