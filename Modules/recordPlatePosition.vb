Sub RecordPlatePositions()
  ' 配置シートに配置されている図形の位置情報を配置記録シートに記録するコード
  ' 配置記録シートの１列目は、社員コードがあることが条件
  ' 出勤プレートは、出勤_記録日、残業プレートは、残業_記録日の列名を追加する
  ' 記録する時に、過去の記録をクリアしてから、新たに記録する
    Dim platePosition As Worksheet, plateWs As Worksheet
    Dim shp As shape
    Dim i As Long, lastRow As Long
    Dim platePositionColumn As Long
    Dim currentDate As String
    Dim employeeCode As String
    
    ' シートの設定
    Set platePosition = ThisWorkbook.Sheets("配置記録")
    Set plateWs = ThisWorkbook.Sheets("配置")
    
    ' 最終行を取得
    lastRow = platePosition.Cells(platePosition.Rows.Count, "A").End(xlUp).Row
    
    ' 過去の記録を消去
    If lastRow > 1 Then
        platePosition.Columns("B:" & Split(platePosition.Cells(1, platePosition.Columns.Count).Address, "$")(1)).Delete
    End If
    
    ' 現在の日付を取得してカラム名に設定
    currentDate = Format(Date, "YYYYMMDD")
    platePositionColumn = platePosition.Cells(1, platePosition.Columns.Count).End(xlToLeft).Column + 1
    platePosition.Cells(1, platePositionColumn).Value = "出勤_" & currentDate
    platePosition.Cells(1, platePositionColumn + 1).Value = "残業_" & currentDate
    
    ' 配置シートの全図形の位置情報を配置記録シートに記録
    For Each shp In plateWs.Shapes
        ' 図形の名前から出勤プレートか残業プレートかを判断し、それぞれの列に位置情報を記録
        If InStr(shp.name, "atd") > 0 Then ' 出勤プレートの場合
            employeeCode = Mid(shp.name, InStr(shp.name, "atd") + 3)
            For i = 2 To lastRow
                If platePosition.Cells(i, 1).Value = employeeCode Then
                    platePosition.Cells(i, platePositionColumn).Value = shp.Left & "," & shp.Top
                    Exit For
                End If
            Next i
        ElseIf InStr(shp.name, "ovt") > 0 Then ' 残業プレートの場合
            employeeCode = Mid(shp.name, InStr(shp.name, "ovt") + 3)
            For i = 2 To lastRow
                If platePosition.Cells(i, 1).Value = employeeCode Then
                    platePosition.Cells(i, platePositionColumn + 1).Value = shp.Left & "," & shp.Top
                    Exit For
                End If
            Next i
        End If
    Next shp
End Sub

