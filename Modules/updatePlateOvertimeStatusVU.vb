Sub UpdatePlateAndOvertimeStatus()
    ' 出勤時に2Dコードリーダーで読み込んだ社員コードの社員のプレートを赤色から黄色へ変更するコード
    ' 残業可否の◯のついた社員は、プレートの枠を太く青色に変更する
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
    ' ここに2Dコードリーダーで読み込んだ社員コードを設定
    employeeCode = InputBox("社員コードを入力してください:")
    
    ' 配置シートから該当するプレートを探す
    foundPlate = False
    For Each shp In plateWs.Shapes
        If shp.Name = "atd" & employeeCode Then
            plateName = shp.Name
            foundPlate = True
            Exit For
        End If
    Next shp
    
    ' 該当するプレートが見つかった場合、色を黄色に変更する
    If foundPlate Then
        With plateWs.Shapes(plateName)
            .Fill.ForeColor.RGB = RGB(255, 255, 0) ' 黄色に変更
        End With
    Else
        MsgBox "該当するプレートが見つかりませんでした。"
    End If
    
    ' 社員データシートから該当する社員の行を探す
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If dataWs.Cells(i, 1).Value = employeeCode Then
            ' 残業可否の部分に◯印を入力する
            dataWs.Cells(i, 3).Value = "◯"
            
            ' 残業可否が◯の社員のプレートの枠を太い青色に変更する
            If dataWs.Cells(i, 3).Value = "◯" Then
                With plateWs.Shapes(plateName)
                    .Line.ForeColor.RGB = RGB(0, 0, 255) ' 青色に変更
                    .Line.Weight = 2.25 ' 枠を太くする
                End With
            End If
            Exit For
        End If
    Next i
End Sub
