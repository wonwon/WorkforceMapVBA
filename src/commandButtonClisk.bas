Private Sub CommandButton1_Click()
    ' エラーハンドリングを有効化（フォームがない場合の対策）
    On Error Resume Next
    
    ' フォームの存在チェック
    Dim frm As Object
    Set frm = UserForm1
    
    ' エラー処理解除
    On Error GoTo 0
    
    ' フォームが存在する場合
    If Not frm Is Nothing Then
        ' フォームが表示されているかチェック
        If Not UserForm1.Visible Then
            ' フォームをモデルレスで表示（他の作業を妨げない）
            UserForm1.Show vbModeless
        Else
            ' フォームを最前面に持ってくる
            UserForm1.ZOrder 0
        End If
    Else
        ' フォームが存在しない場合は新しくロードして表示
        Load UserForm1
        UserForm1.Show vbModeless
    End If
    
    ' オブジェクト変数の解放
    Set frm = Nothing
End Sub