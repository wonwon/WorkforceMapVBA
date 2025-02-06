Private Sub CommandButton1_Click()
    ’　ボタンを押すとウインドウが開くコード、おかしな挙動を無くすコード
    ' フォームが既に表示されているかを確認
    If Not UserForm1.Visible Then
        ' フォームが表示されていない場合は表示
        UserForm1.Show
    Else
        ' フォームが表示されている場合はアクティブにする
        UserForm1.Activate
    End If
End Sub