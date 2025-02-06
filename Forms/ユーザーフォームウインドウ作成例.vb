Private Sub UserForm_Activate()
    ' ユーザーフォームに追加することで
    '　ユーザーに処理中にメッセージ表示できるコード
    '　ユーザーフォームにラベル（Label）を追加し、名前をlblStatusに変更
  
    ' 処理中のメッセージを表示
    lblStatus.Caption = "データを処理しています。お待ちください..."
    
    ' メイン処理を呼び出し
    Call MainProcessing

    ' 処理完了後のメッセージを表示
    lblStatus.Caption = "処理が完了しました！"
    
    ' 3秒後にフォームを閉じる
    Application.OnTime Now + TimeValue("00:00:03"), "CloseStatusForm"
End Sub

Sub ShowStatusForm()
    ' ユーザーフォームを表示するコード
    Load UserForm1
    UserForm1.Show
End Sub

Sub CloseStatusForm()
    ' ユーザーフォームを閉じる
    Unload UserForm1
End Sub

Private Sub CommandButton1_Click()
    ' ステータスフォームを表示するボタンのクリックイベントコード
    Call ShowStatusForm
End Sub
