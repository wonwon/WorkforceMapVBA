' 適切にオブジェクトを解放して、メモリリークを防ぐ
Dim wb As Workbook
Set wb = Workbooks.Open("C:\path\to\file.xlsx")
' 処理後
wb.Close SaveChanges:=False
Set wb = Nothing
  
    
'開いたファイルや使用したリソースを確実に閉じる
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Dim file As Object
Set file = fs.OpenTextFile("C:\path\to\file.txt", 1)
        
' 処理後
file.Close
Set file = Nothing
Set fs = Nothing

' 辞書オブジェクトを解放　メモリーリークを防ぐ
Set lastNameDict = Nothing
Set fullNameDict = Nothing
