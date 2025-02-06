    ' 外部ブックのパスとシート名を設定
    dataWbPath = "C:\path\to\your\folder\社員データ.xlsx"
    dataWsName = "社員データ"
    
    ' 外部ブックを開く
    Set dataWb = Workbooks.Open(dataWbPath)
    Set dataWs = dataWb.Sheets(dataWsName)
    
    ' 外部ブックを保存して閉じる
    dataWb.Close SaveChanges:=True