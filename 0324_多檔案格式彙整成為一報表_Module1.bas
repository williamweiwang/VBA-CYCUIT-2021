Sub 匯入試算表()
op_f = Array(1, 2, 3, 4) '陣列宣告
  For i = 0 To 3 ' 執行三次

    Workbooks.Open Filename:="C:\Users\User\Desktop\20210324\合併\" & op_f(i) & ".xlsx" '打開檔案  & op_f(i) & 為連接字串
    Sheets(1).Copy After:=Workbooks("4合併外部檔案.xlsm").Sheets(1) '複製工作表
    Windows(op_f(i) & ".xlsx").Close '回到原畫面並關閉
  
    Next
End Sub
