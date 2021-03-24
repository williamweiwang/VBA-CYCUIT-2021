Sub 匯入文字檔()
op_f = Array("lisa", "pisa", "visa", "4T") '陣列宣告
For i = 0 To 3

    Workbooks.OpenText Filename:="C:\Users\User\Desktop\20210324\合併\" & op_f(i) & ".txt", _
        Origin:=950, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1)), TrailingMinusNumbers:=True '開文字檔案進行合併
        
    Sheets(1).Copy After:=Workbooks("4合併外部檔案.xlsm").Sheets(1)
    Windows(op_f(i) & ".txt").Close
    
    Next
End Sub
