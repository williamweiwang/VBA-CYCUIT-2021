Sub 分類()

mydata = InputBox("請輸入分類欄號, 例:A,B,C,D", "10842284陳氏鸞")
my_col = Range(mydata & "1").Column

'-----找唯一-----
    Columns(my_col).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Columns( _
        my_col), CopyToRange:=Range("G1"), Unique:=True
        
        
        Range("G2").Select
        
        While ActiveCell.Value <> Empty '當被選擇的欄位不為空白
            new_sht_name = ActiveCell.Value '先抓被選擇欄位的名字
            Sheets.Add after:=Sheets(1) '在表單後創建新表單
            ActiveSheet.Name = new_sht_name '新表單命名(由前面抓的東西)
            Sheets(1).Select '回到最前面表單
            ActiveCell.Offset(1, 0).Select '往下走一格
            
        Wend
            Selection.EntireColumn.Delete '刪除多餘欄位
   
'----------資料移轉----------

For i = 2 To Sheets.Count '到最後一張

    Range("A1").Select 'CTRL+HOME
    Selection.AutoFilter '點資料-篩選
    ActiveSheet.Range("$A$1:$F$312").AutoFilter Field:=my_col, Criteria1:=Sheets(i).Name '依工作表名稱篩選
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select '選範圍 CTRL+SHIFT+END
    Selection.Copy '複製範圍
    Sheets(i).Select '選取特定表單
    ActiveSheet.Paste '貼上
    Application.CutCopyMode = False '退出剪貼簿
    Sheets(1).Select '選取第一張表單
    Range("A1").Select 'CTRL+HOME
    Selection.AutoFilter '再點一次資料-篩選
    
    Next
    
End Sub
