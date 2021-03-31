Sub 分類()
'-----找唯一-----

    Columns("C:C").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Columns( _
        "C:C"), CopyToRange:=Range("G1"), Unique:=True
        
        
        Range("G2").Select
        
        While ActiveCell.Value <> Empty '當被選擇的欄位不為空白
            new_sht_name = ActiveCell.Value '先抓被選擇欄位的名字
            Sheets.Add after:=Sheets(1) '在表單後創建新表單
            ActiveSheet.Name = new_sht_name '新表單命名(由前面抓的東西)
            Sheets(1).Select '回到最前面表單
            ActiveCell.Offset(1, 0).Select '往下走一格
            
        Wend
            Selection.EntireColumn.Delete '刪除多餘欄位
            
End Sub
