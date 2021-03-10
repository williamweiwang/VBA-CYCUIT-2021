Sub 第一個() '巨集名稱

    Sheets.Add before:=Sheets(1) '於第一張前面新增工作表
    'Sheets("工作表1").Select ---這一行要刪掉!!
    Sheets(1).Name = "OK" '第一張工作表命名為OK
    For i = 2 To 5 '迴圈設計
    'Sheets(2)選第二張工作表 這行為了迴圈改成i
    Sheets(i).Select '選第i張工作表
        If i = 2 Then '避開標題重複
        Range("A1").Select '選擇A1儲存格子 CTRL + HOME
            Else
                Range("A2").Select
                End If
            Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select '選擇那一堆資料 CTRL + Shift + end
                Selection.Copy '複製那一堆資料
                Sheets(1).Select '選擇第一張的工作表
                    ActiveSheet.Paste '貼上在第一張工作表
                    Application.CutCopyMode = False '清空剪貼簿
                        Selection.End(xlDown).offset(1,0).Select '到資料最尾端 CTRL + 方向下
                    'offset(1,0): 移至最後1個row的下1row
            Next

End Sub
            
