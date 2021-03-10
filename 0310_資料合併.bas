Sub 第一個() '巨集名稱

    Sheets.Add before:=Sheets(1) '於第一張前面新增工作表
    'Sheets("工作表1").Select ---這一行要刪掉!!
    Sheets(1).Name = "OK" '第一張工作表命名為OK
    Sheets(2).Select '選第二張工作表
    Range("A1").Select '選擇A1儲存格子 CTRL + HOME
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select '選擇那一堆資料 CTRL + Shift + end
    Selection.Copy '複製那一堆資料
    Sheets("OK").Select '選擇名字叫OK的工作表
    ActiveSheet.Paste '貼上在OK工作表
    Application.CutCopyMode = False '清空剪貼簿
    Selection.End(xlDown).Select '到資料最尾端 CTRL + 方向下
End Sub
