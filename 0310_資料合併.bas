Sub 第一個() '巨集名稱

    Sheets.Add before:=Sheets(1) '於第一張前面新增工作表
    'Sheets("工作表1").Select ---這一行要刪掉!!
    Sheets(1).Name = "OK" '第一張工作表命名為OK
    Sheets("三月").Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("OK").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
End Sub
