Sub 合併分析()
 For i = 2 To Sheets.Count


    Sheets(i).Select
    If i = 2 Then
     Range("A1").Select
    Else
     Range("A2").Select
    End If
    
    
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets(1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Offset(1, 0).Select
 Next
End Sub
