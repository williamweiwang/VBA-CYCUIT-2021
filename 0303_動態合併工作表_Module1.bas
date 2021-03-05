Sub 動態單張工作表() '巨集名稱
        Attribute 動態單張工作表.VB_ProcData.VB_Invoke_Func = " \n14" 
        Range("A1").FormulaR1C1 = "=COUNTIF(C[4],R[1]C[4])" '計算次數
        myc = Range("A1").Value  
    
    Rows("1:5").Delete Shift:=xlUp '刪除1到5列

        For i = 1 To myc 
        Cells.Find(What:="下", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _  
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
                , MatchByte:=False, SearchFormat:=False).Activate 
        x = ActiveCell.Row ' "下"的作用儲存格的row 
        y = x & ":" & x + 10 '12到22, 22到32, 32到42 
        Rows(y).Delete Shift:=xlUp 
Next 
Range("A1").Select 
End Sub 
