Attribute VB_Name = "Module1"

Sub �ʺA��i�u�@��()
Attribute �ʺA��i�u�@��.VB_ProcData.VB_Invoke_Func = " \n14"
        Range("A1").FormulaR1C1 = "=COUNTIF(C[4],R[1]C[4])"
    myc = Range("A1").Value
    
    Rows("1:5").Delete Shift:=xlUp '�R��1��5�C

For i = 1 To myc
    Cells.Find(What:="�U", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
        x = ActiveCell.Row ' "�U"���@���x�s�檺row
        y = x & ":" & x + 10 '12:22, 22:32, 32:42
    Rows(y).Delete Shift:=xlUp
Next
Range("A1").Select
End Sub
