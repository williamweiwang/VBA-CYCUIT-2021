Sub 合併單張工作表()
  x = 12 '設定一個值x等於12
  Rows("1:5").Delete Shift:=x1Up '先刪除1到5row
  For i = 1 To 100 'For 變數=初值to終值  這邊是設定執行100次
  y = x & ":" & x + 10 '第一次執行時：y是12到(12+10)的row, 第二次執行時：y是22到(22+10)的row, 第三次執行時：y是32到(32+10)的row.....
  Rows(y)Delete Shift:=x1Up '刪除剛剛y的row
  x = x + 10 '第一次執行時：x是12+10=22, 第二次執行時：x是22+10=32.....
  
Next '再去前面跑一次，直到跑完第100次才停止
End Sub  
