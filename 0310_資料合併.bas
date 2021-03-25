Sub 第一個() '巨集名稱
    pwd = Inputbox("請輸入密碼", "資料合併系統_陳小彎") '由使用者提供資訊
    if pwd = "10842284" Then '判定是否符合要求 如果密碼 = 10842284 就會執行下面
        
    Sheets.Add before:=Sheets(1) '於第一張前面新增工作表
    'Sheets("工作表1").Select ---這一行要刪掉!!
    Sheets(1).Name = "OK" '第一張工作表命名為OK
    'For i = 2 To 5 '迴圈設計 (這是限制資料表筆數=5)
    For i = 2 To Sheets.Count '自動偵測工作表合併數量
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
                Else 'else可以不用打
                End if
                Range("A1").Select '選擇A1儲存格子 CTRL + HOME
        
'-----以下為樞紐分析-----

    Sheets.Add
    ActiveSheet.Name = "分析"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "OK!R1C1:R1294C5", Version:=6).CreatePivotTable TableDestination:= _
        "分析!R3C1", TableName:="樞紐分析表2", DefaultVersion:=6
    Sheets("分析").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("樞紐分析表2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("樞紐分析表2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("樞紐分析表2").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("樞紐分析表2").PivotFields("店家")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("樞紐分析表2").AddDataField ActiveSheet.PivotTables("樞紐分析表2" _
        ).PivotFields("數量"), "加總 - 數量", xlSum
End Sub

            
