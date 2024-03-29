Sub Macro1()
'
' Macro1 Macro
' 巨集由 01561 錄製，時間: 2021/08/02
'
    Selection.AutoFilter
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$1", Visible:=False
    Range("H1").Select
    Range("A1:AK21149").AutoFilter Field:=8, Criteria1:=Array("BSHSC-Hub Billing", "SOLSC-資材專用訂單", "SORSC-RMA-出貨", "RORSC-RMA-退貨", "SPPSC-試產銷貨訂單", "SNGSC-資材NG+Scrap", "CBOSC-更正-出貨", "CBRSC-更正-退貨", "RODSC-DOA-退貨", "SODSC-DOA-出貨", "BMHSC-ISBU Hub Billing", "CBRSC-跨期更正-退貨", "CBOSC-跨期更正-出貨", "BBHSC-BU3 Hub Billing", ""), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$21149", Visible:=False
    Range("A2:A21149").Select

    Selection.Delete Shift:=xlShiftUp

    Range("C5898").Select
    Range("A1:AK5897").AutoFilter Field:=8
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$5897", Visible:=False
    Range("A1:AK5897").AutoFilter Field:=6, Criteria1:="<>0*", Operator:=xlAnd, Criteria2:="<>98*"
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$5897", Visible:=False
    Range("A1:AK5897").AutoFilter Field:=6, Criteria1:="<>0*", Operator:=xlAnd, Criteria2:="<>98*"
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$5897", Visible:=False
    Rows("2:5781").Select
    ActiveWindow.ScrollRow = 4037
    Selection.Delete Shift:=xlShiftUp
    ActiveWindow.ScrollRow = 1
    Range("A1:AK3573").AutoFilter Field:=6
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3573", Visible:=False
    Range("F1").Select
    ActiveWindow.ScrollColumn = 14
    Range("A1:AK3573").AutoFilter Field:=28, Criteria1:=Array("SN904"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3573", Visible:=False
    Rows("7:304").Select
    Range("N7").Activate
    Range("A1:AK3567").AutoFilter Field:=28
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3567", Visible:=False
    Range("X5").Select
    Range("A1:AK3567").AutoFilter Field:=28, Criteria1:=Array("TP03", "S2SN904", "TP31", "RD001", "S1001", "TP32", "TP51", "TP82MKI"), Operator:=xlFilterValues
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3567", Visible:=False
    Rows("7:2273").Select
    Range("N2273").Activate
    Selection.Delete Shift:=xlShiftUp
    Range("A1:AK3543").AutoFilter Field:=28
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3543", Visible:=False
    Range("Z11").Select
    Range("A1:AK3543").AutoFilter Field:=29
    ActiveWorkbook.Names.Add Name:="'VES_SOMX5063_65401429_20200201-'!_FilterDatabase", RefersTo:="='VES_SOMX5063_65401429_20200201-'!$A$1:$AK$3543", Visible:=False
    Columns("R:R").Select
    Selection.Insert Shift:=xlShiftToRight
    Columns("T:T").Select
    Selection.Insert Shift:=xlShiftToRight
    Range("R1").Select
    Selection.FormulaR1C1 = "R Month"
    Range("T1").Select
    Selection.FormulaR1C1 = "S Month"
    Range("R2").Select
    Selection.FormulaR1C1 = "=IF(AND(Q2>=" & """" & "2020/08" & """" & ",Q2<" & """" & "2021/08" & """" & "),MONTH(Q2),0)"
    Selection.AutoFill Destination:=Range("R2").EntireColumn, Type:=xlFillDefault
    Range("R2:R3543").Select
    Selection.FormulaR1C1 = "=IF(AND(S2>=" & """" & "2020/08" & """" & ",S2<" & """" & "2021/08" & """" & "),MONTH(S2),0)"
    Selection.AutoFill Destination:=Range("T2").EntireColumn, Type:=xlFillDefault
    Range("T2:T3543").Select
    Columns("U:U").Select
    Selection.Insert Shift:=xlShiftToRight
    Range("U1").Select
    Selection.FormulaR1C1 = "OTD"
    Range("U2").Select
    Selection.FormulaR1C1 = "=IF(AND(S2>Q2,T2<>R2), " & """" & "Delay" & """" & ", IF(T2=R2," & """" & "OTD" & """" & "," & """" & "Advance" & """" & "))"
    Selection.AutoFill Destination:=Range("U2:U3543"), Type:=xlFillDefault
    Range("U2:U3543").Select
    ActiveWindow.ScrollRow = 1
    Range("N1").Select
    ActiveWindow.ScrollColumn = 4
    With ActiveWindow
        .ScrollRow = 3518
        .ScrollColumn = 20
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    ActiveWorkbook.PivotCaches.Add SourceType:=xlDatabase, SourceData:="='VES_SOMX5063_65401429_20200201-'!R1C1:R3543C40"
    Sheets.Add
    ActiveWorkbook.PivotCache.CreatePivotTable TableDestination:="Sheet1!R3C1", TableName:="樞紐分析表1", DefaultVersion:=1
    Range("A3").Activate
    With ActiveSheet.PivotTable.PivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTable.PivotField
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTable.AddDataField Field:=ActiveSheet.PivotTable.PivotField
End Sub


