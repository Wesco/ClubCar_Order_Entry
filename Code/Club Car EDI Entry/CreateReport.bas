Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateJitPiv()
    Dim PivCache As PivotCache
    Dim PivTable As PivotTable

    Sheets("JIT Report").Select

    'Create pivot table cache
    Set PivCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                                       SourceData:=ActiveSheet.UsedRange, _
                                                       Version:=xlPivotTableVersion14)

    'Create pivot table from cache
    Set PivTable = PivCache.CreatePivotTable(TableDestination:=Sheets("JIT Pivot").Range("A1"), _
                                               TableName:="PivotTable1", _
                                               DefaultVersion:=xlPivotTableVersion14)

    Sheets("JIT Pivot").Select
    Range("A1").Select

    With PivTable
        .PivotFields("Item Nbr").Orientation = xlRowField
        .PivotFields("Item Nbr").LayoutForm = xlTabular
        .PivotFields("item Nbr").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("Item Desc").Orientation = xlRowField
        .PivotFields("Item Desc").LayoutForm = xlTabular
        
        .AddDataField .PivotFields("Short Qty"), "Sum of Short Qty", xlSum
    End With

    PivTable.ColumnGrand = False
End Sub
