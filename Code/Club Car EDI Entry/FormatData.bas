Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatJitReport()
    Dim i As Long

    Sheets("JIT Report").Select

    'Delete the report header
    Range("1:2").Delete

    'Remove superfluous columns
    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Cells(1, i).Value <> "Item Nbr" And _
           Cells(1, i).Value <> "Item Desc" And _
           Cells(1, i).Value <> "Short Qty" Then
            Columns(i).Delete
        End If
    Next
End Sub

Sub FormatJitPiv()
    Dim PivData As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer

    Sheets("JIT Pivot").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    PivData = ActiveSheet.UsedRange

    'Delete the pivot table
    Cells.Delete
    
    'Read array out to sheet
    With Range(Cells(1, 1), Cells(TotalRows, TotalCols))
        .NumberFormat = "@"
        .Value = PivData
    End With
    
    'Fix column headers
    Range("A1").Value = "Item"
    Range("B1").Value = "Description"
    Range("C1").Value = "Qty"
    
    'Short Qty
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).NumberFormat = "#,0"
End Sub
