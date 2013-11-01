Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatJitRep()
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
    Dim TotalRows As Long
    Dim TotalCols As Integer

    Sheets("JIT Pivot").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Fix column headers
    Range("A1").Value = "Item"
    Range("B1").Value = "Description"
    Range("C1").Value = "Qty"

    'Short Qty
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).NumberFormat = "0"
End Sub

Sub FormatEDIOrd()
    Dim TotalRows As Long

    '    A        B      C       D       E    F       G         H     I       J       K         L      M      N
    '    1        2      3       4       5    6       7         8     9       10      11        12     13     14
    'PO_NUMBER , Branch, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, ShipTo, NOTE1, NOTE2

    Sheets("EDI Order").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove invalid characters from part descriptions
    With Range(Cells(1, 10), Cells(TotalRows, 10))
        .Replace ",", ""
        .Replace ";", ""
        .Replace "/", "-"
    End With

    'Remove column headers
    Rows(1).Delete
End Sub
