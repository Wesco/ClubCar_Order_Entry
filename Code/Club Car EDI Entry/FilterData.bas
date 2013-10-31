Attribute VB_Name = "FilterData"
Option Explicit

Sub FilterRemovedItems()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim ColHeaders As Variant

    Sheets("Removed Items").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))

    'Remove duplicate column headers
    ActiveSheet.UsedRange.AutoFilter 1, "=PO_NUMBER"
    Cells.Delete
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)) = ColHeaders
    
    'Remove unrelated note columns
    Columns("M:N").Delete
End Sub

Sub FilterEDIOrd()
    Dim TotalRows As Long
    Dim TotalCols As Integer

    '\\idxexchange-new\EDI\Spreadsheet_PO\
    '
    '    A        B      C       D       E    F       G         H     I       J       K         L      M      N      O      P      Q      R
    '    1        2      3       4       5    6       7         8     9       10      11        12     13     14     15     16     17     18
    'PO_NUMBER , Branch, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, ShipTo, NOTE1, NOTE2, NOTE1, NOTE2, NOTE3, NOTE4

    Sheets("EDI Order").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 4

    'Add formula to show if items are on master
    Range("O1").Value = "NOTE1"
    Range("O2:O" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(I2,Master!A:B,2,FALSE)=0,""Not On Master"",""""),""Not On Master"")"

    'Add formula to show if items are on blanket
    Range("P1").Value = "NOTE2"
    Range("P2:P" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(I2,Blanket!B:B,1,FALSE)=0,""Not On Blanket"",""""),""Not On Blanket"")"

    'Add formula to show if items have bin size
    Range("Q1").Value = "NOTE3"
    Range("Q2:Q" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(I2,Master!A:D,4,FALSE)=0,""No Bin Size"", """"),""No Bin Size"")"

    'Add formula to show if items have qty per bin
    Range("R1").Value = "NOTE4"
    Range("R2:R" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(I2,Master!A:E,5,FALSE)=0,""No Qty Per Bin"",""""),""No Qty Per Bin"")"


    'Remove items without a qty per bin
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter 18, "<>"
    CopyRemove
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove items without a bin size
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter 17, "<>"
    CopyRemove
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove items not on blanket
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter 16, "<>"
    CopyRemove
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove items not on master
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter 15, "<>"
    CopyRemove
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Remove tape and breaker
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter 9, "=1014471", xlOr, "=102838101"
    CopyRemove

    'Remove additional note columns
    Columns("O:R").Delete
End Sub

'---------------------------------------------------------------------------------------
' Proc : CopyRemove
' Date : 10/29/2013
' Desc : Copies filtered data to the removed items sheet and then deletes it from
'---------------------------------------------------------------------------------------
Private Sub CopyRemove()
    Dim RIRows As Long
    Dim ColHeaders As Variant
    Dim PrevDispAlert As Boolean
    Dim TotalRows As Long
    Dim TotalCols As Integer

    Sheets("EDI Order").Select
    PrevDispAlert = Application.DisplayAlerts
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    RIRows = Sheets("Removed Items").UsedRange.Rows.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))

    'Copy data to removed items list
    If RIRows = 1 Then
        ActiveSheet.UsedRange.Copy Destination:=Sheets("Removed Items").Range("A1")
    Else
        ActiveSheet.UsedRange.Copy Destination:=Sheets("Removed Items").Cells(RIRows + 1, 1)
    End If

    'Remove copied items from edi order
    Application.DisplayAlerts = False
    ActiveSheet.UsedRange.Cells.Delete
    Application.DisplayAlerts = PrevDispAlert

    'Insert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)) = ColHeaders
End Sub
