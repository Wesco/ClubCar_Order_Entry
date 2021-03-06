Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateJitPiv()
    Dim PivCache As PivotCache
    Dim PivTable As PivotTable
    Dim PivData As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer

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
    PivData = ActiveSheet.UsedRange
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    'Delete the pivot table
    Cells.Delete

    'Read array out to sheet
    With Range(Cells(1, 1), Cells(TotalRows, TotalCols))
        .NumberFormat = "@"
        .Value = PivData
    End With

    Range("D1").Value = "Qty"
    Range("D2:D" & TotalRows).Formula = "=CEILING(C2/IFERROR(IF(VLOOKUP(A2,Master!A:E,5,FALSE)=0,1,VLOOKUP(A2,Master!A:E,5,FALSE)),1),1)*IFERROR(IF(VLOOKUP(A2,Master!A:E,5,FALSE)=0,1,VLOOKUP(A2,Master!A:E,5,FALSE)),1)"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value
    Columns(3).Delete
End Sub

Sub CreateEDIOrd()
    Dim TotalRows As Long
    Dim EDIHeaders As Variant

    '\\idxexchange-new\EDI\Spreadsheet_PO\
    '
    '    A        B      C       D       E    F       G         H     I       J       K         L      M      N
    '    1        2      3       4       5    6       7         8     9       10      11        12     13     14
    'PO_NUMBER , Branch, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, ShipTo, NOTE1, NOTE2

    Sheets("JIT Pivot").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Copy Item Number to EDI order
    Range("A1:A" & TotalRows).Copy Destination:=Sheets("EDI Order").Range("I1")

    'Copy Description to EDI order
    Range("B1:B" & TotalRows).Copy Destination:=Sheets("EDI Order").Range("J1")

    'Copy Quantity to EDI order
    Range("C1:C" & TotalRows).Copy Destination:=Sheets("EDI Order").Range("E1")

    Sheets("EDI Order").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Create column headers
    EDIHeaders = Array("PO_NUMBER", "BRANCH", "DPC", "CUST_LINE", "QTY", "UOM", "UNIT_PRICE", _
                       "SIM", "PART_NO", "DESC", "SHIP_DATE", "SHIP_TO", "NOTE1", "NOTE2")
    Range(Cells(1, 1), Cells(1, UBound(EDIHeaders) + 1)) = EDIHeaders

    'PO Number
    Range("A2:A" & TotalRows).Value = CreatePONumber()

    'Branch
    Range("B2:B" & TotalRows).Value = "3615"

    'DPC
    Range("C2:C" & TotalRows).Value = "14940"

    'Unit Price
    Range("G2:G" & TotalRows).Formula = "=IFERROR(VLOOKUP(I2,Master!A:C,3,FALSE),0)"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value

    'SIM
    Range("H2:H" & TotalRows).Formula = "=IFERROR(VLOOKUP(I2,Master!A:B,2,FALSE),"""")"
    Range("H2:H" & TotalRows).NumberFormat = "@"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value

    'UOM
    Range("F2:F" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(H2,Gaps!A:AJ,36,FALSE)=0,""E"",VLOOKUP(H2,Gaps!A:AJ,36,FALSE)),""E"")"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    'Ship To
    Range("L2:L" & TotalRows).Value = "1"

    'Note 1
    Range("M2:M" & TotalRows).Formula = "=VLOOKUP(I2,Master!A:D,4,FALSE)"
    Range("M2:M" & TotalRows).Value = Range("M2:M" & TotalRows).Value

    'Note 2
    Range("N2:N" & TotalRows).Formula = "=VLOOKUP(I2,Master!A:E,5,FALSE)"
    Range("N2:N" & TotalRows).Value = Range("N2:N" & TotalRows).Value
End Sub

Private Function CreatePONumber()
    Const ABC As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim FilePath As String
    Dim PONumber As String
    Dim Modifier As String
    Dim i As Long

    FilePath = "\\7938-HP02\Shared\club car\PO's dropped into EDI\" & Format(Date, "yyyy-mm-dd") & "\"
    PONumber = Sheets("Master").Range("F2").Value & "-JIT-" & Format(Date, "yymmdd")

    'If the file exists add a modifier to the end of the PO number
    Do While FileExists(FilePath & PONumber & ".csv")
        i = i + 1
        If i < 27 Then
            Modifier = "-" & Mid(ABC, i, 1)
        Else
            Modifier = "-" & i - 26
        End If
        PONumber = Sheets("Master").Range("F2").Value & "-JIT-" & Format(Date, "yymmdd") & Modifier
    Loop

    CreatePONumber = PONumber
End Function
