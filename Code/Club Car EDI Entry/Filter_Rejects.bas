Attribute VB_Name = "Filter_Rejects"
Option Explicit

Sub FilterRejects()
    Dim aSheets As Variant
    Dim aRejects As Variant
    Dim s As Variant        'For Each loop object
    Dim iRows As Long
    Dim iRows2 As Long
    Dim iCols As Long
    Dim i As Long           'Do While loop counter


    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In")

    For Each s In aSheets
        Sheets(s).Select
        If Range("A2").Value <> "" Then
            Range("A:A").EntireColumn.Insert
            Range("A1").Value = "SIM"
            Range("A2").Formula = "=IFERROR(VLOOKUP(B2,Master!A:B,2,FALSE),"""")"
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count
            Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
            Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value

            i = 2
            Do While i <= ActiveSheet.UsedRange.Rows.Count
                If Cells(i, 1).Value = "" Then
                    aRejects = Range(Cells(i, 1), Cells(i, iCols))
                    Rows(i).Delete
                    Sheets("Not On Blanket").Select
                    Range("A1:L1") = Array("SIM", "Part", "Description", "Value Stream", _
                                           "Station Address", "VS Route", "Bin Size", "# Bins", _
                                           "Qty Per Bin", "Station Name", "Supermarket Address", "Order")
                    iRows2 = ActiveSheet.UsedRange.Rows.Count + 1
                    Range(Cells(iRows2, 1), Cells(iRows2, UBound(aRejects, 2))) = aRejects
                    Sheets(s).Select
                Else
                    i = i + 1
                End If
            Loop
        End If
    Next

    Sheets("Not On Blanket").Select
    Columns(1).Delete
    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit
End Sub

