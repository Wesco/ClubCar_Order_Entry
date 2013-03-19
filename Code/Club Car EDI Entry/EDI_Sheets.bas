Attribute VB_Name = "EDI_Sheets"
Option Explicit

Sub CreateEDI()
    Dim aSheets As Variant
    Dim aRange As Variant
    Dim s As Variant

    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In")

    For Each s In aSheets
        Sheets(s).Select

        Select Case Sheets(s).Name
            Case "AWD Drop In"
                If Range("A2").Value <> "" Then
                    SetupEDI Sheets("Master").Range("F2") & "-AWD-" & Format(Date, "mmddyy"), "2"
                End If

            Case "DS Drop In"
                If Range("A2").Value <> "" Then
                    SetupEDI Sheets("Master").Range("F2") & "-DS-" & Format(Date, "mmddyy"), "1"
                End If

            Case "PREC Drop In"
                If Range("A2").Value <> "" Then
                    SetupEDI Sheets("Master").Range("F2") & "-PREC-" & Format(Date, "mmddyy"), "4"
                End If

            Case "UTIL Drop In"
                If Range("A2").Value <> "" Then
                    SetupEDI Sheets("Master").Range("F2") & "-UTIL-" & Format(Date, "mmddyy"), "3"
                End If

        End Select
    Next
End Sub


Sub SetupEDI(PO As String, ShipTo As String)
    Dim iRows As Long
    Const Branch As String = "3615"
    Const DPC As String = "14940"
    Dim i As Long

    iRows = ActiveSheet.UsedRange.Rows.Count

    Columns("A:G").EntireColumn.Insert
    Range(Cells(2, 1), Cells(iRows, 1)).Value = PO
    Range(Cells(2, 2), Cells(iRows, 2)).Value = Branch
    Range(Cells(2, 3), Cells(iRows, 3)).Value = DPC
    Columns(19).Cut Destination:=Columns(5)
    Columns("K:S").Delete
    Range("A1:N1").Value = Array("PO_NUMBER", "BRANCH", "DPC", _
                                 "CUST_LINE", "QTY", "UOM", "UNIT_PRICE", _
                                 "SIM", "PART_NO", "DESC", "SHIP_DATE", _
                                 "SHIP_TO", "NOTE1", "NOTE2")

    Range("F2").Formula = "=IFERROR(VLOOKUP(H2,Gaps!A:AJ,36,FALSE),"""")"
    Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(iRows, 6))

    Range("G2").Formula = "=IFERROR(VLOOKUP(H2,Master!B:C,2,FALSE),"""")"
    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(iRows, 7))

    Range(Cells(2, 12), Cells(iRows, 12)).Value = ShipTo    'Column L

    Range("M2").Formula = "=IFERROR(VLOOKUP(H2,Master!B:D,3,FALSE),"""")"
    Range("M2").AutoFill Destination:=Range(Cells(2, 13), Cells(iRows, 13))

    Range("N2").Formula = "=IFERROR(VLOOKUP(H2,Master!B:E,4,FALSE),"""")"
    Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(iRows, 14))

    i = 1
    Do While i < ActiveSheet.UsedRange.Rows.Count
        If InStr(Cells(i, 10).Value, ",") Then
            Cells(i, 10).Value = Replace(Cells(i, 10).Value, ",", "")
        End If
        i = i + 1
    Loop

    With ActiveSheet.UsedRange
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        .Value = .Value
    End With

    Rows(1).Delete
End Sub

















