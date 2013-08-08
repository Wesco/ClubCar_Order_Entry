Attribute VB_Name = "Filter_Rejects"
Option Explicit

Sub FilterRejects()
    Dim aSheets As Variant  'List of sheets to loop through
    Dim aRejects As Variant
    Dim TotalRows As Long   'Number of columns on the active sheet
    Dim TotalCols As Long   'Number of rows on the active sheet
    Dim NOBRows As Long     'Number of rows on "Not On Blanket"
    Dim NOMRows As Long     'Number of rows on "Not On Master"
    Dim s As Variant        'For Each loop object
    Dim i As Long           'Do While loop counter

    'List of sheets to filter
    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In")

    For Each s In aSheets
        Sheets(s).Select

        'Get the number of rows and columns
        TotalRows = Rows(Rows.Count).End(xlUp).Row
        TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 2    'Add two columns since 2 will be added later

        'Check if the current sheet contains any items to order
        If Range("A2").Value <> "" Then
            'Insert columns for "SIM" and "On Blanket"
            Range("A:B").Insert

            'Add SIM numbers
            Range("B1").Value = "SIM"
            With Range("B2:B" & TotalRows)
                .Formula = "=IFERROR(VLOOKUP(C2,Master!A:B,2,FALSE),"""")"
                .NumberFormat = "@"
                .Value = .Value
            End With

            'Add On Blanket
            Range("A1").Value = "On Blanket"
            With Range("A2:A" & TotalRows)
                .Formula = "=IFERROR(IF(VLOOKUP(C2,Blanket!B:B,1,FALSE)=C2,""YES"",""NO""),""NO"")"
                .Value = .Value
            End With


            'Check for items that need to be removed from the order
            i = 2
            Do While i <= ActiveSheet.UsedRange.Rows.Count
                'If the item is not on the blanket move it to "Not On Blanket"
                If Cells(i, 1).Value = "NO" Then
                    NOBRows = NOBRows + 1
                    Range(Cells(i, 1), Cells(i, TotalCols)).Copy Destination:=Sheets("Not On Blanket").Range("A" & NOBRows)
                    Rows(i).Delete
                'If a SIM was not found move it to "Not On Master"
                ElseIf Cells(i, 2).Value = "" Then
                    NOMRows = NOMRows + 1
                    Range(Cells(i, 1), Cells(i, TotalCols)).Copy Destination:=Sheets("Not On Master").Range("A" & NOMRows)
                    Rows(i).Delete
                Else
                    i = i + 1
                End If
            Loop
        End If

        'Remove "On Blanket"
        Columns(1).Delete
    Next

    'Insert column headers if sheet isn't blank
    Sheets("Not On Master").Select
    If Range("A1").Value <> "" Then
        Rows(1).Insert
        Range("A1:M1") = Array("On Blanket", "SIM", "Part", "Description", "Value Stream", _
                               "Station Address", "VS Route", "Bin Size", "# Bins", _
                               "Qty Per Bin", "Station Name", "Supermarket Address", "Order")
        ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit
    End If

    'Insert column headers if sheet isn't blank
    Sheets("Not On Blanket").Select
    If Range("A1").Value <> "" Then
        Rows(1).Insert
        Range("A1:M1") = Array("On Blanket", "SIM", "Part", "Description", "Value Stream", _
                               "Station Address", "VS Route", "Bin Size", "# Bins", _
                               "Qty Per Bin", "Station Name", "Supermarket Address", "Order")
        ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit
    End If
End Sub

