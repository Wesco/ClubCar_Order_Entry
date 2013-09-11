Attribute VB_Name = "Program"
Option Explicit

Sub Main()

End Sub

Sub Clean()
    Dim aSheets As Variant
    Dim s As Worksheet

    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            Cells.Delete
            Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select
End Sub
