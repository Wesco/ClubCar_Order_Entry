Attribute VB_Name = "Program"
Option Explicit

Sub Macro1()
    Dim Result As Boolean
    Application.ScreenUpdating = False

    If ImportGaps = True And ImportSheets = True Then
        ImportMaster
        FixDropIns
        FilterRejects
        CreateEDI
        SaveEdiCsv
        CleanUp
    End If

    Application.ScreenUpdating = True
End Sub

Sub CleanUp()
    Dim aSheets As Variant
    Dim s As Variant

    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In", "Gaps", "Info", "Not On Blanket", "Master")

    On Error Resume Next
    For Each s In aSheets
        Sheets(s).Select
        ActiveSheet.Cells.Delete
        Range("A1").Select
    Next
    On Error GoTo 0

    Sheets("Macro").Select
End Sub
