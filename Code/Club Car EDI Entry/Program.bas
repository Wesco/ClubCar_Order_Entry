Attribute VB_Name = "Program"
Option Explicit

Sub Macro1()
    Dim Result As Boolean
    Application.ScreenUpdating = False

    On Error GoTo IMPORT_FAILED
    ImportGaps
    ImportMaster
    ImportBlanket
    
    On Error GoTo USER_ABORTED
    ImportSheets
    On Error GoTo 0
    
    FixDropIns
    FilterRejects
    CreateEDI
    SaveEdiCsv
    CleanUp

    Application.ScreenUpdating = True
    Exit Sub

USER_ABORTED:
    CleanUp
    MsgBox "User canceled, macro aborted!", vbOKOnly, "User Aborted"
    Exit Sub

IMPORT_FAILED:
    CleanUp
    MsgBox Err.Description, vbOKOnly, Err.Source
    Exit Sub
End Sub

Sub CleanUp()
    Dim aSheets As Variant
    Dim s As Variant

    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In", _
                    "Gaps", "Info", "Not On Blanket", "Not On Master", "Blanket", "Master")

    On Error Resume Next
    For Each s In aSheets
        Sheets(s).Select
        ActiveSheet.Cells.Delete
        Range("A1").Select
    Next
    On Error GoTo 0

    Sheets("Macro").Select
End Sub
