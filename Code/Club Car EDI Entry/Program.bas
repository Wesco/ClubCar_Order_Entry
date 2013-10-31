Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.0"

Sub Main()
    ImportGaps      '\\br3615gaps\gaps\3615 Gaps Download\
    ImportBlanket   '\\br3615gaps\gaps\Club Car\Master\
    ImportMaster    '\\br3615gaps\gaps\Club Car\Master\

    MsgBox "Please select the 'JIT Report'"
    UserImportFile Sheets("JIT Report").Range("A1")
    
    FormatJitRep
    CreateJitPiv
    FormatJitPiv

    CreateEDIOrd
    FilterEDIOrd
    FormatEDIOrd
    
    FilterRemovedItems
    ExportRemovedItems
End Sub

'---------------------------------------------------------------------------------------
' Proc : Clean
' Date : 9/11/2013
' Desc : Removes all data from the macro
'---------------------------------------------------------------------------------------
Sub Clean()
    Dim PrevAlrt As Boolean
    Dim PrevScrn As Boolean
    Dim s As Worksheet

    'Stores current state
    PrevAlrt = Application.DisplayAlerts
    PrevScrn = Application.ScreenUpdating

    'Disables alerts/screen updating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'Activate the macro workbook
    ThisWorkbook.Activate

    'Removes data
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            Cells.Delete
            Range("A1").Select
        End If
    Next

    'Selects the cell under the macro button
    Sheets("Macro").Select
    Range("C7").Select

    'Resets alerts/screen updating
    Application.DisplayAlerts = PrevAlrt
    Application.ScreenUpdating = PrevScrn
End Sub
