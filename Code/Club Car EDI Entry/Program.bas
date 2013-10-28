Attribute VB_Name = "Program"
Option Explicit

Sub Main()

    MsgBox "Please select the 'JIT Report'"

    'If an error occurs while importing the JIT Report
    'display a message describing the error and remove
    'any data that was created at run-time
    On Error GoTo Import_Error
    ImportGaps
    ImportBlanket
    ImportMaster
    
    UserImportFile Sheets("JIT Report").Range("A1")
    FormatJitReport
    
    CreateJitPiv
    FormatJitPiv
    
    

    On Error GoTo 0

    Exit Sub

Import_Error:
    If Err.Number = Errors.USER_INTERRUPT And Err.Source = "UserImportFile" Then
        MsgBox "User canceled JIT import."
    Else
        MsgBox "Error " & Err.Number & " LN " & Erl & "(" & Err.Description & ") in procedure Clean of Module Program"
    End If
    Clean

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
