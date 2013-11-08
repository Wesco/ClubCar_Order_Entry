Attribute VB_Name = "Imports"
Option Explicit

Sub ImportBlanket()
    Dim FilePath As String
    Dim FileName As String
    Dim PrevDispAlert As Boolean

    FilePath = "\\br3615gaps\gaps\Club Car\Master\"
    FileName = "Blanket " & Format(Date, "yyyy") & ".xlsx"
    PrevDispAlert = Application.DisplayAlerts

    Workbooks.Open FilePath & FileName
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Blanket").Range("A1")

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

Sub ImportMaster()
    Dim FilePath As String
    Dim FileName As String
    Dim TotalRows As Long
    Dim PrevDispAlert As Boolean

    FilePath = "\\br3615gaps\gaps\Club Car\Master\"
    FileName = "Club Car Master " & Format(Date, "yyyy") & ".xlsx"
    PrevDispAlert = Application.DisplayAlerts

    Workbooks.Open FilePath & FileName
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Store part numbers as text
    Columns(1).Insert
    Range("A1").Value = "Part Number"
    With Range(Cells(2, 1), Cells(TotalRows, 1))
        .Formula = "=""="" & """""""" & B2 & """""""""
        .Value = .Value
    End With
    Columns(2).Delete
End Sub
