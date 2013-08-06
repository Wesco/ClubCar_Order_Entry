Attribute VB_Name = "Import_Sheets"
Option Explicit

Function ImportSheets() As Boolean
    Dim sPath As String
    Dim StartTime As Double
    Dim TotalRows As Long
    Dim TotalCols As Integer

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    sPath = Application.GetOpenFilename
    StartTime = Timer

    If sPath <> "False" Then
        Workbooks.Open sPath
        Sheets("AWD").Select
        UnhideData
        TotalRows = Rows(Rows.Count).End(xlUp).Row
        TotalCols = Columns(Columns.Count).End(xlToLeft).Column
        Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Copy Destination:=ThisWorkbook.Sheets("AWD Drop In").Range("A1")

        Sheets("DS").Select
        UnhideData
        TotalRows = Rows(Rows.Count).End(xlUp).Row
        TotalCols = Columns(Columns.Count).End(xlToLeft).Column
        Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Copy Destination:=ThisWorkbook.Sheets("DS Drop In").Range("A1")

        Sheets("Prec Cpl").Select
        UnhideData
        TotalRows = Rows(Rows.Count).End(xlUp).Row
        TotalCols = Columns(Columns.Count).End(xlToLeft).Column
        Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Copy Destination:=ThisWorkbook.Sheets("PREC Drop In").Range("A1")

        Sheets("Util Cpl").Select
        UnhideData
        TotalRows = Rows(Rows.Count).End(xlUp).Row
        TotalCols = Columns(Columns.Count).End(xlToLeft).Column
        Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Copy Destination:=ThisWorkbook.Sheets("UTIL Drop In").Range("A1")

        ActiveWorkbook.Close

        Sheets("Info").Select
        Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).Value = "ImportSheets"
        Cells(ActiveSheet.UsedRange.Rows.Count, 3).Value = Timer - StartTime
        ActiveSheet.Columns.EntireColumn.AutoFit
        ImportSheets = True
    Else
        Sheets("Info").Select
        Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).Value = "ImportSheets"
        Cells(ActiveSheet.UsedRange.Rows.Count, 3).Value = "Failed"
        ActiveSheet.Columns.EntireColumn.AutoFit
        ImportSheets = False
    End If

    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Function

Sub ImportMaster()
    Dim sPath As String
    sPath = "\\br3615gaps\gaps\Club Car\Master\Club Car Master " & Format(Date, "yyyy") & ".xlsx"

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
    ActiveWorkbook.Close
End Sub

Sub UnhideData()
    Columns.Hidden = False
    Columns.AutoFit
End Sub















