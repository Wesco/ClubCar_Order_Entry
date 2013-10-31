Attribute VB_Name = "ExportReport"
Option Explicit

Sub ExportEDIOrd()
    Dim FileName As String
    Dim FilePath As String
    Dim CopyFilePath As String
    Dim PrevDispAlert As Boolean

    Sheets("EDI Order").Select
    FileName = Range("A1").Value & ".csv"
    FilePath = "\\idxexchange-new\EDI\Spreadsheet_PO\"
    CopyFilePath = "\\7938-HP02\Shared\club car\PO's dropped into EDI\" & Format(Date, "yyyy-mm-dd") & "\"
    PrevDispAlert = Application.DisplayAlerts

    Sheets("EDI Order").Copy
    ActiveSheet.Name = Range("A1").Value

    If Not FolderExists(CopyFilePath) Then
        RecMkDir CopyFilePath
    End If

    'Send to EDI and save a copy
    ActiveWorkbook.SaveAs FilePath & FileName, xlCSV
    ActiveWorkbook.SaveAs CopyFilePath & FileName, xlCSV

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

Sub ExportRemovedItems()
    Dim FileName As String
    Dim FilePath As String
    Dim PrevDispAlert As Boolean
    
    FileName = "Removed Items " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    FilePath = "\\br3615gaps\gaps\Club Car\Removed Items\"
    PrevDispAlert = Application.DisplayAlerts

    Sheets("Removed Items").Copy
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub














