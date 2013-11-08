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

    'Copy EDI Order to a new workbook
    Sheets("EDI Order").Copy
    ActiveSheet.Name = Range("A1").Value

    'Create the file path if it doesn't exist
    If Not FolderExists(CopyFilePath) Then
        RecMkDir CopyFilePath
    End If

    'Send to EDI and save a copy
    ActiveWorkbook.SaveAs FilePath & FileName, xlCSV
    ActiveWorkbook.SaveAs CopyFilePath & FileName, xlCSV

    'Close the EDI order
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

Sub ExportRemovedItems()
    Dim FileName As String
    Dim FilePath As String
    Dim PrevDispAlert As Boolean
    Dim i As Long

    FileName = "Removed Items " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    FilePath = "\\br3615gaps\gaps\Club Car\Removed Items\"
    PrevDispAlert = Application.DisplayAlerts

    If Sheets("Removed Items").Range("A2").Value <> "" Then
        'Copy Removed Items to a new workbook
        Sheets("Removed Items").Copy
        ActiveSheet.UsedRange.Columns.AutoFit

        'Append a modifier to the filename if the file exists
        Do While FileExists(FilePath & FileName)
            i = i + 1
            FileName = "Removed Items " & Format(Date, "yyyy-mm-dd") & " (" & i & ")" & ".xlsx"
        Loop

        'Save to the network
        ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook

        'Close the Removed Items workbook
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = PrevDispAlert

        'Email Removed Items
        Email SendTo:="ataylor@wesco.com", _
              CC:="JAbercrombie@wesco.com", _
              Subject:="CC Removed Items", _
              Body:="A copy of the removed items report is attached. The report can also be found on the network <a href=""" & FilePath & FileName & """>here</a>.", _
              Attachment:=FilePath & FileName
    End If
End Sub
