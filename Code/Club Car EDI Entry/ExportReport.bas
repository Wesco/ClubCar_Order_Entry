Attribute VB_Name = "ExportReport"
Option Explicit

Sub ExportEDIOrd()
    Dim FileName As String
    Dim FilePath As String
    Dim CopyFilePath As String
    Dim PrevDispAlert As Boolean
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim s As Worksheet
    Dim i As Long

    Sheets("EDI Order").Select
    FilePath = "\\idxexchange-new\EDI\Spreadsheet_PO\"
    CopyFilePath = "\\7938-HP02\Shared\club car\PO's dropped into EDI\" & Format(Date, "yyyy-mm-dd") & "\"
    PrevDispAlert = Application.DisplayAlerts
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    'Copy EDI Order to a new workbook
    Sheets("EDI Order").Copy
    ActiveSheet.Name = Range("A1").Value

    'Create the file path if it doesn't exist
    If Not FolderExists(CopyFilePath) Then
        RecMkDir CopyFilePath
    End If

    'If the order has more than 40 lines split it up into multiple orders
    If TotalRows > 40 Then
        ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count), Count:=WorksheetFunction.RoundUp(TotalRows / 40, 0) - 1
        Sheets(1).Select
        For i = TotalRows To 41 Step -40
            Range(Cells(i - 39, 1), Cells(i, TotalCols)).Cut Destination:=Sheets(WorksheetFunction.RoundUp(i / 40, 0)).Range("A1")
        Next
    End If

    For Each s In ActiveWorkbook.Sheets
        'Change the PO number based on the number of orders in the current batch
        s.Name = CreatePONum(s)
        s.Range("A1:A" & s.Rows(Rows.Count).End(xlUp).Row).Value = s.Name
        s.Copy
        FileName = Range("A1").Value & ".csv"

        'Send to EDI and save a copy
        ActiveWorkbook.SaveAs FilePath & FileName, xlCSV
        ActiveWorkbook.SaveAs CopyFilePath & FileName, xlCSV
        ActiveWorkbook.Saved = True

        MsgBox "PO " & Range("A1").Value & " sent!", vbOKOnly, "PO Sent"

        'Close the EDI order
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = PrevDispAlert
    Next

    'Close the EDI workbook
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
              CC:="acoffey@wesco.com; rsetzer@wesco.com", _
              Subject:="CC Removed Items", _
              Body:="A copy of the removed items report is attached. The report can also be found on the network <a href=""" & FilePath & FileName & """>here</a>.", _
              Attachment:=FilePath & FileName
    End If
End Sub

Private Function CountChar(Source As String, Find As String)
    Dim char As String
    Dim i As Long
    Dim j As Long

    For i = 1 To Len(Source)
        char = Mid(Source, i, 1)
        If char = Find Then j = j + 1
    Next

    CountChar = j
End Function

Private Function CreatePONum(Sheet As Worksheet)
    Const ABC As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim PONumber As String
    Dim Modifier As String
    Dim CharCount As Long
    Dim i As Long
    Dim j As Long

    PONumber = Sheet.Range("A1").Value
    CharCount = CountChar(PONumber, "-")
    j = InStr(ABC, Right(PONumber, Len(PONumber) - InStrRev(PONumber, "-")))
    i = InStr(ABC, Right(PONumber, Len(PONumber) - InStrRev(PONumber, "-"))) + ActiveWorkbook.Sheets.Count - Sheet.Index

    If CharCount = 3 And j > 0 Then
        PONumber = Left(PONumber, InStrRev(PONumber, "-") - 1)
    End If

    If i < 27 And i > 0 And j > 0 Or CharCount = 2 And i < 27 And i > 0 Then
        Modifier = "-" & Mid(ABC, i, 1)
    ElseIf j = 0 And CharCount = 3 Then
        Modifier = "-" & CLng(Right(PONumber, Len(PONumber) - InStrRev(PONumber, "-"))) + ActiveWorkbook.Sheets.Count - Sheet.Index
        PONumber = Left(PONumber, InStrRev(PONumber, "-") - 1)
    ElseIf i > 26 Then
        Modifier = "-" & i - 26
    End If

    CreatePONum = PONumber & Modifier
End Function
