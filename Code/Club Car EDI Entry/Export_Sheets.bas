Attribute VB_Name = "Export_Sheets"
Option Explicit

Sub SaveEdiCsv()
    Dim aSheets As Variant
    Dim s As Variant
    Dim sPath As String

    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In", "Not On Blanket")
    sPath = "\\br3615gaps\gaps\Club Car\Not On Blanket\"

    For Each s In aSheets
        Sheets(s).Select

        Select Case s
            Case "AWD Drop In"
                If Range("A1").Value <> "Part Number" Then
                    CopySave Sheets(s).Name
                End If

            Case "DS Drop In"
                If Range("A1").Value <> "Part Number" Then
                    CopySave Sheets(s).Name
                End If

            Case "PREC Drop In"
                If Range("A1").Value <> "Part Number" Then
                    CopySave Sheets(s).Name
                End If

            Case "UTIL Drop In"
                If Range("A1").Value <> "Part Number" Then
                    CopySave Sheets(s).Name
                End If

            Case "Not On Blanket"
                If Range("A2").Value <> "" Then
                    Application.DisplayAlerts = False
                    Sheets(s).Copy
                    ActiveWorkbook.SaveAs FileName:=sPath & "Not On Blanket " & Format(Date, "m-dd-yy") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
                    ActiveWorkbook.Close
                    ThisWorkbook.Activate
                    Application.DisplayAlerts = True
                    Email SendTo:="ataylor@wesco.com; JAbercrombie@wesco.com;", _
                          Subject:="CC Items Not On Blanket", _
                          Body:="""\\br3615gaps\gaps\Club Car\Not On Blanket\" & "Not On Blanket " & Format(Date, "m-dd-yy") & ".xlsx"""
                End If
        End Select
    Next
End Sub

Sub CopySave(WS As String)
    Dim PO As String
    Dim BackupPath As String
    Const EDIPath As String = "\\idxexchange-new\EDI\Spreadsheet_PO\"

    BackupPath = "\\7938-hp02\shared\club car\PO's dropped into EDI\" & Format(Date, "m-dd-yy") & "\"

    If FolderExists(BackupPath) = False Then
        MkDir BackupPath
    End If

    Sheets(WS).Select
    PO = Range("A1").Value
    Application.DisplayAlerts = False
    Sheets(WS).Copy
    ActiveWorkbook.SaveAs FileName:=EDIPath & PO & ".csv", FileFormat:=xlCSV
    ActiveWorkbook.SaveAs FileName:=BackupPath & PO & ".csv", FileFormat:=xlCSV
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub
