Attribute VB_Name = "Export_Sheets"
Option Explicit

Sub SaveEdiCsv()
    Dim aSheets As Variant
    Dim s As Variant
    Dim sPath As String
    Dim NOMPath As String
    Dim Body As String

    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In", "Not On Blanket", "Not On Master")
    sPath = "\\br3615gaps\gaps\Club Car\Not On Blanket\"
    NOMPath = "\\br3615gaps\gaps\Club Car\Not On Master\"

    For Each s In aSheets
        Sheets(s).Select

        Select Case s
            Case "Not On Blanket"
                If Range("A2").Value <> "" Then
                    Application.DisplayAlerts = False
                    Sheets(s).Copy
                    ActiveWorkbook.SaveAs _
                            FileName:=sPath & "Not On Blanket " & Format(Date, "yyyy-mm-dd") & ".xlsx", _
                            FileFormat:=xlOpenXMLWorkbook
                    ActiveWorkbook.Close
                    ThisWorkbook.Activate
                    Application.DisplayAlerts = True
                    Body = Body & "<br>" & _
                           """\\br3615gaps\gaps\Club Car\Not On Blanket\" & "Not On Blanket " & Format(Date, "yyyy-mm-dd") & ".xlsx"""
                End If

            Case "Not On Master"
                If Range("A2").Value <> "" Then
                    Application.DisplayAlerts = False
                    Sheets(s).Copy
                    ActiveWorkbook.SaveAs _
                            FileName:=NOMPath & "Not On Master " & Format(Date, "yyyy-mm-dd") & ".xlsx", _
                            FileFormat:=xlOpenXMLWorkbook
                    ActiveWorkbook.Close
                    ThisWorkbook.Activate
                    Application.DisplayAlerts = True
                    Body = Body & "<br>" & _
                           """\\br3615gaps\gaps\Club Car\Not On Master\" & "Not On Master " & Format(Date, "yyyy-mm-dd") & ".xlsx"""
                End If

            Case Else
                If Range("A1").Value <> "Part Number" Then
                    CopySave Sheets(s).Name
                End If
        End Select
    Next

    If Body <> "" Then
        Email SendTo:="ataylor@wesco.com; JAbercrombie@wesco.com;", _
              Subject:="CC Macro Notification", _
              Body:=Body
    End If
End Sub

Sub CopySave(WS As String)
    Dim PO As String
    Dim BackupPath As String
    Const EDIPath As String = "\\idxexchange-new\EDI\Spreadsheet_PO\"

    BackupPath = "\\7938-hp02\shared\club car\PO's dropped into EDI\" & Format(Date, "yyyy-mm-dd") & "\"

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
