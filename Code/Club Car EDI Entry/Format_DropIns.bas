Attribute VB_Name = "Format_DropIns"
Option Explicit

Sub FixDropIns()
    Dim StartTime As Double
    Dim aSheets As Variant
    Dim aHeaders As Variant
    Dim s As Variant

    StartTime = Timer
    aSheets = Array("AWD Drop In", "DS Drop In", "PREC Drop In", "UTIL Drop In")

    For Each s In aSheets
        Sheets(s).Select
        FormatData  'Removes irrelevant lines
        Separate    'Separates data into columns
        FixQty      'Multiply order qyt by qty per bin
        RemItmNOO   'Removes items not ordered
        ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
    Next
    
    Sheets("Info").Select
    Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1).Value = "FixDropIns"
    Cells(ActiveSheet.UsedRange.Rows.Count, 3).Value = Timer - StartTime
End Sub

Sub FormatData()
    Dim i As Long
    Dim s As String

    i = 2
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        If InStr(Cells(i, 1).Value, "NEW PARTS") Then
            Rows(i).Delete
        ElseIf InStr(Cells(i, 1), "Part Number") Then
            Rows(i).Delete
        ElseIf Cells(i, 1).Value = "" Then
            Rows(i).Delete
        ElseIf Cells(i, 1).Value = "LOADING" Then
            Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
End Sub

Sub Separate()
    Dim i As Long
    Dim aStr As Variant

    i = 2
    Do While i < ActiveSheet.UsedRange.Rows.Count
        aStr = Split(Cells(i, 1).Value, " ", 2)

        If UBound(aStr) > 0 Then
            If Cells(i, 2).Value = "" Then
                Cells(i, 1).Value = aStr(0)
                Cells(i, 2).Value = aStr(1)
            End If
        Else
            i = i + 1
        End If
    Loop
    On Error GoTo 0
End Sub

Sub FixQty()
    Dim Rng As Range
    Dim aRng As Variant
    Dim iRows As Long
    Dim iCols As Integer
    Dim i As Long
    Dim x As Long
    
    Columns("L:O").Delete
    
    iCols = ActiveSheet.UsedRange.Columns.Count + 1
    iRows = ActiveSheet.UsedRange.Rows.Count
    
    Cells(1, iCols).Value = "Order"
    Cells(2, iCols).Formula = "=IFERROR(IF(K2*H2=0,"""", K2*H2),"""")"
    Cells(2, iCols).AutoFill Destination:=Range(Cells(2, iCols), Cells(iRows, iCols))
    Range(Cells(2, iCols), Cells(iRows, iCols)).Value = Range(Cells(2, iCols), Cells(iRows, iCols)).Value
    Columns("K:K").Delete
    iCols = ActiveSheet.UsedRange.Columns.Count
End Sub

Sub RemItmNOO()
    Dim i As Long

    Columns("L:N").Delete

    i = 2
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        If InStr(Cells(i, 11).Value, " ") Then Cells(i, 11).Value = Replace(Cells(i, 11).Value, " ", "")
        If Cells(i, 11).Value = "" Then
            Rows(i).Delete
        ElseIf Cells(i, 11).Value <> "" Then
            i = i + 1
        End If
    Loop
End Sub











