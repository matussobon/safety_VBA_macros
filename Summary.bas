Attribute VB_Name = "Summary"
Function DoesSheetExists(sh As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sh)
    On Error GoTo 0

    If Not ws Is Nothing Then DoesSheetExists = True

End Function

Sub CopyCellValueToNewSheet()


Dim ws As Worksheet
Dim Rng As Range
Dim i As Long


If DoesSheetExists("Table_fmea") Then

    Sheets("Table_fmea").Cells.Clear
    
    'kopirovanie end effect
    Sheets("Fmea").Activate
    Range("F:F").Copy
    Sheets("Table_fmea").Range("D1").PasteSpecial Paste:=xlPasteValues
    
    'kopirovanie failure rate
    Sheets("Fmea").Activate
    Range("S:S").Copy
    Sheets("Table_fmea").Range("F1").PasteSpecial Paste:=xlPasteValues

    'kopirovanie oznacenia
    Sheets("Fmea").Activate
    Range("A:A").Copy
    Sheets("Table_fmea").Range("A1").PasteSpecial Paste:=xlPasteValues

    Sheets("Fmea").Activate
    Range("B:B").Copy
    Sheets("Table_fmea").Range("B1").PasteSpecial Paste:=xlPasteValues

    'kopirovanie severity category
    Sheets("Fmea").Activate
    Range("G:G").Copy
    Sheets("Table_fmea").Range("E1").PasteSpecial Paste:=xlPasteValues

    'kopirovanie det. method
    Sheets("Fmea").Activate
    Range("H:H").Copy
    Sheets("Table_fmea").Range("G1").PasteSpecial Paste:=xlPasteValues
    
Else
    'kopirovanie end effect
    Sheets("Fmea").Activate
    Range("F:F").Copy
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Table_fmea"
    Sheets("Table_fmea").Range("D1").PasteSpecial Paste:=xlPasteValues
    
    'kopirovanie failure rate
    Sheets("Fmea").Activate
    Range("S:S").Copy
    Sheets("Table_fmea").Range("F1").PasteSpecial Paste:=xlPasteValues
    
    'kopirovanie oznacenia
    Sheets("Fmea").Activate
    Range("A:A").Copy
    Sheets("Table_fmea").Range("A1").PasteSpecial Paste:=xlPasteValues

    Sheets("Fmea").Activate
    Range("B:B").Copy
    Sheets("Table_fmea").Range("B1").PasteSpecial Paste:=xlPasteValues
    
    'kopirovanie severity category
    Sheets("Fmea").Activate
    Range("G:G").Copy
    Sheets("Table_fmea").Range("E1").PasteSpecial Paste:=xlPasteValues

    'kopirovanie det. method
    Sheets("Fmea").Activate
    Range("H:H").Copy
    Sheets("Table_fmea").Range("G1").PasteSpecial Paste:=xlPasteValues

End If


End Sub
Sub delEmptyRows()
Sheets("Table_fmea").Activate
Application.ScreenUpdating = False
Columns("D:D").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Application.ScreenUpdating = True
End Sub


Sub FindEffect()
Dim ws As Worksheet
Dim i As Long

Set ws = Worksheets("Table_fmea")

'Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
Range("D:D").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("L:L"), Unique:=True
End Sub

Sub FindSeverity() 'funkcia na hladanie severity, jednoduchy for loop a if statement
Dim i As Long
Dim j As Long
Dim k As Long

Range("N1").Value = "Severity Category"
Range("O1").Value = "Severity"

For i = 1 To Range("L1000").End(xlUp).Row - 1
    For j = 1 To Range("D1000").End(xlUp).Row - 1
        If Range("L" & i + 1).Value = Range("D" & j + 1).Value Then
            Range("N" & i + 1).Value = Range("E" & j + 1).Value
        End If
    Next
Next

Dim Rng As String
For k = 1 To Range("N1000").End(xlUp).Row - 1
     Rng = Range("N" & k + 1).Value
    Select Case Rng
        Case "1"
            Rng = "Catastrophic"
            Range("O" & k + 1).Value = Rng
        Case "2"
            Rng = "Critical"
            Range("O" & k + 1).Value = Rng
        Case "3"
            Rng = "Marginal"
            Range("O" & k + 1).Value = Rng
        Case "4"
            Rng = "Negligible"
            Range("O" & k + 1).Value = Rng
    End Select
Next

Dim RomRng As String
For k = 1 To Range("N1000").End(xlUp).Row - 1
     RomRng = Range("N" & k + 1).Value
    Select Case RomRng
        Case "1"
            RomRng = "I"
            Range("N" & k + 1).Value = RomRng
        Case "2"
            RomRng = "II"
            Range("N" & k + 1).Value = RomRng
        Case "3"
            RomRng = "III"
            Range("N" & k + 1).Value = RomRng
        Case "4"
            RomRng = "IV"
            Range("N" & k + 1).Value = RomRng
    End Select
Next



        
End Sub

Sub CalculateFailRate() 'funkcia ktora pocita finalny percentualny failure rate

Dim ws As Worksheet
Dim i As Long
Set ws = Worksheets("Table_fmea")

Range("P1").Value = "Failure Rate per Hour"

For i = 1 To Range("L1000").End(xlUp).Row - 1
    Range("P" & i + 1).Value = WorksheetFunction.SumIf(Range("D:D"), "*" & Range("L" & i + 1) & "*", Range("F:F"))
Next

Dim FailRatePerHour As Double
Dim FailRateHundred As Double

Range("Q1").Value = "Percentage Of Failure Rate"

For i = 1 To Range("P1000").End(xlUp).Row - 1
    FailRateHundred = Range("P" & i + 1).Value + FailRateHundred
Next

For i = 1 To Range("P1000").End(xlUp).Row - 1
    FailRatePerHour = Range("P" & i + 1).Value
    Range("Q" & i + 1).Value = (FailRatePerHour * 100) / FailRateHundred
Next
End Sub
Sub DetMethodCalc() ' funkcia na pocitanie deteringu, vyuziva for loop, case a podla vzorcov vyrata vysledok

Dim i As Long
Dim FailureRateSum As Double

For i = 1 To Range("E1000").End(xlUp).Row - 1
    If Range("E" & i + 1).Value <> 4 Then
        FailureRateSum = Range("F" & i + 1).Value + FailureRateSum
    End If
Next
Range("T1").Value = "Det. Method"
Range("T2").Value = 1
Range("T3").Value = 3
Range("T4").Value = 4
Range("T5").Value = 5
Range("T6").Value = 6
Range("T7").Value = 36
Range("U1").Value = "Det. Coverage"

Dim lambda_one As Double
Dim lambda_three As Double
Dim lambda_four As Double
Dim lambda_five As Double
Dim lambda_six As Double
Dim CellValueCheck As Long

For i = 1 To Range("G1000").End(xlUp).Row - 1
   CellValueCheck = Range("G" & i + 1).Value
    Select Case CellValueCheck
        Case 1
            lambda_one = lambda_one + Range("F" & i + 1).Value
        Case 3
            lambda_three = lambda_three + Range("F" & i + 1).Value
        Case 4
            lambda_four = lambda_four + Range("F" & i + 1).Value
        Case 5
            lambda_five = lambda_five + Range("F" & i + 1).Value
        Case 6
            lambda_six = lambda_six + Range("F" & i + 1).Value
        Case 13
            lambda_one = lambda_one + Range("F" & i + 1).Value
            lambda_three = lambda_three + Range("F" & i + 1).Value
    End Select
Next

For i = 1 To Range("T1000").End(xlUp).Row - 1
    Select Case Range("T" & i + 1).Value
        Case 1
            Range("U" & i + 1).Value = (lambda_one / FailureRateSum) * 100
        Case 3
            Range("U" & i + 1).Value = (lambda_three / FailureRateSum) * 100
        Case 4
            Range("U" & i + 1).Value = (lambda_four / FailureRateSum) * 100
        Case 5
            Range("U" & i + 1).Value = (lambda_five / FailureRateSum) * 100
        Case 6
            Range("U" & i + 1).Value = (lambda_six / FailureRateSum) * 100
        Case 36
           Range("U" & i + 1).Value = ((lambda_three / FailureRateSum) * 100) + ((lambda_six / FailureRateSum) * 100)
    End Select
Next






End Sub

Sub SortFR() 'zoradi FailureRate od najmensej hodnoty
Dim lastrow As Long
lastrow = Cells(Rows.Count, 16).End(xlUp).Row
Range("L2:Q" & lastrow).Sort key1:=Range("P2:P" & lastrow), _
   order1:=xlDescending, Header:=xlNo
End Sub

Sub StringJoin()
Dim i As Long

Columns(1).NumberFormat = "@" '<-- nastavi format stlpcov A,B,C na text only aby nedrbalo zlucovaniu textu
Columns(2).NumberFormat = "@"
Columns(3).NumberFormat = "@"

Range("C1").Value = "Failure Mode Identifier"
    For i = 1 To Range("A1000").End(xlUp).Row - 1
        Range("C" & i + 1).Value = Range("A" & i + 1).Value & Range("B" & i + 1).Value
    Next
End Sub

Sub FailureModeAssign()

Dim i As Long
Dim j As Long
Range("M1").Value = "Failure Mode Identifier"
For i = 1 To Range("L1000").End(xlUp).Row - 1
    For j = 1 To Range("D1000").End(xlUp).Row - 1
        If Range("L" & i + 1).Value = Range("D" & j + 1).Value Then
            Range("M" & i + 1).Value = Range("M" & i + 1).Value & Range("C" & j + 1).Value & " "
            End If
        Next
Next
    
End Sub


Sub TableCreate()

Call CopyCellValueToNewSheet
Call delEmptyRows
Call FindEffect
Call StringJoin
Call FailureModeAssign
Call FindSeverity
Call CalculateFailRate
Call SortFR
Call DetMethodCalc

End Sub


