Attribute VB_Name = "Microcircuit_Special"
Sub Split()

Sheets("Component_FR_calc").Activate

Dim r As Integer
Dim c As Integer

r = ActiveCell.Row
c = ActiveCell.Column

Selection.TextToColumns Destination:=Cells(r, c + 1), DataType:=xlDelimited, _
    ConsecutiveDelimiter:=True, Comma:=True, Space:=True

End Sub



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


If DoesSheetExists("Component_FR_calc") Then

    Sheets("Component_FR_calc").Cells.Clear
    Sheets("Fmea").Activate
    Set Rng = Application.Selection
    Rng.Copy
    Sheets("Component_FR_calc").Range("A1").PasteSpecial Paste:=xlPasteValues

Else

    Sheets("Fmea").Activate
    Set Rng = Application.Selection
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Component_FR_calc"
    Rng.Copy
    ws.Paste

End If



End Sub

Sub Transpose_new()

Sheets("Component_FR_calc").Activate
Dim Rng As Range
Dim i As Long

Set Rng = Range("B1", Cells(1, Columns.Count).End(xlToLeft))

For Each Cell In Rng
Cell.Value = WorksheetFunction.Trim(Cell)
Next

Rng.Copy
Range("A3").PasteSpecial Transpose:=True

End Sub

Sub ParameterLookUp()  'v tejto funkcii je nutne nastavit part failure rate stlpec

Sheets("Component_FR_calc").Activate

On Error GoTo CHYBA:

Dim i As Long
Dim ws1 As Worksheet
Set ws1 = Worksheets("Microcircuits")

For i = 1 To Range("A1000").End(xlUp).Row - 2
Range("B" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:AW5000"), 35, 0) 'treba zmenit cislo stlpca v ktorom je part failure rate
Range("C" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:AW5000"), 3, 0)
Next

Exit Sub

CHYBA:

Range("B" & i + 2).Value = "N/A"
Range("C" & i + 2).Value = "N/A"

Resume Next

End Sub

Sub StrFind()

Sheets("Component_FR_calc").Activate

Dim j As Long
Dim k As Long

Dim lin As Double
Dim dig As Double
Dim mem As Double
Dim vysledok As Double

lin = 0
dig = 0
mem = 0


For j = 1 To Range("C1000").End(xlUp).Row - 2

    If InStr(Range("C" & j + 2).Value, "Linear") <> 0 Then
        lin = lin + Range("B" & j + 2).Value
            For k = 1 To Range("A1000").End(xlUp).Row - 2
                If InStr(Range("A" & k + 2).Value, "improper") <> 0 Then
                    vysledok = 0.77 * lin
                ElseIf InStr(Range("A" & k + 2).Value, "no") <> 0 Then
                    vysledok = 0.23 * lin
                ElseIf InStr(Range("A" & k + 2).Value, "functional") And InStr(Range("A" & k + 3).Value, "failure") <> 0 Then
                    vysledok = (2 * lin) + (2 * dig) + (2 * mem)
                ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 And InStr(Range("A" & k + 1).Value, "functional") = 0 Then
                    vysledok = lin + dig + mem
                End If
            Next
                  
    ElseIf InStr(Range("C" & j + 2).Value, "Digital") <> 0 Then
        dig = dig + Range("B" & j + 2).Value
            For k = 1 To Range("A1000").End(xlUp).Row - 2
                If InStr(Range("A" & k + 2).Value, "improper") <> 0 Then
                    vysledok = 0.77 * dig
                ElseIf InStr(Range("A" & k + 2).Value, "no") <> 0 Then
                    vysledok = 0.23 * dig
                ElseIf InStr(Range("A" & k + 2).Value, "functional") And InStr(Range("A" & k + 3).Value, "failure") <> 0 Then
                    vysledok = (2 * lin) + (2 * dig) + (2 * mem)
                ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 And InStr(Range("A" & k + 1).Value, "functional") = 0 Then
                    vysledok = lin + dig + mem
                ElseIf InStr(Range("A" & k + 2).Value, "stuck") And InStr(Range("A" & k + 3).Value, "high") <> 0 Then
                    vysledok = 0.5 * dig
                ElseIf InStr(Range("A" & k + 2).Value, "stuck") And InStr(Range("A" & k + 3).Value, "low") <> 0 Then
                    vysledok = 0.5 * dig
                End If
            Next
            
    ElseIf InStr(Range("C" & j + 2).Value, "Memory") <> 0 Then
        mem = mem + Range("B" & j + 2).Value
        For k = 1 To Range("A1000").End(xlUp).Row - 2
                If InStr(Range("A" & k + 2).Value, "transfer") <> 0 Then
                    vysledok = 0.79 * mem
                ElseIf InStr(Range("A" & k + 2).Value, "bit") <> 0 Then
                    vysledok = 0.21 * mem
                ElseIf InStr(Range("A" & k + 2).Value, "functional") And InStr(Range("A" & k + 3).Value, "failure") <> 0 Then
                    vysledok = (2 * lin) + (2 * dig) + (2 * mem)
                ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 And InStr(Range("A" & k + 1).Value, "functional") = 0 Then
                    vysledok = lin + dig + mem
                End If
            Next
            
    
    End If

Next

For k = 1 To Range("A1000").End(xlUp).Row - 2

    
    
Next

Range("J3").Value = "FailureRate"
Range("K3").Value = vysledok
Range("K3").Copy


End Sub


Sub MC_Special_CALC()
Attribute MC_Special_CALC.VB_ProcData.VB_Invoke_Func = "w\n14"

Call CopyCellValueToNewSheet
Call Split
Call Transpose_new
Call ParameterLookUp
Call StrFind
Sheets("Fmea").Activate

End Sub

