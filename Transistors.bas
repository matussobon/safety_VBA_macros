Attribute VB_Name = "Transistors"
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

Set ws1 = Worksheets("Transistors")

For i = 1 To Range("A1000").End(xlUp).Row - 2
Range("B" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:AI5000"), 32, 0) 'treba zmenit cislo stlpca v ktorom je part failure rate
Range("C" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:AI5000"), 3, 0)
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


Dim bjt As Double
Dim fet As Double
Dim vysledok As Double
bjt = 0
fet = 0

For j = 1 To Range("C1000").End(xlUp).Row - 2

    If InStr(Range("C" & j + 2).Value, "Bipolar") <> 0 Then

        bjt = bjt + Range("B" & j + 2).Value

    ElseIf InStr(Range("C" & j + 2).Value, "MOSFET") <> 0 Then
        fet = fet + Range("B" & j + 2).Value
        
    
    End If

Next


For k = 1 To Range("A1000").End(xlUp).Row - 2
    If InStr(Range("A" & k + 2).Value, "short") <> 0 Then
        vysledok = (0.73 * bjt) + (0.73 * fet)
    ElseIf InStr(Range("A" & k + 2).Value, "open") <> 0 Then
        vysledok = (0.27 * bjt) + (0.27 * fet)
    ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 Then
        vysledok = bjt + fet
    End If
Next

Range("J3").Value = "FailureRate"
Range("K3").Value = vysledok
Range("K3").Copy

End Sub




Sub T_CALC()
Attribute T_CALC.VB_ProcData.VB_Invoke_Func = "s\n14"
Call CopyCellValueToNewSheet
Call Split
Call Transpose_new
Call ParameterLookUp
Call StrFind
Sheets("Fmea").Activate
End Sub


