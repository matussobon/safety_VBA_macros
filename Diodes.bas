Attribute VB_Name = "Diodes"
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
    'Sheets("Component_FR_calc").Paste
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

Set ws1 = Worksheets("Diodes")

For i = 1 To Range("A1000").End(xlUp).Row - 2
Range("B" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:AW5000"), 29, 0) 'treba zmenit cislo stlpca v ktorom je part failure rate
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


Dim tvs As Double
Dim rect As Double
Dim rect1 As Double
Dim smSig As Double
Dim vReg As Double
Dim vysledok As Double

tvs = 0
rect = 0
rect1 = 0
smSig = 0
vReg = 0

For j = 1 To Range("C1000").End(xlUp).Row - 2

    If InStr(Range("C" & j + 2).Value, "Suppressor") <> 0 Then
        tvs = tvs + Range("B" & j + 2).Value

    ElseIf InStr(Range("C" & j + 2).Value, "Rectifier") <> 0 Then
        rect = rect + Range("B" & j + 2).Value
        
         ElseIf InStr(Range("C" & j + 2).Value, "Schottky") <> 0 Then
        rect1 = rect1 + Range("B" & j + 2).Value

    ElseIf InStr(Range("C" & j + 2).Value, "Switching") <> 0 Then
        smSig = smSig + Range("B" & j + 2).Value

    ElseIf InStr(Range("C" & j + 2).Value, "Regulator") <> 0 Then
        vReg = vReg + Range("B" & j + 2).Value
    
    End If

Next


For k = 1 To Range("A1000").End(xlUp).Row - 2
    If InStr(Range("A" & k + 2).Value, "short") <> 0 Then
        vysledok = (1 * tvs) + (0.61 * rect) + (0.47 * smSig) + (0.35 * vReg) + (0.61 * rect1)

    ElseIf InStr(Range("A" & k + 2).Value, "open") <> 0 Then
        vysledok = (0.39 * rect) + (0.53 * smSig) + (0.65 * vReg) + (0.39 * rect1)

    ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 Then
        vysledok = rect + smSig + vReg + rect1
    End If
Next

Range("J3").Value = "FailureRate"
Range("K3").Value = vysledok
Range("K3").Copy

End Sub




Sub D_CALC()
Attribute D_CALC.VB_ProcData.VB_Invoke_Func = "d\n14"
Call CopyCellValueToNewSheet
Call Split
Call Transpose_new
Call ParameterLookUp
Call StrFind
Sheets("Fmea").Activate
End Sub


