Attribute VB_Name = "Capacitors"
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


Sub ParameterLookUp()
Sheets("Component_FR_calc").Activate
On Error GoTo CHYBA:

Dim i As Long

Dim ws1 As Worksheet

Set ws1 = Worksheets("Capacitors")

For i = 1 To Range("A1000").End(xlUp).Row - 2
Range("B" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:W5000"), 19, 0)
Range("C" & i + 2).Value = WorksheetFunction.VLookup(Range("A" & i + 2).Value, ws1.Range("A3:W5000"), 3, 0)
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


Dim tantal As Double
Dim ceramic As Double
Dim vysledok As Double
tantal = 0
ceramic = 0

For j = 1 To Range("C1000").End(xlUp).Row - 2

    If InStr(Range("C" & j + 2).Value, "Tantalum") <> 0 Then

        tantal = tantal + Range("B" & j + 2).Value

    ElseIf InStr(Range("C" & j + 2).Value, "Ceramic") <> 0 Then
        ceramic = ceramic + Range("B" & j + 2).Value
        
    
    End If

Next


For k = 1 To Range("A1000").End(xlUp).Row - 2
    If InStr(Range("A" & k + 2).Value, "short") <> 0 Then
        vysledok = (0.49 * ceramic) + (0.57 * tantal)
    ElseIf InStr(Range("A" & k + 2).Value, "open") <> 0 Then
        vysledok = (0.51 * ceramic) + (0.43 * tantal)
    ElseIf InStr(Range("A" & k + 2).Value, "failure") <> 0 Then
        vysledok = ceramic + tantal
    End If
Next

Range("J3").Value = "FailureRate"
Range("K3").Value = vysledok
Range("K3").Copy

End Sub

Sub SheetDel()
Application.DisplayAlerts = False

Sheets("Component_FR_calc").Delete

Application.DisplayAlerts = True
End Sub



Sub C_CALC()
Attribute C_CALC.VB_ProcData.VB_Invoke_Func = "z\n14"
Call CopyCellValueToNewSheet
Call Split
Call Transpose_new
Call ParameterLookUp
Call StrFind
Sheets("Fmea").Activate
End Sub


