Sub balancete_sienge()

Dim ws As Worksheet
Set ws = ActiveSheet

Application.ScreenUpdating = False

With ws
.Range("A1").Value = "Empresa:"
.Range("A2").Value = "CNPJ:"
.Range("A3").Value = "Período:"
.Range("B1").Value = ws.Range("B4").Value
.Range("B2").Value = ws.Range("B5").Value
.Range("B3").Value = ws.Range("D5").Value
.Rows("4:6").Delete
.Range("A1").Value = ""
.Columns("A").Insert
End With

ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
ws.Range("A5:A" & ultimaLinha).Formula = "=len(B5)"

'Pintar e formatar linhas
For i = 5 To ultimaLinha
    If ws.Cells(i, "A").Value < 21 Then
         ws.Rows(i).Interior.Color = RGB(242, 242, 242)
         ws.Rows(i).Font.Bold = True
    End If
Next i

With ws
.Range("K4").Value = "Relatório"
.Range("L4").Value = "Diferença"
.Range("M4").Value = "Status"
.Range("N4").Value = "Obs.:"
ws.Rows("5:5").Select
End With

ActiveWindow.FreezePanes = True

With ws
.Range("A4:N4").AutoFilter
.Columns("A:N").AutoFit
.Rows("1:4").Font.Bold = True
.Rows("1:4").Interior.Color = RGB(217, 217, 217)
.Activate
.Range("B1").Value = "Empresa:"
.Columns("C").ColumnWidth = 14.1
End With

ActiveWindow.DisplayGridlines = False
Application.ScreenUpdating = True

MsgBox "Limpeza do relatório finalizada"

End Sub
