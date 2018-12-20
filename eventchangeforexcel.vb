'registra la fecha y hora en la columna 2 en el momento que escribas en la columna 1
'funciona para excel en visual basic 
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 Then
        Sheets("Hoja1").Cells(Target.Row, 2) = Now
    End If
    
End Sub
