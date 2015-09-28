# VBACode
VBA Code

'Borrar una hoja determinada, pasada por parámetros
'Delete a Sheets

Private Sub BorrarHoja(HojaABorrar As String)

 Dim Hoja As Worksheet
 Application.DisplayAlerts = False
 'La busca en el libro y si la encuentra la elimina
  For Each Hoja In Worksheets
         If Hoja.Name = HojaABorrar Then
                Hoja.Delete
                Application.DisplayAlerts = True
                Exit Sub
            End If
    Next Hoja
Application.DisplayAlerts = True
End Sub


'Comprueba si existe Hoja
'Check for a Sheet
Function ExisteHoja(ByVal Nombre_Hoja As String) As Boolean
 'comprueba si existe la hoja a crear, devuelve Verdadero si Existe
 'Return TRUE if the Sheets exist
 On Error Resume Next
   ExisteHoja = CBool(Len(Sheets(Nombre_Hoja).Name) > 0)
End Function
