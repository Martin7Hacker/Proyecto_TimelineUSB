Attribute VB_Name = "externosF"
'***************************************************************************
'*
'*
'* Archivos Externos en  Timeline
'*
'*
'***************************************************************************
Dim historial As String 'variable de almacenamiento de historiales
Dim xhistorial As String 'variable de almacenamiento de xhistoriales
Public selecionado As String
Public xselecionado As String
Public ventana As String

Public Sub Abrir_Archivo_Externo()
 On Error GoTo nose
 Open App.Path & "\Historial.ini" For Input As 1
 Do While Not EOF(1)
 Line Input #1, historial
 xhistorial = guardarF.es.desescriptar(historial)
 frmArranque.List1.AddItem xhistorial
 Loop
 Close #1
 AbrirVentana
 frmprograma.WindowState = sistema.ven
nose:
End Sub

Public Sub Abrir_selecionado()
 On Error GoTo nose
 Open App.Path & "\Seleccionado.ini" For Input As 1
 Do While Not EOF(1)
 Line Input #1, selecionado
 xselecionado = guardarF.es.desescriptar(selecionado)
 Loop
 Close #1
 guardar_archivo = xselecionado
nose:
End Sub


Public Sub guardar_Archivo_Externo()
On Error GoTo nose
 Open App.Path & "\Historial.ini" For Output As 1
 Dim g As Integer
 For g = 0 To frmArranque.List1.ListCount - 1
 historial = guardarF.es.escriptar(frmArranque.List1.List(g))
 Print #1, historial
 Next
 Close #1
nose:
End Sub

Public Sub AbrirVentana()
 On Error GoTo nose
 Open App.Path & "\ventana.ini" For Input As 1
 Do While Not EOF(1)
 Line Input #1, ventana
 sistema.ven = CInt(ventana)
 Loop
 Close #1
nose:
End Sub

Public Sub GuardarVentana()
 On Error GoTo nose
 Open App.Path & "\ventana.ini" For Output As 1
 Print #1, sistema.ven
 Close #1
nose:
End Sub

Public Sub guardar_selecionado()
On Error GoTo nose
 Open App.Path & "\Seleccionado.ini" For Output As 1
 selecionado = guardarF.es.escriptar(xselecionado)
 Print #1, selecionado
 Close #1
nose:
End Sub

