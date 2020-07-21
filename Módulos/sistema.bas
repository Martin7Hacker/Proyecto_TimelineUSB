Attribute VB_Name = "sistema"
'***************************************************************************
'*
'*
'* sistema de Timeline
'*
'*
'***************************************************************************
Public comando As String
Public tiempo As String
Public comentario As String
Public textoX As String
Public TextoY As String
Public ven As Byte

Public Sub tomarDatos()
On Error GoTo nose
 'le pasa el comando de disparo
 comando = frmfunciones.devolver_comando
 'le pasa el tiempo t en seg.
 tiempo = frmfunciones.DTPicker1.Minute
 'comentario del apagado
 comentario = frmfunciones.txtd.Text
nose:
End Sub

Public Sub ingresarDatos()
On Error GoTo nose
 'ingresar datos
 With frmprograma
 .liscomando.AddItem (comando)
 .listiempo.AddItem (tiempo)
 .lisdialogo.AddItem (comentario)
 End With
nose:
End Sub

Public Sub modificarDatos() 'funciones para modificar datos del timbre
On Error GoTo nose
 With frmprograma
 .liscomando.List(.liscomando.ListIndex) = sistema.comando
 .lisdialogo.List(.lisdialogo.ListIndex) = sistema.comentario
 .listiempo.List(.listiempo.ListIndex) = sistema.tiempo
 End With
nose:
End Sub

Public Sub eleminarDatos() 'Elimniar datos en memoria
On Error GoTo nose
 With frmprograma
 .liscomando.Clear
 .lisdialogo.Clear
 .listiempo.Clear
 End With
nose:
End Sub

