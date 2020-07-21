Attribute VB_Name = "fc"
'***************************************************************************
'*
'*
'* Procedimiento Abreviado  Timeline
'*
'*
'***************************************************************************

Public Sub comp_clave_fSalir(ByVal camb As Boolean, ByVal cla_num _
 As Byte, ByVal cla_hex As String, ByVal comp_num As Byte, com_hex _
 As String, ByVal ventana As Form)
 On Error GoTo nose
 'modo numerico
 Select Case (camb)
 Case (False)
 If cla_num = comp_num Then
 f_salir ventana
 End If
 If cla_hex = com_hex Then
 f_salir ventana
 End If
 Case (True)
 If cla_num = comp_num And cla_hex = com_hex Then
 f_salir ventana
 End If
End Select
nose:
End Sub

Private Sub f_salir(ByVal vent As Form)
On Error GoTo nose
Unload vent
nose:
End Sub
