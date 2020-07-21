Attribute VB_Name = "puertof"

'Puerto de salisda del pc
Public COM As String

Public pu1, pu2, pu3, pu4, pu5, pu6, pu7, pu8 As Byte

Public Sub disparar_bit()
On Error GoTo nose
puertof.encenderUSB
nose:
End Sub

Private Sub encend(ByVal puerto As Byte)
On Error GoTo nose
 encenderUSB
nose:
End Sub

Private Sub apagar(ByVal puerto As Byte)
On Error GoTo nose
ApagarUSB
nose:
End Sub

Public Sub apagar_puertos()
On Error GoTo nose
ApagarUSB
nose:
End Sub

Private Sub encender()
On Error GoTo nose
encenderUSB
nose:
End Sub

Public Sub encenderUSB()
On Error GoTo nose
With frmprograma
.usb.Output = "1"
End With
nose:
End Sub

Private Sub ApagarUSB()
On Error GoTo nose
With frmprograma
.usb.Output = "0"
End With
nose:
End Sub
