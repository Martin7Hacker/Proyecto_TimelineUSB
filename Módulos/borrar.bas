Attribute VB_Name = "borrar"
'***************************************************************************
'*
'*
'* borrar datos para Timeline
'*
'*
'***************************************************************************
Public Sub borrar()
On Error GoTo nose
 xnombre = ""
 xnombre2 = ""
 xapellido = ""
 xapellido2 = ""
 xdireccion = ""
 xdireccion2 = ""
 xlocalidad = ""
 xPais = ""
 xtelefono = ""
 xcel = ""
 xemail = ""
 xfacebook = ""
 xcomentario_general = ""
 xhora = ""
 xtipo = ""
 xintervalo = ""
 xcomentario = ""
 xlunes = ""
 xmartes = ""
 xmiercoles = ""
 xjueves = ""
 xviernes = ""
 xsabado = ""
 xdomingo = ""
 pu1 = 0
 pu2 = 0
 pu3 = 0
 pu4 = 0
 pu5 = 0
 pu6 = 0
 pu7 = 0
 pu8 = 0
 Dim B As Integer
 For B = 0 To 3
  frmprograma.listado(B).Clear
 Next
 frmprograma.lunes(0).Clear
 frmprograma.martes.Clear
 frmprograma.miercoles.Clear
 frmprograma.jueves.Clear
 frmprograma.viernes.Clear
 frmprograma.sabado.Clear
 frmprograma.domingo.Clear
 frmprograma.Filtro.Clear
nose:
End Sub
