Attribute VB_Name = "detonador"
'***************************************************************************
'*
'*
'* Coleciónes en Memoria para Timeline
'*
'*
'***************************************************************************
Public evento As evento
Public dias As MonthView
Public amacenar As Integer
Public sd As Boolean

Public Sub comparacionVirtual()
On Error GoTo nose
solo_hora
nose:
End Sub

Private Sub solo_hora()
On Error GoTo nose
 sd = False ' no se cumplio la condicion
 Dim x As Integer
 With frmprograma
 For x = 0 To .listado(0).ListCount - 1
 If .listado(0).List(x) = Time Then
 sd = True
 End If
 If sd = True Then
 sd = False
 frmtimbre.timpo_programado = .listado(2).List(x)
 frmtimbre.comentario_general = .listado(3).List(x)
 frmtimbre.Label1(0).Caption = lenguaje_Menu(341) & _
 .listado(2).List(x) & " " & lenguaje_Menu(382)
 frmtimbre.Label1(4).Caption = lenguaje_Menu(345) & " " & .listado(1).List(x)
 frmtimbre.Label1(3).Caption = lenguaje_Menu(344) & _
 " " & .listado(0).List(x) & " " & lenguaje_Menu(350)
 frmprograma.listado(0).ListIndex = x
 frmprograma.listado(1).ListIndex = x
 frmprograma.listado(2).ListIndex = x
 frmprograma.listado(3).ListIndex = x
 frmtimbre.frmsolo_hora.Visible = True: frmtimbre.Show 1
 End If
 Next
 End With
nose:
End Sub

Public Sub hora_y_dia()
 On Error GoTo nose
 sd = False 'no se cumplio la condición
 Dim d As Integer
 With frmprograma
 For d = 0 To .listado(0).ListCount - 1
 'comparacion de hora y dia
 'comparacion de hora y lunes
 If .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .lunes(0).List(d) Then
 funcion_activar d, 0
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .martes.List(d) Then
 funcion_activar d, 1
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .miercoles.List(d) Then
 funcion_activar d, 2
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .jueves.List(d) Then
 funcion_activar d, 3
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .viernes.List(d) Then
 funcion_activar d, 4
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .sabado.List(d) Then
 funcion_activar d, 5
 ElseIf .listado(0).List(d) = Time And .MonthView1.DayOfWeek = _
 .domingo.List(d) Then
 funcion_activar d, 6
 End If
 Next
 End With
nose:
End Sub

Private Sub funcion_activar(ByVal x As Integer, ByVal dia As Byte)
On Error GoTo nose
 With frmprograma
 sd = False
 frmtimbre.timpo_programado = .listado(2).List(x)
 frmtimbre.comentario_general = .listado(3).List(x)
 frmtimbre.Label1(0).Caption = lenguaje_Menu(341) & .listado(2).List(x) & " " & lenguaje_Menu(350)
 frmtimbre.Label1(4).Caption = lenguaje_Menu(345) & " " & .listado(1).List(x)
 frmtimbre.Label1(3).Caption = lenguaje_Menu(344) & _
 " " & .listado(0).List(x) & " " & lenguaje_Menu(350)
 frmprograma.listado(0).ListIndex = x
 frmprograma.listado(1).ListIndex = x
 frmprograma.listado(2).ListIndex = x
 frmprograma.listado(3).ListIndex = x
 If frmtimbre.Visible = False Then
 Select Case dia
 Case (0)
 borrar 0, 6
 frmtimbre.Check1(0).Value = 1
 Case (1)
 borrar 0, 6
 frmtimbre.Check1(1).Value = 1
 Case (2)
 borrar 0, 6
 frmtimbre.Check1(2).Value = 1
 Case (3)
 borrar 0, 6
 frmtimbre.Check1(3).Value = 1
 Case (4)
 borrar 0, 6
 frmtimbre.Check1(4).Value = 1
 Case (5)
 borrar 0, 6
 frmtimbre.Check1(5).Value = 1
 Case (6)
 borrar 0, 6
 frmtimbre.Check1(6).Value = 1
End Select
frmtimbre.fram_dias.Visible = True: frmtimbre.Show 1
End If
 End With
nose:
End Sub

Private Sub borrar(ByVal principio As Integer, ByVal fin As Integer)
On Error GoTo nose
 Dim x As Integer
 For x = principio To fin
 frmtimbre.Check1(x).Value = 0
 Next
nose:
End Sub

