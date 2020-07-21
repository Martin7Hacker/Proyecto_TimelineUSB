Attribute VB_Name = "guardarF"
Public es As New escripta
'***************************************************************************
'*
'*
'* Procedimiento Abreviado  Guardar Ficheros Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Option Explicit
'variables de memoria
Public guardar_archivo As String
'Datos ---------------------------------------------------------------
Dim Nombre As String             ' nombre de la persona
Dim Nombre2 As String            ' segundo nombre de la persona
Dim apellido As String           ' apellido de la persona
Dim apellido2 As String          ' segundo apellido
Dim direccion As String          ' dirección donde vive
Dim direccion2 As String         ' dirección alternativa
Dim localidad As String          ' localidad ej : Canelones
Dim Pais As String               ' pais donde reside
Dim telefono As String           ' telefono de linea
Dim cel As String                ' celular inalambrico
Dim email As String              ' email correo electrónico
Dim facebook As String           ' red social facebook
Dim comentario_general As String ' comentario general obsional
'Evento ------------------------------------------------------------
Dim hora As String               'hora programada
Dim Tipo As String               'tipo : entada o salida
Dim intervalo As String          'intervalo ej : 5 seg
Dim comentario As String         'comentario acerca de el evento
Dim Filtro As String             'filtro solo hora o hora y dia
'Dias  -------------------------------------------------------------
Dim lunes As String              'si se activa el lunes
Dim martes As String             'si se activa el martes
Dim miercoles As String          'si se activa el miercoles
Dim jueves As String             'si se activa el jueves
Dim viernes As String            'si se activa el viernes
Dim sabado As String             'si se activa el sabado
Dim domingo As String            'si se activa el domingo
'Pines de Salida ---------------------------------------------------
Dim p1 As String                 ' especifica la salida en 5v en p1
Dim p2 As String                 ' especifica la salida en 5v en p2
Dim p3 As String                 ' especifica la salida en 5v en p3
Dim p4 As String                 ' especifica la salida en 5v en p4
Dim p5 As String                 ' especifica la salida en 5v en p5
Dim p6 As String                 ' especifica la salida en 5v en p6
Dim p7 As String                 ' especifica la salida en 5v en p7
Dim p8 As String                 ' especifica la salida en 5v en p8
'-----------------------------------------------------------------
Dim commando   As String         ' espesifica el comando de apagado
Dim comentario1 As String        ' espesifica el comentario
Dim tiempo     As String         ' espesifica el tiempo

Public Sub Almacenar_Fichero(ByVal variable As String)
 On Error GoTo no_se
 Dim g As Integer ' variable para el for que carga los datos
 Open variable For Output As 1
 ' Esrivimos el Archivo a Guardar escriptandolo
 For g = 0 To frmprograma.listado(0).ListCount - 1 'la cantidad de archivos - la pocición sin memoria
 frmprograma.mostrar_menu False
 'Datos ------------------------------------------------------------------
 Nombre = es.escriptar(xnombre)
 Print #1, Nombre
 Nombre2 = es.escriptar(xnombre2)
 Print #1, Nombre2
 apellido = es.escriptar(xapellido)
 Print #1, apellido
 apellido2 = es.escriptar(xapellido2)
 Print #1, apellido2
 direccion = es.escriptar(xdireccion)
 Print #1, direccion
 direccion2 = es.escriptar(xdireccion2)
 Print #1, direccion2
 localidad = es.escriptar(xlocalidad)
 Print #1, localidad
 Pais = es.escriptar(xPais)
 Print #1, Pais
 telefono = es.escriptar(xtelefono)
 Print #1, telefono
 cel = es.escriptar(xcel)
 Print #1, cel
 email = es.escriptar(xemail)
 Print #1, email
 facebook = es.escriptar(xfacebook)
 Print #1, facebook
 comentario_general = es.escriptar(xcomentario_general)
 Print #1, comentario_general
'Evento -----------------------------------------------------------------
 hora = es.escriptar(frmprograma.listado(0).List(g))
 Print #1, hora
 Tipo = es.escriptar(frmprograma.listado(1).List(g))
 Print #1, Tipo
 intervalo = es.escriptar(frmprograma.listado(2).List(g))
 Print #1, intervalo
 comentario = es.escriptar(frmprograma.listado(3).List(g))
 Print #1, comentario
 Filtro = es.escriptar(frmprograma.Filtro.List(g))
 Print #1, Filtro
'Dias -------------------------------------------------------------------
 lunes = es.escriptar(frmprograma.lunes(0).List(g))
 Print #1, lunes
 martes = es.escriptar(frmprograma.martes.List(g))
 Print #1, martes
 miercoles = es.escriptar(frmprograma.miercoles.List(g))
 Print #1, miercoles
 jueves = es.escriptar(frmprograma.jueves.List(g))
 Print #1, jueves
 viernes = es.escriptar(frmprograma.viernes.List(g))
 Print #1, viernes
 sabado = es.escriptar(frmprograma.sabado.List(g))
 Print #1, sabado
 domingo = es.escriptar(frmprograma.domingo.List(g))
 Print #1, domingo
'Puerto -----------------------------------------------------------------
 p1 = es.escriptar(Str(pu1))
 Print #1, p1
 p2 = es.escriptar(Str(pu2))
 Print #1, p2
 p3 = es.escriptar(Str(pu3))
 Print #1, p3
 p4 = es.escriptar(Str(pu4))
 Print #1, p4
 p5 = es.escriptar(Str(pu5))
 Print #1, p5
 p6 = es.escriptar(Str(pu6))
 Print #1, p6
 p7 = es.escriptar(Str(pu7))
 Print #1, p7
 p8 = es.escriptar(Str(pu8))
 Print #1, p8
'S.O------------------------------------------------------------------------
 commando = es.escriptar(frmprograma.liscomando.List(g))
 Print #1, commando
 comentario1 = es.escriptar(frmprograma.lisdialogo.List(g))
 Print #1, comentario1
 tiempo = es.escriptar(frmprograma.listiempo.List(g))
 Print #1, tiempo
'---------------------------------------------------------------------------
 Next g
 Close #1
 frmprograma.mostrar_menu True
no_se:
End Sub

