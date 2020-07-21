Attribute VB_Name = "abrirF"
'***************************************************************************
'*
'*
'* Módulo para Abrir el Archivo en Timeline
'*
'*
'***************************************************************************
Option Explicit

      'variables de memoria
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
           Dim comentario1 As String         ' espesifica el comentario
           Dim tiempo     As String         ' espesifica el tiempo
'*************************************************************************
'
'
'
' La diferencia entre las variables de arriba y las variables de abajo es
' que le asigne una x para que no alla conflico las que no tienen x abren
' el archivo mientras que las que si lo tienen cargan y descargan los datos
' del fichero .
'
'*************************************************************************


   
           'variables de memoria con < x
      'Datos ---------------------------------------------------------------
           Public xnombre As String             ' nombre de la persona
           Public xnombre2 As String            ' segundo nombre de la persona
           Public xapellido As String           ' apellido de la persona
           Public xapellido2 As String          ' segundo apellido
           Public xdireccion As String          ' dirección donde vive
           Public xdireccion2 As String         ' dirección alternativa
           Public xlocalidad As String          ' localidad ej : Canelones
           Public xPais As String               ' pais donde reside
           Public xtelefono As String           ' telefono de linea
           Public xcel As String                ' celular inalambrico
           Public xemail As String              ' email correo electrónico
           Public xfacebook As String           ' red social facebook
           Public xcomentario_general As String ' comentario general obsional
        'Evento ------------------------------------------------------------
           Public xhora As String               'hora programada
           Public xtipo As String               'tipo : entada o salida
           Public xintervalo As String          'intervalo ej : 5 seg
           Public xcomentario As String         'comentario acerca de el evento
           Public xfiltro As String             'filtro solo hora o hora y dia
        'Dias  -------------------------------------------------------------
           Public xlunes As String              'si se activa el lunes
           Public xmartes As String             'si se activa el martes
           Public xmiercoles As String          'si se activa el miercoles
           Public xjueves As String             'si se activa el jueves
           Public xviernes As String            'si se activa el viernes
           Public xsabado As String             'si se activa el sabado
           Public xdomingo As String            'si se activa el domingo
        'Pines de Salida ---------------------------------------------------
           Public xp1 As String                 ' especifica la salida en 5v en p1
           Public xp2 As String                 ' especifica la salida en 5v en p2
           Public xp3 As String                 ' especifica la salida en 5v en p3
           Public xp4 As String                 ' especifica la salida en 5v en p4
           Public xp5 As String                 ' especifica la salida en 5v en p5
           Public xp6 As String                 ' especifica la salida en 5v en p6
           Public xp7 As String                 ' especifica la salida en 5v en p7
           Public xp8 As String                 ' especifica la salida en 5v en p8
           '-----------------------------------------------------------------
           Public xcommando As String           ' espesifica el comando de apagado
           Public xcomentario1 As String         ' espesifica el comentario
           Public xtiempo     As String         ' espesifica el tiempo
Public Sub Abrir_Fichero(ByRef variable As String)
On Error GoTo nose
  
Open variable For Input As 1
 Do While Not EOF(1)
  '
  '
  'Datos ----------------------------------------------
       Line Input #1, Nombre
                      xnombre = guardarF.es.desescriptar(Nombre)
       Line Input #1, Nombre2
                      xnombre2 = guardarF.es.desescriptar(Nombre2)
       Line Input #1, apellido
                      xapellido = guardarF.es.desescriptar(apellido)
       Line Input #1, apellido2
                      xapellido2 = guardarF.es.desescriptar(apellido2)
       Line Input #1, direccion
                      xdireccion = guardarF.es.desescriptar(direccion)
       Line Input #1, direccion2
                      xdireccion2 = guardarF.es.desescriptar(direccion2)
       Line Input #1, localidad
                      xlocalidad = guardarF.es.desescriptar(localidad)
       Line Input #1, Pais
                      xPais = guardarF.es.desescriptar(Pais)
       Line Input #1, telefono
                      xtelefono = guardarF.es.desescriptar(telefono)
       Line Input #1, cel
                      xcel = guardarF.es.desescriptar(cel)
       Line Input #1, email
                      xemail = guardarF.es.desescriptar(email)
       Line Input #1, facebook
                      xfacebook = guardarF.es.desescriptar(facebook)
       Line Input #1, comentario_general
                      xcomentario_general = guardarF.es.desescriptar(comentario_general)
  'Evento --------------------------------------------
  
       Line Input #1, hora
                      xhora = guardarF.es.desescriptar(hora)
                      frmprograma.listado(0).AddItem xhora
       Line Input #1, Tipo
                      xtipo = guardarF.es.desescriptar(Tipo)
                      frmprograma.listado(1).AddItem xtipo
       Line Input #1, intervalo
                      xintervalo = guardarF.es.desescriptar(intervalo)
                      frmprograma.listado(2).AddItem xintervalo
       Line Input #1, comentario
                      xcomentario = guardarF.es.desescriptar(comentario)
                      frmprograma.listado(3).AddItem xcomentario
       Line Input #1, Filtro
                      xfiltro = guardarF.es.desescriptar(Filtro)
                      frmprograma.Filtro.AddItem xfiltro
  'Dias ----------------------------------------------
       Line Input #1, lunes
                      xlunes = guardarF.es.desescriptar(lunes)
                      frmprograma.lunes(0).AddItem xlunes
       Line Input #1, martes
                      xmartes = guardarF.es.desescriptar(martes)
                      frmprograma.martes.AddItem xmartes
       Line Input #1, miercoles
                      xmiercoles = guardarF.es.desescriptar(miercoles)
                      frmprograma.miercoles.AddItem xmiercoles
       Line Input #1, jueves
                      xjueves = guardarF.es.desescriptar(jueves)
                      frmprograma.jueves.AddItem xjueves
       Line Input #1, viernes
                      xviernes = guardarF.es.desescriptar(viernes)
                      frmprograma.viernes.AddItem xviernes
       Line Input #1, sabado
                      xsabado = guardarF.es.desescriptar(sabado)
                      frmprograma.sabado.AddItem xsabado
       Line Input #1, domingo
                      xdomingo = guardarF.es.desescriptar(domingo)
                      frmprograma.domingo.AddItem xdomingo

  'Puerto --------------------------------------------
       Line Input #1, p1
                      pu1 = guardarF.es.desescriptar(p1)
       Line Input #1, p2
                      pu2 = guardarF.es.desescriptar(p2)
       Line Input #1, p3
                      pu3 = guardarF.es.desescriptar(p3)
       Line Input #1, p4
                      pu4 = guardarF.es.desescriptar(p4)
       Line Input #1, p5
                      pu5 = guardarF.es.desescriptar(p5)
       Line Input #1, p6
                      pu6 = guardarF.es.desescriptar(p6)
       Line Input #1, p7
                      pu7 = guardarF.es.desescriptar(p7)
       Line Input #1, p8
                      pu8 = guardarF.es.desescriptar(p8)
 'apagado ----------------------------------------------------
       Line Input #1, commando
                      xcommando = guardarF.es.desescriptar(commando)
                      frmprograma.liscomando.AddItem (xcommando)
       Line Input #1, comentario1
                      xcomentario1 = guardarF.es.desescriptar(comentario1)
                      frmprograma.lisdialogo.AddItem (xcomentario1)
       Line Input #1, tiempo
                      xtiempo = guardarF.es.desescriptar(tiempo)
                      frmprograma.listiempo.AddItem (xtiempo)
Loop

  Close #1
  pasarpin
nose:
End Sub
Private Sub pasarpin()
On Error GoTo nose
abrirF.xp1 = puertof.pu1
abrirF.xp2 = puertof.pu2
abrirF.xp3 = puertof.pu3
abrirF.xp4 = puertof.pu4
abrirF.xp5 = puertof.pu5
abrirF.xp6 = puertof.pu6
abrirF.xp7 = puertof.pu7
abrirF.xp8 = puertof.pu8
 frmprograma.mostrar_menu True
nose:
End Sub

