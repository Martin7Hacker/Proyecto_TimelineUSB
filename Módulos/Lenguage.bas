Attribute VB_Name = "Lenguage"
'***************************************************************************
'*
'*
'* Lenguaje con Timeline
'*
'*
'***************************************************************************
Public lenguaje_Menu(385) As String              ' representa un vector de cadena de Menú.
Public sel As String
Public Sub definir_lenguage_opciones() 'estructura para ejecutar el lenguage del programa
                                       'es desir hace posible que el programa dentro
                                       'del menú opciónes allan diferentes opciones.
                                       'de idioma.
On Error GoTo nose
                                       
Call cmdcargar

lenguaje_Menu(0) = "Archivo"
lenguaje_Menu(1) = "Nuevo"
lenguaje_Menu(2) = "Abrir"
lenguaje_Menu(3) = "Guardar"
lenguaje_Menu(4) = "Guardar Como"
lenguaje_Menu(5) = "Salir"
lenguaje_Menu(6) = "&Ver"
lenguaje_Menu(7) = "&Panel de Dias"
lenguaje_Menu(8) = "Opciones"
lenguaje_Menu(9) = "Registros"
lenguaje_Menu(10) = "Nuevo"
lenguaje_Menu(11) = "Modificar"
lenguaje_Menu(12) = "Eliminación"
lenguaje_Menu(13) = "Eliminar Todo"
lenguaje_Menu(14) = "Eliminar Selecionado"
lenguaje_Menu(15) = "Desplazar"
lenguaje_Menu(16) = "Anterior"
lenguaje_Menu(17) = "Siguiente"
lenguaje_Menu(18) = "Salida"
lenguaje_Menu(19) = "Puerto Paralelo"
lenguaje_Menu(20) = "Opciones de Archivo"
lenguaje_Menu(21) = "Rutas de Archivo"
lenguaje_Menu(22) = "Automatizar Programa"
lenguaje_Menu(23) = "Ejecutar cuando inicie el pc"
lenguaje_Menu(24) = "Usuario"
lenguaje_Menu(25) = "Datos personales"
lenguaje_Menu(26) = "Utilizar Manualmente"
lenguaje_Menu(27) = "Dispositivo"
lenguaje_Menu(28) = "Calendario"
lenguaje_Menu(29) = "Generador de Rutinas de Eventos programables"

lenguaje_Menu(30) = "Visor"
lenguaje_Menu(31) = "Ventana"
lenguaje_Menu(32) = "Ayuda"
lenguaje_Menu(33) = "Temas de ayuda"
lenguaje_Menu(34) = "Configurar Idioma"
lenguaje_Menu(35) = "Acerca de Virtual Martin Temporize"
lenguaje_Menu(36) = "Circuito Electrònio"
lenguaje_Menu(37) = "Personalizar Idioma"
lenguaje_Menu(38) = "---> Donativo  <---"
lenguaje_Menu(39) = "&Mostrar todos los Meses"
lenguaje_Menu(40) = "&Solo definidos Actuales"
lenguaje_Menu(41) = "&Historial"
'// CARGAR IDIOMA EN EL MENÚ



 '################################################################
 'define las opciones dentro  de Opciones de Modificado
 'generador de timbres
 '################################################################
 lenguaje_Menu(42) = "Opciónes de modificado"
 lenguaje_Menu(43) = "&" & "Modificado de datos"
 lenguaje_Menu(44) = "&" & "Oprimiendo los bótones"
 lenguaje_Menu(45) = "&" & "Hora"
 lenguaje_Menu(46) = "&" & "Tipo"
 lenguaje_Menu(47) = "&" & "Filtro"
 lenguaje_Menu(48) = "&" & "Int Entrada"
 lenguaje_Menu(49) = "&" & "Int Salida"
 lenguaje_Menu(50) = "&" & "Texto Entrada"
 lenguaje_Menu(51) = "&" & "Texto Salida"
 lenguaje_Menu(52) = "&" & "Lunes"
 lenguaje_Menu(53) = "&" & "Martes"
 lenguaje_Menu(54) = "&" & "Miercoles"
 lenguaje_Menu(55) = "&" & "Jueves"
 lenguaje_Menu(56) = "&" & "Viernes"
 lenguaje_Menu(57) = "&" & "Sabados"
 lenguaje_Menu(58) = "&" & "Domingos"
 lenguaje_Menu(59) = "&" & "Tipo de Aplicado:"
 lenguaje_Menu(60) = "Se niegan las opciones oprimidas"
 lenguaje_Menu(61) = "Se niegan las opciónes no oprimidas"
 lenguaje_Menu(62) = "&" & "Aplicar"
 lenguaje_Menu(63) = "&" & "Salir"
 lenguaje_Menu(64) = "&" & "C: on/off"
 lenguaje_Menu(65) = "&" & "Restaurar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'generador de timbres
 '################################################################
 lenguaje_Menu(66) = "Generador de Rutinas"
 lenguaje_Menu(67) = "&" & "Lista Desplegable"
 lenguaje_Menu(68) = "&" & "Programa Desplegable"
 lenguaje_Menu(69) = "&" & "Opciones de Modificado"
 lenguaje_Menu(70) = "&" & "Hora:"
 lenguaje_Menu(71) = "&" & "Tipo:"
 lenguaje_Menu(72) = "&" & "Filtro:"
 lenguaje_Menu(73) = "&" & "Int:                                  [Entrada]"
 lenguaje_Menu(74) = "&" & "Int:                                  [Salida]"
 lenguaje_Menu(75) = "&" & "DIAS"
 lenguaje_Menu(76) = "&" & "Lunes"
 lenguaje_Menu(77) = "&" & "Martes"
 lenguaje_Menu(78) = "&" & "Miercoles"
 lenguaje_Menu(79) = "&" & "Jueves"
 lenguaje_Menu(80) = "&" & "Viernes"
 lenguaje_Menu(81) = "&" & "Sabados"
 lenguaje_Menu(82) = "&" & "Domingos"
 lenguaje_Menu(83) = "&" & "No existe ningun elemento de evento."
 lenguaje_Menu(84) = "&" & "Existen actualmente"
 lenguaje_Menu(85) = "&" & "elementos de evento"
 lenguaje_Menu(86) = "&" & "elemento de evento"
 lenguaje_Menu(87) = "&" & "[" & "ENTRADA" & "]"
 lenguaje_Menu(88) = "&" & "[" & "SALIDA" & "]"
 lenguaje_Menu(89) = "&" & " Cancelar"
 lenguaje_Menu(90) = "&" & "Crear Evento:              "
 lenguaje_Menu(91) = "&" & "Crear Eventos:             "
 lenguaje_Menu(92) = "&" & "Modificar"
 lenguaje_Menu(93) = "No hacer nada *"
 lenguaje_Menu(94) = "Apagar el Equipo"
 lenguaje_Menu(95) = "Apagar y reiniciar el equipo"
 lenguaje_Menu(96) = "Anular el Apagado de equipo"
 lenguaje_Menu(97) = "Equipo que se apagara / reiniciara / anulara"
 lenguaje_Menu(98) = "Establecer el tiempo de espera de apagado."
 lenguaje_Menu(99) = "Comentario de apagado máximo, 127 caracteres"
 lenguaje_Menu(100) = "Forzar el cierre de todas las aplicaciones sin advertir"
 lenguaje_Menu(101) = "Ingrese descripción máximo, 127 caracteres"
 lenguaje_Menu(102) = "&" & "Sin dialogo..."
 lenguaje_Menu(103) = "&" & "Tiempo ="
 lenguaje_Menu(104) = "encendido."
 
lenguaje_Menu(105) = "Hora:           programada"
lenguaje_Menu(106) = "Tipo:           Entrada o Salida"
lenguaje_Menu(107) = "Filtro:         Entrada o Salida o Aleatorio"
lenguaje_Menu(108) = "Intervalo:      Entrada"
lenguaje_Menu(109) = "Intervalo:      Salida"
lenguaje_Menu(110) = "Dias:           lunes"
lenguaje_Menu(111) = "Dias:           martes"
lenguaje_Menu(112) = "Dias:           miercoles"
lenguaje_Menu(113) = "Dias:           jueves"
lenguaje_Menu(114) = "Dias:           viernes"
lenguaje_Menu(115) = "Dias:           Sabados"
lenguaje_Menu(116) = "Dias:           Domingos"
lenguaje_Menu(117) = "Comentarios:    Entada"
lenguaje_Menu(118) = "Comentarios:    Salida"
lenguaje_Menu(119) = "Auto:           Apagar Encender etc*"
 
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'rutas de Archivos
 '################################################################
 lenguaje_Menu(120) = "Historial de Rutas"
 lenguaje_Menu(121) = "&" & "Historial de Archivos definidos"
 lenguaje_Menu(122) = "&" & "Cancelar"
 lenguaje_Menu(123) = "&" & "Cargar"
 lenguaje_Menu(124) = "&" & "Borrar Selección"
 lenguaje_Menu(125) = "&" & "Borrar Todo"
 lenguaje_Menu(126) = "&" & "Usar Archivo"
 lenguaje_Menu(127) = "&" & "Des Usar Archivo"
 lenguaje_Menu(128) = "&" & "Aceptar"
 lenguaje_Menu(129) = "Quieres utilizar este archivo con  Timeline" & " "
 lenguaje_Menu(130) = "¿ Quieres eliminar el Archivo usado de Memoria ?"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'Inicar Windows
 '################################################################
 lenguaje_Menu(131) = "Iniciar con Windows Automaticamente"
 lenguaje_Menu(132) = "&" & "¿ Arrancar con Windows ?"
 lenguaje_Menu(133) = "&" & "Arrancar"
 lenguaje_Menu(134) = "&" & "No Arrancar"
 lenguaje_Menu(135) = "&" & "Aceptar"
 lenguaje_Menu(136) = "Cuando Inicie o Reinicie Windows Virtual Martin temporize Arrancara con Windows"
 lenguaje_Menu(137) = " Hubo un error, Para Iniciar Con Windows S.O"
 lenguaje_Menu(138) = "Se elimino el Arranque Automatico en Windows S.O"
 
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'circuito impreso
 '################################################################
 lenguaje_Menu(139) = "Circuito Electrónico"
 lenguaje_Menu(140) = "Esqemas"
 lenguaje_Menu(141) = "&" & "Imprimir"
 lenguaje_Menu(142) = "&" & "Aceptar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'estados de la ventana principal
 '################################################################
 lenguaje_Menu(143) = "Estado Grafico de la Ventana Principal"
 lenguaje_Menu(144) = "Estado:"
 lenguaje_Menu(145) = "&" & "Cancelar"
 lenguaje_Menu(146) = "&" & "Aplicar"
 lenguaje_Menu(147) = "Ventana Restaurada"
 lenguaje_Menu(148) = "Ventana Minimizada"
 lenguaje_Menu(149) = "Ventana Maximizada"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'datos del creador
 '################################################################
 lenguaje_Menu(150) = "Personalizar datos"
 lenguaje_Menu(151) = "Datos del Creador de los Timbres"
 lenguaje_Menu(152) = "Aceptar"
 lenguaje_Menu(153) = "Nombre :"
 lenguaje_Menu(154) = "Segundo Nombre :"
 lenguaje_Menu(155) = "Apellido :"
 lenguaje_Menu(156) = "Segundo Apellido :"
 lenguaje_Menu(157) = "Dirección :"
 lenguaje_Menu(158) = "Segunda Dirección :"
 lenguaje_Menu(159) = "Localidad :"
 lenguaje_Menu(160) = "Pais :"
 lenguaje_Menu(161) = "Teléfono :"
 lenguaje_Menu(162) = "Celular :"
 lenguaje_Menu(163) = "Correo Electrónico :"
 lenguaje_Menu(164) = "Facebook :"
 lenguaje_Menu(165) = "Comentario General :"
 lenguaje_Menu(166) = "Cancelar"
 lenguaje_Menu(167) = "Limpiar"
 lenguaje_Menu(168) = "&Aceptar"
 lenguaje_Menu(169) = " ¿ Quieres Limpiar Todos los Datos en Pantalla ?"
 lenguaje_Menu(170) = "Los datos se guardaron en memoria con éxito"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'funciones
 '################################################################
 lenguaje_Menu(171) = "Funciones al Sistema"
 lenguaje_Menu(172) = "Funciones de Sistema Operables"
 lenguaje_Menu(173) = "comentarios:"
 lenguaje_Menu(174) = "No hacer nada *"
 lenguaje_Menu(175) = "Apagar el Equipo"
 lenguaje_Menu(176) = "Apagar y reiniciar el equipo"
 lenguaje_Menu(177) = "Anular el Apagado de equipo"
 lenguaje_Menu(178) = "Equipo que se apagara / reiniciara / anulara"
 lenguaje_Menu(179) = "Establecer el tiempo de espera de apagado."
 lenguaje_Menu(180) = "Comentario de apagado máximo, 127 caracteres"
 lenguaje_Menu(181) = "Forzar el cierre de todas las aplicaciones sin advertir"
 lenguaje_Menu(182) = "Ingrese descripción máximo, 127 caracteres"
 lenguaje_Menu(183) = "&" & "Sin dialogo..."
 lenguaje_Menu(184) = "&" & "Tiempo ="
 lenguaje_Menu(185) = "encendido."
 lenguaje_Menu(186) = "Cancelar"
 lenguaje_Menu(187) = "Aplicar"
 lenguaje_Menu(188) = "Sin Dialogo"
 lenguaje_Menu(189) = "Tiempo:"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'Impresor por Cantidad
 '################################################################
 lenguaje_Menu(190) = "Impresor por Cantidad"
 lenguaje_Menu(191) = "Copias:"
 lenguaje_Menu(192) = "-"
 lenguaje_Menu(193) = "+"
 lenguaje_Menu(194) = "Cancelar"
 lenguaje_Menu(195) = "Mandar a Imprimir las Copias"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'archivos en Memoria
 '################################################################
 lenguaje_Menu(196) = "Archivos actuales en la Memoria del Software"
 lenguaje_Menu(197) = "¿Existen Archivos en memoria que desea Hacer?"
 lenguaje_Menu(198) = "id"
 lenguaje_Menu(199) = "Hora"
 lenguaje_Menu(200) = "Tipo"
 lenguaje_Menu(201) = "Segundos"
 lenguaje_Menu(202) = "Comentario"
 lenguaje_Menu(203) = "Salir"
 lenguaje_Menu(204) = "Guardar y Salir"
 lenguaje_Menu(205) = "Cancelar"
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'crear modificar
 '################################################################
 lenguaje_Menu(206) = "titulo programa"
 lenguaje_Menu(207) = "Agregar Nuevo Evento"
 lenguaje_Menu(208) = "Funciones al sistema"
 lenguaje_Menu(209) = "Hora :"
 lenguaje_Menu(210) = "Tipo :"
 lenguaje_Menu(211) = "Intervalo :"
 lenguaje_Menu(212) = "Filtro :"
 lenguaje_Menu(213) = "Lunes"
 lenguaje_Menu(214) = "Martes"
 lenguaje_Menu(215) = "Miercoles"
 lenguaje_Menu(216) = "Jueves"
 lenguaje_Menu(217) = "Viernes"
 lenguaje_Menu(218) = "Sabados"
 lenguaje_Menu(219) = "Domingos"
 lenguaje_Menu(220) = "Comentario:"
 lenguaje_Menu(221) = "comentarios:"
 lenguaje_Menu(222) = "Cancelar"
 lenguaje_Menu(223) = "color de:Pintagrama"
 lenguaje_Menu(224) = "Crear"
 lenguaje_Menu(225) = "Modificar"
 lenguaje_Menu(226) = "Modificar Evento."
 lenguaje_Menu(227) = "Entrada"
 lenguaje_Menu(228) = "Salida"
 lenguaje_Menu(229) = "Solo Hora"
 lenguaje_Menu(230) = "Hora y Dia"
 lenguaje_Menu(231) = "Quieres Aplicar las Modificaciones del Evento."
 
 lenguaje_Menu(232) = "Personalizar Idioma"
 lenguaje_Menu(233) = "Archivo"
 lenguaje_Menu(234) = "Valor"
 lenguaje_Menu(235) = "Renombrar"
 lenguaje_Menu(236) = "Idioma"
 lenguaje_Menu(237) = "Cancelar"
 lenguaje_Menu(238) = "Cargar Idioma"
 lenguaje_Menu(239) = "Guardar Archivo"
 lenguaje_Menu(240) = "Aplicar Idioma"
 
 lenguaje_Menu(241) = "Pizarrón de Horarios Programados."
 lenguaje_Menu(242) = "Pizarrón de Tipo Entrada o Salida."
 lenguaje_Menu(243) = "Pizarrón de Tiempo en segundos."
 lenguaje_Menu(244) = "Pizarrón de Comentarios."
 
 lenguaje_Menu(245) = "enero"
 lenguaje_Menu(246) = "Febrero"
 lenguaje_Menu(247) = "Marzo"
 lenguaje_Menu(248) = "Abril"
 lenguaje_Menu(249) = "Mayo"
 lenguaje_Menu(250) = "Junio"
 lenguaje_Menu(251) = "Julio"
 lenguaje_Menu(252) = "Agosto"
 lenguaje_Menu(253) = "Septiembre"
 lenguaje_Menu(254) = "Oct / Nob / Dic"
 lenguaje_Menu(255) = "Ir al mes actual"
 
 lenguaje_Menu(256) = "Visor de Eventos Programados Actualmente"
 lenguaje_Menu(257) = "id"
 lenguaje_Menu(258) = "Hora"
 lenguaje_Menu(259) = "Tipo"
 lenguaje_Menu(260) = "Segundos"
 lenguaje_Menu(261) = "Comentario"
 lenguaje_Menu(262) = "Imprimir"
 lenguaje_Menu(263) = "Imprimir Más"
 lenguaje_Menu(264) = "Evento"
 
 lenguaje_Menu(265) = "Añadir comentarios"
 lenguaje_Menu(266) = "Comentario"
 lenguaje_Menu(267) = "Añadir"
 lenguaje_Menu(268) = "Cargar"
 lenguaje_Menu(269) = "Cancelar"
 lenguaje_Menu(270) = "Eliminar Seleciónado"
 lenguaje_Menu(271) = "Eliminar Todo"
 lenguaje_Menu(272) = "Guardar"
 lenguaje_Menu(273) = "¿Quieres eliminar este Comentario de la Lista?"
 lenguaje_Menu(274) = "¿Quieres eliminar Todos los Comentarios?"
 lenguaje_Menu(275) = "tiempo"
 lenguaje_Menu(276) = "Ingrese una Opción del Sistema"
 
 '################################################################
 'define las opciones de lenguage dentro del formulario
 'Acerca de:
 '################################################################
 
  lenguaje_Menu(277) = "ver los derechos de este software en el Ambito Legal."
  lenguaje_Menu(278) = "ver quienes participaron en Virtual Martin Temporize v1.7"
  lenguaje_Menu(279) = "Versión:"
  lenguaje_Menu(280) = "La información del sistema no está disponible en este momento"
  lenguaje_Menu(281) = ".:: Acerca de Virtual Martin Temporize v1.7 ::."
  lenguaje_Menu(282) = "Compilado"
  lenguaje_Menu(283) = "software pensado y creado "
  lenguaje_Menu(284) = "Programación"
  lenguaje_Menu(285) = "Diseño gráfico"
  lenguaje_Menu(286) = "Ide 's"
  lenguaje_Menu(287) = "Estructuras"
  lenguaje_Menu(288) = "Estadísticas"
  lenguaje_Menu(289) = "Análisis"
  lenguaje_Menu(290) = "Herramienta Pizarrón"
  lenguaje_Menu(291) = "Herramienta Generador dinámico de Horarios"
  lenguaje_Menu(292) = "Herramienta Meses Virtuales"
  lenguaje_Menu(293) = "Comparación"
  lenguaje_Menu(294) = "Artilugios gráficos para API"
  lenguaje_Menu(295) = "Entrada y Salida de Archivos"
  lenguaje_Menu(296) = "Algoritmos"
  lenguaje_Menu(297) = "Traducción y Idiomas por"
  lenguaje_Menu(298) = "Traductor de Google (c)"
  lenguaje_Menu(299) = "Librerías y cabeceras de I/0 "
  lenguaje_Menu(300) = "Microsoft Corporation (c)"
  lenguaje_Menu(301) = "Dinamismo y estructurado"
  lenguaje_Menu(302) = ".exe"
  lenguaje_Menu(303) = "Tipo de versión"
  lenguaje_Menu(304) = "usuarios en general"
  lenguaje_Menu(305) = "Advertencia este programa no esta protegido por leyes de derechos de autor ni otros tratados internacionales la reproducción o distribución no autorizada de este programa o de cualquier parte del mismo no da a responsabilidades sibiles ni criminales y no serán perseguidas"
  lenguaje_Menu(306) = "&Info. del sistema..."
  lenguaje_Menu(307) = "&Aceptar"
  lenguaje_Menu(308) = "Recursos"
  lenguaje_Menu(309) = "Autores"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Donativos
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(310) = "Virtual Martin Temporize v1.7"
  lenguaje_Menu(311) = "para cumplir mi sueño de ir a EE:UU"
  lenguaje_Menu(312) = "Amo mucho a EE:UU"
  lenguaje_Menu(313) = "con cuenta propia..."
  lenguaje_Menu(314) = "Con tarjetas de créditos"
  lenguaje_Menu(315) = "&Colaborar"
  lenguaje_Menu(316) = "&Aceptar"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Calendario
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(317) = "Virtual Martin temporize: Calendario"
  lenguaje_Menu(318) = "&Calendario Grafico"
  lenguaje_Menu(319) = "&Ir a la fecha de Hoy"
  lenguaje_Menu(320) = "&Salir"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Utilizar Timbre Manualmente
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(321) = "Utilizar Timbre Manualmente"
  lenguaje_Menu(322) = "&Encendido"
  lenguaje_Menu(323) = "&Apagado"
  lenguaje_Menu(324) = "Estado: Apagado"
  lenguaje_Menu(325) = "Estado: Encendido"
  lenguaje_Menu(326) = "&Aceptar"
  lenguaje_Menu(327) = "Led Que Muestra el Estado del Timbre"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'frmTimbre
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(328) = "Evento en ejecución"
  lenguaje_Menu(329) = "El Timbre se esta ejecutando..."
  lenguaje_Menu(330) = "Comentarios"
  lenguaje_Menu(331) = "Cargando..."
  lenguaje_Menu(332) = "Lunes"
  lenguaje_Menu(333) = "Martes"
  lenguaje_Menu(334) = "Miercoles"
  lenguaje_Menu(335) = "Jueves"
  lenguaje_Menu(336) = "Viernes"
  lenguaje_Menu(337) = "Sabados"
  lenguaje_Menu(338) = "domingos"
  lenguaje_Menu(339) = "DIAS"
  lenguaje_Menu(340) = "Solo Hora."
  lenguaje_Menu(341) = "Timpo Total :"
  lenguaje_Menu(342) = "Timpo Trascurrido :"
  lenguaje_Menu(343) = "Timpo Restante :"
  lenguaje_Menu(344) = "Hora que se Activo :"
  lenguaje_Menu(345) = "Tipo :"
  lenguaje_Menu(346) = "&Cerrar"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'frmReloj Digital
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(347) = "Virtual Martin Temporize:  Reloj Digital"
  lenguaje_Menu(348) = "Reloj del Sistema"
  lenguaje_Menu(349) = "fecha:"
  lenguaje_Menu(350) = "hs"
  lenguaje_Menu(351) = "&Aceptar"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  'pin de datos
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
  lenguaje_Menu(352) = "Puerto de Salida"
  lenguaje_Menu(353) = "Usted Tiene que tener conocimiento antes de realizar algun cambio Aquí."
  lenguaje_Menu(354) = "&Salida 5v"
  lenguaje_Menu(355) = "Pin 1"
  lenguaje_Menu(356) = "Pin 2"
  lenguaje_Menu(357) = "Pin 3"
  lenguaje_Menu(358) = "Pin 4"
  lenguaje_Menu(359) = "Pin 5"
  lenguaje_Menu(360) = "Pin 6"
  lenguaje_Menu(361) = "Pin 7"
  lenguaje_Menu(362) = "Pin 8"
  lenguaje_Menu(363) = "&Cancelar"
  lenguaje_Menu(364) = "&Normal"
  lenguaje_Menu(365) = "&Aceptar"
  ''''''''''''''''''''''''''''''''''
  ' ficheros y configuración
  ''''''''''''''''''''''''''''''''''
  lenguaje_Menu(366) = "¿ Quieres Guardar los Cambios ?"
  lenguaje_Menu(367) = " Abrir Archivo"
  lenguaje_Menu(368) = "Timeline USB"
  lenguaje_Menu(369) = "Todos los Archivos"
  lenguaje_Menu(370) = "el Año minimo es 1753"
  lenguaje_Menu(371) = "Deseas eliminar todos los timbres con eventos Programados"
  lenguaje_Menu(372) = "Opciones de Eliminación"
  lenguaje_Menu(373) = "Deseas eliminar este timbre con el Evento"
  lenguaje_Menu(374) = "Opciones de Eliminación"
  lenguaje_Menu(375) = "para poder liminar seleccione un evento"
  lenguaje_Menu(376) = " v1.7"
  lenguaje_Menu(377) = " Guardar Archivo"
  lenguaje_Menu(378) = "todos los Archivos"
  lenguaje_Menu(379) = "Nuevo"
  lenguaje_Menu(380) = "No le asignaste un nombre de archivo"
  lenguaje_Menu(381) = "Aleatorio"
  lenguaje_Menu(382) = "seg. "
  lenguaje_Menu(383) = "D LMMJV  S"
  lenguaje_Menu(384) = " ------------------------------------------"
  lenguaje_Menu(385) = " ------------------------------------------"
 Dim int_c As Integer
 Dim cargar As String

Open App.Path & "\" & Lenguage.sel For Input As 1
 Do While Not EOF(1)
  
       Line Input #1, cargar
       lenguaje_Menu(int_c) = cargar
       int_c = int_c + 1
       Loop
       Close #1
       int_c = 0
nose:
End Sub

Private Sub cmdcargar()
On Error GoTo nose
Dim cargar As String
Open App.Path & "\idiomas\inicio.inf" For Input As 1
 Do While Not EOF(1)
  Line Input #1, Lenguage.sel
 Loop
 Close #1
nose:
End Sub


