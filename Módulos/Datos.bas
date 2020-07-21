Attribute VB_Name = "Datos"
'***************************************************************************
'*
'*
'* Coleciónes en Memoria para Timeline
'*
'*
'***************************************************************************
Public Type importancia

 noImportante As Coleccionados.coleccionDatos
 espocoimportante As Coleccionados.coleccionDatos
 muyimportante As Coleccionados.coleccionDatos
 destacado As Coleccionados.coleccionDatos
 nose As Coleccionados.coleccionDatos
 eseventoestrella As Coleccionados.coleccionDatos
 importanteparalosalumnos As Coleccionados.coleccionDatos

End Type

Public Type EventoDato

 xcomentario As Coleccionados.coleccionDatos
 xtipo       As Coleccionados.coleccionDatos
 xintervalo  As Coleccionados.coleccionDatos
 xlunes      As Coleccionados.coleccionDatos
 xmartes     As Coleccionados.coleccionDatos
 xmiercoles  As Coleccionados.coleccionDatos
 xjueves     As Coleccionados.coleccionDatos
 xviernes    As Coleccionados.coleccionDatos
 xsabados    As Coleccionados.coleccionDatos
 xdomingos   As Coleccionados.coleccionDatos
 xautomatico As Coleccionados.coleccionDatos
 enero       As Coleccionados.coleccionDatos
 xfebrero    As Coleccionados.coleccionDatos
 xmarzo      As Coleccionados.coleccionDatos
 xabril      As Coleccionados.coleccionDatos
 xmayo       As Coleccionados.coleccionDatos
 xjunio      As Coleccionados.coleccionDatos
 xjulio      As Coleccionados.coleccionDatos
 xagosto     As Coleccionados.coleccionDatos
 xsetiembre  As Coleccionados.coleccionDatos
 xoctubre    As Coleccionados.coleccionDatos
 xnobiembre  As Coleccionados.coleccionDatos
 xdiciembre  As Coleccionados.coleccionDatos
 xfiltro     As Coleccionados.coleccionDatos

End Type

Public Type segundos

 evento As EventoDato
 segundos As Coleccionados.coleccionDatos

End Type

Public Type horas

 hora As Coleccionados.coleccionDatos
 minutos As Coleccionados.coleccionDatos
 seg As segundos
 setsegundo As Coleccionados.coleccionDatos

End Type

Public Type DiasDeLaSemana

 dLunes As horas
 dMartes As horas
 dMiercoles As horas
 dJueves As horas
 dViernes As horas
 dSabados As horas
 dDomingo As horas
 fecha As horas
 mes As horas
 anio As horas
 importa As importancia

End Type

Public Type dias

 d1 As DiasDeLaSemana
 d2 As DiasDeLaSemana
 d3 As DiasDeLaSemana
 d4 As DiasDeLaSemana
 d5 As DiasDeLaSemana
 d6 As DiasDeLaSemana
 d7 As DiasDeLaSemana
 d8 As DiasDeLaSemana
 d9 As DiasDeLaSemana
 d10 As DiasDeLaSemana
 d11 As DiasDeLaSemana
 d12 As DiasDeLaSemana
 d13 As DiasDeLaSemana
 d14 As DiasDeLaSemana
 d15 As DiasDeLaSemana
 d16 As DiasDeLaSemana
 d17 As DiasDeLaSemana
 d18 As DiasDeLaSemana
 d19 As DiasDeLaSemana
 d20 As DiasDeLaSemana
 d21 As DiasDeLaSemana
 d22 As DiasDeLaSemana
 d23 As DiasDeLaSemana
 d24 As DiasDeLaSemana
 d25 As DiasDeLaSemana
 d26 As DiasDeLaSemana
 d27 As DiasDeLaSemana
 d28 As DiasDeLaSemana
 d29 As DiasDeLaSemana
 d30 As DiasDeLaSemana
 d31 As DiasDeLaSemana

End Type

Public Type dato

 Nombre As Coleccionados.coleccionDatos
 segundonombre As Coleccionados.coleccionDatos
 apellido As Coleccionados.coleccionDatos
 segundoapellido As Coleccionados.coleccionDatos
 direccion As Coleccionados.coleccionDatos
 segundadireccion As Coleccionados.coleccionDatos
 localidad As Coleccionados.coleccionDatos
 Pais As Coleccionados.coleccionDatos
 telefono As Coleccionados.coleccionDatos
 celular  As Coleccionados.coleccionDatos
 correoelectronico As Coleccionados.coleccionDatos
 facebook As Coleccionados.coleccionDatos
 comentariogeneral As Coleccionados.coleccionDatos

End Type

Public Type Años
 
 hora As horas
 segundos As segundos
 DiasDeLaSemana As DiasDeLaSemana
 dias As dias
 evento As EventoDato
 Datosp As dato
End Type



