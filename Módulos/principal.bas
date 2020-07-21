Attribute VB_Name = "principal"
'***************************************************************************
'*
'*
'* Configuración prinicipal de Timeline
'*
'*
'***************************************************************************
Option Explicit
Global i&, j&, k&                             ' Contadores
Global Msg$, MsgErr$, NumErr&                 ' Variables de control de error
Global Cont%, opc%                            ' Otros contadores
Global Const Format_Fech1$ = "YYYYMMDD"       ' Formato de fecha 1
Global Const Format_Fech2$ = "DD/MM/YYYY"     ' Formato de fecha 2
Global Const Format_Fech3$ = "YYYY-MM-DD"     ' Formato de fecha 3
Global Const Format_Hour$ = "HH:MM:SS"        ' Formato de hora general AM/FM
Global Const Format_Money$ = "$#0.00"         ' Formato moneda corta
Global Const Format_Real1$ = "#0.00"          ' Formato numerico con dos decimales
Global Const Format_Real2$ = "#0.0000"        ' Formato numerico con cuatro decimales
Global VgTime1$, VgTime2$                     ' Variables de captura de Horas
Global VgDate1$, VgDate2$                     ' Variables de captura de fechas
Global VgDateTime1$, VgDateTime2$             ' Variables de fecha/hora completa

' Enumeracion de seleccion de reportes
Public Enum ShowReport

  RptFactura      ' Reporte de Factura de Venta
  RptFacturaRes   ' Reporte de Resumen de Factura de Venta
  RptInventario   ' Reporte de Inventario de Productos

End Enum

Public RptShower As ShowReport
Public Skn$

' Funcion principal de la aplicación
Public Sub Main()
On Error GoTo nose
 Inicio.Inicio 'iniciar el idioma
 If (App.PrevInstance) Then
 MsgBox LoadResString(19), vbCritical, LoadResString(1)
 End
 End If
 Call InitCommonControlsVB
 frmprograma.Show
 Exit Sub
Falla:
 MsgBox LastError$, vbCritical, LoadResString(1)
 End
nose:
End Sub
