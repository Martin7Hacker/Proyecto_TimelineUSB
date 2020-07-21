Attribute VB_Name = "ModImprimir"
'***************************************************************************
'*
'*
'* Imprimir Guardar Ficheros con Timeline
'*
'*
'***************************************************************************
'A esta función se le envía el control LV a imprimir
Public Sub Imprimir_ListView()
On Error GoTo nose
 Dim i As Integer, AnchoCol As Single, espacio As Integer, x As Integer
 AnchoCol = 0
 'Recorremos desde la primer columna hasta la última para almacenar el ancho total
 For i = 1 To frmVisorEventos.ListView1.ColumnHeaders.Count
 AnchoCol = AnchoCol + frmVisorEventos.ListView1.ColumnHeaders(i).Width
 Next
 espacio = 0
 'Encabezado de ejemplo
 Printer.Print lenguaje_Menu(384)
 Printer.Print " Timeline USB* " & lenguaje_Menu(376)
 Printer.Print lenguaje_Menu(385)
 Printer.Print
 'Imprime una línea
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 With frmVisorEventos.ListView1
 Printer.Print "ID"
 'Acá se imprimen los encabezados del ListView
 For i = 1 To .ColumnHeaders.Count
 espacio = espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
 Printer.Print frmVisorEventos.ListView1.ColumnHeaders(i).Text;
 Printer.CurrentX = espacio
 Next
 'Printer.Print frmVisorEventos.ListView1.ColumnHeaders(1).Text & frmVisorEventos.ListView1.ColumnHeaders(2).Text & frmVisorEventos.ListView1.ColumnHeaders(3).Text & frmVisorEventos.ListView1.ColumnHeaders(4).Text & frmVisorEventos.ListView1.ColumnHeaders(5).Text;
 Printer.Print
 'Imprime una línea
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 'Imprime Línea en blanco
 Printer.Print
 'Este bucle recorre los items y subitems del ListView  y los imprime
 For i = 1 To .ListItems.Count
 espacio = 0
 Printer.Print frmVisorEventos.ListView1.ListItems.Item(i).Text;
 'Recorremos las columnas
 For x = 1 To .ColumnHeaders.Count - 1
 espacio = espacio + CInt(.ColumnHeaders(x).Width * _
 Printer.ScaleWidth / AnchoCol)
 Printer.CurrentX = espacio
 Printer.Print frmVisorEventos.ListView1.ListItems.Item(i).SubItems(x);
 Next
 'Otro espacio en blanco
 Printer.Print
 Next
 End With
 Printer.Print
 'Imprime la línea de final de impresión
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 Printer.Print
 'Texto del pie
 Printer.Print "Timeline USB* - Martinsoft 2017"
 Printer.Print lenguaje_Menu(349) & Date & " - " & lenguaje_Menu(258) & ": " & Time
 Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
 Printer.Print
 Printer.Print "++++++++++++++++++++++++ Fin de la impresión ++++++++++++++++++++++++ "
 'Comenzamos la impresión
 Printer.EndDoc
nose:
End Sub

