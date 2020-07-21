Attribute VB_Name = "Inicio"
'***************************************************************************
'*
'*
'* Iniciar Programa con Timeline
'*
'*
'***************************************************************************

Public Sub Inicio() 'inicio del programa para que se pueda
                     'configurar previamente y luego
                     'armarse para despues crearlo
                     'graficamente
On Error GoTo nose
 Lenguage.definir_lenguage_opciones 'carga el lenguage previo
 frmprograma.Show                   'carga el programa
nose:
End Sub
