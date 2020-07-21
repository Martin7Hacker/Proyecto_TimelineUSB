Attribute VB_Name = "OtrasF"
'***************************************************************************
'*
'*
'* Comparar Fechas con Timeline
'*
'*
'***************************************************************************
Option Explicit

Public Function ComparaFechas(ByVal F1 As Date, ByVal F2 As Date) As Boolean
 On Error GoTo nose
 ComparaFechas = False
 If (CDate(F1) < CDate(F2)) Then Exit Function
 ComparaFechas = True
nose:
End Function

Public Function LastError$()
On Error GoTo nose
 LastError$ = "Se ha generado el Código de Error No. " _
 & NumErr& & vbCrLf & vbCrLf & _
 "Motivo: " & MsgErr$
nose:
End Function

