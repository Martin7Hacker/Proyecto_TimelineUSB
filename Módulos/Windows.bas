Attribute VB_Name = "Windows"
'***************************************************************************
'*
'*
'* Windows con Timeline
'*
'*
'***************************************************************************
Option Explicit
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE1 = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const REG_SZ1 As Long = 1
Public Const REG_DWORD As Long = 4
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long
'Abre una clave
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
Alias "RegOpenKeyExA" _
       (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long) As Long
'Establece un valor de tipo cadena
Private Declare Function RegSetValueExString Lib "advapi32.dll" _
Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
         ByVal lpValueName As String, _
         ByVal Reserved As Long, _
         ByVal dwType As Long, _
         ByVal lpValue As String, _
         ByVal cbData As Long) As Long
'Establece un valor de tipo entero
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
         "RegSetValueExA" (ByVal hKey As Long, _
         ByVal lpValueName As String, _
         ByVal Reserved As Long, _
         ByVal dwType As Long, _
         lpValue As Long, _
         ByVal cbData As Long) As Long
'Elimina una clave
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias _
"RegDeleteKeyA" _
        (ByVal hKey As Long, _
         ByVal lpSubKey As String)
'Elimina un valor del registro
Private Declare Function RegDeleteValue& Lib "advapi32.dll" _
Alias "RegDeleteValueA" _
(ByVal hKey As Long, _
ByVal lpValueName As String)

Function EliminarClave(clave As Long, Nombre_clave As String)
 On Error GoTo nose
 Dim Ret As Long
 ' Eliminar
 Ret = RegDeleteKey(clave, Nombre_clave)
nose:
End Function

Function EliminarValor(clave As Long, _
 Nombre_clave As String, _
 Nombre_valor As String) As Boolean
 On Error GoTo nose
 Dim Ret As Long
 Dim Handle_clave As Long
 ' Abre la clave del registro indicada
 Ret = RegOpenKeyEx(clave, Nombre_clave, 0, KEY_ALL_ACCESS, Handle_clave)
 'si el valor de retorno es distinto de 0 es por que hubo un error
 If Ret <> 0 Then
 EliminarValor = False
 Exit Function
 End If
 'Elimina el valor del registro
 Ret = RegDeleteValue(Handle_clave, Nombre_valor)
 If Ret <> 0 Then
 EliminarValor = False
 Exit Function
 End If
 'Cierra la vlave del registro abierta
 RegCloseKey (Handle_clave)
 ' OK
 EliminarValor = True
nose:
End Function

Function EstablecerValor(clave As Long, _
 Nombre_clave As String, _
 Nombre_valor As String, _
 el_Valor As Variant, _
 Tipo_Valor As Long) As Boolean
 On Error GoTo nose
 Dim Ret As Long
 Dim Handle_clave As Long
 'Abre la clave del registro indicada
 Ret = RegOpenKeyEx(clave, Nombre_clave, 0, KEY_ALL_ACCESS, Handle_clave)
 'si el valor de retorno es distinto de 0 es por que hubo un error
 If Ret <> 0 Then
 EstablecerValor = False
 Exit Function
 End If
 'Establece el nuevo dato
 Ret = SetValueEx(Handle_clave, Nombre_valor, Tipo_Valor, el_Valor)
 If Ret <> 0 Then
 EstablecerValor = False
 Exit Function
 End If
 'cierra la clave abierta
 RegCloseKey (Handle_clave)
 'Ok
 EstablecerValor = True
nose:
 End Function
 
Private Function SetValueEx(ByVal Handle_clave As Long, _
 Nombre_valor As String, _
 Tipo As Long, _
 el_Valor As Variant) As Long
 On Error GoTo nose
 Dim Ret As Long
 Dim sValue As String
 Select Case Tipo
 ' Valor de tipo cadena
 Case REG_SZ
 sValue = el_Valor
 SetValueEx = RegSetValueExString(Handle_clave, _
 Nombre_valor, 0&, _
 Tipo, sValue, Len(sValue))
 'Valor Entero
 Case REG_DWORD
 Ret = el_Valor
 SetValueEx = RegSetValueExLong(Handle_clave, Nombre_valor, 0&, Tipo, Ret, 4)
 End Select
nose:
End Function
