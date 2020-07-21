Attribute VB_Name = "archivoF"
'***************************************************************************
'*
'*
'* M�dulo para Abrir el Archivo en Timeline
'*
'*
'***************************************************************************
Option Explicit
'Funciones Api para leer, abrir, cerrar y escribir en el registro
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, _
ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, _
lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal _
lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
'Constantes varias para las funciones Api del registro
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0

Public Sub CrearAsociacion(RutadelExe As String, EXT As String _
, Descripci�n As String, LibreriaIcono As String)
On Error GoTo nose
 Dim sPath As String
 sPath = App.Path & "\" & App.EXEName & " %1"
 CreateNewKey "." & EXT, HKEY_CLASSES_ROOT
 SetKeyValue "." & EXT, "", RutadelExe, REG_SZ
 CreateNewKey RutadelExe & "\shell\open\command", HKEY_CLASSES_ROOT
 CreateNewKey RutadelExe & "\DefaultIcon", HKEY_CLASSES_ROOT
 SetKeyValue RutadelExe, "", Descripci�n, REG_SZ
 SetKeyValue RutadelExe & "\shell\open\command", "", sPath, REG_SZ
 SetKeyValue RutadelExe & "\DefaultIcon", "", LibreriaIcono, REG_SZ
nose:
End Sub

Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
 On Error GoTo nose
 'handle para la nueva clave
 Dim hKey As Long
 'retorno de la funci�n RegCreateKeyEx
 Dim r As Long
 r = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, _
 REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, r)
 Call RegCloseKey(hKey)
nose:
End Sub

Public Sub SetKeyValue(sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long)
On Error GoTo nose
 'retorno de funcion SetValueEx
 Dim r As Long
 'handle
 Dim hKey As Long
 'Abrimos la clave especifica
 r = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, hKey)
 r = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
 'cerramos la clave abierta pasandole el handle
 Call RegCloseKey(hKey)
nose:
End Sub

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, _
 lType As Long, vValue As Variant) As Long
 On Error GoTo nose
 Dim nValue As Long, sValue As String
 Select Case lType
 'Valor de Cadena
 Case REG_SZ
 sValue = vValue & Chr$(0)
 'Establecemos el valor en el registro
 SetValueEx = RegSetValueExString(hKey, sValueName, _
 0&, lType, sValue, Len(sValue))
 'Valor entero
 Case REG_DWORD
 nValue = vValue
 'Establecer el valor en el registro
 SetValueEx = RegSetValueExLong(hKey, sValueName, _
 0&, lType, nValue, 4)
 End Select
nose:
End Function
