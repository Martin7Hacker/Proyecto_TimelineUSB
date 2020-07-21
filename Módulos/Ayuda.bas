Attribute VB_Name = "Ayuda"
'***************************************************************************
'*
'*
'* Mostrar Ayuda en Timeline
'*
'*
'***************************************************************************
Option Explicit
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu _
As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd _
As Long, ByVal bRevert As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock _
As Long) As Long
Public Langs As String
Public FilterB As String
Public AllSupport As String
Public ClrString As String
Public Declare Function DeleteFile Lib "kernel32" _
Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" _
Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String _
, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const HH_DISPLAY_TOC = &H1
Public StopClose As Boolean

Public Sub DisableX(TheForm As Form)
On Error GoTo nose
 Dim lngMenu As Long
 lngMenu = GetSystemMenu(TheForm.hwnd, False)
 DeleteMenu lngMenu, 6, MF_BYPOSITION
nose:
End Sub

Public Sub HHShowContents(lhWnd As Long)
On Error GoTo nose
 HTMLHelp lhWnd, App.Path & "\Ayuda.chm" _
 & "", HH_DISPLAY_TOC, 0
nose:
End Sub


