VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmArranque 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de rutas de archivos"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmArranque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -120
      Picture         =   "frmArranque.frx":57E2
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   8
      Top             =   0
      Width           =   8160
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Historial de Archivos definidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog cdgAbrir 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7575
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":148A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcargar 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cargar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":148C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdborrar 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Borrar Selección"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":148DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdborrartodo 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Borrar Todo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":148F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdusar 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Usar Archivo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":14914
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdaceptar 
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArranque.frx":14930
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Historial de Archivos definidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7515
   End
End
Attribute VB_Name = "frmArranque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Iniciar Archivo con el  programa Timeline
'* Historial de Rutas de Archivo
'*
'***************************************************************************
Private Sub cmdAceptar_Click()
On Error GoTo nose
 externosF.guardar_Archivo_Externo
 frmprograma.Enabled = True
 Unload Me
nose:
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdborrar_Click()
On Error GoTo nose
If Not (List1.ListIndex = -1) Then
 Select Case MsgBox(lenguaje_Menu(130) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   List1.RemoveItem (List1.ListIndex)
 End Select
End If
nose:
End Sub

Private Sub cmdborrar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdborrartodo_Click()
On Error GoTo nose
 Select Case MsgBox("Quieres eliminar todos los Archivos definidos en el Historial" _
 , vbYesNo + vbInformation)
  Case (vbYes)
  List1.Clear
 End Select
nose:
End Sub

Private Sub cmdborrartodo_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo nose
 Unload Me
nose:
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdcargar_Click()
On Error GoTo nose
With cdgAbrir
 If .CancelError = False Then
 .DialogTitle = "Timeline USB*: Cargar Archivo"
 .Filter = "Timeline USB* (*.tml)|*.tml|todos los Archivos (*.*)|*.*|"
 .ShowOpen
 If .FileName = "" Then
 MsgBox "Tienes que seleccionar un Archivo para poder Abrirlo", vbInformation
 End If
 If .FileName <> "" Then
 List1.AddItem .FileName
 End If
 End If
End With
nose:
End Sub

Private Sub cmdcargar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
salir_op KeyAscii
nose:
End Sub

Private Sub cmdusar_Click()
On Error GoTo nose
If cmdusar.Caption = Lenguage.lenguaje_Menu(126) Then
 If Not (List1.ListIndex = -1) Then
 MsgBox Lenguage.lenguaje_Menu(129) & "" & List1.List(List1.ListIndex)
 externosF.xselecionado = List1.List(List1.ListIndex)
 externosF.guardar_selecionado
 End If
 cmdusar.Caption = Lenguage.lenguaje_Menu(127)
 ElseIf cmdusar.Caption = Lenguage.lenguaje_Menu(127) Then
 Select Case MsgBox(Lenguage.lenguaje_Menu(130), vbYesNo + vbInformation)
  Case (vbYes)
   externosF.xselecionado = ""
   externosF.guardar_selecionado
   End Select
   cmdusar.Caption = Lenguage.lenguaje_Menu(126)
 End If
nose:
End Sub

Private Sub cmdusar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
 Me.Icon = frmprograma.Icon
 externosF.Abrir_Archivo_Externo
 cargar_lenguage ' carga el lenguage
 Label2.Caption = Label1.Caption
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmArranque
nose:
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cargar_lenguage()
On Error GoTo nose
 Me.Caption = Lenguage.lenguaje_Menu(120)
 Label1.Caption = Lenguage.lenguaje_Menu(121)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(122)
 cmdcargar.Caption = Lenguage.lenguaje_Menu(123)
 cmdborrar.Caption = Lenguage.lenguaje_Menu(124)
 cmdborrartodo.Caption = Lenguage.lenguaje_Menu(125)
 cmdusar.Caption = Lenguage.lenguaje_Menu(126)
 cmdAceptar.Caption = Lenguage.lenguaje_Menu(128)
nose:
End Sub
