VERSION 5.00
Begin VB.Form frmutilizarManual 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilizar Timbre Manualmente"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "frmutilizarManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "frmutilizarManual.frx":57E2
      ScaleHeight     =   735
      ScaleWidth      =   3615
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000C0&
         BorderWidth     =   10
         Height          =   375
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   165
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "Led Que Muestra el Estado del Timbre ."
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   -600
      ScaleHeight     =   0
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdencendido 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Encendido"
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
      MICON           =   "frmutilizarManual.frx":15798
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdapagado 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Apagado"
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
      MICON           =   "frmutilizarManual.frx":157B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAceptar 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmutilizarManual.frx":157D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmutilizarManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Detonador Manual de Timeline
'*
'*
'***************************************************************************

Private Sub cmdAceptar_Click()
On Error GoTo nose
 frmprograma.Enabled = True
 Unload Me
nose:
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdapagado_Click()
On Error GoTo nose
 frmprograma.cargarPuerto False
 Shape1.BackColor = &H80FF80
 Shape1.BorderColor = &HC000&
 Label1.Caption = lenguaje_Menu(324)
 frmtimbre.Finalizar ' Apaga todos los Puertos
 ActivarLed 1, 5
nose:
End Sub

Private Sub cmdapagado_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdencendido_Click()
On Error GoTo nose
 frmprograma.cargarPuerto True
 Shape1.BackColor = &HFF&
 Shape1.BorderColor = &HC0&
 Label1.Caption = lenguaje_Menu(325)
 puertof.disparar_bit ' Enciendo el Timbre
 ActivarLed 1, 6
nose:
End Sub

Private Sub cmdencendido_KeyPress(KeyAscii As Integer)
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
 cmdapagado_Click
 Call cargarIdioma
nose:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo nose
 cmdapagado_Click
 ActivarLed 1, 5
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmutilizarManual
nose:
End Sub

Private Sub cargarIdioma()
On Error GoTo nose
 Me.Caption = lenguaje_Menu(321)
 cmdencendido.Caption = lenguaje_Menu(322)
 cmdapagado.Caption = lenguaje_Menu(323)
 Label1.Caption = lenguaje_Menu(324)
 Label2.ToolTipText = lenguaje_Menu(327)
nose:
End Sub

Public Sub ActivarLed(ByVal muestro As Byte, ByVal recurso As Byte)
On Error GoTo nose
 frmprograma.StatusBar1.Panels(muestro).Picture = _
  frmprograma.StatusBar1.Panels(recurso).Picture
nose:
End Sub

