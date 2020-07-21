VERSION 5.00
Begin VB.Form frmDatos 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personalizar"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAceptar 
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      Top             =   6720
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
      MICON           =   "frmDatos.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdlimpiar 
      Height          =   375
      Left            =   1080
      TabIndex        =   29
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Limpiar"
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
      MICON           =   "frmDatos.frx":57FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6720
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
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDatos.frx":581A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "frmDatos.frx":5836
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   26
      Top             =   0
      Width           =   8160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Perzonalizar Datos"
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
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1155
      Index           =   12
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   5400
      Width           =   4815
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   11
      Left            =   960
      TabIndex        =   23
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   10
      Left            =   1560
      TabIndex        =   21
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   9
      Left            =   840
      TabIndex        =   19
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   8
      Left            =   960
      TabIndex        =   17
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   7
      Left            =   840
      TabIndex        =   15
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   6
      Left            =   960
      TabIndex        =   13
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   5
      Left            =   1680
      TabIndex        =   11
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   4
      Left            =   960
      TabIndex        =   9
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtdato 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtdato 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario General :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   5085
      Width           =   1485
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facebook :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   4725
      Width           =   810
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electrónico :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   4365
      Width           =   1395
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Celular :"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   4005
      Width           =   570
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   3645
      Width           =   720
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pais :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   3285
      Width           =   390
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2925
      Width           =   780
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Segunda Dirección :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2205
      Width           =   765
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Segundo Apellido :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1845
      Width           =   1335
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1485
      Width           =   645
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Segundo Nombre :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   645
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Datos de los creadores de Timeline
'*
'*
'***************************************************************************

Private Sub cmdAceptar_Click()
On Error GoTo nose
 guardar
 frmprograma.Enabled = True
 Unload Me
nose:
End Sub

Private Sub guardar()
On Error GoTo nose
 abrirF.xnombre = txtdato(0).Text
 abrirF.xnombre2 = txtdato(1).Text
 abrirF.xapellido = txtdato(2).Text
 abrirF.xapellido2 = txtdato(3).Text
 abrirF.xdireccion = txtdato(4).Text
 abrirF.xdireccion2 = txtdato(5).Text
 abrirF.xlocalidad = txtdato(6).Text
 abrirF.xPais = txtdato(7).Text
 abrirF.xtelefono = txtdato(8).Text
 abrirF.xcel = txtdato(9).Text
 abrirF.xemail = txtdato(10).Text
 abrirF.xfacebook = txtdato(11).Text
 abrirF.xcomentario_general = txtdato(12).Text
 MsgBox Lenguage.lenguaje_Menu(170), vbInformation
nose:
End Sub

Private Sub mostrar()
On Error GoTo nose
 txtdato(0).Text = abrirF.xnombre
 txtdato(1).Text = abrirF.xnombre2
 txtdato(2).Text = abrirF.xapellido
 txtdato(3).Text = abrirF.xapellido2
 txtdato(4).Text = abrirF.xdireccion
 txtdato(5).Text = abrirF.xdireccion2
 txtdato(6).Text = abrirF.xlocalidad
 txtdato(7).Text = abrirF.xPais
 txtdato(8).Text = abrirF.xtelefono
 txtdato(9).Text = abrirF.xcel
 txtdato(10).Text = abrirF.xemail
 txtdato(11).Text = abrirF.xfacebook
 txtdato(12).Text = abrirF.xcomentario_general
nose:
End Sub

Private Sub cmdAceptar_KeyPress(KeyAscii As Integer)
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

Private Sub cmdlimpiar_Click()
On Error GoTo nose
 Select Case MsgBox(Lenguage.lenguaje_Menu(169) _
 , vbYesNo + vbInformation)
 Case (vbYes)
  Dim l As Byte
  For l = 0 To 12
   txtdato(l).Text = ""
  Next
 End Select
nose:
End Sub

Private Sub cmdlimpiar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
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
 mostrar
 Me.Icon = frmprograma.Icon
 cargar_lenguage ' cargar lenguage
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmDatos
nose:
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cargar_lenguage()
On Error GoTo nose
 Me.Caption = Lenguage.lenguaje_Menu(150)
 Label1.Caption = Lenguage.lenguaje_Menu(151)
 cmdAceptar.Caption = Lenguage.lenguaje_Menu(152)
 labdatos(0).Caption = Lenguage.lenguaje_Menu(153)
 labdatos(1).Caption = Lenguage.lenguaje_Menu(154)
 labdatos(2).Caption = Lenguage.lenguaje_Menu(155)
 labdatos(3).Caption = Lenguage.lenguaje_Menu(156)
 labdatos(4).Caption = Lenguage.lenguaje_Menu(157)
 labdatos(5).Caption = Lenguage.lenguaje_Menu(158)
 labdatos(6).Caption = Lenguage.lenguaje_Menu(159)
 labdatos(7).Caption = Lenguage.lenguaje_Menu(160)
 labdatos(8).Caption = Lenguage.lenguaje_Menu(161)
 labdatos(9).Caption = Lenguage.lenguaje_Menu(162)
 labdatos(10).Caption = Lenguage.lenguaje_Menu(163)
 labdatos(11).Caption = Lenguage.lenguaje_Menu(164)
 labdatos(12).Caption = Lenguage.lenguaje_Menu(165)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(166)
 cmdlimpiar.Caption = Lenguage.lenguaje_Menu(167)
 cmdAceptar.Caption = Lenguage.lenguaje_Menu(168)
nose:
End Sub
