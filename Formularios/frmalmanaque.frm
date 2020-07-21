VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmalmanaque 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Martin temporize: Calendario"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmalmanaque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Salir"
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
      MICON           =   "frmalmanaque.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdFechaHoy 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ir a la fecha de Hoy"
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
      MICON           =   "frmalmanaque.frx":57FE
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
      Left            =   -1200
      Picture         =   "frmalmanaque.frx":581A
      ScaleHeight     =   420
      ScaleWidth      =   8130
      TabIndex        =   2
      Top             =   0
      Width           =   8160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Calendario Grafico"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   1620
      TabIndex        =   1
      Top             =   4680
      Width           =   1620
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   4710
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   8308
      _Version        =   393216
      ForeColor       =   4210752
      BackColor       =   -2147483646
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483646
      StartOfWeek     =   107216898
      TitleBackColor  =   12582912
      TitleForeColor  =   -2147483634
      TrailingForeColor=   12583104
      CurrentDate     =   40784
   End
End
Attribute VB_Name = "frmalmanaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Calendario Grafico para el programa Timeline
'*
'*
'***************************************************************************
Private Sub cmdsalir_Click()
 On Error GoTo nose
 frmprograma.Enabled = True
 Unload Me
nose:
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdFechaHoy_Click()
On Error GoTo nose
 MonthView1.Value = Date
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
 On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmalmanaque
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
 Me.Icon = frmprograma.Icon
 cmdFechaHoy_Click
 Call cargarIdioma
nose:
End Sub



Private Sub MonthView1_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub
Private Sub cargarIdioma()
On Error GoTo nose
Me.Caption = lenguaje_Menu(317)
Label1.Caption = lenguaje_Menu(318)
cmdFechaHoy.Caption = lenguaje_Menu(319)
cmdSalir.Caption = lenguaje_Menu(320)
nose:
End Sub
