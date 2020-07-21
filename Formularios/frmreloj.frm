VERSION 5.00
Begin VB.Form frmreloj 
   BackColor       =   &H80000002&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Virtual Martin Temporize:  Reloj Digital"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   Icon            =   "frmreloj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdaceptar 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmreloj.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Picture         =   "frmreloj.frx":57FE
         ScaleHeight     =   975
         ScaleWidth      =   3855
         TabIndex        =   3
         Top             =   480
         Width           =   3855
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   75
         End
         Begin VB.Label lab_reloj 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   3855
         End
      End
      Begin VB.Label labdata 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1245
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reloj del Sistema."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmreloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Reloj de Timeline
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

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
 Me.Icon = frmprograma.Icon
 Call cargarIdioma
nose:
End Sub

Private Sub Timer1_Timer()
On Error GoTo nose
 lab_reloj.Caption = Time & " " & lenguaje_Menu(350)
 Label2.Caption = lenguaje_Menu(349) & " " & Date
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmreloj
nose:
End Sub

Private Sub cargarIdioma()
On Error GoTo nose
  Me.Caption = lenguaje_Menu(347)
  Label1.Caption = lenguaje_Menu(348)
  cmdAceptar.Caption = lenguaje_Menu(351)
nose:
End Sub
