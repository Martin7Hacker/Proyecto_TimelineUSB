VERSION 5.00
Begin VB.Form frmDonativos 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donativos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "frmDonativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pdonar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   600
      MouseIcon       =   "frmDonativos.frx":57E2
      Picture         =   "frmDonativos.frx":5AEC
      ScaleHeight     =   1155
      ScaleWidth      =   2925
      TabIndex        =   3
      Top             =   970
      Width           =   2955
   End
   Begin VB.PictureBox ptargeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      Picture         =   "frmDonativos.frx":10C0A
      ScaleHeight     =   225
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcolaborar 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Colaborar"
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
      MICON           =   "frmDonativos.frx":11405
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
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      MICON           =   "frmDonativos.frx":11421
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amo mucho a EE:UU"
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
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "para cumplir mi sueño de ir a EE:UU"
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
      Height          =   315
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "con cuenta propia..."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   2745
   End
   Begin VB.Label lblcard 
      BackStyle       =   0  'Transparent
      Caption         =   "Con tarjetas de créditos"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   2985
   End
End
Attribute VB_Name = "frmDonativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Para realizar donacíones para el proyecto Timeline
'*
'*
'***************************************************************************

Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAceptar_Click()
On Error GoTo nose
 Unload Me
nose:
End Sub

Private Sub cmdcolaborar_Click()
On Error GoTo nose
 ptargeta_Click
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
 Me.Icon = frmprograma.Icon
 Call cargarIdioma
nose:
End Sub

Private Sub Label1_Click()
On Error GoTo nose
 ptargeta_Click
nose:
End Sub

Private Sub lblcard_Click()
On Error GoTo nose
 ptargeta_Click
nose:
End Sub

Private Sub pdonar_Click()
On Error GoTo nose
 ptargeta_Click
nose:
End Sub

Private Sub ptargeta_Click()
On Error GoTo nose
 Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "http://martinsoft0.blogspot.com/p/donar.html", _
 &O0, &O0, 0)
 Unload Me
nose:
End Sub
Private Sub cargarIdioma()
On Error GoTo nose
Me.Caption = lenguaje_Menu(310)
Label2.Caption = lenguaje_Menu(311)
Label3.Caption = lenguaje_Menu(312)
Label1.Caption = lenguaje_Menu(313)
lblcard.Caption = lenguaje_Menu(314)
cmdcolaborar.Caption = lenguaje_Menu(315)
cmdAceptar.Caption = lenguaje_Menu(316)
nose:
End Sub

