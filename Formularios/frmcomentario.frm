VERSION 5.00
Begin VB.Form frmcomentario 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir comentarios"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "frmcomentario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstComentario 
      BackColor       =   &H80000002&
      ForeColor       =   &H00404040&
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8655
   End
   Begin VB.TextBox txtComentario 
      BackColor       =   &H80000002&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAniadir 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Añadir"
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
      MICON           =   "frmcomentario.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdeliminarselecionado 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Eliminar Seleciónado"
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
      MICON           =   "frmcomentario.frx":57FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdEliminarTodo 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Eliminar Todo"
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
      MICON           =   "frmcomentario.frx":581A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAplicar 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Guardar"
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
      MICON           =   "frmcomentario.frx":5836
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmcomentario.frx":5852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCargarComentarios 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cargar "
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
      MICON           =   "frmcomentario.frx":586E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   200
      Width           =   1215
   End
End
Attribute VB_Name = "frmcomentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* comentarios en  Timeline
'*
'*
'***************************************************************************

Private Sub cmdAniadir_Click()
On Error GoTo nose
If Not (txtComentario.Text = "") Then
lstComentario.AddItem txtComentario.Text
End If
txtComentario.Text = ""
nose:
End Sub

Private Sub cmdAplicar_Click()
On Error GoTo nose
Dim r As Integer
Open "comentarios.txt" For Output As 1
 For r = 0 To lstComentario.ListCount - 1
 Print #1, lstComentario.List(r)
 Next r
Close #1
Unload Me
nose:
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo nose
Unload Me
nose:
End Sub

Private Sub cmdCargarComentarios_Click()
On Error GoTo nose
lstComentario.Clear
Dim cargar As String

Open "comentarios.txt" For Input As 1
 Do While Not EOF(1)
       Line Input #1, cargar
       lstComentario.AddItem cargar
       Loop
       Close #1
nose:
End Sub

Private Sub cmdeliminarselecionado_Click()
On Error GoTo nose
If Not (lstComentario.ListIndex = -1) Then
 Select Case MsgBox(lenguaje_Menu(273) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.RemoveItem (lstComentario.ListIndex)
 End Select
End If
nose:
End Sub

Private Sub cmdEliminarTodo_Click()
On Error GoTo nose
If Not (lstComentario.ListIndex <= -1) Then
 Select Case MsgBox(lenguaje_Menu(274) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.Clear
 End Select
End If
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
Me.Icon = frmprograma.Icon
cmdCargarComentarios_Click
cargarIdioma
nose:
End Sub
Private Sub cargarIdioma()
On Error GoTo nose
Me.Caption = lenguaje_Menu(265)
lblComentario.Caption = lenguaje_Menu(266)
cmdAniadir.Caption = lenguaje_Menu(267)
cmdCargarComentarios.Caption = lenguaje_Menu(268)
cmdCancelar.Caption = lenguaje_Menu(269)
cmdeliminarselecionado.Caption = lenguaje_Menu(270)
cmdEliminarTodo.Caption = lenguaje_Menu(271)
cmdAplicar.Caption = lenguaje_Menu(272)
nose:
End Sub

Private Sub lstComentario_Click()
On Error GoTo nose
txtComentario.Text = lstComentario.List(lstComentario.ListIndex)
nose:
End Sub

Private Sub lstComentario_Scroll()
On Error GoTo nose
lstComentario_Click
nose:
End Sub
