VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProgramacon 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programación de Audio Digital "
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmProgramacon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd 
      Left            =   2880
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2760
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   7095
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCargarAudio 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cargar Audio..."
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdx 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdy 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdEliminarSelecionado 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Eliminar Seleciónado"
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0D1E
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
      Left            =   5040
      TabIndex        =   7
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Eliminar "
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0D3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancel 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0D56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdGuradar 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Guardar Programación"
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProgramacon.frx":0D72
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
      Caption         =   $"frmProgramacon.frx":0D8E
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6840
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   7095
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12515
      _cy             =   1085
   End
End
Attribute VB_Name = "frmProgramacon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Programación de Audio de Virtual Martin temporize v1.0
'*
'*
'***************************************************************************

Private Sub cmdcancel_Click()
 Unload Me
End Sub

Private Sub cmdCargarAudio_Click()
 With cd
 .DialogTitle = "Selecióne un Audio a Ser Cargado para Virtual Martin temporize v1.0"
 .Filter = "Archivos de Audio en formato MP3|*.mp3|todos los Archivos (*.*)|*.*|"
 .ShowOpen
 If .FileName <> "" Then
 List1.AddItem .FileName
 End If
 End With
End Sub

Private Sub cmdGuradar_Click()
 guardar_archivo
 Unload Me
End Sub

Private Sub cmdx_Click()
 If Not List1.ListIndex = 0 Then
 On Error GoTo nose
 List1.ListIndex = List1.ListIndex - 1
 End If
nose:
End Sub

Private Sub cmdy_Click()
If Not List1.ListIndex = -1 Then
 If Not List1.ListIndex = List1.ListCount - 1 Then
 List1.ListIndex = List1.ListIndex + 1
 End If
End If
End Sub

Private Sub cmdEliminarSelecionado_Click()
 If Not List1.ListIndex = -1 Then
 Select Case MsgBox("¿ Quieres eliminar el Archivo seleciónado de la Lista de Reprodución ?", vbYesNo + vbInformation, "Opciones de Eliminación")
 Case vbYes
 List1.RemoveItem (List1.ListIndex)
 End Select
 End If
End Sub

Private Sub cmdEliminarTodo_Click()
 If Not List1.ListIndex = -1 Then
 Select Case MsgBox("¿ Quieres eliminar todos los Audios en Lista del Reproductor ?", vbExclamation + vbYesNo, "Microtime v1.0")
 Case vbYes
 List1.Clear
 End Select
 End If
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 wmp.settings.volume = 100
 wmp.Controls.play
 wmp.settings.playCount = 1000
 Abrir_Archivo
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Sub List1_Click()
 wmp.URL = List1.List(List1.ListIndex)
End Sub

Private Sub List1_Scroll()
 List1_Click
End Sub

Public Sub guardar_archivo()
 Dim mus_x As String
 On Error GoTo no_se
 Open App.Path & "\archivoAudio.txt" For Output As 1
 Dim g As Integer
 For g = 0 To List1.ListCount - 1
 mus_x = guardarF.es.escriptar(List1.List(g))
 Print #1, mus_x
 Next
 Close #1
no_se:
 guardar_Archivo_indice
End Sub

Public Sub guardar_Archivo_indice()
 Dim mus_x As String
 On Error GoTo no_se
 Open App.Path & "\archivoIndice.txt" For Output As 1
 Dim g As Integer
 mus_x = guardarF.es.escriptar(List1.ListIndex)
 Print #1, mus_x
 Close #1
no_se:
End Sub

Public Sub Abrir_Archivo()
 Dim dato_xr As String
 On Error GoTo no_se
 Open App.Path & "\archivoAudio.txt" For Input As 1
 Do While Not EOF(1)
 Line Input #1, dato_xr
 List1.AddItem guardarF.es.desescriptar(dato_xr)
 Loop
 Close #1
no_se:
 Abrir_Archivo_indice
End Sub

Public Sub Abrir_Archivo_indice()
 Dim dato_xr As String
 On Error GoTo no_se
 Open App.Path & "\archivoIndice.txt" For Input As 1
 Do While Not EOF(1)
 Line Input #1, dato_xr
 List1.ListIndex = guardarF.es.desescriptar(dato_xr)
 Loop
 Close #1
no_se:
End Sub

