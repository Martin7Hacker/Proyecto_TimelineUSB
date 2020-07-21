VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmmensage 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "frmmensage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   4210752
      BackColor       =   -2147483646
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmmensage.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdguardarysalir 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Guardar y Salir"
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
      MICON           =   "frmmensage.frx":57FE
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
      Left            =   9600
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmmensage.frx":581A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label labdatos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¿Existen Archivos en memoria que desea Hacer?"
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
      Height          =   195
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   4170
   End
End
Attribute VB_Name = "frmmensage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Existen Archivos para Timeline
'*
'*
'***************************************************************************
Dim d As Long

Private Sub cmdCancelar_Click()
On Error GoTo nose
 frmprograma.Enabled = True
 Unload Me
nose:
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdguardarysalir_Click()
On Error GoTo nose
 frmprograma.guardard_Click
 End
nose:
End Sub

Private Sub cmdguardarysalir_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdsalir_Click()
On Error GoTo nose
 End
nose:
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
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
 cargar_datos
 cargar_lenguage
nose:
End Sub

Private Sub cargar_datos()
On Error GoTo nose
 Const espacio As String = "                               "
  With frmprograma
  Dim ah As Integer
  Dim v As String
  Dim et As ListItem
  With ListView1.ColumnHeaders
 .Add , , lenguaje_Menu(257)
 .Add , , lenguaje_Menu(258)
 .Add , , lenguaje_Menu(259)
 .Add , , lenguaje_Menu(260)
 .Add , , lenguaje_Menu(261)
 .Add , , lenguaje_Menu(383)
  End With
  With ListView1
  ' Las pruebas serán en modo "detalle"
  .View = lvwReport
  .LabelEdit = lvwManual
  ' Permitir múltiple selección
  .MultiSelect = False
  ' Para que al perder el foco,
  ' se siga viendo el que está seleccionado
  .HideSelection = False
  End With
  ListView1.View = lvwReport
  For ah = 0 To .listado(0).ListCount - 1
  If .listado(1).List(ah) = lenguaje_Menu(18) Then
  v = "   "
  Else
  v = ""
  End If
  d = Int(ah) + 1
   With ListView1.ListItems.Add(, , lenguaje_Menu(264) & "_____ " & d)
 .SubItems(1) = frmprograma.listado(0).List(ah)
 .SubItems(2) = frmprograma.listado(1).List(ah)
 .SubItems(3) = lenguaje_Menu(382) & frmprograma.listado(2).List(ah)
 .SubItems(4) = frmprograma.listado(3).List(ah)
 .SubItems(5) = " " & frmprograma.domingo.List(ah) & " " & _
 frmprograma.lunes(0).List(ah) & " " & _
 frmprograma.martes.List(ah) & " " & _
 frmprograma.miercoles.List(ah) & " " & _
 frmprograma.jueves.List(ah) & " " & _
 frmprograma.viernes.List(ah) & " " & _
 frmprograma.sabado.List(ah)
 End With
 Next ah
 End With
nose:
End Sub
Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
'sale oprimendo Esc
fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmmensage
nose:
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cargar_lenguage()
On Error GoTo nose
 Me.Caption = Lenguage.lenguaje_Menu(196)
 labdatos.Caption = Lenguage.lenguaje_Menu(197)
 cmdSalir.Caption = Lenguage.lenguaje_Menu(203)
 cmdguardarysalir.Caption = Lenguage.lenguaje_Menu(204)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(205)
nose:
End Sub
