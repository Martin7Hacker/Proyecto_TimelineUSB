VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnuevoevento 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   Icon            =   "frmnuevoevento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcomentarios 
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   645
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&comentarios:"
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
      MICON           =   "frmnuevoevento.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.TextBox Text1 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   2055
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "Domingos."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000080FF&
         Caption         =   "Sabado ."
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Viernes ."
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Jueves ."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Miercoles."
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Martes ."
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Lunes ."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   2
         ItemData        =   "frmnuevoevento.frx":57FE
         Left            =   840
         List            =   "frmnuevoevento.frx":5800
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         ItemData        =   "frmnuevoevento.frx":5802
         Left            =   840
         List            =   "frmnuevoevento.frx":5804
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   1050
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         ItemData        =   "frmnuevoevento.frx":5806
         Left            =   600
         List            =   "frmnuevoevento.frx":5808
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   295
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         Format          =   107282434
         UpDown          =   -1  'True
         CurrentDate     =   0.805555555555556
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   650
         Width           =   600
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn boton 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1095
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmnuevoevento.frx":580A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   20
      Top             =   0
      Width           =   0
   End
   Begin Virtual_Martin_temporize.ChameleonBtn boton 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   21
      Top             =   3240
      Width           =   1095
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmnuevoevento.frx":5826
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdfunct 
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&Funciones al sistema"
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
      MICON           =   "frmnuevoevento.frx":5842
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   120
      X2              =   2520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label labinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar Nuevo Evento"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1950
   End
   Begin VB.Menu comentar 
      Caption         =   "menú"
      Visible         =   0   'False
      Begin VB.Menu mc 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmnuevoevento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Nuevo Evento y Modificación para Timeline
'*
'*
'***************************************************************************
Dim nuevoEvento As evento

Private Sub boton_Click(Index As Integer)
On Error GoTo nose
 frmprograma.Enabled = True
 With frmprograma
 Select Case Index
  Case (0)
  Unload Me
  Case (1)
 If boton(1).Caption = Lenguage.lenguaje_Menu(224) Then
 nuevo_evento_de_dias
 Crear ' crea un nunevo evento de timbre
 sistema.ingresarDatos
 ElseIf boton(1).Caption = Lenguage.lenguaje_Menu(225) Then
 'selección
 Select Case MsgBox(Lenguage.lenguaje_Menu(231) _
 , vbYesNo + vbInformation, lenguaje_Menu(8))
  Case (vbYes)

  labinfo.Caption = Lenguage.lenguaje_Menu(226)
  
  .listado(0).List(.listado(0).ListIndex) = DTPicker1.Value
  .listado(1).List(.listado(1).ListIndex) = Combo1(0).Text
  .listado(2).List(.listado(2).ListIndex) = Combo1(1).Text
  .listado(3).List(.listado(3).ListIndex) = Text1.Text
  .liscomando.List(.liscomando.ListIndex) = frmfunciones.devolver_comando
  'dias Set'
 set_dias ' cambia los dias de la semana
 .Filtro.List(.Filtro.ListIndex) = Combo1(2).ListIndex
 Unload Me
 End Select
 End If
 End Select
 End With
nose:
End Sub

Private Sub Crear()
On Error GoTo nose
 Set nuevoEvento = New evento
 With nuevoEvento
 .vHora.Add DTPicker1.Value
 .vTipo.Add Combo1(0).Text
 .vIntervalo.Add Combo1(1).Text
 .vtipod.Add Combo1(2).Text
 .vDescripcion.Add Text1.Text
 End With
 With frmprograma
 Dim recor As Integer
 For recor = 1 To nuevoEvento.vHora.Count
 .listado(0).AddItem nuevoEvento.vHora(recor)
 .listado(1).AddItem nuevoEvento.vTipo(recor)
 .listado(2).AddItem nuevoEvento.vIntervalo(recor)
 .listado(3).AddItem nuevoEvento.vDescripcion(recor)
 Next
 End With
 Unload Me
nose:
End Sub

Private Sub boton_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdobsiones_Click()
On Error GoTo nose
 PopupMenu obsiones
nose:
End Sub

Private Sub cmdobsiones_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cmdcomentarios_Click()
On Error GoTo nose
 PopupMenu comentar
nose:
End Sub

Private Sub cmdfunct_Click()
On Error GoTo nose
 If boton(1).Caption = lenguaje_Menu(92) Then
 frmfunciones.cmdAplicar.Caption = lenguaje_Menu(92)
 End If
 frmfunciones.Show 1
nose:
End Sub

Private Sub cmdfunct_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Combo1_Click(Index As Integer)
On Error GoTo nose
 Select Case Index
  Case (2)
 Select Case Combo1(2).ListIndex
  Case (0)
  visiblex False
  activado 0
  Dim td As Byte
  For td = 0 To 6
  Check1(CInt(td)).Value = 1
  Next
 Case (1)
 visiblex True
 activado 0
 End Select
 End Select
nose:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
Call cmdCargarComentarios: Call cargarIdioma
 Me.Icon = frmprograma.Icon
 Combo1(2).ListIndex = CInt(MemoriaF.numero)
 visiblex CInt(MemoriaF.numero)
 DTPicker1.Value = Time
 agregar_elementos
 If MemoriaF.dias = True Then
 devolver_dias
 End If
nose:
End Sub

Private Sub agregar_elementos()
On Error GoTo nose
 Dim numero As Integer
 Combo1(0).ListIndex = 0
 For numero = 1 To 77
 Combo1(1).AddItem (numero)
 Next
 Combo1(1).ListIndex = 4
nose:
End Sub

Private Sub visiblex(ByVal visilblex As Boolean)
On Error GoTo nose
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Enabled = visilblex
 Next
nose:
End Sub

Private Sub activado(ByVal activado As Byte)
On Error GoTo nose
 Dim rx As Integer
 For rx = 0 To 6
 Check1(rx).Value = activado
 Next
nose:
End Sub

Private Sub almanaque_Click()
On Error GoTo nose
 frmalmanaque.Show 1
nose:
End Sub

Private Sub nuevo_evento_de_dias()
On Error GoTo nose
 Const nulo As String = "0"      'nulo
 Const lunes As String = "2"     'lunes
 Const martes As String = "3"    'martes
 Const miercoles As String = "4" 'miercoles
 Const jueves As String = "5"    'jueves
 Const viernes As String = "6"   'viernes
 Const sabado As String = "7"    'sabado
 Const domingo As String = "1"   'domingo
 With frmprograma
 Select Case Check1(0).Value     ' Lunes
  Case (1)
  .lunes(0).AddItem lunes
  Case (0)
  .lunes(0).AddItem nulo
 End Select
Select Case Check1(1).Value      ' Martes
 Case (1)
 .martes.AddItem martes          ' Martes
 Case (0)
 .martes.AddItem nulo
End Select
Select Case Check1(2).Value ' Miercoles
 Case (1)
 .miercoles.AddItem miercoles
 Case (0)
 .miercoles.AddItem nulo
End Select
Select Case Check1(3).Value ' Jueves
 Case (1)
 .jueves.AddItem jueves
 Case (0)
 .jueves.AddItem nulo
End Select
Select Case Check1(4).Value ' Viernes
 Case (1)
 .viernes.AddItem viernes
 Case (0)
 .viernes.AddItem nulo
End Select
Select Case Check1(5).Value ' Sabado
 Case (1)
 .sabado.AddItem sabado
 Case (0)
 .sabado.AddItem nulo
End Select
Select Case Check1(6).Value ' Domingo
 Case (1)
 .domingo.AddItem domingo
 Case (0)
 .domingo.AddItem nulo
End Select
'***************'> Asignacion de Filtro <******************'
.Filtro.AddItem Combo1(2).ListIndex
 End With
nose:
End Sub

Public Sub devolver_dias()
On Error GoTo nose
 Dim dev As Integer
 For dev = 0 To frmprograma.listado(0).ListCount
 With frmprograma
 'lunes
 Select Case .lunes(0).List(.lunes(0).ListIndex)
  Case (2)
  Check1(0).Value = 1
  Case (0)
  Check1(0).Value = 0
 End Select
'martes
Select Case .martes.List(.martes.ListIndex)
 Case (3)
 Check1(1).Value = 1
 Case (0)
 Check1(1).Value = 0
End Select
'miercoles
Select Case .miercoles.List(.miercoles.ListIndex)
 Case (4)
 Check1(2).Value = 1
 Case (0)
 Check1(2).Value = 0
End Select
'jueves
Select Case .jueves.List(.jueves.ListIndex)
 Case (5)
 Check1(3).Value = 1
 Case (0)
 Check1(3).Value = 0
End Select
'viernes
Select Case .viernes.List(.viernes.ListIndex)
 Case (6)
 Check1(4).Value = 1
 Case (0)
 Check1(4).Value = 0
End Select
'sabado
Select Case .sabado.List(.sabado.ListIndex)
 Case (7)
 Check1(5).Value = 1
 Case (0)
 Check1(5).Value = 0
End Select
'domingo
Select Case .domingo.List(.domingo.ListIndex)
 Case (1)
 Check1(6).Value = 1
 Case (0)
 Check1(6).Value = 0
End Select
End With
Next dev
nose:
End Sub

Private Sub set_dias()
On Error GoTo nose
 With frmprograma
 'lunes
 Select Case Check1(0).Value
  Case (1)
  .lunes(0).List(.lunes(0).ListIndex) = 2
  Case (0)
  .lunes(0).List(.lunes(0).ListIndex) = 0
 End Select
 'martes
 Select Case Check1(1).Value
  Case (1)
  .martes.List(.martes.ListIndex) = 3
  Case (0)
  .martes.List(.martes.ListIndex) = 0
  End Select
 'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.List(.miercoles.ListIndex) = 4
 Case (0)
 .miercoles.List(.miercoles.ListIndex) = 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.List(.jueves.ListIndex) = 5
 Case (0)
 .jueves.List(.jueves.ListIndex) = 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.List(.viernes.ListIndex) = 6
 Case (0)
 .viernes.List(.viernes.ListIndex) = 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.List(.sabado.ListIndex) = 7
 Case (0)
 .sabado.List(.sabado.ListIndex) = 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.List(.domingo.ListIndex) = 1
 Case (0)
 .domingo.List(.domingo.ListIndex) = 0
End Select
End With
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmnuevoevento
nose:
End Sub

Private Sub mc_Click(Index As Integer)
On Error GoTo nose
 Text1.Text = mc.Item(Index).Caption
nose:
End Sub

Private Sub cargarIdioma()
On Error GoTo nose
 labinfo.Caption = lenguaje_Menu(207)
 boton(1).Caption = Lenguage.lenguaje_Menu(224)
 cmdfunct.Caption = lenguaje_Menu(208)
 etiqueta(0).Caption = Lenguage.lenguaje_Menu(209)
 etiqueta(1).Caption = Lenguage.lenguaje_Menu(210)
 etiqueta(2).Caption = Lenguage.lenguaje_Menu(211)
 etiqueta(3).Caption = Lenguage.lenguaje_Menu(212)
 Check1(0).Caption = Lenguage.lenguaje_Menu(213)
 Check1(1).Caption = Lenguage.lenguaje_Menu(214)
 Check1(2).Caption = Lenguage.lenguaje_Menu(215)
 Check1(3).Caption = Lenguage.lenguaje_Menu(216)
 Check1(4).Caption = Lenguage.lenguaje_Menu(217)
 Check1(5).Caption = Lenguage.lenguaje_Menu(218)
 Check1(6).Caption = Lenguage.lenguaje_Menu(219)
 etiqueta(4).Caption = Lenguage.lenguaje_Menu(220)
 cmdcomentarios.Caption = Lenguage.lenguaje_Menu(221)
 boton(0).Caption = Lenguage.lenguaje_Menu(222)
 Combo1(0).AddItem Lenguage.lenguaje_Menu(227)
 Combo1(0).AddItem Lenguage.lenguaje_Menu(228)
 Combo1(2).AddItem Lenguage.lenguaje_Menu(229)
 Combo1(2).AddItem Lenguage.lenguaje_Menu(230)
nose:
End Sub

Private Sub cmdCargarComentarios()
On Error GoTo nose
Dim cargar As String
Dim r As Integer
r = 1
Open "comentarios.txt" For Input As 1
 Do While Not EOF(1)
       Line Input #1, cargar
       If mc(0).Caption = "" Then
          mc(0).Caption = cargar
       Else
       Load frmnuevoevento.mc(r)
       mc(r).Caption = cargar
       mc(r).Visible = True
       r = r + 1
       End If
       Loop
       Close #1
       r = 0
nose:
End Sub
