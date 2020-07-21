VERSION 5.00
Begin VB.Form frmopciones 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciónes de Módificado"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9120
   Icon            =   "frmopciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9120
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdrestaurar 
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   2490
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
      MICON           =   "frmopciones.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdSalir 
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Top             =   3000
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
      MICON           =   "frmopciones.frx":57FE
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
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aplicar"
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
      MICON           =   "frmopciones.frx":581A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   14
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame fratipos 
      BackColor       =   &H80000002&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   8895
      Begin VB.ComboBox cbotipo 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   200
         Width           =   4815
      End
      Begin VB.Label labtipos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Aplicado:"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   13
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   12
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   11
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   10
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   9
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   8
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   7
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   6
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   4
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   3
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
      BackColor       =   &H80000002&
      Height          =   375
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox cheControles 
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
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame fratipos 
      BackColor       =   &H80000002&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.Label labtipos 
         BackStyle       =   0  'Transparent
         Caption         =   "Oprimiendo los bótones."
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
         Index           =   2
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label labtipos 
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado de Datos:"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmopciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Opciones de Modificado de Timeline
'*
'*
'***************************************************************************
Private est1(17) As Boolean

Private Sub cmdAplicar_Click()
On Error GoTo nose
 Select Case (cbotipo.ListIndex)
 Case (0)
 pasar_aBoolean True
 Case (1)
 pasar_aBoolean False
 End Select
 optener_estado
nose:
End Sub

Private Sub cmdrestaurar_Click()
On Error GoTo nose
 Dim elemento As Byte
 For elemento = 0 To 14
 cheControles(elemento).Value = 0
 Next elemento
nose:
End Sub

Private Sub cmdsalir_Click() 'se ejecuta cuando se hace un Click sobre el bóton salir
On Error GoTo nose
 salir_formulario
nose:
End Sub

Private Sub Form_Load() 'se ejecuta cuando se carga el formulario
On Error GoTo nose
 formulario_cargar
 pasar_control
nose:
End Sub

Private Sub formulario_cargar()  'procedimiento para aplicar al programas las opciones adecuadas
On Error GoTo nose
 modificar_controles             'llamada a la procedimiento módificar controles
 tipo_aplicado
 pasar_lenguage                  'carga el lenguage
nose:
End Sub

Private Sub modificar_controles() 'procedimiento de módificado de dats
On Error GoTo nose
 Me.Icon = frmprograma.Icon       'de generador de datos del programa
nose:
End Sub                           'pasar el icono del programa principal
                                  'a nuestro programa
                                  
Private Sub salir_formulario() 'se utiliza este procedimiento para
On Error GoTo nose
 Unload Me                     'salir del formulario de opciónes de módificado
nose:
End Sub                        'salir
                               'sale de este formulario

Private Sub tipo_aplicado()                'se utiliza este procedimiento para
On Error GoTo nose
 With cbotipo                              'selecónar una opción de modificado
 .Clear                                    'de que control pertenece
 .AddItem (Lenguage.lenguaje_Menu(60))
 .AddItem (Lenguage.lenguaje_Menu(61)) ' borro el selector para que no se sobrecargen las opciónes
 .ListIndex = 0                            ' agrego elementos de seleción
 End With                                  ' de desimos que opción seleciónar
nose:
End Sub                                    'seleciónar el indice 0
                                          
Private Sub pasar_lenguage()
On Error GoTo nose
 Me.Caption = Lenguage.lenguaje_Menu(42)
 Me.labtipos(0).Caption = Lenguage.lenguaje_Menu(43)
 Me.labtipos(1).Caption = Lenguage.lenguaje_Menu(44)
 cheControles(0).Caption = Lenguage.lenguaje_Menu(45)
 cheControles(1).Caption = Lenguage.lenguaje_Menu(46)
 cheControles(2).Caption = Lenguage.lenguaje_Menu(47)
 cheControles(3).Caption = Lenguage.lenguaje_Menu(48)
 cheControles(4).Caption = Lenguage.lenguaje_Menu(49)
 cheControles(5).Caption = Lenguage.lenguaje_Menu(50)
 cheControles(6).Caption = Lenguage.lenguaje_Menu(51)
 cheControles(7).Caption = Lenguage.lenguaje_Menu(52)
 cheControles(8).Caption = Lenguage.lenguaje_Menu(53)
 cheControles(9).Caption = Lenguage.lenguaje_Menu(54)
 cheControles(10).Caption = Lenguage.lenguaje_Menu(55)
 cheControles(11).Caption = Lenguage.lenguaje_Menu(56)
 cheControles(12).Caption = Lenguage.lenguaje_Menu(57)
 cheControles(13).Caption = Lenguage.lenguaje_Menu(58)
 labtipos(1).Caption = Lenguage.lenguaje_Menu(59) '18
 cmdAplicar.Caption = Lenguage.lenguaje_Menu(62) '
 cmdSalir.Caption = Lenguage.lenguaje_Menu(63) '
 cheControles(14).Caption = Lenguage.lenguaje_Menu(64)
 cmdrestaurar.Caption = Lenguage.lenguaje_Menu(65)
nose:
End Sub

Private Sub control_activo(ByVal control As Object _
, ByVal estado As Boolean)
On Error GoTo nose
 control.Enabled = estado
nose:
End Sub

Private Sub modifico_oNo(ByVal control1 As Boolean, _
 ByVal control2 As Boolean, ByVal control3 As Boolean, _
 ByVal control4 As Boolean, ByVal control5 As Boolean, _
 ByVal control6 As Boolean, ByVal control7 As Boolean, _
 ByVal control8 As Boolean, ByVal control9 As Boolean, _
 ByVal control10 As Boolean, ByVal control11 As Boolean, _
 ByVal control12 As Boolean, ByVal control13 As Boolean, _
 ByVal control14 As Boolean, ByVal control15 As Boolean, _
 ByVal control16 As Boolean)
 On Error GoTo nose
  With frmentradasalida
  control_activo .DTPicker1, control1
  control_activo .cobd(0), control2
  control_activo .cobd(1), control3
  control_activo .cobd(2), control4
  control_activo .cobd(3), control5
  control_activo .Check1(0), control6
  control_activo .Check1(1), control7
  control_activo .Check1(2), control8
  control_activo .Check1(3), control9
  control_activo .Check1(4), control10
  control_activo .Check1(5), control11
  control_activo .Check1(6), control12
  control_activo .Text1(0), control13
  control_activo .Text1(1), control14
  control_activo .cob1, control15
  control_activo .txtd, control16
  End With
nose:
End Sub

Private Sub pasar_aBoolean(ByVal estado As Boolean)
On Error GoTo nose
 With frmentradasalida
 Dim control As Byte
 Select Case (cheControles.Item(0).Value)
  Case (1)
  Select Case (estado)
   Case (True)
   .DTPicker1.Enabled = False
   Case (False)
   .DTPicker1.Enabled = True
  End Select
  Case (0)
  Select Case (estado)
   Case (True)
   .DTPicker1.Enabled = True
   Case (False)
   .DTPicker1.Enabled = False
  End Select
 End Select
Select Case (cheControles.Item(1).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .cobd(0).Enabled = False
 Case (False)
 .cobd(0).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .cobd(0).Enabled = True
 Case (False)
 .cobd(0).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(2).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .cobd(1).Enabled = False
 Case (False)
 .cobd(1).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .cobd(1).Enabled = True
 Case (False)
 .cobd(1).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(3).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .cobd(2).Enabled = False
 Case (False)
 .cobd(2).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .cobd(2).Enabled = True
 Case (False)
 .cobd(2).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(4).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .cobd(3).Enabled = False
 Case (False)
 .cobd(3).Enabled = True
End Select
Case (0)
 Select Case (estado)
 Case (True)
 .cobd(3).Enabled = True
 Case (False)
 .cobd(3).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(5).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Text1(0).Enabled = False
 Case (False)
 .Text1(0).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Text1(0).Enabled = True
 Case (False)
 .Text1(0).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(6).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Text1(1).Enabled = False
 Case (False)
 .Text1(1).Enabled = True
 End Select
Case (0)
Select Case (estado)
 Case (True)
 .Text1(1).Enabled = True
 Case (False)
 .Text1(1).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(7).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(0).Enabled = False
 Case (False)
 .Check1(0).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Check1(0).Enabled = True
 Case (False)
 .Check1(0).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(8).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(1).Enabled = False
 Case (False)
 .Check1(1).Enabled = True
 End Select
 Case (0)
 Select Case (estado)
 Case (True)
 .Check1(1).Enabled = True
 Case (False)
 .Check1(1).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(9).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(2).Enabled = False
 Case (False)
 .Check1(2).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Check1(2).Enabled = True
 Case (False)
 .Check1(2).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(10).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(3).Enabled = False
 Case (False)
 .Check1(3).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Check1(3).Enabled = True
 Case (False)
 .Check1(3).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(11).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(4).Enabled = False
 Case (False)
 .Check1(4).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Check1(4).Enabled = True
 Case (False)
 .Check1(4).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(12).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(5).Enabled = False
 Case (False)
 .Check1(5).Enabled = True
 End Select
Case (0)
 Select Case (estado)
 Case (True)
 .Check1(5).Enabled = True
 Case (False)
 .Check1(5).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(13).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .Check1(6).Enabled = False
 Case (False)
 .Check1(6).Enabled = True
 End Select
 Case (0)
 Select Case (estado)
 Case (True)
 .Check1(6).Enabled = True
 Case (False)
 .Check1(6).Enabled = False
 End Select
End Select
Select Case (cheControles.Item(14).Value)
Case (1)
 Select Case (estado)
 Case (True)
 .cob1.Enabled = False
 .txtd.Enabled = False
 Case (False)
 .cob1.Enabled = True
 .txtd.Enabled = True
 End Select
 Case (0)
 Select Case (estado)
 Case (True)
 .cob1.Enabled = True
 .txtd.Enabled = True
 Case (False)
 .cob1.Enabled = False
 .txtd.Enabled = False
 End Select
End Select
End With
nose:
End Sub

Private Sub selecionar_enLista()
On Error GoTo nose
 pasar_aBoolean True
nose:
End Sub

Private Sub optener_estado()
On Error GoTo nose
 Dim elemento As Byte
 For elemento = 0 To 14
 MemoriaF.estado_opciones(elemento) = cheControles(elemento).Value
 Next elemento
nose:
End Sub

Private Sub pasar_control()
On Error GoTo nose
 Dim elemento As Byte
 For elemento = 0 To 14
 cheControles(elemento).Value = MemoriaF.estado_opciones(elemento)
 Next elemento
nose:
End Sub
