VERSION 5.00
Begin VB.Form frmidioma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instancia de Idioma"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdrenombrar 
      Caption         =   "Renombrar"
      Height          =   285
      Left            =   6960
      TabIndex        =   11
      Top             =   450
      Width           =   975
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2790
      Left            =   7680
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdidioma 
      Caption         =   $"idioma.frx":0000
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   7815
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Idioma.idi"
      Top             =   90
      Width           =   6015
   End
   Begin VB.CommandButton cmdcargar 
      Caption         =   "Cargar Idioma"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cargar el Idioma"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtvalor 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   450
      Width           =   4935
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "&Guardar Archivo"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ListBox lstidioma 
      BackColor       =   &H8000000F&
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Archivo:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   405
   End
End
Attribute VB_Name = "frmidioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdcargar_Click()
idioma
End Sub

Private Sub cmdguardar_Click()
Dim fileS As Byte
Open Lenguage.Lfile For Output As 1
 For fileS = 0 To lstidioma.ListCount - 2
 Print #1, lstidioma.List(fileS)
 Next fileS
Close #1
End Sub

Private Sub cmdrenombrar_Click()
lstidioma.List(lstidioma.ListIndex) = txtvalor.Text
End Sub

Private Sub Form_Load()
Call cargarPrograma
Call idioma
Me.Icon = frmprograma.Icon
VScroll1_Change
txtnombre.Text = Lfile
End Sub

Private Sub cargarPrograma()
Me.Icon = frmprograma.Icon
End Sub

Public Sub idioma()
With lstidioma
Dim inc_x As Byte
For inc_x = 0 To 23
 .AddItem lenguage_opciones(inc_x)
Next inc_x
For inc_x = 0 To 40
 .AddItem lenguage_opciones_generador(inc_x)
Next inc_x
For inc_x = 0 To 10
 .AddItem lenguage_rutas(inc_x)
 Next inc_x

For inc_x = 0 To 7
 .AddItem lenguage_iniciarwindows(inc_x)
 Next inc_x

 For inc_x = 0 To 3
 .AddItem lenguage_circuito(inc_x)
 Next inc_x

 For inc_x = 0 To 6
 .AddItem lenguage_estVentana(inc_x)
 Next inc_x

 For inc_x = 0 To 7
 .AddItem lenguage_datosCreador(inc_x)
 Next inc_x

 For inc_x = 0 To 20
 .AddItem lenguage_fichaCreador(inc_x)
 Next inc_x

 For inc_x = 0 To 17
 .AddItem lenguage_estFunciones(inc_x)
 Next inc_x

 For inc_x = 0 To 5
 .AddItem lenguage_estcopias(inc_x)
 Next inc_x

 For inc_x = 0 To 10
 .AddItem lenguage_memoria(inc_x)
 Next inc_x

 For inc_x = 0 To 25
 .AddItem lenguage_crearModificar(inc_x)
 Next inc_x
.AddItem "Idioma"
.AddItem "Archivo"
.AddItem "Idioma"
End With
lstidioma.Clear

Dim r_idioma As Integer
Dim cargar As String
Open Lenguage.Lfile For Input As 1
 Do While Not EOF(1)
  
   Line Input #1, cargar
   lstidioma.AddItem cargar
  Loop
Close #1

 For inc_x = 0 To lstidioma.ListCount + 1
 Lenguage.c_lenguage(inc_x) = lstidioma.List(inc_x)
 Next inc_x


End Sub

Private Sub lstidioma_Click()
lstidioma_Scroll
VScroll1.Value = lstidioma.ListIndex
End Sub

Private Sub lstidioma_Scroll()
txtvalor.Text = lstidioma.List(lstidioma.ListIndex)
VScroll1.Value = lstidioma.ListIndex

End Sub

Private Sub VScroll1_Change()
 On Error GoTo nose
 With VScroll1
 .Max = lstidioma.ListCount - 1
 .Min = 0
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 End With
nose:
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub
