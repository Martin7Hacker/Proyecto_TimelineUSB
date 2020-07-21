VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmentradasalida 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Rutinas"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   Icon            =   "frmentradasalida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdmas 
      Height          =   360
      Left            =   3600
      TabIndex        =   43
      Top             =   4080
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
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
      MICON           =   "frmentradasalida.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Lista Desplegable."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   32
      Top             =   50
      Width           =   2055
   End
   Begin VB.ComboBox cobd 
      BackColor       =   &H80000002&
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   4
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4890
      Width           =   975
   End
   Begin VB.ComboBox cobd 
      BackColor       =   &H80000002&
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Text            =   "cobd"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame frmRutina 
      BackColor       =   &H80000002&
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8535
      Begin VB.PictureBox Picture4 
         Height          =   60
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   4875
         TabIndex        =   27
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000002&
         ForeColor       =   &H0000FF00&
         Height          =   1575
         Index           =   1
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000002&
         ForeColor       =   &H000000FF&
         Height          =   1575
         Index           =   0
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   480
         Width           =   3375
      End
      Begin VB.PictureBox Picture2 
         Height          =   4095
         Left            =   4920
         ScaleHeight     =   4035
         ScaleWidth      =   0
         TabIndex        =   21
         Top             =   120
         Width           =   60
         Begin VB.PictureBox Picture3 
            Height          =   4695
            Left            =   -6840
            ScaleHeight     =   4635
            ScaleWidth      =   6795
            TabIndex        =   26
            Top             =   -2280
            Width           =   6855
         End
      End
      Begin VB.Frame fram_dias 
         BackColor       =   &H80000002&
         ForeColor       =   &H00008000&
         Height          =   2175
         Left            =   3600
         TabIndex        =   12
         ToolTipText     =   "Listado de Progrmación de los dias o el dia que queres activar el timbre."
         Top             =   280
         Width           =   1215
         Begin VB.CheckBox Check1 
            BackColor       =   &H0000FF00&
            Caption         =   "domingo"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   19
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Sabado"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   18
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Viernes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Jueves"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Miercoles"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Martes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Lunes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DIAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   20
            Top             =   1845
            Width           =   735
         End
      End
      Begin VB.ComboBox cobd 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   3
         Left            =   840
         TabIndex        =   10
         Text            =   "cobd"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cobd 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Text            =   "cobd"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox cobd 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Text            =   "cobd"
         Top             =   1200
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744448
         CalendarTitleBackColor=   0
         CalendarTrailingForeColor=   0
         Format          =   106954754
         UpDown          =   -1  'True
         CurrentDate     =   0.805555555555556
      End
      Begin Virtual_Martin_temporize.ChameleonBtn cmdcomentario 
         Height          =   255
         Left            =   5040
         TabIndex        =   45
         Top             =   165
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
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
         MICON           =   "frmentradasalida.frx":57FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Virtual_Martin_temporize.ChameleonBtn cmdComentariox 
         Height          =   255
         Left            =   5040
         TabIndex        =   46
         Top             =   2200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
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
         MICON           =   "frmentradasalida.frx":581A
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1035
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   3000
         Width           =   3540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "[Salida]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   7
         Left            =   6420
         TabIndex        =   25
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Entrada]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   24
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Int:                                  [Salida]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   11
         Top             =   2205
         Width           =   3585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Int:                                (Entrada]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   1725
         Width           =   3180
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro :"
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
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1250
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
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
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   800
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.Image Image1 
         Height          =   1470
         Left            =   0
         Picture         =   "frmentradasalida.frx":5836
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Frame frmframe 
      BackColor       =   &H80000002&
      Height          =   2775
      Left            =   120
      TabIndex        =   31
      Top             =   5400
      Width           =   8535
      Begin VB.TextBox txtd 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   1455
         Left            =   3960
         MaxLength       =   127
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000002&
         Height          =   1815
         Left            =   3840
         TabIndex        =   34
         Top             =   600
         Width           =   4335
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   420
            Left            =   1800
            TabIndex        =   35
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "s"
            Format          =   106954755
            UpDown          =   -1  'True
            CurrentDate     =   0.805555555555556
         End
         Begin VB.Label lbld 
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   960
            TabIndex        =   37
            Top             =   960
            Width           =   795
         End
         Begin VB.Label labinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Sin dialogo..."
            Height          =   195
            Left            =   1800
            TabIndex        =   36
            Top             =   960
            Width           =   1785
         End
      End
      Begin VB.ComboBox cob1 
         BackColor       =   &H80000002&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   4335
      End
      Begin VB.Image Image2 
         Height          =   2295
         Left            =   120
         Picture         =   "frmentradasalida.frx":102C0
         Top             =   240
         Width           =   3525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   1920
         Width           =   45
      End
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdmod 
      Height          =   495
      Left            =   7440
      TabIndex        =   40
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Modificar"
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
      MICON           =   "frmentradasalida.frx":2AA26
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCrearEventos 
      Height          =   495
      Left            =   2400
      TabIndex        =   41
      Top             =   4800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
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
      MICON           =   "frmentradasalida.frx":2AA42
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
      Height          =   495
      Left            =   120
      TabIndex        =   42
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
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
      BCOL            =   -2147483646
      BCOLO           =   -2147483646
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmentradasalida.frx":2AA5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdfiltro 
      Height          =   255
      Left            =   6240
      TabIndex        =   44
      Top             =   50
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Opciones  de modificado"
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
      MICON           =   "frmentradasalida.frx":2AA7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label labrutina 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   30
      Top             =   50
      Width           =   8460
   End
   Begin VB.Menu Filtro 
      Caption         =   "filtro"
      Visible         =   0   'False
      Begin VB.Menu menu 
         Caption         =   "Hora:                              programada"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^H
      End
      Begin VB.Menu menu 
         Caption         =   "Tipo:                               Entrada o Salida"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu menu 
         Caption         =   "Filtro:                              Entrada o Salida o Aleatorio"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu menu 
         Caption         =   "Intervalo:                       Entrada"
         Checked         =   -1  'True
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu menu 
         Caption         =   "Intervalo:                       Salida"
         Checked         =   -1  'True
         Index           =   4
         Shortcut        =   ^I
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Lunes"
         Checked         =   -1  'True
         Index           =   5
         Shortcut        =   ^L
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Martes"
         Checked         =   -1  'True
         Index           =   6
         Shortcut        =   ^M
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Miercoles"
         Checked         =   -1  'True
         Index           =   7
         Shortcut        =   ^N
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Jueves"
         Checked         =   -1  'True
         Index           =   8
         Shortcut        =   ^J
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Viernes"
         Checked         =   -1  'True
         Index           =   9
         Shortcut        =   ^V
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Sabados"
         Checked         =   -1  'True
         Index           =   10
         Shortcut        =   ^S
      End
      Begin VB.Menu menu 
         Caption         =   "Dias:                               Domingos"
         Checked         =   -1  'True
         Index           =   11
         Shortcut        =   ^D
      End
      Begin VB.Menu menu 
         Caption         =   "Comentarios:                  Entada"
         Checked         =   -1  'True
         Index           =   12
         Shortcut        =   ^C
      End
      Begin VB.Menu menu 
         Caption         =   "Comentarios:                  Salida"
         Checked         =   -1  'True
         Index           =   13
         Shortcut        =   ^K
      End
      Begin VB.Menu menu 
         Caption         =   "Auto: Apagar Encender etc*"
         Checked         =   -1  'True
         Index           =   14
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu esp 
         Caption         =   "-"
      End
   End
   Begin VB.Menu comentario 
      Caption         =   "comentario"
      Visible         =   0   'False
      Begin VB.Menu mc 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmentradasalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Generador de Rutinas de Timeline
'*
'*
'***************************************************************************
Dim aleatorio As Boolean: Public devolver_comando As String
Dim con_c As Boolean
Private Sub cmdcomentario_Click()
On Error GoTo nose
con_c = True
PopupMenu comentario
nose:
End Sub

Private Sub cmdComentariox_Click()
On Error GoTo nose
con_c = False
PopupMenu comentario
nose:
End Sub

Private Sub cmdCrearEventos_Click()
On Error GoTo nose
 Dim d As Integer
  For d = 1 To cobd(4).List(cobd(4).ListIndex)
  crearTimbre
  Next
  Unload Me
nose:
End Sub

Private Sub Check1_KeyPress(Index As Integer, _
 KeyAscii As Integer)
 On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub Check2_Click()
On Error GoTo nose
 cargar_idioma
nose:
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo nose
 Unload Me
nose:
End Sub

Private Sub cmdfiltro_Click()
On Error GoTo nose
 Select Case (Check2.Value)
  Case (1)
   PopupMenu Filtro, , _
   cmdfiltro.Left
  Case (0)
   frmopciones.Show 1
 End Select
nose:
End Sub

Private Sub cmdmas_Click()
On Error GoTo nose
 If opciones.mover_flecha(0) = True Then
 Me.Height = 5790
 Me.cmdmas.Caption = "6"
 opciones.mover_flecha(0) = False
 ElseIf opciones.mover_flecha(0) = False Then
 Me.Height = 8835
 Me.cmdmas.Caption = "5"
 frmframe.Top = 5400
 opciones.mover_flecha(0) = True
 End If
nose:
End Sub

Private Sub cmdmod_Click()
 On Error GoTo nose
 modificarDatos DTPicker1, cobd.Item(0), _
 cobd.Item(1), cobd.Item(2), cobd.Item(3), _
 Check1.Item(0), Check1.Item(1), Check1.Item(2), _
 Check1.Item(3), Check1.Item(4), Check1.Item(5), _
 Check1.Item(6), Text1.Item(0), Text1.Item(1), cob1
 Unload Me
nose:
End Sub

Private Sub cob1_Change()
On Error GoTo nose
 cob1_Click
nose:
End Sub

Private Sub cob1_Click()
On Error GoTo nose
 If cob1.ListIndex = 5 Then
 labinfo.Visible = False
 Frame2.Visible = True
 txtd.Visible = False
 lbld.Visible = True
 DTPicker2.Visible = True
 ElseIf cob1.ListIndex = 6 Then
 labinfo.Visible = False
 Frame2.Visible = True
 txtd.Visible = True
 DTPicker2.Visible = False
 lbld.Visible = False
 Else
 labinfo.Visible = True
 Frame2.Visible = True
 txtd.Visible = False
 DTPicker2.Visible = False
 lbld.Visible = False
 End If
nose:
End Sub

Private Sub cob1_DblClick()
On Error GoTo nose
 cob1_Click
nose:
End Sub

Private Sub cob1_Scroll()
On Error GoTo nose
 cob1_Click
nose:
End Sub

Private Sub cobd_Click(Index As Integer)
On Error GoTo nose
 Select Case (Index)
  Case (0)
  Select Case (cobd(0).ListIndex)
  Case (0)
  avilitarControles True, True, True, True
  Case (1)
  avilitarControles True, False, True, False
  Case (2)
  avilitarControles False, True, False, True
  End Select
  Case (1)
 Select Case (cobd(1).ListIndex)
  Case (0)
  fram_dias.Enabled = False
  Dim i As Byte
   For i = 0 To 6
   Check1(i).Value = 1
   Next
   Case (1)
   fram_dias.Enabled = True
   For i = 0 To 6
   Check1(i).Value = 0
   Next
 End Select
 End Select
 unoobarios
nose:
End Sub

Private Sub unoobarios()
 On Error GoTo nose
 If cobd(4).ListIndex = 0 Then
 cmdCrearEventos.Caption = Lenguage.lenguaje_Menu(90)
 Else
 cmdCrearEventos.Caption = Lenguage.lenguaje_Menu(91)
 End If
nose:
End Sub

Private Sub Command1_Click()
 On Error GoTo nose
 cmdmas_Click
nose:
End Sub

Private Sub Command2_Click()
On Error GoTo nose
 cmdmas_Click
nose:
End Sub

Private Sub Command3_Click()
On Error GoTo nose
 cmdmas_Click
nose:
End Sub

Private Sub Form_Load()
On Error GoTo nose
 Me.Icon = frmprograma.Icon
 cargar_controles
 carga_datos
 Cantidad_elementos
 cob1.ListIndex = 0
 Me.Height = 5790
 unoobarios
 cargar_idioma ' cargar idioma
 cmdCargarComentarios
 labrutina.Caption = Me.Caption
nose:
End Sub

Private Sub cargar_controles()
On Error GoTo nose
 With cob1
 .AddItem Lenguage.lenguaje_Menu(93)
 .AddItem Lenguage.lenguaje_Menu(94)
 .AddItem Lenguage.lenguaje_Menu(95)
 .AddItem Lenguage.lenguaje_Menu(96)
 .AddItem Lenguage.lenguaje_Menu(97)
 .AddItem Lenguage.lenguaje_Menu(98)
 .AddItem Lenguage.lenguaje_Menu(99)
 .AddItem Lenguage.lenguaje_Menu(100)
 End With
nose:
End Sub

Private Sub devolverString()
On Error GoTo nose
 Select Case cob1.ListIndex
  Case (0)
  devolver_comando = ""   '// sin opcion
  Case (1)
  devolver_comando = "so.dll -s -f" '// Apagar el equipo
  Case (2)
  devolver_comando = "so.dll -r"    '// reiniciar el equipo
  Case (3)
  devolver_comando = "so.dll -a"    '// anular el apagado del equipo
  Case (4)
  devolver_comando = "so.dll -m"    '// equipo que se / apagara / reiniciara / anulara
  Case (5)
  devolver_comando = "so.dll -t"    '// establecer el tiempo de cierre de apagado en xx segundos
  Case (6)
  devolver_comando = "so.dll -c"    '// le puedes aplicar comentarios
  Case (7)
  devolver_comando = "so.dll -f"    '// fuerza el cierre de aplicaciones sin advertir
  End Select
nose:
End Sub

Private Sub carga_datos()
 On Error GoTo nose
 'TIPO DE INTERVALO
 With cobd(0)
 .AddItem lenguaje_Menu(381)
 .AddItem lenguaje_Menu(227)
 .AddItem lenguaje_Menu(228)
 .ListIndex = 0
 End With
 'TIPO DE INTERVALO ENTRADA
 Dim x As Integer
 With cobd(1)
 .AddItem lenguaje_Menu(229)
 .AddItem lenguaje_Menu(230)
 .ListIndex = 1 ' dia y hora
 End With
 'TIPO DE INTERVALO ENTRADA
 With cobd(2)
 For x = 0 To 77
 .AddItem (x)
 Next
 .ListIndex = 7
 End With
 'TIPO DE INTERVALO SALIDA
 With cobd(3)
 For x = 0 To 77
 .AddItem (x)
 Next
 .ListIndex = 5
 End With
 'Cargar elementos de Carga de Memoria
 Dim ce As Integer
 For ce = 1 To 100 ' crea un maximo de CIEN timbres
 cobd(4).AddItem (ce)
 Next
 cobd(4).ListIndex = 0
nose:
End Sub

Private Sub Cantidad_elementos()
On Error GoTo nose
 If frmprograma.listado(0).ListCount = 0 Then
 Label1(8).Caption = Lenguage.lenguaje_Menu(83): cmdmod.Enabled = False
 ElseIf frmprograma.listado(0).ListCount = 1 Then
 Label1(8).Caption = Lenguage.lenguaje_Menu(84) _
 & " " & frmprograma.listado(0).ListCount & " " & _
 Lenguage.lenguaje_Menu(85): cmdmod.Enabled = True
 Else
 Label1(8).Caption = Lenguage.lenguaje_Menu(84) _
 & " " & frmprograma.listado(0).ListCount & " " & _
 Lenguage.lenguaje_Menu(85): cmdmod.Enabled = True
 End If
nose:
End Sub

Private Sub modificarDatos(ByVal control0 As Object, ByVal control1 As Object, _
 ByVal control2 As Object, ByVal control3 As Object, ByVal control4 As Object, _
 ByVal control5 As Object, ByVal control6 As Object, ByVal control7 As Object, _
 ByVal control8 As Object, ByVal control9 As Object, ByVal control10 As Object, _
 ByVal control11 As Object, ByVal control12 As Object, ByVal control13 As Object, _
 ByVal control14 As Object)
 On Error GoTo nose
  Dim it As Integer
    With frmprograma
    For it = 0 To frmprograma.listado(0).ListCount - 1
    If cobd(0).ListIndex = 0 Then
    'seleciono el tipo
    Select Case (aleatorio)
    Case (False)
    If control0.Enabled = True Then
    .listado(0).List(it) = DTPicker1.Value
    End If
    If control1.Enabled = True Then
    .listado(1).List(it) = lenguaje_Menu(227)  ' entrada | salida
    End If
    If control3.Enabled = True Then
    .listado(2).List(it) = cobd(2).Text
    End If
    If control12.Enabled = True Then
    .listado(3).List(it) = Text1(0).Text
    End If
    'le aplica el filtro correspondiente
    If control2.Enabled = True Then
    .Filtro.List(it) = cobd(1).ListIndex
    '# SETEO LOS DIAS DE LA SEMANA
    End If
    'lunes
    If control5.Enabled = True Then
     Select Case Check1(0).Value
     Case (1)
     .lunes(0).List(it) = 2
     Case (0)
     .lunes(0).List(it) = 0
     End Select
    End If
    'martes
    If control6.Enabled = True Then
     Select Case Check1(1).Value
     Case (1)
     .martes.List(it) = 3
     Case (0)
     .martes.List(it) = 0
     End Select
    End If
    'miercoles
    If control7.Enabled = True Then
     Select Case Check1(2).Value
     Case (1)
     .miercoles.List(it) = 4
     Case (0)
     .miercoles.List(it) = 0
     End Select
    End If
'jueves
If control8.Enabled = True Then
 Select Case Check1(3).Value
 Case (1)
 .jueves.List(it) = 5
 Case (0)
 .jueves.List(it) = 0
 End Select
End If
'viernes
If control9.Enabled = True Then
 Select Case Check1(4).Value
 Case (1)
 .viernes.List(it) = 6
 Case (0)
 .viernes.List(it) = 0
 End Select
End If
'sabado
If control10.Enabled = True Then
 Select Case Check1(5).Value
 Case (1)
 .sabado.List(it) = 7
 Case (0)
 .sabado.List(it) = 0
 End Select
End If
'domingo
If control11.Enabled = True Then
 Select Case Check1(6).Value
 Case (1)
 .domingo.List(it) = 1
 Case (0)
 .domingo.List(it) = 0
 End Select
End If
' para avilitar o no la modificacion del sistema
If control14.Enabled = True Then
.liscomando.List(it) = devolver_comando
.lisdialogo.List(it) = txtd.Text
.listiempo.List(it) = DTPicker2.Second
End If
aleatorio = True
Case (True)
 If control0.Enabled = True Then
 .listado(0).List(it) = DTPicker1.Value
 End If
 If control1.Enabled = True Then
 .listado(1).List(it) = lenguaje_Menu(228)  ' entrada | salida
 End If
 If control3.Enabled = True Then
 .listado(2).List(it) = cobd(3).Text
 End If
 If control12.Enabled = True Then
 .listado(3).List(it) = Text1(1).Text
 End If
 'le aplica el filtro correspondiente
 If control2.Enabled = True Then
 .Filtro.List(it) = cobd(1).ListIndex
 '# SETEO LOS DIAS DE LA SEMANA
 End If
 'lunes
 If control5.Enabled = True Then
  Select Case Check1(0).Value
  Case (1)
  .lunes(0).List(it) = 2
  Case (0)
  .lunes(0).List(it) = 0
 End Select
End If
'martes
If control6.Enabled = True Then
 Select Case Check1(1).Value
 Case (1)
 .martes.List(it) = 3
 Case (0)
 .martes.List(it) = 0
 End Select
End If
'miercoles
If control7.Enabled = True Then
 Select Case Check1(2).Value
 Case (1)
 .miercoles.List(it) = 4
 Case (0)
 .miercoles.List(it) = 0
 End Select
End If
'jueves
If control8.Enabled = True Then
 Select Case Check1(3).Value
 Case (1)
 .jueves.List(it) = 5
 Case (0)
 .jueves.List(it) = 0
 End Select
End If
'viernes
If control9.Enabled = True Then
 Select Case Check1(4).Value
 Case (1)
 .viernes.List(it) = 6
 Case (0)
 .viernes.List(it) = 0
 End Select
End If
'sabado
If control10.Enabled = True Then
 Select Case Check1(5).Value
 Case (1)
 .sabado.List(it) = 7
 Case (0)
 .sabado.List(it) = 0
 End Select
End If
'domingo
If control11.Enabled = True Then
 Select Case Check1(6).Value
 Case (1)
 .domingo.List(it) = 1
 Case (0)
 .domingo.List(it) = 0
 End Select
End If
' para avilitar o no la modificacion del sistema
If control14.Enabled = True Then
.liscomando.List(it) = devolver_comando
.lisdialogo.List(it) = txtd.Text
.listiempo.List(it) = DTPicker2.Second
End If
aleatorio = False
End Select
ElseIf cobd(0).ListIndex = 1 Then
'ENTRADA
If control0.Enabled = True Then
.listado(0).List(it) = DTPicker1.Value
End If
If control1.Enabled = True Then
.listado(1).List(it) = lenguaje_Menu(227) ' entrada | salida
End If
If control3.Enabled = True Then
.listado(2).List(it) = cobd(2).Text
End If
If control12.Enabled = True Then
.listado(3).List(it) = Text1(0).Text
End If
'le aplica el filtro correspondiente
If control2.Enabled = True Then
.Filtro.List(it) = cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
End If
'lunes
If control5.Enabled = True Then
 Select Case Check1(0).Value
 Case (1)
 .lunes(0).List(it) = 2
 Case (0)
 .lunes(0).List(it) = 0
 End Select
End If
'martes
If control6.Enabled = True Then
 Select Case Check1(1).Value
 Case (1)
 .martes.List(it) = 3
 Case (0)
 .martes.List(it) = 0
 End Select
End If
'miercoles
If control7.Enabled = True Then
 Select Case Check1(2).Value
 Case (1)
 .miercoles.List(it) = 4
 Case (0)
 .miercoles.List(it) = 0
 End Select
End If
'jueves
If control8.Enabled = True Then
 Select Case Check1(3).Value
 Case (1)
 .jueves.List(it) = 5
 Case (0)
 .jueves.List(it) = 0
 End Select
End If
'viernes
If control9.Enabled = True Then
 Select Case Check1(4).Value
 Case (1)
 .viernes.List(it) = 6
 Case (0)
 .viernes.List(it) = 0
 End Select
End If
'sabado
If control10.Enabled = True Then
 Select Case Check1(5).Value
 Case (1)
 .sabado.List(it) = 7
 Case (0)
 .sabado.List(it) = 0
 End Select
End If
'domingo
If control11.Enabled = True Then
 Select Case Check1(6).Value
 Case (1)
 .domingo.List(it) = 1
 Case (0)
 .domingo.List(it) = 0
 End Select
End If
' para avilitar o no la modificacion del sistema
If control14.Enabled = True Then
   .liscomando.List(it) = devolver_comando
   .lisdialogo.List(it) = txtd.Text
   .listiempo.List(it) = DTPicker2.Second
End If
ElseIf cobd(0).ListIndex = 2 Then
'SALIDA
If control0.Enabled = True Then
.listado(0).List(it) = DTPicker1.Value
End If
If control1.Enabled = True Then
.listado(1).List(it) = lenguaje_Menu(228)  ' entrada | salida
End If
If control3.Enabled = True Then
.listado(2).List(it) = cobd(3).Text
End If
If control12.Enabled = True Then
.listado(3).List(it) = Text1(1).Text
End If
'le aplica el filtro correspondiente
If control2.Enabled = True Then
.Filtro.List(it) = cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
End If
'lunes
If control5.Enabled = True Then
 Select Case Check1(0).Value
 Case (1)
 .lunes(0).List(it) = 2
 Case (0)
 .lunes(0).List(it) = 0
 End Select
End If
'martes
If control6.Enabled = True Then
 Select Case Check1(1).Value
 Case (1)
 .martes.List(it) = 3
 Case (0)
 .martes.List(it) = 0
 End Select
End If
'miercoles
If control7.Enabled = True Then
 Select Case Check1(2).Value
 Case (1)
 .miercoles.List(it) = 4
 Case (0)
 .miercoles.List(it) = 0
 End Select
End If
'jueves
If control8.Enabled = True Then
 Select Case Check1(3).Value
 Case (1)
 .jueves.List(it) = 5
 Case (0)
 .jueves.List(it) = 0
 End Select
End If
'viernes
If control9.Enabled = True Then
 Select Case Check1(4).Value
 Case (1)
 .viernes.List(it) = 6
 Case (0)
 .viernes.List(it) = 0
 End Select
End If
'sabado
If control10.Enabled = True Then
 Select Case Check1(5).Value
 Case (1)
 .sabado.List(it) = 7
 Case (0)
 .sabado.List(it) = 0
 End Select
End If
'domingo
If control11.Enabled = True Then
 Select Case Check1(6).Value
 Case (1)
 .domingo.List(it) = 1
 Case (0)
 .domingo.List(it) = 0
 End Select
End If
' para avilitar o no la modificacion del sistema
If control14.Enabled = True Then
.liscomando.List(it) = devolver_comando
.lisdialogo.List(it) = txtd.Text
.listiempo.List(it) = DTPicker2.Second
End If
aleatorio = False
ElseIf cobd(0).ListIndex = 1 Then
'ENTRADA
If control0.Enabled = True Then
.listado(0).List(it) = DTPicker1.Value
End If
If control1.Enabled = True Then
.listado(1).List(it) = lenguaje_Menu(227)  ' entrada | salida
End If
If control3.Enabled = True Then
.listado(2).List(it) = cobd(2).Text
End If
If control12.Enabled = True Then
.listado(3).List(it) = Text1(0).Text
End If
'le aplica el filtro correspondiente
If control2.Enabled = True Then
.Filtro.List(it) = cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
End If
'lunes
If control5.Enabled = True Then
  Select Case Check1(0).Value
  Case (1)
  .lunes(0).List(it) = 2
  Case (0)
  .lunes(0).List(it) = 0
  End Select
End If
'martes
If control6.Enabled = True Then
 Select Case Check1(1).Value
 Case (1)
 .martes.List(it) = 3
 Case (0)
 .martes.List(it) = 0
 End Select
End If
'miercoles
If control7.Enabled = True Then
 Select Case Check1(2).Value
 Case (1)
 .miercoles.List(it) = 4
 Case (0)
 .miercoles.List(it) = 0
 End Select
End If
'jueves
If control8.Enabled = True Then
 Select Case Check1(3).Value
 Case (1)
 .jueves.List(it) = 5
 Case (0)
 .jueves.List(it) = 0
 End Select
End If
'viernes
If control9.Enabled = True Then
 Select Case Check1(4).Value
 Case (1)
 .viernes.List(it) = 6
 Case (0)
 .viernes.List(it) = 0
 End Select
End If
'sabado
If control10.Enabled = True Then
 Select Case Check1(5).Value
 Case (1)
 .sabado.List(it) = 7
 Case (0)
 .sabado.List(it) = 0
 End Select
End If
'domingo
If control11.Enabled = True Then
 Select Case Check1(6).Value
 Case (1)
 .domingo.List(it) = 1
 Case (0)
 .domingo.List(it) = 0
 End Select
End If
' para avilitar o no la modificacion del sistema
If control14.Enabled = True Then
.liscomando.List(it) = devolver_comando
.lisdialogo.List(it) = txtd.Text
.listiempo.List(it) = DTPicker2.Second
End If
End If
Next
End With
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
 Select Case Check1(0).Value ' Lunes
  Case (1)
  .lunes(0).AddItem lunes
  Case (0)
  .lunes(0).AddItem nulo
 End Select
 Select Case Check1(1).Value ' Martes
  Case (1)
  .martes.AddItem martes 'martes
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
 .Filtro.AddItem cobd(1).ListIndex
 End With
nose:
End Sub

Private Sub avilitarControles(ByVal texto0 As Boolean, _
ByVal texto1 As Boolean, ByVal int0 As Boolean, ByVal int1 As Boolean)
On Error GoTo nose
Text1(0).Enabled = texto0
Text1(1).Enabled = texto1
cobd(2).Enabled = int0
cobd(3).Enabled = int1
nose:
End Sub

Private Sub crearTimbre()
On Error GoTo nose
 Dim it As Integer
 With frmprograma
 If cobd(0).ListIndex = 0 Then
 'selecio el tipo
 Select Case (aleatorio)
  Case (False)
  .listado(0).AddItem DTPicker1.Value
  .listado(1).AddItem lenguaje_Menu(227) ' entrada | salida
  .listado(2).AddItem cobd(2).Text
  .listado(3).AddItem Text1(0).Text
  'le aplica el filtro correspondiente
  .Filtro.AddItem cobd(1).ListIndex
  '# SETEO LOS DIAS DE LA SEMANA
  'lunes
 Select Case Check1(0).Value
  Case (1)
  .lunes(0).AddItem 2
  Case (0)
  .lunes(0).AddItem 0
 End Select
'martes
Select Case Check1(1).Value
 Case (1)
 .martes.AddItem 3
 Case (0)
 .martes.AddItem 0
End Select
'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.AddItem 4
 Case (0)
 .miercoles.AddItem 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.AddItem 5
 Case (0)
 .jueves.AddItem 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.AddItem 6
 Case (0)
 .viernes.AddItem 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.AddItem 7
 Case (0)
 .sabado.AddItem 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.AddItem 1
 Case (0)
 .domingo.AddItem 0
End Select
crear_evento_automatico ' esto es para la libreria so.dll
aleatorio = True
Case (True)
.listado(0).AddItem DTPicker1.Value
.listado(1).AddItem lenguaje_Menu(228) ' entrada | salida
.listado(2).AddItem cobd(3).Text
.listado(3).AddItem Text1(1).Text
'le aplica el filtro correspondiente
.Filtro.AddItem cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
'lunes
Select Case Check1(0).Value
 Case (1)
 .lunes(0).AddItem 2
 Case (0)
 .lunes(0).AddItem 0
End Select
'martes
Select Case Check1(1).Value
 Case (1)
 .martes.AddItem 3
 Case (0)
 .martes.AddItem 0
End Select
'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.AddItem 4
 Case (0)
 .miercoles.AddItem 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.AddItem 5
 Case (0)
 .jueves.AddItem 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.AddItem 6
 Case (0)
 .viernes.AddItem 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.AddItem 7
 Case (0)
 .sabado.AddItem 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.AddItem 1
 Case (0)
 .domingo.AddItem 0
End Select
crear_evento_automatico ' esto es para la libreria so.dll
aleatorio = False
End Select
ElseIf cobd(0).ListIndex = 1 Then
'ENTRADA
.listado(0).AddItem DTPicker1.Value
.listado(1).AddItem lenguaje_Menu(227)  ' entrada | salida
.listado(2).AddItem cobd(2).Text
.listado(3).AddItem Text1(0).Text
'le aplica el filtro correspondiente
.Filtro.AddItem cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
'lunes
Select Case Check1(0).Value
 Case (1)
 .lunes(0).AddItem 2
 Case (0)
 .lunes(0).AddItem 0
End Select
'martes
Select Case Check1(1).Value
 Case (1)
 .martes.AddItem 3
 Case (0)
 .martes.AddItem 0
End Select
'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.AddItem 4
 Case (0)
 .miercoles.AddItem 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.AddItem 5
 Case (0)
 .jueves.AddItem 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.AddItem 6
 Case (0)
 .viernes.AddItem 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.AddItem 7
 Case (0)
 .sabado.AddItem 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.AddItem 1
 Case (0)
 .domingo.AddItem 0
End Select
crear_evento_automatico ' esto es para la libreria so.dll
ElseIf cobd(0).ListIndex = 2 Then
'SALIDA
.listado(0).AddItem DTPicker1.Value
.listado(1).AddItem lenguaje_Menu(228)  ' entrada | salida
.listado(2).AddItem cobd(3).Text
.listado(3).AddItem Text1(1).Text
'le aplica el filtro correspondiente
.Filtro.AddItem cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
'lunes
Select Case Check1(0).Value
 Case (1)
 .lunes(0).AddItem 2
 Case (0)
 .lunes(0).AddItem 0
 End Select
 'martes
Select Case Check1(1).Value
 Case (1)
 .martes.AddItem 3
 Case (0)
 .martes.AddItem 0
End Select
'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.AddItem 4
 Case (0)
 .miercoles.AddItem 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.AddItem 5
 Case (0)
 .jueves.AddItem 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.AddItem 6
 Case (0)
 .viernes.AddItem 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.AddItem 7
 Case (0)
 .sabado.AddItem 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.AddItem 1
 Case (0)
 .domingo.AddItem 0
End Select
crear_evento_automatico ' esto es para la libreria so.dll
aleatorio = False
ElseIf cobd(0).ListIndex = 1 Then
'ENTRADA
.listado(0).AddItem DTPicker1.Value
.listado(1).AddItem lenguaje_Menu(227)  ' entrada | salida
.listado(2).AddItem cobd(2).Text
.listado(3).AddItem Text1(0).Text
'le aplica el filtro correspondiente
.Filtro.AddItem cobd(1).ListIndex
'# SETEO LOS DIAS DE LA SEMANA
'lunes
Select Case Check1(0).Value
 Case (1)
 .lunes(0).AddItem 2
 Case (0)
 .lunes(0).AddItem 0
End Select
'martes
Select Case Check1(1).Value
 Case (1)
 .martes.AddItem 3
 Case (0)
 .martes.AddItem 0
End Select
'miercoles
Select Case Check1(2).Value
 Case (1)
 .miercoles.AddItem 4
 Case (0)
 .miercoles.AddItem 0
End Select
'jueves
Select Case Check1(3).Value
 Case (1)
 .jueves.AddItem 5
 Case (0)
 .jueves.AddItem 0
End Select
'viernes
Select Case Check1(4).Value
 Case (1)
 .viernes.AddItem 6
 Case (0)
 .viernes.AddItem 0
End Select
'sabado
Select Case Check1(5).Value
 Case (1)
 .sabado.AddItem 7
 Case (0)
 .sabado.AddItem 0
End Select
'domingo
Select Case Check1(6).Value
 Case (1)
 .domingo.AddItem 1
 Case (0)
 .domingo.AddItem 0
End Select
crear_evento_automatico
 End If
 End With
nose:
End Sub

Private Sub crear_evento_automatico()
On Error GoTo nose
 devolverString
 With frmprograma
 .liscomando.AddItem devolver_comando
 .lisdialogo.AddItem txtd.Text
 .listiempo.AddItem DTPicker1.Value
 End With
nose:
End Sub

'menu virtual del programa
Private Sub menu_Click(Index As Integer)
On Error GoTo nose
 ' si el chequed del menu selecionado esta activado se desactiva sino se activa
 If menu.Item(Index).Checked = True Then
 menu.Item(Index).Checked = False
 ElseIf menu.Item(Index).Checked = False Then
 menu.Item(Index).Checked = True
 End If
  Select Case (Index)
   Case 0
   control_activo DTPicker1, menu.Item(Index).Checked
   Case 1
   control_activo cobd.Item(0), menu.Item(Index).Checked
   Case 2
   control_activo cobd.Item(1), menu.Item(Index).Checked
   Case 3
   control_activo cobd.Item(2), menu.Item(Index).Checked
   Case 4
   control_activo cobd.Item(3), menu.Item(Index).Checked
   Case 5
   control_activo Check1.Item(0), menu.Item(Index).Checked
   Case 6
   control_activo Check1.Item(1), menu.Item(Index).Checked
   Case 7
   control_activo Check1.Item(2), menu.Item(Index).Checked
   Case 8
   control_activo Check1.Item(3), menu.Item(Index).Checked
   Case 9
   control_activo Check1.Item(4), menu.Item(Index).Checked
   Case 10
   control_activo Check1.Item(5), menu.Item(Index).Checked
   Case 11
   control_activo Check1.Item(6), menu.Item(Index).Checked
   Case 12
   control_activo Text1.Item(0), menu.Item(Index).Checked
   Case 13
   control_activo Text1.Item(1), menu.Item(Index).Checked
   Case 14
   control_activo cob1, menu.Item(Index).Checked
   control_activo txtd, menu.Item(Index).Checked
   control_activo DTPicker2, menu.Item(Index).Checked
 End Select
nose:
End Sub

Private Sub control_activo(ByVal control As Object, ByVal estado As Boolean)
 On Error GoTo nose
 control.Enabled = estado
nose:
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub salir_op(ByVal dig As Byte)
On Error GoTo nose
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmentradasalida
nose:
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo nose
 salir_op KeyAscii
nose:
End Sub

Private Sub cargar_idioma()
On Error GoTo nose
 Me.Caption = Lenguage.lenguaje_Menu(66)
 If Check2.Value = 0 Then
 Me.Check2.Caption = Lenguage.lenguaje_Menu(68)
 ElseIf Check2.Value = 1 Then
 Me.Check2.Caption = Lenguage.lenguaje_Menu(67)
 End If
 cmdfiltro.Caption = Lenguage.lenguaje_Menu(69)
 Label1.Item(0).Caption = Lenguage.lenguaje_Menu(70)
 Label1.Item(1).Caption = Lenguage.lenguaje_Menu(71)
 Label1.Item(2).Caption = Lenguage.lenguaje_Menu(72)
 Label1.Item(3).Caption = Lenguage.lenguaje_Menu(73)
 Label1.Item(4).Caption = Lenguage.lenguaje_Menu(74)
 Label1.Item(5).Caption = Lenguage.lenguaje_Menu(75)
 Check1.Item(0).Caption = Lenguage.lenguaje_Menu(76)
 Check1.Item(1).Caption = Lenguage.lenguaje_Menu(77)
 Check1.Item(2).Caption = Lenguage.lenguaje_Menu(78)
 Check1.Item(3).Caption = Lenguage.lenguaje_Menu(79)
 Check1.Item(4).Caption = Lenguage.lenguaje_Menu(80)
 Check1.Item(5).Caption = Lenguage.lenguaje_Menu(81)
 Check1.Item(6).Caption = Lenguage.lenguaje_Menu(82)
 Label1.Item(6).Caption = Lenguage.lenguaje_Menu(87)
 Label1.Item(7).Caption = Lenguage.lenguaje_Menu(88)
 cmdCancelar.Caption = Lenguage.lenguaje_Menu(89)
 cmdmod.Caption = Lenguage.lenguaje_Menu(92)
 txtd.Text = Lenguage.lenguaje_Menu(101)
 labinfo.Caption = Lenguage.lenguaje_Menu(102)
 lbld.Caption = Lenguage.lenguaje_Menu(103)
 Label2.Caption = Lenguage.lenguaje_Menu(104)
 
 menu(0).Caption = Lenguage.lenguaje_Menu(105)
 menu(1).Caption = Lenguage.lenguaje_Menu(106)
 menu(2).Caption = Lenguage.lenguaje_Menu(107)
 menu(3).Caption = Lenguage.lenguaje_Menu(108)
 menu(4).Caption = Lenguage.lenguaje_Menu(109)
 menu(5).Caption = Lenguage.lenguaje_Menu(110)
 menu(6).Caption = Lenguage.lenguaje_Menu(111)
 menu(7).Caption = Lenguage.lenguaje_Menu(112)
 menu(8).Caption = Lenguage.lenguaje_Menu(113)
 menu(9).Caption = Lenguage.lenguaje_Menu(114)
 menu(10).Caption = Lenguage.lenguaje_Menu(115)
 menu(11).Caption = Lenguage.lenguaje_Menu(116)
 menu(12).Caption = Lenguage.lenguaje_Menu(117)
 menu(13).Caption = Lenguage.lenguaje_Menu(118)
 menu(14).Caption = Lenguage.lenguaje_Menu(119)
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
       Load mc(r)
       mc(r).Caption = cargar
       mc(r).Visible = True
       r = r + 1
       End If
       Loop
       Close #1
       r = 0
nose:
End Sub
Private Sub mc_Click(Index As Integer)
On Error GoTo nose
If con_c = True Then
 Text1(0).Text = mc.Item(Index).Caption
ElseIf con_c = False Then
 Text1(1).Text = mc.Item(Index).Caption
End If
nose:
End Sub

'fin del generador de horarios programables by Martin Grasso
