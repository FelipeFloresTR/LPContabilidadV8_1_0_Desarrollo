VERSION 5.00
Begin VB.Form FrmActivoFijo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activo Fijo - Detalle Tributario"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   Icon            =   "FrmActivoFijo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Area Negocio / Centro de Gestion"
      Height          =   2355
      Left            =   7920
      TabIndex        =   103
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox Cb_CGestion 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   960
         Width           =   1875
      End
      Begin VB.ComboBox Cb_CNegocio 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestion:"
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area Negocio:"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Fr_DepLey21256 
      Caption         =   "Depreciación Ley 21.256 Art. 3 (22 Bis TTO Ley 21.210)"
      Height          =   975
      Left            =   5640
      TabIndex        =   100
      Top             =   3480
      Width           =   5775
      Begin VB.Frame Fr_AcogeLey21256 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   180
         TabIndex        =   101
         Top             =   300
         Width           =   2175
         Begin VB.OptionButton Op_AcogeLey21256 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton Op_AcogeLey21256 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   20
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Se acoge:"
            Height          =   195
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Tx_requisitosLey21256 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3300
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton Bt_InfoAdic 
      Caption         =   "Info, Adicional"
      Height          =   615
      Left            =   11760
      Picture         =   "FrmActivoFijo.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Frame Fr_DepLey21210 
      Caption         =   "Depreciación Régimen Ley 21.210"
      Height          =   975
      Left            =   1440
      TabIndex        =   93
      Top             =   3480
      Width           =   4095
      Begin VB.Frame Fr_TipoDepLey21210 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   180
         TabIndex        =   96
         Top             =   600
         Width           =   3735
         Begin VB.OptionButton Op_TipoDepLey21210 
            Caption         =   "Instantánea e Inmediata *"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   2115
         End
         Begin VB.OptionButton Op_TipoDepLey21210 
            Caption         =   "Araucanía *"
            Height          =   255
            Index           =   2
            Left            =   2580
            TabIndex        =   17
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.Frame Fr_AcogeLey21210 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   180
         TabIndex        =   94
         Top             =   300
         Width           =   2175
         Begin VB.OptionButton Op_AcogeLey21210 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   15
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton Op_AcogeLey21210 
            Caption         =   "Si"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   14
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Se acoge:"
            Height          =   195
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Tx_requisitosLey21210 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3060
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crédito Art. 33 bis "
      Height          =   795
      Left            =   1440
      TabIndex        =   79
      Top             =   2580
      Width           =   9975
      Begin VB.CheckBox Ch_Cred4Porc 
         Caption         =   "Créd. art. 33 bis"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Tx_ValCred33 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   12
         ToolTipText     =   "Valor Crédito art. 33 bis forzado para no pasar tope"
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   1620
         TabIndex        =   90
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Tx_requisitosCred33bis 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7500
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Index           =   18
         Left            =   2100
         TabIndex        =   85
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(en blanco: el sistema lo calcula)"
         Height          =   195
         Index           =   19
         Left            =   3960
         TabIndex        =   80
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Bt_DetCosto 
      Caption         =   "Det. Financiero"
      Height          =   615
      Left            =   11760
      Picture         =   "FrmActivoFijo.frx":04E9
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   300
      Picture         =   "FrmActivoFijo.frx":0925
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   76
      Top             =   300
      Width           =   885
   End
   Begin VB.Frame Fr_DepAnoAct 
      Caption         =   "Meses a depreciar año actual"
      Height          =   1515
      Left            =   1440
      TabIndex        =   73
      Top             =   4560
      Width           =   9975
      Begin VB.TextBox Tx_DepDecimaParte2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   32
         Top             =   1020
         Width           =   615
      End
      Begin VB.OptionButton Op_TipoDep 
         Caption         =   "Art. 31, 5 bis  Inc. 1 1/10 *"
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   31
         Top             =   1020
         Width           =   2715
      End
      Begin VB.TextBox Tx_VidaUtilAnos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "Vida útil en años para la depreciación Acelerada Esp. (1/10) vida útil"
         Top             =   1020
         Width           =   615
      End
      Begin VB.OptionButton Op_TipoDep 
         Caption         =   "Décima Parte Art. 31, 5 bis inc. 2° *"
         Height          =   315
         Index           =   4
         Left            =   3120
         TabIndex        =   28
         Top             =   660
         Width           =   2835
      End
      Begin VB.OptionButton Op_TipoDep 
         Caption         =   "Instant. Art. 31,5 bis, inc.1° *"
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   24
         Top             =   300
         Width           =   2775
      End
      Begin VB.TextBox Tx_DepDecimaParte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   29
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox Tx_DepInstant 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Tx_DepNormal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Tx_DepAcelerada 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   660
         Width           =   615
      End
      Begin VB.OptionButton Op_TipoDep 
         Caption         =   "Normal:"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton Op_TipoDep 
         Caption         =   "Acelerada:"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   26
         Left            =   6660
         TabIndex        =   98
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Tx_RequisitosDecimaParte2 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7500
         TabIndex        =   97
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vida útil total:"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   92
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "años"
         Height          =   195
         Index           =   25
         Left            =   2040
         TabIndex        =   91
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   5040
         TabIndex        =   89
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Tx_RequisitosDecParte 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7500
         TabIndex        =   55
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Tx_RequisitosInstant 
         AutoSize        =   -1  'True
         Caption         =   "Requisitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7500
         TabIndex        =   54
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   21
         Left            =   6660
         TabIndex        =   78
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   20
         Left            =   6660
         TabIndex        =   77
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   6
         Left            =   2040
         TabIndex        =   75
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   8
         Left            =   2040
         TabIndex        =   74
         Top             =   720
         Width           =   450
      End
   End
   Begin VB.Frame Fr_DepAnoAnt 
      Caption         =   "Depreciación realizada al 31 de Diciembre del año anterior"
      Height          =   1815
      Left            =   1440
      TabIndex        =   69
      Top             =   6180
      Width           =   9975
      Begin VB.TextBox Tx_DepDecimaparte2Hist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   43
         Top             =   1260
         Width           =   615
      End
      Begin VB.OptionButton Op_TipoDepHist 
         Caption         =   "Acelerada Esp. (1/10) vida útil  Art. 31, 5 bis inc. 1° LIR"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   42
         Top             =   1260
         Width           =   2775
      End
      Begin VB.OptionButton Op_TipoDepHist 
         Caption         =   "Acelerada Esp. (1/10) vida útil  Art. 31, 5 bis inc. 2° LIR"
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   39
         Top             =   780
         Width           =   2775
      End
      Begin VB.TextBox Tx_DepDecimaParteHist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   40
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox Tx_DepInstantHist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6000
         TabIndex        =   36
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Tx_DepAceleradaHist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox Tx_DepNormalHist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton Op_TipoDepHist 
         Caption         =   "Normal:"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   33
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton Op_TipoDepHist 
         Caption         =   "Acelerada:"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox Tx_DepAcumHist 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   41
         ToolTipText     =   "Dep. histórica acumulada al 31 dic. del año anterior:"
         Top             =   1260
         Width           =   1155
      End
      Begin VB.OptionButton Op_TipoDepHist 
         Caption         =   "Instantánea Art. 31,5 bis, inc.1° LIR"
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   35
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   255
         Index           =   27
         Left            =   6660
         TabIndex        =   99
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   23
         Left            =   6660
         TabIndex        =   87
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   22
         Left            =   6660
         TabIndex        =   86
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   10
         Left            =   1980
         TabIndex        =   72
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   12
         Left            =   1980
         TabIndex        =   71
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dep. hist. acum.:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   70
         Top             =   1320
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clasificación Activo Fijo"
      Height          =   1215
      Left            =   1440
      TabIndex        =   65
      Top             =   9000
      Width           =   9975
      Begin VB.ComboBox Cb_Cuenta 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Obs. :    Seleccione una cuenta para clasificar los activos fijos en el  informe."
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   67
         Top             =   780
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Act. Fijo:"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   66
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Venta o Baja"
      Height          =   795
      Left            =   1440
      TabIndex        =   60
      Top             =   8100
      Width           =   9975
      Begin VB.TextBox Tx_FechaVenta 
         Height          =   315
         Left            =   3000
         TabIndex        =   45
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton Bt_FechaVenta 
         Caption         =   "?"
         Height          =   315
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   300
         Width           =   215
      End
      Begin VB.TextBox Tx_IVAVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7080
         TabIndex        =   48
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_NetoVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5160
         TabIndex        =   47
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox Cb_TipoMovAF 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   9
         Left            =   2460
         TabIndex        =   64
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Index           =   15
         Left            =   6720
         TabIndex        =   63
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Neto:"
         Height          =   195
         Index           =   14
         Left            =   4740
         TabIndex        =   62
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Fr_Compra 
      Caption         =   "Compra"
      Height          =   2355
      Left            =   1440
      TabIndex        =   56
      Top             =   120
      Width           =   6435
      Begin VB.TextBox Tx_Neto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   1860
         Width           =   1155
      End
      Begin VB.TextBox Tx_IVA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2940
         TabIndex        =   9
         Top             =   1860
         Width           =   1155
      End
      Begin VB.TextBox Tx_VidaUtil 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4980
         TabIndex        =   10
         Top             =   1860
         Width           =   795
      End
      Begin VB.CommandButton Bt_FechaUtilizacion 
         Height          =   315
         Left            =   5580
         Picture         =   "FrmActivoFijo.frx":0F0D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   780
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   2460
         Picture         =   "FrmActivoFijo.frx":1217
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   780
         Width           =   230
      End
      Begin VB.CheckBox Ch_TotalmenteDepreciado 
         Caption         =   "Activo Fijo Totalmente Depreciado"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox Ch_NoDepreciable 
         Caption         =   "Activo Fijo No Depreciable"
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox Tx_FechaUtilizacion 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox Tx_Descrip 
         Height          =   315
         Left            =   1320
         MaxLength       =   80
         TabIndex        =   7
         Top             =   1500
         Width           =   4455
      End
      Begin VB.TextBox Tx_Cantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1140
         Width           =   1155
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Neto:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   84
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   83
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vida útil:"
         Height          =   195
         Index           =   7
         Left            =   4320
         TabIndex        =   82
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "meses"
         Height          =   195
         Index           =   11
         Left            =   5820
         TabIndex        =   81
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de utilización:"
         Height          =   195
         Index           =   13
         Left            =   2940
         TabIndex        =   68
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   59
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   58
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha compra:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   57
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   11760
      TabIndex        =   53
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   11760
      TabIndex        =   52
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nota: (*) Está sujeto a que usted cumpla con las condiciones para imputarlo."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1440
      TabIndex        =   88
      Top             =   10440
      Width           =   5385
   End
End
Attribute VB_Name = "FrmActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFecha As Long
Dim lNeto As Double
Dim lIVA As Double
Dim lDescrip As String

Dim lIdCuentaExento As Long
Dim lIdCuentaAfecto As Long

Dim lIdDoc As Long
'Dim lIdDocVenta As Long

Dim lidComp As Long
Dim lIdMovComp As Long

Dim lTipoLib As Integer

Dim lIdActFijo As Long

Dim lRc As Integer
Dim lOper As Integer

Dim lInLoad As Boolean

Dim ModFechaUtil As Boolean

Dim lModif As Boolean

Dim lTipoDep As Integer, lTipoDepHist As Integer
Dim lInClearDep As Boolean

Dim lPatenteRol As String
Dim lNombreProy As String
Dim lFechaProy As Long

'2861733
Dim lArea As Long
Dim lCentro As Long
'2861733

Private Sub Bt_Cancel_Click()
   
   lRc = vbCancel
   Unload Me
   
End Sub


Private Sub Bt_Componentes_Click()
   Dim Frm As FrmAFCompsFicha
   
   Set Frm = New FrmAFCompsFicha
   Call Frm.FEdit(lIdActFijo)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_DetCosto_Click()
   Dim Frm As FrmAFFicha
   
   If lModif Then
      If MsgBox1("Se grabarán los datos antes de ver el detalle financiero." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   
      If Valida() Then
         
         Call SaveAll
      Else
         Exit Sub
         
      End If
   End If
   
   Set Frm = New FrmAFFicha
   Call Frm.FEdit(lIdActFijo)
   Set Frm = Nothing

End Sub

Private Sub Bt_InfoAdic_Click()
   Dim Frm As FrmActFijoInfoAdic
   
   If lModif Then
      If MsgBox1("Se grabarán los datos antes de ver el detalle financiero." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   
      If Valida(False) Then
         
         Call SaveAll
      Else
         Exit Sub
         
      End If
   End If

   Set Frm = New FrmActFijoInfoAdic
   
   If lOper = O_VIEW Then
      Call Frm.FView(lIdActFijo)
      
   Else
      Call Frm.FEdit(lIdActFijo, Tx_Descrip, Op_TipoDepLey21210(DEP_LEY21210_INST), lPatenteRol, lNombreProy, lFechaProy)
      
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()
   
   If Valida() Then
      
      Call SaveAll
      lRc = vbOK
      
      Unload Me
      
   End If
   
End Sub
Private Sub Cb_Cuenta_Click()
   lModif = True

End Sub

Private Sub Cb_TipoMovAF_Click()

   If Cb_TipoMovAF = "" Then
      Call SetTxDate(Tx_FechaVenta, 0)
      Tx_NetoVenta = ""
      Tx_IVAVenta = ""
   End If
   lModif = True

End Sub

Private Sub Ch_Cred4Porc_Click()
   If Ch_Cred4Porc <> 0 Then
      Tx_ValCred33.Enabled = True
   Else
      Tx_ValCred33.Enabled = False
      Tx_ValCred33 = ""
   End If
   lModif = True

End Sub

Private Sub Ch_NoDepreciable_Click()

   Call EnableFieldsNoDepTotDep
   
   lModif = True
      
End Sub
Private Sub EnableFieldsNoDepTotDep()
   Dim i As Integer
   Dim Enable As Boolean
   
   Enable = True
   
   If Ch_NoDepreciable <> 0 Or Ch_TotalmenteDepreciado <> 0 Then
      Enable = False
      Ch_Cred4Porc = 0
      Tx_ValCred33 = ""
      Tx_FechaUtilizacion = ""
      Tx_VidaUtil = ""
      Call ClearDep(0)
      Call ClearDepHist(0)
      Call SetTxRO(Tx_VidaUtilAnos, True)
      lTipoDep = 0
      
   End If
      
   Ch_Cred4Porc.Enabled = Enable
   Tx_ValCred33.Enabled = Enable
   Tx_FechaUtilizacion.Enabled = Enable
   Bt_FechaUtilizacion.Enabled = Enable
   Tx_VidaUtil.Enabled = Enable
   Fr_AcogeLey21210.Enabled = Enable
   Fr_TipoDepLey21210.Enabled = Enable
   
   
   For i = 1 To DEP_DECIMAPARTE
      Op_TipoDep(i).Enabled = Enable
      Op_TipoDepHist(i).Enabled = Enable
   Next i
   Tx_DepNormal.Enabled = Enable
   Tx_DepAcelerada.Enabled = Enable
   Tx_DepInstant.Enabled = Enable
   Tx_DepDecimaParte.Enabled = Enable
   Tx_DepDecimaParte2.Enabled = Enable
   Tx_DepNormalHist.Enabled = Enable
   Tx_DepAceleradaHist.Enabled = Enable
   Tx_DepInstantHist.Enabled = Enable
   Tx_DepDecimaParteHist.Enabled = Enable
   Tx_DepDecimaparte2Hist.Enabled = Enable
      
   Call SetTxRO(Tx_VidaUtilAnos, True)
   If lTipoDep = DEP_DECIMAPARTE Then
      Call SetTxRO(Tx_VidaUtilAnos, Not Enable)
   End If
      

   If Ch_NoDepreciable <> 0 Then
      Tx_DepAcumHist = ""
      Tx_DepAcumHist.Enabled = False
   Else
      Tx_DepAcumHist.Enabled = True
   End If

End Sub

Private Sub Ch_TotalmenteDepreciado_Click()

   Call EnableFieldsNoDepTotDep
   lModif = True
   
End Sub


Private Sub Form_Load()
   Dim i As Integer
   Dim TipoMov As Integer
   Dim Q1

   lInLoad = True
   
   Call BtFechaImg(Bt_Fecha)
   Call BtFechaImg(Bt_FechaUtilizacion)
   Call BtFechaImg(Bt_FechaVenta)
   
   '2861733 tema 2
   Q1 = "SELECT Descripcion,idCCosto FROM CentroCosto WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbAddItem(Cb_CGestion, " ", 0)
   Call FillCombo(Cb_CGestion, DbMain, Q1, -1)

   Q1 = "SELECT Descripcion, idAreaNegocio FROM AreaNegocio WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbAddItem(Cb_CNegocio, " ", 0)
   Call FillCombo(Cb_CNegocio, DbMain, Q1, -1)


    If lArea > 0 Then
         Call SelItem(Cb_CNegocio, lArea)
    End If

    If lCentro > 0 Then
         Call SelItem(Cb_CGestion, lCentro)
    End If

'2861733    tema 2

   'esta combo sólo se usa para la venta o baja
   Cb_TipoMovAF.AddItem ""
   Cb_TipoMovAF.ItemData(Cb_TipoMovAF.NewIndex) = 0
   
   For i = MOVAF_VENTA To MAX_TIPOMOVAF
      Cb_TipoMovAF.AddItem gMovActivoFijo(i)
      Cb_TipoMovAF.ItemData(Cb_TipoMovAF.NewIndex) = i
   Next i
   
   Cb_TipoMovAF.ListIndex = 0
   
   Q1 = "SELECT Descripcion, IdCuenta FROM Cuentas WHERE Atrib" & ATRIB_ACTIVOFIJO & "<> 0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(Cb_Cuenta, DbMain, Q1, 0, True)
   
   Call EnableForm(Me, Not lOper = O_VIEW)
   Bt_DetCosto.Enabled = True
   
   If lOper = O_NEW Then
   
      If lFecha > 0 Then
         Call SetTxDate(Tx_Fecha, lFecha)
      Else
         Call SetTxDate(Tx_Fecha, Now)
      End If
      
      If lTipoLib = LIB_COMPRAS Then
         TipoMov = 0   'sólo se usa para la venta o baja
      ElseIf lTipoLib = LIB_VENTAS Then
         TipoMov = MOVAF_VENTA
      Else
         TipoMov = -1
      End If
         
      For i = 0 To Cb_TipoMovAF.ListCount - 1
         If Cb_TipoMovAF.ItemData(i) = TipoMov Then
            Cb_TipoMovAF.ListIndex = i
            Exit For
         End If
      Next i
      
      Tx_Cantidad = 1
      
      Op_TipoDep(DEP_NORMAL) = 1
      Op_TipoDepHist(DEP_NORMAL) = 1
      
      
      Tx_Descrip = lDescrip
      
      If lTipoLib = LIB_VENTAS Then
      
         Tx_NetoVenta = Format(lNeto, BL_NUMFMT)
         Tx_IVAVenta = Format(lIVA, BL_NUMFMT)
         
      Else        'LibCompras o Comprobante
      
         Tx_Neto = Format(lNeto, BL_NUMFMT)
         
         If lIVA = 0 And lTipoLib = 0 And lNeto > 0 Then 'viene de comprobante, cuando viene de libro de compras se manda el IVA si es afecto
            Tx_IVA = Format(lNeto * gIVA, BL_NUMFMT)
         Else
            Tx_IVA = Format(lIVA, BL_NUMFMT)
         End If
         
      End If
      
      If lIdCuentaExento > 0 And vFmt(Tx_IVA) = 0 Then
         Call SelItem(Cb_Cuenta, lIdCuentaExento)
      ElseIf lIdCuentaAfecto > 0 Then
         Call SelItem(Cb_Cuenta, lIdCuentaAfecto)
      End If
   
      Call EnableForm(Me, gEmpresa.FCierre = 0)
      Call SetTxRO(Tx_VidaUtilAnos, True)
      Bt_DetCosto.Enabled = True

   Else
         
      Call LoadAll
      
      If lOper = O_VIEW Then
         Call EnableForm(Me, False)
         Bt_DetCosto.Enabled = True
      End If
         
   End If
   
   Call SetupPriv
   
   ModFechaUtil = False
   lInLoad = False
   
   If lOper = O_NEW Then
      lModif = True
            
'   ElseIf lNeto > 0 Then  'viene de doc o comprobante
'      lModif = True

   Else
      lModif = False
   End If
   
   
End Sub
'crea un nuevo mov. de activo fijo asociado a un IdDoc con datos propuestos a partir del doc.
Public Function FNewFromDoc(ByVal IdDoc As Long, ByVal Fecha As Long, ByVal Neto As Double, ByVal IVA As Double, ByVal Descrip As String, ByVal TipoLib As Integer, ByVal IdCuentaAfecto As Long, ByVal IdCuentaExento As Long)
   
   lOper = O_NEW
   
   lTipoLib = TipoLib
   
   lIdDoc = 0
   If lTipoLib = LIB_COMPRAS Then   'uno u otro
      lIdDoc = IdDoc
'   Else
'      lIdDocVenta = 0    'IdDoc
   End If
   
   lidComp = 0
   lIdMovComp = 0
   
   lFecha = Fecha
   lNeto = Neto
   lIVA = IVA
   lDescrip = Descrip
   
   lIdCuentaExento = IdCuentaExento
   lIdCuentaAfecto = IdCuentaAfecto

   Me.Show vbModal
   
   FNewFromDoc = lRc
   
End Function

'2861733
'crea un nuevo mov. de activo fijo asociado a un IdDoc con datos propuestos a partir del doc.
Public Function FNewFromDocActFijo(ByVal IdDoc As Long, ByVal Fecha As Long, ByVal Neto As Double, ByVal IVA As Double, ByVal Descrip As String, ByVal TipoLib As Integer, ByVal IdCuentaAfecto As Long, ByVal IdCuentaExento As Long, ByVal IdArea As Long, ByVal IdCentro As Long)

   lOper = O_NEW

   lTipoLib = TipoLib

   lIdDoc = 0
   If lTipoLib = LIB_COMPRAS Then   'uno u otro
      lIdDoc = IdDoc
'   Else
'      lIdDocVenta = 0    'IdDoc
   End If

   lidComp = 0
   lIdMovComp = 0

   lFecha = Fecha
   lNeto = Neto
   lIVA = IVA
   lDescrip = Descrip

   lIdCuentaExento = IdCuentaExento
   lIdCuentaAfecto = IdCuentaAfecto

   lArea = IdArea
   lCentro = IdCentro

   Me.Show vbModal

   FNewFromDocActFijo = lRc

End Function
'2861733

'crea un nuevo mov. de activo fijo asociado a un IdDoc con datos propuestos a partir del doc.
Public Function FNewFromComp(ByVal idcomp As Long, ByVal IdMovComp As Long, ByVal Fecha As Long, ByVal Neto As Double, ByVal IVA As Double, ByVal Descrip As String, ByVal IdCuenta As Long)
   
   lOper = O_NEW
   
   lIdDoc = 0          'uno u otro
   lidComp = idcomp
   lIdMovComp = IdMovComp
   
   lFecha = Fecha
   lNeto = Neto
   lIVA = IVA
   lDescrip = Descrip
   
   lIdCuentaExento = IdCuenta  'ponemos la misma ya que es una sola
   lIdCuentaAfecto = IdCuenta

   Me.Show vbModal
   
   FNewFromComp = lRc
   
End Function
Public Function FEdit(ByVal IdActFijo As Long)
   
   lOper = O_EDIT
   
   lIdActFijo = IdActFijo
      
   Me.Show vbModal
   
   FEdit = lRc
   
End Function
Public Sub FView(ByVal IdActFijo As Long)
   
   lOper = O_VIEW
   
   lIdActFijo = IdActFijo
   
   Me.Show vbModal
      
End Sub

Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha)
   lModif = True
      
   Set Frm = Nothing
   
   Call EnableDepEspecial
  
End Sub
Private Sub Bt_FechaUtilizacion_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaUtilizacion)
   
   Call EnableDepEspecial
   Call CalcMesesDep
   lModif = True

   Set Frm = Nothing
End Sub
Private Sub Bt_FechaVenta_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaVenta)
   
   Set Frm = Nothing
   
   lModif = True

End Sub


Private Sub Op_AcogeLey21210_Click(Index As Integer)
   Call EnableDepEspecial
         
   If Op_AcogeLey21210(VAL_SI) Then
      Op_AcogeLey21256(VAL_NO) = True
   End If

End Sub

Private Sub Op_AcogeLey21256_Click(Index As Integer)

   Call EnableDepEspecial
   
   If Op_AcogeLey21256(VAL_SI) Then
      Op_AcogeLey21210(VAL_NO) = True
      Op_TipoDepLey21210(DEP_LEY21210_INST) = False
      Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = False
      Fr_DepAnoAct.Enabled = False
      Call ClearDep(0)
      lTipoDep = 0
   Else
      Fr_DepAnoAct.Enabled = True
   End If
   
End Sub

Private Sub Op_TipoDep_Click(Index As Integer)

   lModif = True
   lTipoDep = Index
   Call ClearDep(Index)
   
   If Index = DEP_INSTANTANEA Then
      Tx_VidaUtil = 12
   End If

   If Index <> DEP_DECIMAPARTE And Index <> DEP_DECIMAPARTE2 Then
      Tx_VidaUtilAnos = ""
      Call SetTxRO(Tx_VidaUtilAnos, True)
   Else
      Call SetTxRO(Tx_VidaUtilAnos, False)
   End If
   
   Call CalcMesesDep
   
End Sub

Private Sub Op_TipoDepHist_Click(Index As Integer)
   lModif = True
   lTipoDepHist = Index
   Call ClearDepHist(Index)
   
End Sub

Private Sub Op_TipoDepLey21210_Click(Index As Integer)
   Dim i As Integer

   If Op_TipoDepLey21210(DEP_LEY21210_INST) Then
      Fr_DepAnoAct.Enabled = True
   Else                                   'Araucanía se deprecia 100% en período
      Fr_DepAnoAct.Enabled = False
      Call ClearDep(0)
      lTipoDep = 0
   End If

End Sub

Private Sub Tx_Cantidad_Change()
   lModif = True

End Sub

Private Sub Tx_Cantidad_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub


Private Sub Tx_DepAcelerada_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDep(DEP_ACELERADA) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepAceleradaHist_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDepHist(DEP_ACELERADA) = 1
   End If
   
   lModif = True

End Sub

Private Sub Tx_DepAceleradaHist_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_DepAcumHist_GotFocus()
   Call NumGotFocus(Tx_DepAcumHist)
   lModif = True

End Sub

Private Sub Tx_DepDecimaParte_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDep(DEP_DECIMAPARTE) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepDecimaParte2_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDep(DEP_DECIMAPARTE2) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepDecimaParte2Hist_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDepHist(DEP_DECIMAPARTE2) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepDecimaParteHist_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDepHist(DEP_DECIMAPARTE) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepDecimaParteHist_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_DepInstant_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDep(DEP_INSTANTANEA) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepInstantHist_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDepHist(DEP_INSTANTANEA) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepInstantHist_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_DepNormal_Change()

   If Not lInLoad And Not lInClearDep Then
      Op_TipoDep(DEP_NORMAL) = 1
   End If
   lModif = True

End Sub

Private Sub Tx_DepNormalHist_Change()
   If Not lInLoad And Not lInClearDep Then
      Op_TipoDepHist(DEP_NORMAL) = True
   End If
   lModif = True

End Sub

Private Sub Tx_DepNormalHist_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_Descrip_Change()
   lModif = True

End Sub

Private Sub Tx_Fecha_GotFocus()
   Call DtGotFocus(Tx_Fecha)
End Sub

Private Sub Tx_FechaUtilizacion_Change()
   ModFechaUtil = True
   lModif = True
   
'   If Not lInLoad Then
'      Call CalcMesesDep
'   End If

End Sub

Private Sub Tx_FechaVenta_Change()
   lModif = True

End Sub

Private Sub Tx_FechaVenta_GotFocus()
   Call DtGotFocus(Tx_FechaVenta)
End Sub
Private Sub Tx_FechaUtilizacion_GotFocus()
   Call DtGotFocus(Tx_FechaUtilizacion)
End Sub

Private Sub Tx_Fecha_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_Fecha) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_Fecha)
   
   Call DtLostFocus(Tx_Fecha)
   
   Call EnableDepEspecial
     
End Sub
Private Sub Tx_FechaVenta_LostFocus()
   
   If Trim$(Tx_FechaVenta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FechaVenta)
   
End Sub
Private Sub Tx_FechaUtilizacion_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_FechaUtilizacion) = "" Then
      If ModFechaUtil Then
         Call CalcMesesDep
      End If
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FechaUtilizacion)
   
   Fecha = GetTxDate(Tx_FechaUtilizacion)
   
   If Val(Tx_FechaUtilizacion.Tag) <> Fecha Then  'cambió la fecha
      Call CalcMesesDep
      Tx_FechaUtilizacion.Tag = Fecha
   End If

   Call EnableDepEspecial
   
End Sub

Private Sub Tx_Fecha_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FechaVenta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FechaUtilizacion_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub Tx_IVA_Change()
   lModif = True

End Sub

Private Sub Tx_IVA_GotFocus()
   Call NumGotFocus(Tx_IVA)
End Sub

Private Sub Tx_IVA_LostFocus()
   Tx_IVA = Format(vFmt(Tx_IVA), NUMFMT)
End Sub

Private Sub Tx_IVAVenta_Change()
   lModif = True

End Sub

Private Sub Tx_IVAVenta_GotFocus()
   Call NumGotFocus(Tx_IVAVenta)
End Sub

Private Sub Tx_IVAVenta_LostFocus()
   Tx_IVAVenta = Format(vFmt(Tx_IVAVenta), NUMFMT)
End Sub

Private Sub Tx_Neto_Change()
   lModif = True

End Sub

Private Sub Tx_Neto_GotFocus()
   Call NumGotFocus(Tx_Neto)
End Sub

Private Sub Tx_Neto_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_NetoVenta_Change()
   lModif = True

End Sub

Private Sub Tx_NetoVenta_GotFocus()
   Call NumGotFocus(Tx_NetoVenta)
End Sub


Private Sub Tx_requisitosCred33bis_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewCredArt33bis
   Set Frm = Nothing

End Sub

Private Sub Tx_RequisitosDecimaParte2_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewDecimaParte2
   Set Frm = Nothing

End Sub

Private Sub Tx_RequisitosDecParte_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewDecimaParte
   Set Frm = Nothing

End Sub

Private Sub Tx_RequisitosInstant_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewDepInstant
   Set Frm = Nothing
   
End Sub



Private Sub Tx_requisitosLey21210_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewLey21210
   Set Frm = Nothing

End Sub

Private Sub Tx_requisitosLey21256_Click()
   Dim Frm As FrmRequisitos
   
   Set Frm = New FrmRequisitos
   Call Frm.FViewLey21256
   Set Frm = Nothing

End Sub

Private Sub Tx_ValCred33_Change()
   lModif = True

End Sub

Private Sub Tx_ValCred33_GotFocus()
   Call NumGotFocus(Tx_ValCred33)

End Sub

Private Sub Tx_ValCred33_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_ValCred33_LostFocus()
   Tx_ValCred33 = Trim(Tx_ValCred33)
   If Tx_ValCred33 <> "" Then
      Tx_ValCred33 = Format(vFmt(Tx_ValCred33), NUMFMT)
   End If

End Sub

Private Sub Tx_VidaUtil_Change()
   lModif = True

End Sub

Private Sub Tx_VidaUtil_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_NetoVenta_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_IVA_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_IVAVenta_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_DepNormal_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_DepAcelerada_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_DepAcumHist_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Ctrl As Control
         
   If lIdActFijo <= 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT IdDoc, IdComp, IdMovComp, TipoMovAF, Fecha, Cantidad, Descrip, FechaUtilizacion, NoDepreciable "
   Q1 = Q1 & ", Neto, IVA, Cred4Porc, ValCred33, DepNormal, DepAcelerada, TipoDep, DepNormalHist, DepAceleradaHist "
   Q1 = Q1 & ", DepAcumHist, TipoDepHist, NetoVenta, IVAVenta, FechaVentaBaja, IdCuenta, VidaUtil, FImported"
   Q1 = Q1 & ", DepInstant, DepInstantHist, DepDecimaParte, DepDecimaParteHist, TipoDepLey21210, DepLey21256, VidaUtilAnos"
   Q1 = Q1 & ", DepDecimaParte2, DepDecimaParte2Hist, TotalmenteDepreciado, PatenteRol, NombreProy, FechaProy "
   '2861733 tema 2
   Q1 = Q1 & ", idCCosto, IdAreaNeg "
   '2861733 tema 2
   
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & " WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      lIdDoc = vFld(Rs("IdDoc"))
      'lIdDocVenta = vFld(Rs("IdDocVenta"))
   
      lidComp = vFld(Rs("IdComp"))
      lIdMovComp = vFld(Rs("IdMovComp"))
      
      Cb_TipoMovAF.ListIndex = 0
      If vFld(Rs("TipoMovAF")) >= MOVAF_VENTA Then
         Cb_TipoMovAF.ListIndex = FindItem(Cb_TipoMovAF, vFld(Rs("TipoMovAF")))
      End If
      Call SetTxDate(Tx_FechaUtilizacion, vFld(Rs("FechaUtilizacion")))
      Call SetTxDate(Tx_Fecha, vFld(Rs("Fecha")))
      Call SetTxDate(Tx_FechaVenta, vFld(Rs("FechaVentaBaja")))
      Tx_Cantidad = Format(vFld(Rs("Cantidad")), NUMFMT)
      Tx_Descrip = vFld(Rs("Descrip"), True)
      Tx_Neto = Format(vFld(Rs("Neto")), NUMFMT)
      Tx_IVA = Format(vFld(Rs("IVA")), NUMFMT)
      Tx_NetoVenta = Format(vFld(Rs("NetoVenta")), NUMFMT)
      Tx_IVAVenta = Format(vFld(Rs("IVAVenta")), NUMFMT)
      If IsNull(Rs("Valcred33")) Or vFld(Rs("Valcred33")) < 0 Then
         Tx_ValCred33 = ""
      Else
         Tx_ValCred33 = Format(vFld(Rs("Valcred33")), NUMFMT)
      End If
      Ch_Cred4Porc = Abs(vFld(Rs("Cred4Porc")) > 0)
      Tx_VidaUtil = Format(vFld(Rs("VidaUtil")), NUMFMT)
      
      lTipoDep = vFld(Rs("TipoDep"))
      If lTipoDep > 0 Then
         Op_TipoDep(lTipoDep) = True
      End If
      
      If vFld(Rs("TipoDepLey21210")) <> 0 Then
         Op_AcogeLey21210(VAL_SI) = True
         Op_TipoDepLey21210(vFld(Rs("TipoDepLey21210"))) = True
      Else
         Op_AcogeLey21210(VAL_NO) = True
      End If
         
      If vFld(Rs("DepLey21256")) <> 0 Then
         Op_AcogeLey21256(VAL_SI) = True
      Else
         Op_AcogeLey21256(VAL_NO) = True
      End If
        
      Tx_DepNormal = IIf(lTipoDep = DEP_NORMAL, vFld(Rs("DepNormal")), "")
      Tx_DepAcelerada = IIf(lTipoDep = DEP_ACELERADA, vFld(Rs("DepAcelerada")), "")
      Tx_DepInstant = IIf(lTipoDep = DEP_INSTANTANEA, vFld(Rs("DepInstant")), "")
      Tx_DepDecimaParte = IIf(lTipoDep = DEP_DECIMAPARTE, vFld(Rs("DepDecimaParte")), "")
      Tx_DepDecimaParte2 = IIf(lTipoDep = DEP_DECIMAPARTE2, vFld(Rs("DepDecimaParte2")), "")
      
      lTipoDepHist = vFld(Rs("TipoDepHist"))
      If lTipoDepHist > 0 Then
         Op_TipoDepHist(lTipoDepHist) = True
      Else
         Op_TipoDepHist(DEP_NORMAL) = False
      End If
      Tx_DepNormalHist = vFld(Rs("DepNormalHist"))
      Tx_DepAceleradaHist = vFld(Rs("DepAceleradaHist"))
      Tx_DepInstantHist = vFld(Rs("DepInstantHist"))
      Tx_DepDecimaParteHist = vFld(Rs("DepDecimaParteHist"))
      Tx_DepDecimaparte2Hist = vFld(Rs("DepDecimaParte2Hist"))
      
      Tx_DepAcumHist = Format(vFld(Rs("DepAcumHist")), NUMFMT)
      
      Call SelItem(Cb_Cuenta, vFld(Rs("IdCuenta")))
      
      Call EnableForm(Me, gEmpresa.FCierre = 0)   'tiene que estar antes de las deshabilitaciones siguientes, si no, se pierden
      Bt_DetCosto.Enabled = True
      
      
      If GetTxDate(Tx_FechaUtilizacion) < gFechaInicioDepInstantanea Then
         Call SetTxRO(Tx_DepInstant, True)
         Call SetTxRO(Tx_DepDecimaParte, True)
         Op_TipoDep(DEP_INSTANTANEA).Enabled = False
         Op_TipoDep(DEP_DECIMAPARTE).Enabled = False
         
      Else
         Call SetTxRO(Tx_DepInstant, False)
         Call SetTxRO(Tx_DepDecimaParte, False)
         Op_TipoDep(DEP_INSTANTANEA).Enabled = True
         Op_TipoDep(DEP_DECIMAPARTE).Enabled = True
      
      End If
        
      Call SetTxRO(Tx_VidaUtilAnos, True)
      If lTipoDep = DEP_DECIMAPARTE Then
         Call SetTxRO(Tx_VidaUtilAnos, False)
         Tx_VidaUtilAnos = vFld(Rs("VidaUtilAnos"))
      End If
    
      Call EnableDepEspecial(False)

      
      Ch_NoDepreciable = Abs(vFld(Rs("NoDepreciable")) > 0)       'estas dos tienen que estar después de la cargaa de los otros datos, porque algunos se ponen en cero dependiendo de estos dos Checks
      Ch_TotalmenteDepreciado = Abs(vFld(Rs("TotalmenteDepreciado")) > 0)

      lPatenteRol = vFld(Rs("PatenteRol"))
      lNombreProy = vFld(Rs("NombreProy"))
      lFechaProy = vFld(Rs("FechaProy"))
      
      If vFld(Rs("FImported")) <> 0 Then    'viene del año anterior
         'Fr_Compra.Enabled = False
         Ch_TotalmenteDepreciado.Enabled = False
         Ch_NoDepreciable.Enabled = False
         Call SetTxRO(Tx_Fecha, True)
         Bt_Fecha.Enabled = False
         Call SetTxRO(Tx_FechaUtilizacion, True)
         Bt_FechaUtilizacion.Enabled = False
         Call SetTxRO(Tx_Cantidad, True)
         Call SetTxRO(Tx_Descrip, True)
         Call SetTxRO(Tx_Neto, True)
         Call SetTxRO(Tx_IVA, True)
         Call SetTxRO(Tx_VidaUtil, True)
         Call SetTxRO(Tx_ValCred33, True)
         
         'Fr_DepAnoAnt.Enabled = False
         For i = 1 To DEP_DECIMAPARTE
            Op_TipoDepHist(i).Enabled = False
         Next i
         Call SetTxRO(Tx_DepNormalHist, True)
         Call SetTxRO(Tx_DepAceleradaHist, True)
         Call SetTxRO(Tx_DepInstantHist, True)
         Call SetTxRO(Tx_DepDecimaParteHist, True)
         
      End If
         
         '2861733 tema 2
         Call SelItem(Cb_CGestion, vFld(Rs("idCCosto")))
         Call SelItem(Cb_CNegocio, vFld(Rs("IdAreaNeg")))
         '2861733 tema 2

         
   End If
   
   Call CloseRs(Rs)

End Sub

Private Sub SaveAll()
   Dim Q1 As String
   Dim FldArray(2) As AdvTbAddNew_t

   If lIdActFijo = 0 Then   'new

      FldArray(0).FldName = "IdDoc"
      FldArray(0).FldValue = lIdDoc
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "IdEmpresa"
      FldArray(1).FldValue = gEmpresa.id
      FldArray(1).FldIsNum = True
                  
      FldArray(2).FldName = "Ano"
      FldArray(2).FldValue = gEmpresa.Ano
      FldArray(2).FldIsNum = True
      
      lIdActFijo = AdvTbAddNewMult(DbMain, "MovActivoFijo", "IdActFijo", FldArray)
      
      If lIdActFijo = 0 Then
         Exit Sub
      End If
      
   End If
   
   If vFmt(Tx_DepNormalHist) = 0 And vFmt(Tx_DepAceleradaHist) = 0 And vFmt(Tx_DepInstantHist) = 0 And vFmt(Tx_DepDecimaParteHist) = 0 Then
      lTipoDepHist = 0
   End If
   
   Q1 = "UPDATE MovActivoFijo SET "
   Q1 = Q1 & "  IdDoc = " & lIdDoc
   'Q1 = Q1 & ", IdDocVenta = " & lIdDocVenta
   Q1 = Q1 & ", IdComp = " & lidComp
   Q1 = Q1 & ", IdMovComp = " & lIdMovComp
   Q1 = Q1 & ", TipoMovAF = " & IIf(ItemData(Cb_TipoMovAF) <= 0, MOVAF_COMPRA, ItemData(Cb_TipoMovAF))
   Q1 = Q1 & ", Fecha = " & GetTxDate(Tx_Fecha)
   Q1 = Q1 & ", FechaUtilizacion = " & GetTxDate(Tx_FechaUtilizacion)
   Q1 = Q1 & ", FechaVentaBaja = " & GetTxDate(Tx_FechaVenta)
   Q1 = Q1 & ", Cantidad = " & vFmt(Tx_Cantidad)
   Q1 = Q1 & ", Descrip = '" & ParaSQL(Tx_Descrip) & "'"
   Q1 = Q1 & ", Neto = " & vFmt(Tx_Neto)
   Q1 = Q1 & ", VidaUtil = " & vFmt(Tx_VidaUtil)
   Q1 = Q1 & ", IVA = " & vFmt(Tx_IVA)
   Q1 = Q1 & ", NetoVenta = " & vFmt(Tx_NetoVenta)
   Q1 = Q1 & ", IVAVenta = " & vFmt(Tx_IVAVenta)
   Q1 = Q1 & ", Cred4Porc = " & Abs(Ch_Cred4Porc <> 0)
   If Trim(Tx_ValCred33) = "" Then
      Q1 = Q1 & ", ValCred33 = -1"
   Else
      Q1 = Q1 & ", ValCred33 = " & vFmt(Tx_ValCred33)
   End If
   Q1 = Q1 & ", NoDepreciable = " & Abs(Ch_NoDepreciable <> 0)
   Q1 = Q1 & ", TotalmenteDepreciado = " & Abs(Ch_TotalmenteDepreciado <> 0)
   Q1 = Q1 & ", TipoDep = " & lTipoDep
   Q1 = Q1 & ", TipoDepLey21210 = " & IIf(Op_AcogeLey21210(VAL_SI), IIf(Op_TipoDepLey21210(DEP_LEY21210_INST), DEP_LEY21210_INST, DEP_LEY21210_ARAUCANIA), 0)
   Q1 = Q1 & ", DepLey21256 = " & IIf(Op_AcogeLey21256(VAL_SI), 1, 0)
   Q1 = Q1 & ", DepNormal = " & vFmt(Tx_DepNormal)
   Q1 = Q1 & ", DepAcelerada = " & vFmt(Tx_DepAcelerada)
   Q1 = Q1 & ", DepInstant = " & vFmt(Tx_DepInstant)
   Q1 = Q1 & ", DepDecimaParte = " & vFmt(Tx_DepDecimaParte)
   Q1 = Q1 & ", DepDecimaParte2 = " & vFmt(Tx_DepDecimaParte2)
   If lTipoDep = DEP_DECIMAPARTE Or lTipoDep = DEP_DECIMAPARTE2 Then
      Q1 = Q1 & ", VidaUtilAnos = " & vFmt(Tx_VidaUtilAnos)
   Else
      Q1 = Q1 & ", VidaUtilAnos = 0"
   End If
   Q1 = Q1 & ", DepNormalHist = " & vFmt(Tx_DepNormalHist)
   Q1 = Q1 & ", DepAceleradaHist = " & vFmt(Tx_DepAceleradaHist)
   Q1 = Q1 & ", DepInstantHist = " & vFmt(Tx_DepInstantHist)
   Q1 = Q1 & ", DepDecimaParteHist = " & vFmt(Tx_DepDecimaParteHist)
   Q1 = Q1 & ", DepDecimaParte2Hist = " & vFmt(Tx_DepDecimaParteHist)
   Q1 = Q1 & ", DepAcumHist = " & vFmt(Tx_DepAcumHist)
   Q1 = Q1 & ", TipoDepHist = " & lTipoDepHist
   Q1 = Q1 & ", PatenteRol = '" & ParaSQL(lPatenteRol) & "'"
   Q1 = Q1 & ", NombreProy = '" & ParaSQL(lNombreProy) & "'"
   Q1 = Q1 & ", FechaProy = " & lFechaProy
   Q1 = Q1 & ", IdCuenta = " & IIf(ItemData(Cb_Cuenta) < 0, 0, ItemData(Cb_Cuenta))
   
   '2861733
   Q1 = Q1 & ", idCCosto = " & IIf(ItemData(Cb_CGestion) < 0, 0, ItemData(Cb_CGestion))
   Q1 = Q1 & ", IdAreaNeg = " & IIf(ItemData(Cb_CNegocio) < 0, 0, ItemData(Cb_CNegocio))
   '2861733
   
   Q1 = Q1 & " WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   lModif = False
   ModFechaUtil = False

   
End Sub

Private Function Valida(Optional ByVal ValidaInfoAdic As Boolean = True) As Boolean
   Dim Dep As Integer
   Dim DepHist As Integer
   Dim FVenta As Long, FechaUtiliz As Long
   
   Valida = False
   
   If GetTxDate(Tx_Fecha) = 0 Then
      MsgBox1 "Falta ingresar la fecha de compra.", vbExclamation
      Exit Function
   End If
   
'   If Year(GetTxDate(Tx_Fecha)) > gEmpresa.Ano Or Year(GetTxDate(Tx_FechaUtilizacion)) > gEmpresa.Ano Then    'Cambio solicitado por Nicolás Catrin 31/01/2019
   If Year(GetTxDate(Tx_Fecha)) > gEmpresa.Ano Then
      MsgBox1 "Este activo fijo corresponde al período siguiente.", vbExclamation
      Exit Function
   End If
      
   If Year(GetTxDate(Tx_Fecha)) < Year(Now) < 2 Then
      If MsgBox1("Recuerde que si ingresa activos fijos de años anteriores, deberá tener actualizados los valores de IPC de esos años." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
      
      
   If Ch_NoDepreciable = 0 And Ch_TotalmenteDepreciado = 0 Then
      If GetTxDate(Tx_FechaUtilizacion) = 0 Then
         MsgBox1 "Falta ingresar la fecha de utilización.", vbExclamation
         Exit Function
      End If
      
      If GetTxDate(Tx_FechaUtilizacion) < GetTxDate(Tx_Fecha) Then
         MsgBox1 "La fecha de utilización debe ser posterior a la fecha de compra.", vbExclamation
         Exit Function
      End If
   End If
   
   If GetTxDate(Tx_FechaVenta) > 0 Then
      If ItemData(Cb_TipoMovAF) <= 0 Then
         MsgBox1 "Falta indicar el movimiento de Venta o Baja del activo.", vbExclamation
         Exit Function
      End If
   
      If GetTxDate(Tx_FechaVenta) < GetTxDate(Tx_Fecha) Then
         MsgBox1 "La fecha de venta o baja debe ser posterior a la fecha de compra.", vbExclamation
         Exit Function
      End If
   End If
   
   If vFmt(Tx_Cantidad) <= 0 Then
      MsgBox1 "Falta ingresar la cantidad.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Descrip) = "" Then
      MsgBox1 "Falta ingresar la descripción.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_Neto) <= 0 Then
      MsgBox1 "Falta ingresar el valor neto.", vbExclamation
      Exit Function
   End If
   
'   If ValidaInfoAdic And lPatenteRol = "" Then
'      MsgBox1 "Falta ingresar Placa Patente, Rol o Inscripción (según proceda)." & vbCrLf & vbCrLf & "Utilice el botón Info. Adicional para ingresarla.", vbExclamation
'      Exit Function
'   End If
'
   
   If Ch_NoDepreciable = 0 And Ch_TotalmenteDepreciado = 0 Then
   
      If vFmt(Tx_VidaUtil) <= 0 Then
         MsgBox1 "Falta ingresar la vida útil del activo.", vbExclamation
         Exit Function
      End If
      
      FechaUtiliz = GetTxDate(Tx_FechaUtilizacion)
      
      If Op_AcogeLey21210(VAL_SI) Then
      
         If FechaUtiliz < gFechaInicioDepLey21210 Or FechaUtiliz > gFechaTerminoDepLey21210 Then
            MsgBox1 "La fecha de utilización del bien no permite acogerse a Depreciación Ley 21.210.", vbExclamation
            Exit Function
         End If
     
         If Op_TipoDepLey21210(DEP_LEY21210_INST) = 0 And Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = 0 And Op_AcogeLey21256(VAL_SI) = 0 Then
            MsgBox1 "Falta seleccionar el tipo de Depreciación Ley 21.210 a la que se acogerá.", vbExclamation
            Exit Function
         End If
         
         If ValidaInfoAdic And Op_TipoDepLey21210(DEP_LEY21210_INST) = True And (lNombreProy = "" Or lFechaProy = 0) Then
            MsgBox1 "Falta ingresar información del proyecto." & vbCrLf & vbCrLf & "Utilice el botón Info. Adicional para ingresarla.", vbExclamation
            Exit Function
         End If
                   
      End If
      
      If Op_AcogeLey21256(VAL_SI) Then
      
         If FechaUtiliz < gFechaInicioDepLey21256 Or FechaUtiliz > gFechaTerminoDepLey21256 Then
            MsgBox1 "La fecha de utilización del bien no permite acogerse a Depreciación Ley 21.256.", vbExclamation
            Exit Function
         End If
              
'         If ValidaInfoAdic And Op_TipoDepLey21210(DEP_LEY21210_INST) = True And (lNombreProy = "" Or lFechaProy = 0) Then
'            MsgBox1 "Falta ingresar información del proyecto." & vbCrLf & vbCrLf & "Utilice el botón Info. Adicional para ingresarla.", vbExclamation
'            Exit Function
'         End If
                   
      End If
           
      If Op_AcogeLey21256(VAL_SI) And Op_AcogeLey21210(VAL_SI) Then
         MsgBox1 "No puede acogerse a Ley 21.210 y Ley 21.256 al mismo tiempo. Debe seleccionar una de las dos", vbExclamation
         Exit Function
      End If
           
      If Op_TipoDep(DEP_NORMAL) = 0 And Op_TipoDep(DEP_ACELERADA) = 0 And Op_TipoDep(DEP_INSTANTANEA) = 0 And Op_TipoDep(DEP_DECIMAPARTE) = 0 And Op_TipoDep(DEP_DECIMAPARTE2) = 0 Then
         If Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = 0 Then
            If Op_TipoDepLey21210(DEP_LEY21210_INST) = 0 Then
               If Op_AcogeLey21256(VAL_SI) = 0 Then
                  MsgBox1 "Falta seleccionar el tipo de Depreciación que se aplicará.", vbExclamation
                  Exit Function
               End If
            Else
               MsgBox1 "Falta seleccionar el tipo de Depreciación que se aplicará para el 50% restante del bien, que no se deprecia Instantánea e Inmediatamente.", vbExclamation
               Exit Function
            End If
         End If
      End If
      
      If vFmt(Tx_DepNormal) = 0 And Op_TipoDep(DEP_NORMAL) <> 0 Then
         If MsgBox1("Atención: Falta ingresar la Depreciación Normal." & vbCrLf & vbCrLf & "¿Desea continuar?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      If vFmt(Tx_DepAcelerada) = 0 And Op_TipoDep(DEP_ACELERADA) <> 0 Then
         If MsgBox1("Atención: Falta ingresar la Depreciación Acelerada." & vbCrLf & vbCrLf & "¿Desea continuar?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      If Op_TipoDep(DEP_INSTANTANEA) <> 0 Or Op_TipoDep(DEP_DECIMAPARTE) <> 0 Then
         If GetTxDate(Tx_Fecha) < gFechaInicioDepInstantanea Then
            MsgBox1 "No es posible seleccionar este tipo de depreciación" & vbCrLf & "dado que la compra es anterior al 1 de Octubre 2014.", vbExclamation
            Exit Function
         End If
      End If
      
      If Op_TipoDep(DEP_DECIMAPARTE2) <> 0 Then
      
         If GetTxDate(Tx_Fecha) < gFechaInicioDepDecimaParte2 Then
            MsgBox1 "No es posible seleccionar este tipo de depreciación" & vbCrLf & "dado que la compra es anterior al 1 de enero 2020.", vbExclamation
            Exit Function
         End If
         
      End If
     
      If vFmt(Tx_DepInstant) = 0 And Op_TipoDep(DEP_INSTANTANEA) <> 0 Then
         If MsgBox1("Atención: Falta ingresar la Depreciación Instantánea." & vbCrLf & vbCrLf & "¿Desea continuar?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      If Op_TipoDep(DEP_INSTANTANEA) <> 0 And vFmt(Tx_VidaUtil) <> 12 Then
         MsgBox1 "La vida útil de un activo fijo para depreciación instantánea es de 12 meses", vbExclamation
         Exit Function
      End If
      
      If vFmt(Tx_DepDecimaParte) = 0 And Op_TipoDep(DEP_DECIMAPARTE) <> 0 Then
         If MsgBox1("Atención: Falta ingresar la Depreciación Décima Parte." & vbCrLf & vbCrLf & "¿Desea continuar?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      If vFmt(Tx_DepDecimaParte2) = 0 And Op_TipoDep(DEP_DECIMAPARTE2) <> 0 Then
         If MsgBox1("Atención: Falta ingresar la Depreciación Art. 31, 5 bis  Inc. 1 1/10." & vbCrLf & vbCrLf & "¿Desea continuar?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      If Op_TipoDep(DEP_NORMAL) <> 0 Then
         Dep = vFmt(Tx_DepNormal)
      ElseIf Op_TipoDep(DEP_ACELERADA) <> 0 Then
         Dep = vFmt(Tx_DepAcelerada)
      End If

      If GetTxDate(Tx_FechaUtilizacion) < DateSerial(gEmpresa.Ano - 1, 12, 31) Then
         
         If vFmt(Tx_DepNormalHist) = 0 And vFmt(Tx_DepAceleradaHist) = 0 And vFmt(Tx_DepInstantHist) = 0 And vFmt(Tx_DepDecimaParteHist) = 0 Then
            MsgBox1 "Falta ingresar la Depreciación Histórica.", vbExclamation
            Exit Function
              
         Else
            If vFmt(Tx_DepNormalHist) = 0 And Op_TipoDepHist(DEP_NORMAL) <> 0 Then
               MsgBox1 "Falta ingresar la Depreciación Normal Histórica.", vbExclamation
               Exit Function
            End If
            
            If vFmt(Tx_DepAceleradaHist) = 0 And Op_TipoDepHist(DEP_ACELERADA) <> 0 Then
               MsgBox1 "Falta ingresar la Depreciación Acelerada Histórica.", vbExclamation
               Exit Function
            End If
            
            If Op_TipoDepHist(DEP_INSTANTANEA) <> 0 Or Op_TipoDepHist(DEP_DECIMAPARTE) <> 0 Then
               If GetTxDate(Tx_Fecha) < gFechaInicioDepInstantanea Then
                  MsgBox1 "No es posible seleccionar este tipo de depreciación histórica" & vbCrLf & "dado que la compra es anterior al 1 de Octubre 2014.", vbExclamation
                  Exit Function
               End If
            End If

            If vFmt(Tx_DepInstantHist) = 0 And Op_TipoDepHist(DEP_INSTANTANEA) <> 0 Then
               MsgBox1 "Falta ingresar la Depreciación Instantánea Histórica.", vbExclamation
               Exit Function
            End If
            
            If vFmt(Tx_DepDecimaParteHist) = 0 And Op_TipoDepHist(DEP_DECIMAPARTE) <> 0 Then
               MsgBox1 "Falta ingresar la Depreciación Acelerada Esp. (1/10) Histórica.", vbExclamation
               Exit Function
            End If
            
            If Op_TipoDepHist(DEP_NORMAL) <> 0 Then
               DepHist = vFmt(Tx_DepNormalHist)
            ElseIf Op_TipoDepHist(DEP_ACELERADA) <> 0 Then
               DepHist = vFmt(Tx_DepAceleradaHist)
            End If

            If vFmt(Tx_DepAcumHist) <= 0 Then
               MsgBox1 "Falta ingresar la Depreciación Histórica Acumulada.", vbExclamation
               Exit Function
            End If
            
         End If
         
      ElseIf GetTxDate(Tx_FechaUtilizacion) > DateSerial(gEmpresa.Ano - 1, 12, 31) Then
         If vFmt(Tx_DepNormalHist) <> 0 Or vFmt(Tx_DepAceleradaHist) <> 0 Or vFmt(Tx_DepInstantHist) <> 0 Or vFmt(Tx_DepDecimaParteHist) <> 0 Or vFmt(Tx_DepAcumHist) <> 0 Then
            MsgBox1 "No corresponde ingresar Depreciación Histórica, dado que la fecha de utilización del bien corresponde a este año.", vbExclamation
            Exit Function
         End If
         
      End If
   
'      If vFmt(Tx_VidaUtil) < Dep + DepHist Then
'         MsgBox1 "El total de meses de vida útil del activo es menor que la suma de los meses a depreciar en el año actual y los meses depreciados históricamente.", vbExclamation
'         Exit Function
'      End If
'
   End If
   
   If Ch_TotalmenteDepreciado <> 0 And gEmpresa.Region <> 9 Then
      If Abs(vFmt(Tx_Neto) - vFmt(Tx_DepAcumHist)) > 1 Then
         MsgBox1 "Si este bien está totalmente depreciado, la diferencia entre el Valor Neto y la Depreciación Acumulada Histórica no debe ser superior a 1.", vbExclamation
         Exit Function
      End If
      
'      FVenta = GetTxDate(Tx_FechaVenta)
'      If FVenta > 0 And FVenta > DateSerial(gEmpresa.Ano - 1, 12, 31) Then
'         MsgBox1 "Si el activo fijo está totalmente depreciado, no puede tener fecha de venta o baja este año.", vbExclamation
'         Exit Function
'      End If
      
   End If
   
   If ItemData(Cb_TipoMovAF) <> 0 And GetTxDate(Tx_FechaVenta) = 0 Then
      MsgBox1 "Falta ingresar la fecha de venta o baja.", vbExclamation
      Exit Function
   End If
   
   If ItemData(Cb_Cuenta) = 0 Then
      MsgBox1 "Falta seleccionar la cuenta para clasificar los Activos Fijos en el informe.", vbExclamation
      Exit Function
   End If
   
   Valida = True
   
End Function

Private Sub Tx_Neto_LostFocus()

   Tx_Neto = Format(vFmt(Tx_Neto), NUMFMT)
   Tx_IVA = Format(vFmt(Tx_Neto) * gIVA, NUMFMT)
   
End Sub
Private Sub Tx_VidaUtil_LostFocus()

   Tx_VidaUtil = Format(vFmt(Tx_VidaUtil), NUMFMT)
   Call CalcAjusteVidaUtil
End Sub
Private Sub Tx_NetoVenta_LostFocus()

   Tx_NetoVenta = Format(vFmt(Tx_NetoVenta), NUMFMT)
   Tx_IVAVenta = Format(vFmt(Tx_NetoVenta) * gIVA, NUMFMT)
   
End Sub
Private Sub Tx_DepAcumHist_LostFocus()

   Tx_DepAcumHist = Format(vFmt(Tx_DepAcumHist), NUMFMT)
   
End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_DOCS Or PRV_ADM_COMP) Then
      Call EnableForm(Me, False)
      Bt_DetCosto.Enabled = True
   End If
   
End Function

Private Sub CalcMesesDep()
   Dim Meses As Long
   
   If Tx_FechaUtilizacion = "" Then
      If Not lInLoad Then
         MsgBox1 "Debe ingresar la fecha de utilización del activo.", vbExclamation
      End If
      Exit Sub
   End If
   
   Meses = DateDiff("m", GetTxDate(Tx_FechaUtilizacion), DateSerial(gEmpresa.Ano - 1, 12, 31)) + 1

   If Op_TipoDep(DEP_INSTANTANEA) <> 0 Then
   
      If Meses > 0 Then          'lo empezó a usar el año pasado
         
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano - 1, 12, 31) Then
            Tx_DepInstantHist = 0
            
         ElseIf month(GetTxDate(Tx_FechaUtilizacion)) = 12 And Day(GetTxDate(Tx_FechaUtilizacion)) = 31 Then
            Tx_DepInstantHist = Meses - 1
            
         Else
            Tx_DepInstantHist = Meses
         End If
         
         Op_TipoDepHist(DEP_INSTANTANEA) = True
         Tx_DepInstant = 12 - vFmt(Tx_DepInstantHist)
         If vFmt(Tx_DepInstant) < 0 Then
            Tx_DepInstant = 0
         End If
         
      Else
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano, 12, 31) Then
            Tx_DepInstant = 0
         Else        'lo empezó a usar el año pasado
            Tx_DepInstant = DateDiff("m", GetTxDate(Tx_FechaUtilizacion), DateSerial(gEmpresa.Ano, 12, 31)) + 1
         End If
         If vFmt(Tx_DepInstant) < 0 Then
            Tx_DepInstant = 0
         End If
      End If
   
   
   ElseIf Op_TipoDep(DEP_DECIMAPARTE) <> 0 Then   'Idem Dep Instantánea
   
      If Meses > 0 Then          'lo empezó a usar el año pasado
         
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano - 1, 12, 31) Then
            Tx_DepInstantHist = 0
            
         ElseIf month(GetTxDate(Tx_FechaUtilizacion)) = 12 And Day(GetTxDate(Tx_FechaUtilizacion)) = 31 Then
            Tx_DepDecimaParteHist = Meses - 1
            
         Else
            Tx_DepDecimaParteHist = Meses
         End If
         
         Op_TipoDepHist(DEP_DECIMAPARTE) = True
         Tx_DepDecimaParte = 12 - vFmt(Tx_DepDecimaParteHist)
         If vFmt(Tx_DepDecimaParte) < 0 Then
            Tx_DepDecimaParte = 0
         End If
         
      Else         'este año
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano, 12, 31) Then  'justo a fin de este año
            Tx_DepDecimaParte = 0
         Else
            Tx_DepDecimaParte = DateDiff("m", GetTxDate(Tx_FechaUtilizacion), DateSerial(gEmpresa.Ano, 12, 31)) + 1
         End If
         If vFmt(Tx_DepDecimaParte) < 0 Then
            Tx_DepDecimaParte = 0
         End If
      End If
   
   ElseIf Op_TipoDep(DEP_DECIMAPARTE2) <> 0 Then   'Idem Dep Instantánea y Decima Parte
   
      If Meses > 0 Then          'lo empezó a usar el año pasado
         
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano - 1, 12, 31) Then
            Tx_DepInstantHist = 0
            
         ElseIf month(GetTxDate(Tx_FechaUtilizacion)) = 12 And Day(GetTxDate(Tx_FechaUtilizacion)) = 31 Then
            Tx_DepDecimaparte2Hist = Meses - 1
            
         Else
            Tx_DepDecimaparte2Hist = Meses
         End If
         
         Op_TipoDepHist(DEP_DECIMAPARTE2) = True
         Tx_DepDecimaParte2 = 12 - vFmt(Tx_DepDecimaparte2Hist)
         If vFmt(Tx_DepDecimaParte2) < 0 Then
            Tx_DepDecimaParte2 = 0
         End If
         
      Else         'este año
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano, 12, 31) Then  'justo a fin de este año
            Tx_DepDecimaParte2 = 0
         Else
            Tx_DepDecimaParte2 = DateDiff("m", GetTxDate(Tx_FechaUtilizacion), DateSerial(gEmpresa.Ano, 12, 31)) + 1
         End If
         If vFmt(Tx_DepDecimaParte2) < 0 Then
            Tx_DepDecimaParte2 = 0
         End If
      End If


   ElseIf Op_TipoDep(DEP_NORMAL) <> 0 Then
         
      If Meses > 0 Then          'lo empezó a usar el año pasado
      
         Tx_DepNormal = 12
   
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano - 1, 12, 31) Then
            Tx_DepNormalHist = 0
            
         ElseIf month(GetTxDate(Tx_FechaUtilizacion)) = 12 And Day(GetTxDate(Tx_FechaUtilizacion)) = 31 Then
            Tx_DepNormalHist = Meses - 1
            
         Else
            Tx_DepNormalHist = Meses
         End If
         
         Op_TipoDepHist(DEP_NORMAL) = True
         
         Call CalcAjusteVidaUtil
        
      Else        'este año
         If GetTxDate(Tx_FechaUtilizacion) = DateSerial(gEmpresa.Ano, 12, 31) Then
            Tx_DepNormal = 0
         Else
            Tx_DepNormal = Abs(DateDiff("m", GetTxDate(Tx_FechaUtilizacion), DateSerial(gEmpresa.Ano, 12, 31))) + 1
         End If
      
      End If
         
   End If
   
End Sub

Private Sub CalcAjusteVidaUtil()
   Dim Meses As Long
   Dim YaDep As Long
   
   If vFmt(Tx_VidaUtil) <= 0 Or lTipoDepHist = 0 Then
      Exit Sub
   End If
   
   Select Case lTipoDepHist
      Case DEP_NORMAL
         YaDep = vFmt(Tx_DepNormalHist)
      Case DEP_ACELERADA
         YaDep = vFmt(Tx_DepAceleradaHist)
      Case DEP_INSTANTANEA
         YaDep = vFmt(Tx_DepInstantHist)
      Case DEP_DECIMAPARTE
         YaDep = vFmt(Tx_DepDecimaParteHist)
   End Select
      
   Meses = vFmt(Tx_VidaUtil) - YaDep
   If Meses >= 0 And Meses <= 12 Then
      Tx_DepNormal = Meses
'   ElseIf Meses >= 12 Then
'      Tx_DepNormal = 12
   End If
      
End Sub

Private Sub ClearDep(ByVal Index)
   Dim i As Integer
   
   lInClearDep = True
   
   Select Case Index
      Case DEP_NORMAL
         Tx_DepAcelerada = ""
         Tx_DepInstant = ""
         Tx_DepDecimaParte = ""
         Tx_DepDecimaParte2 = ""
      Case DEP_ACELERADA
         Tx_DepNormal = ""
         Tx_DepInstant = ""
         Tx_DepDecimaParte = ""
         Tx_DepDecimaParte2 = ""
      Case DEP_INSTANTANEA
         Tx_DepNormal = ""
         Tx_DepAcelerada = ""
         Tx_DepDecimaParte = ""
         Tx_DepDecimaParte2 = ""
      Case DEP_DECIMAPARTE
         Tx_DepNormal = ""
         Tx_DepAcelerada = ""
         Tx_DepInstant = ""
         Tx_DepDecimaParte2 = ""
      Case DEP_DECIMAPARTE2
         Tx_DepNormal = ""
         Tx_DepAcelerada = ""
         Tx_DepInstant = ""
         Tx_DepDecimaParte = ""
      Case Else
         Tx_DepNormal = ""
         Tx_DepAcelerada = ""
         Tx_DepInstant = ""
         Tx_DepDecimaParte = ""
         Tx_DepDecimaParte2 = ""
         Tx_VidaUtilAnos = ""
         For i = 1 To DEP_DECIMAPARTE
            Op_TipoDep(i) = 0
         Next i

   End Select
   
   lInClearDep = False
End Sub
Private Sub ClearDepHist(ByVal Index)
   Dim i As Integer

   lInClearDep = True
   Select Case Index
      Case DEP_NORMAL
         Tx_DepAceleradaHist = ""
         Tx_DepInstantHist = ""
         Tx_DepDecimaParteHist = ""
         Tx_DepDecimaparte2Hist = ""
      Case DEP_ACELERADA
         Tx_DepNormalHist = ""
         Tx_DepInstantHist = ""
         Tx_DepDecimaParteHist = ""
         Tx_DepDecimaparte2Hist = ""
      Case DEP_INSTANTANEA
         Tx_DepNormalHist = ""
         Tx_DepAceleradaHist = ""
         Tx_DepDecimaParteHist = ""
         Tx_DepDecimaparte2Hist = ""
      Case DEP_DECIMAPARTE
         Tx_DepNormalHist = ""
         Tx_DepAceleradaHist = ""
         Tx_DepInstantHist = ""
         Tx_DepDecimaparte2Hist = ""
      Case DEP_DECIMAPARTE2
         Tx_DepNormalHist = ""
         Tx_DepAceleradaHist = ""
         Tx_DepInstantHist = ""
         Tx_DepDecimaParteHist = ""
      Case Else
         Tx_DepNormalHist = ""
         Tx_DepAceleradaHist = ""
         Tx_DepInstantHist = ""
         Tx_DepDecimaParteHist = ""
         Tx_DepDecimaparte2Hist = ""
         Op_AcogeLey21210(VAL_SI) = False
         Op_AcogeLey21210(VAL_NO) = False
         Op_TipoDepLey21210(DEP_LEY21210_INST) = False
         Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = False
         For i = 1 To DEP_DECIMAPARTE
            Op_TipoDepHist(i) = 0
         Next i

   End Select
   
   lInClearDep = False

End Sub



Private Sub Tx_VidaUtilAnos_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_VidaUtilAnos_LostFocus()
   If lTipoDep = DEP_DECIMAPARTE Or lTipoDep = DEP_DECIMAPARTE2 Then
      Tx_VidaUtil = Int(vFmt(Tx_VidaUtilAnos) * 12 / 10)   'no se debe aproximar sino cortar los decimales
      If vFmt(Tx_VidaUtil) < 12 Then
         Tx_VidaUtil = 12
      End If
   End If
End Sub

Private Sub EnableDepEspecial(Optional ByVal CambiarTipoDep As Boolean = True)
   Dim Fecha As Long, FechaUtiliz As Long
   Static MsgLey21210 As Boolean
   Static EnEnableDepEspecial As Boolean
   
   If EnEnableDepEspecial Then
      Exit Sub
   End If
   
   EnEnableDepEspecial = True

   Fecha = GetTxDate(Tx_Fecha)
   FechaUtiliz = GetTxDate(Tx_FechaUtilizacion)
   
   If FechaUtiliz > 0 And (FechaUtiliz < gFechaInicioDepInstantanea Or FechaUtiliz > gFechaTerminoDepInstantanea) Then
      
      If CambiarTipoDep Then
         Op_TipoDep(DEP_NORMAL) = True
      End If
   
      Call SetTxRO(Tx_DepInstant, True)
      Call SetTxRO(Tx_DepDecimaParte, True)
      Op_TipoDep(DEP_INSTANTANEA).Enabled = False
      Op_TipoDep(DEP_DECIMAPARTE).Enabled = False
      
      Call SetTxRO(Tx_DepInstantHist, True)
      Call SetTxRO(Tx_DepDecimaParteHist, True)
      Op_TipoDepHist(DEP_INSTANTANEA).Enabled = False
      Op_TipoDepHist(DEP_DECIMAPARTE).Enabled = False
      
   ElseIf FechaUtiliz >= gFechaInicioDepInstantanea And FechaUtiliz <= gFechaTerminoDepInstantanea Then
      Call SetTxRO(Tx_DepInstant, False)
      Call SetTxRO(Tx_DepDecimaParte, False)
      Op_TipoDep(DEP_INSTANTANEA).Enabled = True
      Op_TipoDep(DEP_DECIMAPARTE).Enabled = True
      
      Call SetTxRO(Tx_DepInstantHist, False)
      Call SetTxRO(Tx_DepDecimaParteHist, False)
      Op_TipoDepHist(DEP_INSTANTANEA).Enabled = True
      Op_TipoDepHist(DEP_DECIMAPARTE).Enabled = True
   
   End If
   
   
   'Décima Parte 2  (a partir del 1/1/2020)
   If Fecha < gFechaInicioDepDecimaParte2 Then    'aquí se usa la fecha de compra (Víctor Morales 6 may 2020)
      
      Call SetTxRO(Tx_DepDecimaParte2, True)
      Op_TipoDep(DEP_DECIMAPARTE2).Enabled = False
      
      Call SetTxRO(Tx_DepDecimaparte2Hist, True)
      Op_TipoDepHist(DEP_DECIMAPARTE2).Enabled = False
      
   ElseIf Fecha >= gFechaInicioDepDecimaParte2 Then
      Call SetTxRO(Tx_DepDecimaParte2, False)
      Op_TipoDep(DEP_DECIMAPARTE2).Enabled = True
      
      Call SetTxRO(Tx_DepDecimaparte2Hist, False)
      Op_TipoDepHist(DEP_DECIMAPARTE2).Enabled = True
   
   End If
   
   If Fecha > 0 Then
      If Fecha < gFechaInicioDepLey21210 Or Fecha > gFechaTerminoDepLey21210 Then  'aquí se usa la fecha de compra (Víctor Morales 6 may 2020)
         
         Fr_AcogeLey21210.Enabled = False
         Fr_TipoDepLey21210.Enabled = False
         Op_AcogeLey21210(VAL_NO) = False
         Op_AcogeLey21210(VAL_SI) = False
         Op_TipoDepLey21210(DEP_LEY21210_INST) = False
         Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = False
         
      ElseIf Fecha >= gFechaInicioDepLey21210 Or Fecha <= gFechaTerminoDepLey21210 Then
      
         Fr_AcogeLey21210.Enabled = True
         Fr_TipoDepLey21210.Enabled = True
   
         
         If gEmpresa.Region <> 9 Then    'región de la Araucanía
            Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA).Enabled = False
         Else
            Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA).Enabled = True
            
            If lOper = O_NEW Then
               If Not MsgLey21210 Then
                  MsgLey21210 = True
                  If MsgBox1("Dado que la empresa está en la Región de la Araucanía, se puede acoger a la Ley 21210." & vbCrLf & vbCrLf & "Esto requiere que los bienes estén fisicamente y sean utilizados en la  Region de la Araucania." & vbCrLf & vbCrLf & "¿Desea acogerse a esta Ley?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                     Op_AcogeLey21210(VAL_SI) = True
                     Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = True
                     MsgBox1 "Al acogerse a la Ley 21210, Depreciación Araucanía, el 100% del valor del bien se deprecia en el período.", vbOKOnly + vbInformation
                  End If
               End If
               If Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = True Then
                  Fr_DepAnoAct.Enabled = False
                  Call ClearDep(0)
                  lTipoDep = 0
               End If
            End If
         End If
         
         If Op_AcogeLey21210(VAL_SI) Then
            Fr_TipoDepLey21210.Enabled = True
         Else
            Fr_TipoDepLey21210.Enabled = False
            Op_TipoDepLey21210(DEP_LEY21210_INST) = False
            Op_TipoDepLey21210(DEP_LEY21210_ARAUCANIA) = False
            Fr_DepAnoAct.Enabled = True
   
         End If
         
      End If
      
      If Fecha < gFechaInicioDepLey21256 Or Fecha > gFechaTerminoDepLey21256 Then  'aquí se usa la fecha de compra (Víctor Morales 6 may 2020)
         
         Fr_AcogeLey21256.Enabled = False
         Op_AcogeLey21256(VAL_NO) = False
         Op_AcogeLey21256(VAL_SI) = False
         
      ElseIf Fecha >= gFechaInicioDepLey21256 Or Fecha <= gFechaTerminoDepLey21256 Then
      
         Fr_AcogeLey21256.Enabled = True
         
      End If
      
   End If
   
   EnEnableDepEspecial = False
   
End Sub


