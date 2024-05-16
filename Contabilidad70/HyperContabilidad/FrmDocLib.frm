VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmDocLib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Documento"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12945
   Icon            =   "FrmDocLib.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_VerDTE 
      Caption         =   "Ver DTE"
      Height          =   315
      Left            =   11760
      TabIndex        =   37
      ToolTipText     =   "Ver Imagen de DTE si está disponible"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Bt_DocCuotas 
      Caption         =   "Pago a Crï¿½dito"
      Height          =   915
      Left            =   11760
      Picture         =   "FrmDocLib.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ver/agregar detalle de Activo Fijo asociado al documento seleccionado"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   9060
      Width           =   12495
   End
   Begin VB.CommandButton Bt_AddImpAdic 
      Caption         =   "Agregar Imp. Adicionales"
      Height          =   915
      Left            =   11760
      Picture         =   "FrmDocLib.frx":04E9
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Eliminar TODOS los movimientos del documento"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Bt_CleanFExport 
      Caption         =   "Volver a exportar"
      Height          =   915
      Left            =   11760
      Picture         =   "FrmDocLib.frx":0B2C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Permitir que el sistema vuelva a exportar este documento al año siguiente cuando se realice la apertura"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Bt_ClearMov 
      Caption         =   "Eliminar Movimientos"
      Height          =   915
      Left            =   11760
      Picture         =   "FrmDocLib.frx":0EDB
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Eliminar todos los movimientos del documento"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame 
      Height          =   615
      Left            =   60
      TabIndex        =   68
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   11760
         TabIndex        =   55
         Top             =   180
         Width           =   957
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   10740
         TabIndex        =   54
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   9720
         TabIndex        =   53
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton Bt_Calendar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         Picture         =   "FrmDocLib.frx":153D
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Sum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Picture         =   "FrmDocLib.frx":1966
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_ConvMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Picture         =   "FrmDocLib.frx":1A0A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Convertir moneda"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6420
         Picture         =   "FrmDocLib.frx":1DA8
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4860
         Picture         =   "FrmDocLib.frx":2109
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Plan de Cuentas"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Duplicate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Picture         =   "FrmDocLib.frx":24CA
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Duplicar movimiento seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Picture         =   "FrmDocLib.frx":28AA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Eliminar movimiento seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_MoveUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1020
         Picture         =   "FrmDocLib.frx":2CA6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Mover hacia arriba"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_MoveDown 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Picture         =   "FrmDocLib.frx":2D27
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Mover hacia abajo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Paste 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2820
         Picture         =   "FrmDocLib.frx":2DA8
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Pegar dato copiado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Copy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Picture         =   "FrmDocLib.frx":3191
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Copiar dato"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cut 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         Picture         =   "FrmDocLib.frx":3571
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Cortar dato"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cuadrar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3780
         Picture         =   "FrmDocLib.frx":392A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Cuadrar comprobante"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         Picture         =   "FrmDocLib.frx":3CAF
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_ActivoFijo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Picture         =   "FrmDocLib.frx":40F4
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Ver/agregar detalle de Activo Fijo asociado al documento seleccionado"
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.TextBox Tx_TitGrid 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   66
      Text            =   "Detalle Documento"
      Top             =   5220
      Width           =   11505
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   3015
      Left            =   60
      TabIndex        =   32
      Top             =   5640
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   5318
      Cols            =   16
      Rows            =   2
      FixedCols       =   2
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   1
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Fr_Header 
      Caption         =   "Encabezado Documento"
      ForeColor       =   &H00FF0000&
      Height          =   4395
      Left            =   60
      TabIndex        =   56
      Top             =   720
      Width           =   11535
      Begin VB.TextBox Tx_FechaExport 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8460
         MaxLength       =   15
         TabIndex        =   89
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox Ch_CompraBienRaiz 
         Caption         =   "Compra Bien Raï¿½z"
         Height          =   195
         Left            =   7140
         TabIndex        =   30
         Top             =   4080
         Width           =   1635
      End
      Begin VB.TextBox Tx_NumCuotas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1980
         Width           =   495
      End
      Begin VB.ComboBox Cb_PropIVA 
         Height          =   315
         Left            =   10200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   4020
         Width           =   1035
      End
      Begin VB.CheckBox Ch_DTEDocAsoc 
         Caption         =   "DTE"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         ToolTipText     =   "DTE Doc Acociado"
         Top             =   3300
         Width           =   675
      End
      Begin VB.ComboBox Cb_TipoDocAsoc 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox Tx_NumDocAsoc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   25
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Tx_CantBoletas 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   81
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox Tx_NumFiscImpr 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Tx_NumInformeZ 
         Height          =   315
         Left            =   4620
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Tx_VentasAcumInfZ 
         Height          =   315
         Left            =   4620
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2820
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Sucursal 
         Height          =   315
         Left            =   8460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox Ch_DelGiro 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc. del Giro"
         Height          =   195
         Left            =   4380
         TabIndex        =   4
         Top             =   840
         Width           =   1515
      End
      Begin VB.ComboBox Cb_TipoReten 
         Height          =   315
         Left            =   9780
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox Cb_Impto 
         Height          =   315
         Left            =   8460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3600
         Width           =   795
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   5940
         Picture         =   "FrmDocLib.frx":44F2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2400
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2760
         Picture         =   "FrmDocLib.frx":47FC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2400
         Width           =   230
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Top             =   1140
         Width           =   255
      End
      Begin VB.TextBox Tx_ValTotal 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   21
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox Tx_CorrInterno 
         Height          =   315
         Left            =   8460
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Ch_DTE 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc. Electrï¿½nico"
         Height          =   255
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Tx_FEmisionOri 
         Height          =   315
         Left            =   4620
         TabIndex        =   17
         ToolTipText     =   "Fecha de emisión del documento, impresa en el mismo"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Bt_NewEnt 
         Caption         =   "Nueva Entidad..."
         Height          =   315
         Left            =   9840
         Picture         =   "FrmDocLib.frx":4B06
         TabIndex        =   9
         ToolTipText     =   "Crear nueva entidad"
         Top             =   1140
         Width           =   1395
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   8460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2820
         Width           =   1575
      End
      Begin VB.TextBox Tx_Descrip 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   29
         Top             =   3660
         Width           =   5175
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   4755
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   2655
      End
      Begin VB.TextBox Tx_NumDocHasta 
         Height          =   315
         Left            =   4620
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1980
         Width           =   1335
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   8460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2775
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   8460
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "77.049.060-K"
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox Tx_FEmision 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "Fecha de recepción del docuemtno o de ingreso al libro"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Tx_FVenc 
         Height          =   315
         Left            =   8460
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Index           =   2
         Left            =   9780
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2400
         Width           =   230
      End
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Export:"
         Height          =   195
         Index           =   19
         Left            =   7140
         TabIndex        =   90
         Top             =   3300
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuotas:"
         Height          =   195
         Index           =   18
         Left            =   9900
         TabIndex        =   88
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label Lb_PropIVA 
         AutoSize        =   -1  'True
         Caption         =   "IVA Prop.:"
         Height          =   195
         Left            =   9360
         TabIndex        =   85
         Top             =   4080
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N°"
         Height          =   195
         Left            =   4380
         TabIndex        =   84
         Top             =   3300
         Width           =   180
      End
      Begin VB.Label Lb_NotaCred 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Asociado:"
         Height          =   195
         Left            =   300
         TabIndex        =   83
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Boletas:"
         Height          =   195
         Index           =   17
         Left            =   7140
         TabIndex        =   82
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Fisc. Impres:"
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   80
         Top             =   1620
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Informe Z:"
         Height          =   195
         Index           =   15
         Left            =   3480
         TabIndex        =   79
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vtas. Acum Inf. Z:"
         Height          =   195
         Index           =   13
         Left            =   3300
         TabIndex        =   78
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Index           =   2
         Left            =   7140
         TabIndex        =   77
         Top             =   780
         Width           =   660
      End
      Begin VB.Label Lb_TipoReten 
         AutoSize        =   -1  'True
         Caption         =   "Ret.:"
         Height          =   195
         Left            =   9360
         TabIndex        =   76
         Top             =   3660
         Width           =   345
      End
      Begin VB.Label Lb_Impto 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto:"
         Height          =   195
         Left            =   7140
         TabIndex        =   75
         Top             =   3660
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Electrónico"
         Height          =   195
         Index           =   0
         Left            =   4380
         TabIndex        =   74
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   7140
         TabIndex        =   73
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
         Height          =   195
         Index           =   12
         Left            =   300
         TabIndex        =   72
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Correlat. Interno:"
         Height          =   195
         Index           =   11
         Left            =   7140
         TabIndex        =   71
         Top             =   1620
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión:"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   70
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   9
         Left            =   7140
         TabIndex        =   69
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   67
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Razón Social:"
         Height          =   195
         Index           =   7
         Left            =   300
         TabIndex        =   64
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad:"
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   63
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc. Hasta:"
         Height          =   195
         Index           =   8
         Left            =   3480
         TabIndex        =   62
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento:"
         Height          =   195
         Index           =   0
         Left            =   7140
         TabIndex        =   61
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   60
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Recep.:"
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   59
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vencim.:"
         Height          =   195
         Index           =   5
         Left            =   7140
         TabIndex        =   58
         Top             =   2460
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Libro:"
         Height          =   255
         Index           =   14
         Left            =   300
         TabIndex        =   57
         Top             =   360
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   8640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   16
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmDocLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDMOV = 0
Const C_ORDEN = 1
Const C_IDCUENTA = 2
Const C_CODCUENTA = 3
Const C_CUENTA = 4
Const C_LSTCUENTA = 5
Const C_DEBE = 6
Const C_HABER = 7
Const C_TIPOVALLIB = 8
Const C_IDTIPOVALLIB = 9
Const C_TASA = 10
Const C_ESRECUPERABLE = 11
Const C_GLOSA = 12
Const C_AREANEG = 13
Const C_IDAREANEG = 14
Const C_CCOSTO = 15
Const C_IDCCOSTO = 16
Const C_ATRIBUTO = 17
Const C_CODSIIDTE = 18
Const C_UPDATE = 19

Const NCOLS = C_UPDATE

'Matrix de lcbNombre
Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

'Matrix de lcbTipoValLib
Const M_CODSIIDTE = 2
Const M_IDCUENTA = 3



Dim lRc As Integer
Dim lOper As Integer
Dim lcbNombre As ClsCombo
Dim lIdDoc As Long
Dim lMultiplesDocs As Boolean
Dim lEditEnable As Boolean

Dim lMes As Integer
Dim lOldFEmision As Long
Dim lInLoad As Boolean
Dim lMsgAfecto As Boolean

Dim lCurTipoLib As Long
Dim lCurTipoDoc As Long
Dim lIdDocAsoc As Long
Dim lOldIdDocAsoc As Long
Dim lTipoDocAsoc As Integer
Dim lOldTipoDocAsoc As Integer
Dim lNumDocAsoc As String
Dim lOldNumDocAsoc As String
Dim lDTEDocAsoc As Boolean
Dim lOldDTEDocAsoc As Boolean
Dim lEstadoDocAsocPagado As Boolean
Dim lNetoDoc As Double
Dim lcbTipoValLib As ClsCombo
Dim lMovEdited As Boolean

Dim lUrlDTE As String

Private Sub Bt_AddImpAdic_Click()
   Dim Frm As FrmConfigImpAdic
   Dim ImpAdic() As ImpAdic_t
   Dim Rc As Integer
   
   ReDim ImpAdic(0)
   
   Set Frm = New FrmConfigImpAdic
   Rc = Frm.FSelect(lCurTipoLib, lCurTipoDoc, ImpAdic())
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Call AppendImpAdic(ImpAdic, Tx_Descrip)
      MsgBox1 "ATENCIÓN:" & vbCrLf & "Se agregaron los registros con los impuestos adicionales seleccionados." & vbCrLf & vbCrLf & "Recuerde eliminar los registros que ya no son válidos y cuadrar las columnas debe y haber.", vbInformation
   End If
   
   
End Sub

Private Sub Bt_Cancel_Click()

   If lMultiplesDocs = True And lOper = O_NEW And gEmpresa.FCierre = 0 Then
     Call ClearForm
     lIdDoc = 0
   Else
   
      If Bt_Cancel.Caption = "Cerrar" Then   'se permite modificar documento adjunto aunque el docuemtno no sea editable por el estado
         If Not ValidaDocAsoc Then
            Exit Sub
         End If
         If lOldIdDocAsoc <> lIdDocAsoc Or lTipoDocAsoc <> lOldTipoDocAsoc Or lNumDocAsoc <> lOldNumDocAsoc Or lDTEDocAsoc <> lOldDTEDocAsoc Then
            If MsgBox1("¿Desea guardar la información del documento asociado? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
               Call SaveDocAsoc
            End If
         End If
      End If
      
      lRc = vbCancel
      Unload Me
   End If

End Sub

Private Sub Bt_CleanFExport_Click()
   Dim Q1 As String
   
   If MsgBox1("¿Esta seguro que desea volver a importar este documento desde el año siguiente cuando haga la apertura?" & vbCrLf & vbCrLf & "Atención: Si este documento ya existe en el año siguiente, y usted genera el comprobante de apertura, éste quedará duplicado." & vbCrLf & vbCrLf & "Si desea volverlo a importar. elimínelo en el año siguiente, antes de generar el comprobante de apertura.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Q1 = "UPDATE Documento SET FExported = 0 WHERE IdDoc = " & lIdDoc & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   MsgBox1 "El documento podrá ser importado nuevamente desde el año siguiente, al momento de generar el Comprobante de Apertura de ese año.", vbInformation
   
End Sub

Private Sub Bt_ClearMov_Click()
   Dim i As Integer
   
   If MsgBox1("¿Está seguro que desea borrar TODOS los movimientos de este documento?", vbYesNo + vbQuestion) = vbNo Then
      Exit Sub
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
   
      Call FGrModRow(Grid, i, FGR_D, C_IDMOV, C_UPDATE)
   
      Grid.Row = i
      Grid.FlxGrid.Col = C_LSTCUENTA
      Set Grid.CellPicture = LoadPicture()
      
      Grid.RowHeight(i) = 0
      Grid.rows = Grid.rows + 1
   
   Next i
   
   Call CalcTot
   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_ORDEN) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If MsgBox1("¿Está seguro que desea borrar este movimiento?", vbYesNo + vbQuestion) = vbNo Then
      Exit Sub
   End If
      
   Grid.FlxGrid.Row = Row
   Grid.FlxGrid.Col = C_LSTCUENTA
   Set Grid.FlxGrid.CellPicture = LoadPicture()
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDMOV, C_UPDATE)
   
'   Grid.RowHeight(Row) = 0
'   Grid.Rows = Grid.Rows + 1
      
   Call CalcTot
   
End Sub

Private Sub Bt_DocCuotas_Click()
   Dim Frm As FrmDocCuotas
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim Msg As String
   Dim FVenc As Long
   Dim NumCuotas As Integer
            
   If (lOper <> O_EDIT And lOper <> O_NEW) Or lEditEnable = False Then
      Set Frm = New FrmDocCuotas
      Call Frm.FView(IdDoc)
      Set Frm = Nothing
      Exit Sub
   End If
         
   Set Frm = New FrmDocCuotas
   Call Frm.FEdit(lIdDoc, FVenc, NumCuotas)
   Set Frm = Nothing
   
   Call SetTxDate(Tx_FVenc, FVenc)
   Tx_NumCuotas = IIf(NumCuotas > 0, NumCuotas, "")

   
End Sub

Private Sub Bt_Duplicate_Click()
   Dim i As Integer
   Dim Row As Integer
   
   If Grid.TextMatrix(Grid.Row, C_ORDEN) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Row = 0
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Row = i
         Exit For
      End If
   Next i

   If Row = 0 Then
      Grid.rows = Grid.rows + 1
      Row = Grid.rows - 1
   End If

   Grid.TextMatrix(Row, C_ORDEN) = Val(Grid.TextMatrix(Row - 1, C_ORDEN)) + 1
   Grid.TextMatrix(Row, C_IDCUENTA) = Grid.TextMatrix(Grid.Row, C_IDCUENTA)
   Grid.TextMatrix(Row, C_CODCUENTA) = Grid.TextMatrix(Grid.Row, C_CODCUENTA)
   Grid.TextMatrix(Row, C_CUENTA) = Grid.TextMatrix(Grid.Row, C_CUENTA)
   Grid.TextMatrix(Row, C_DEBE) = Grid.TextMatrix(Grid.Row, C_DEBE)
   Grid.TextMatrix(Row, C_HABER) = Grid.TextMatrix(Grid.Row, C_HABER)
   Grid.TextMatrix(Row, C_GLOSA) = Grid.TextMatrix(Grid.Row, C_GLOSA)
   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   Call FGrSetPicture(Grid, Row, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
   
   Call CalcTot
      
   Grid.Row = Row
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = 0
   Grid.ColSel = Grid.Cols - 1
End Sub

Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_FEmision)
   ElseIf Index = 1 Then
      Call Frm.TxSelDate(Tx_FEmisionOri)
   Else
      Call Frm.TxSelDate(Tx_FVenc)
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_Entidades_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   
   Set Frm = New FrmEntidades
   Call Frm.FEdit
   Set Frm = Nothing

End Sub

Private Sub Bt_NewEnt_Click()
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Entidad As Entidad_t
   Dim i As Integer
   Dim Rc As Integer
    
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   
   If Cb_Entidad.ListIndex >= 0 Then
      Entidad.Clasif = ItemData(Cb_Entidad)
   Else
      Entidad.Clasif = SIN_CLASLST
   End If
      
   Rc = Frm.FNew(Entidad)
   If Rc <> vbCancel Then
   
      If Cb_Entidad.ListIndex >= 0 Then
   
         If Entidad.Clasif = ItemData(Cb_Entidad) Then
            
            If Rc = vbOK Then
               Call lcbNombre.AddItem(Entidad.Nombre, Entidad.id, vFmtCID(Entidad.Rut))
               lcbNombre.ListIndex = lcbNombre.NewIndex
            Else
               lcbNombre.SelItem Entidad.id
            End If
            Tx_Rut.Text = Entidad.Rut
            
         Else
            MsgBox1 "La clasificación de la nueva entidad no coincide con la que está seleccionada. Vuelva a seleccionar el tipo de entidad para que la muestre en la lista.", vbOKOnly + vbInformation
         End If
         
      Else
         Cb_Entidad.ListIndex = FindItem(Cb_Entidad, Entidad.Clasif)
         
         Call lcbNombre.AddItem(Entidad.Nombre, Entidad.id, vFmtCID(Entidad.Rut))
         lcbNombre.ListIndex = lcbNombre.NewIndex
         Tx_Rut.Text = Entidad.Rut
      End If
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
End Sub
Private Sub Bt_NewEnt_Click_Old()
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Entidad As Entidad_t
   Dim i As Integer
 
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   Entidad.Clasif = ItemData(Cb_Entidad)
   If Frm.FNew(Entidad) = vbOK Then
   
      If Entidad.Clasif = ItemData(Cb_Entidad) Then
         
         Call lcbNombre.AddItem(Entidad.Nombre, Entidad.id, vFmtCID(Entidad.Rut))
         lcbNombre.ListIndex = lcbNombre.NewIndex
         Tx_Rut.Text = Entidad.Rut
         
      Else
         MsgBox1 "La clasificación de la nueva entidad no coincide con la que está seleccionada. Vuelva a seleccionar el tipo de entidad para que la muestre en la lista.", vbOKOnly + vbInformation
      End If
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
End Sub

Private Sub Bt_OK_Click()
   
   If Not valida() Then
      Exit Sub
   End If
      
   Call SaveAll
   If lMultiplesDocs = True And lOper = O_NEW Then
      'Limpio la grilla
      Call ClearForm
   Else
      'Para Doc Tipo sirve así
      lRc = vbOK
      Unload Me
   End If
   
End Sub

Public Function FNew(IdDoc As Long, Optional ByVal MultiplesDocs As Boolean = True, Optional ByVal Mes As Integer = 0) As Integer
   
   lOper = O_NEW
   lIdDoc = 0
   
   lMultiplesDocs = MultiplesDocs
   lMes = Mes
   
   Me.Show vbModal
   
   IdDoc = lIdDoc
   FNew = lRc
   
End Function
Public Function FEdit(ByVal IdDoc As Long) As Integer
   lOper = O_EDIT
   lIdDoc = IdDoc
   
   Me.Show vbModal
   
   FEdit = lRc
   
End Function

Public Sub FView(ByVal IdDoc As Long)
   
   lOper = O_VIEW
   lIdDoc = IdDoc
   
   Me.Show vbModal
      
End Sub

Private Sub Bt_Close_Click()

   If (Tx_Descrip <> "" Or Grid.TextMatrix(Grid.FixedRows, C_ORDEN) <> "") And lOper = O_NEW Then
   'If lIdDoc <> 0 And lOper = O_NEW Then
      If MsgBox1("¿Desea guardar el documento actual?", vbYesNo + vbQuestion) = vbYes Then
         If valida() Then
            Call SaveAll
         End If
      Else
         lRc = vbCancel
         Unload Me
      End If
   Else
      lRc = vbCancel
      Unload Me
   End If
   
End Sub

Private Sub Bt_VerDTE_Click()
      
      If lUrlDTE = "" Then
         MsgBox1 "No es posible ver el PDF de este documento.", vbExclamation
         Exit Sub
      End If
      
'      Me.MousePointer = vbHourglass
'      If gConectData.Proveedor = PROV_LP Then
'         Call GetPdfDTE(Val(Grid.TextMatrix(Grid.Row, C_FOLIO)), Val(Grid.TextMatrix(Grid.Row, C_CODDOCSII)), Val(Grid.TextMatrix(Grid.Row, C_LNGFECHA)))
'      ElseIf Grid.TextMatrix(Grid.Row, C_URLDTE) = "" Then
'         MsgBox1 "No se encuentra disponible el DTE para ser impreso." & vbCrLf & vbCrLf & "Verifique el estado del DTE.", vbExclamation
'      Else
'         Call AcpShowDTE(Me, lUrlDTE)
'      End If
      
      Call ShellExecute(Me.hWnd, "open", lUrlDTE, "", "", 1)
      
      Me.MousePointer = vbDefault

End Sub

Private Sub Cb_Entidad_Click()
   
   If Cb_Entidad.ListIndex >= 0 Then
      Call SelCbEntidad(ItemData(Cb_Entidad))
   Else
      lcbNombre.Clear
   End If

End Sub


Private Sub cb_Nombre_Click()
   
   Tx_Rut = ""
   
   If lcbNombre.ListIndex >= 0 Then
      If lcbNombre.Matrix(M_RUT) <> "" Then
         Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
         Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
      End If
   End If
   
End Sub



Private Sub Cb_TipoLib_Click()
   Dim Q1 As String
   Dim i As Integer
   Dim TipoLib As Long
   
   Cb_TipoDoc.Clear
   
   TipoLib = CbItemData(Cb_TipoLib)
   
   If TipoLib <= 0 Then
      Exit Sub
   End If
      
   If lCurTipoLib <> TipoLib Then
      Call Cb_TipoDocAsoc.Clear
      Tx_NumDocAsoc = ""
      Ch_DTEDocAsoc = 0
   
      lCurTipoLib = TipoLib
   End If
      
   Call FillTipoDoc(Cb_TipoDoc, TipoLib, True, True)
   
   'Call FillTipoValLib(Grid.CbList(C_TIPOVALLIB), tipolib, True, True, "", ItemData(Cb_TipoDoc))
   
   If TipoLib <> LIB_VENTAS Then
      Ch_DelGiro.visible = False
   Else
      Ch_DelGiro.visible = True
   End If
   
   If TipoLib <> LIB_COMPRAS Then
      Ch_CompraBienRaiz.visible = False
   Else
      Ch_CompraBienRaiz.visible = True
   End If
   
   
   If TipoLib = LIB_RETEN Then
      Tx_Descrip.Width = Cb_Nombre.Width
      Lb_Impto.visible = True
      Cb_Impto.visible = True
      Lb_TipoReten.visible = True
      Cb_TipoReten.visible = True
   Else
      'Tx_Descrip.Width = Cb_TipoReten.Left + Cb_TipoReten.Width - Tx_Descrip.Left
      Lb_Impto.visible = False
      Cb_Impto.visible = False
      Lb_TipoReten.visible = False
      Cb_TipoReten.visible = False
   End If
   
   If TipoLib <> LIB_COMPRAS Then
      Cb_PropIVA.visible = False
      Lb_PropIVA.visible = False
   Else
      Cb_PropIVA.visible = gFunciones.ProporcionalidadIVA
      Lb_PropIVA.visible = gFunciones.ProporcionalidadIVA
   End If
   
   
   Call FrmEnable
   

   
End Sub

Private Sub Cb_TipoDoc_Click()
   Dim DimDoc As String
   Dim TipoLib As Long
   Dim TipoDoc As Integer
   Dim TipoDocAsoc As Integer
   Dim Idx As Long
   

   If CbItemData(Cb_TipoLib) > 0 Then
      TipoLib = CbItemData(Cb_TipoLib)
   Else
      Exit Sub
   End If
   
   If CbItemData(Cb_TipoDoc) > 0 Then
      TipoDoc = CbItemData(Cb_TipoDoc)
   Else
      Exit Sub
   End If

   'Call FillTipoValLib(Grid.CbList(C_TIPOVALLIB), TipoLib, True, True, "", TipoDoc)
   Call FillClsTipoValLib(lcbTipoValLib, TipoLib, True, True, "", TipoDoc, gOcultarImpAdicDescont)
   
   DimDoc = GetDiminutivoDoc(TipoLib, TipoDoc)
   
   Idx = GetTipoDoc(TipoLib, TipoDoc)
   
   If DimDoc = "FAV" Or DimDoc = "FVE" Or DimDoc = "NCV" Or DimDoc = "NDV" Then
      If Not lInLoad And lOper = O_EDIT Then
         Ch_DelGiro = 1
      End If
      Ch_DelGiro.Enabled = True
   Else
      Ch_DelGiro = 0
      Ch_DelGiro.Enabled = False
   End If
   
   If DimDoc = "FAC" Or DimDoc = "FCE" Or DimDoc = "NCC" Or DimDoc = "NDC" Then
      Ch_CompraBienRaiz.Enabled = True
   Else
      Ch_CompraBienRaiz = 0
      Ch_CompraBienRaiz.Enabled = False
   End If
   
   
   If DimDoc = TDOC_MAQREGISTRADORA Then
      Call SetTxRO(Tx_NumFiscImpr, False)
      Call SetTxRO(Tx_NumInformeZ, False)
      Call SetTxRO(Tx_VentasAcumInfZ, False)
   Else
      Call SetTxRO(Tx_NumFiscImpr, True)
      Call SetTxRO(Tx_NumInformeZ, True)
      Call SetTxRO(Tx_VentasAcumInfZ, True)
   End If
   
   '2814014 pipe
   If DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLVENTAEX Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_MAQREGISTRADORA Or DimDoc = TDOC_VALVENTAEX Then 'VPEE
   'If DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLVENTAEX Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_MAQREGISTRADORA Then
   'fin 2814014
   Call SetTxRO(Tx_NumDocHasta, False)
   Else
      Call SetTxRO(Tx_NumDocHasta, True)
      Tx_NumDocHasta = ""
   End If
   
   If lCurTipoDoc <> TipoDoc Then
      Call Cb_TipoDocAsoc.Clear
      Tx_NumDocAsoc = ""
      Ch_DTEDocAsoc = 0
   
      lCurTipoDoc = TipoDoc
   End If

   
   If (DimDoc = "NCC" Or DimDoc = "NDC" Or DimDoc = "NCV" Or DimDoc = "NDV" Or DimDoc = "NCE" Or DimDoc = "NDE") And Not lEstadoDocAsocPagado Then
      Cb_TipoDocAsoc.Locked = False
      Call SetTxRO(Tx_NumDocAsoc, False)
      Ch_DTEDocAsoc.Enabled = True
   Else
      Cb_TipoDocAsoc.Locked = True
      Call SetTxRO(Tx_NumDocAsoc, True)
      Ch_DTEDocAsoc.Enabled = False
   End If
   
   If Cb_TipoDocAsoc.ListCount <= 0 Then
      If DimDoc = "NCC" Or DimDoc = "NDC" Then
         Call CbAddItem(Cb_TipoDocAsoc, " ", 0)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FAC")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FCE")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FCC")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "LFC")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
            
      ElseIf DimDoc = "NCV" Or DimDoc = "NDV" Then
         Call CbAddItem(Cb_TipoDocAsoc, " ", 0)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FAV")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FVE")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "FCV")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         TipoDocAsoc = FindTipoDoc(TipoLib, "LFV")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
                 
         If DimDoc = "NCV" Then
            TipoDocAsoc = FindTipoDoc(TipoLib, TDOC_BOLVENTA)
            Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
            TipoDocAsoc = FindTipoDoc(TipoLib, TDOC_BOLEXENTA)
            Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
            TipoDocAsoc = FindTipoDoc(TipoLib, TDOC_BOLVENTAEX)
            Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
            '2814014 pipe
             TipoDocAsoc = FindTipoDoc(TipoLib, TDOC_VALVENTAEX)
            Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
            'fin 2814014
            
            TipoDocAsoc = FindTipoDoc(TipoLib, TDOC_VALEPAGOELECTR)
            Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
         End If
         
      ElseIf DimDoc = "NCE" Or DimDoc = "NDE" Then
         Call CbAddItem(Cb_TipoDocAsoc, " ", 0)
         TipoDocAsoc = FindTipoDoc(TipoLib, "EXP")
         Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
      
      End If
      
      
   End If
   
   If TipoLib <> LIB_COMPRAS Or Not gTipoDoc(Idx).AceptaPropIVA Then
      Cb_PropIVA.ListIndex = 0
      Cb_PropIVA.Enabled = False
   End If
      
   If Not lInLoad Then
      If lOper = O_EDIT And (ItemData(Cb_TipoLib) = LIB_COMPRAS Or ItemData(Cb_TipoLib) = LIB_VENTAS Or ItemData(Cb_TipoLib) = LIB_RETEN) Then
         MsgBox1 "¡¡¡Atención!!!" & vbNewLine & vbNewLine & "Recuerde que si cambia el tipo de documento, es posible que deba modificar los valores de las columnas Debe y Haber en el detalle del documento.", vbInformation + vbOKOnly
      End If
   End If
   
   Call FrmEnable
End Sub



Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim Idx As Long
   
   lInLoad = True
   lMsgAfecto = False
      
   For i = 0 To 2
      Call BtFechaImg(Bt_Fecha(i))
   Next i
   
   'SE LLENA COMBOS
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Call FillCb
   
   'Creamos combo para columna TipoValLib
   Set lcbTipoValLib = New ClsCombo
   Call lcbTipoValLib.SetControl(Grid.CbList(C_TIPOVALLIB))

         
   Bt_Close.visible = (lMultiplesDocs = True And lOper = O_NEW)
   
   If Not lMultiplesDocs Or lOper <> O_NEW Then
      Bt_OK.Left = Bt_Cancel.Left
      Bt_Cancel.Left = Bt_Close.Left
   End If
   
   Bt_DocCuotas.visible = gFunciones.DocCuotas
   Bt_DocCuotas.Enabled = gFunciones.DocCuotas

   Call SetUpGrid
   Call ClearForm
   Call LoadAll
   Call FrmEnable
   
   'volvemos a llamar a  TipoDoc_Click para que haga las habilitaciones y deshabilitaciones correspondientes al tipo doc que se anulan con FrmEnable
   Call Cb_TipoDoc_Click
   
   Bt_AddImpAdic.visible = False
   
   Select Case lOper
      Case O_NEW
         Caption = "Nuevo Documento"
'         If CbItemData(Cb_TipoLib) = LIB_COMPRAS Then
            Bt_AddImpAdic.visible = True
'         End If
         
      Case O_EDIT
         Caption = "Editar Documento"
'         If CbItemData(Cb_TipoLib) = LIB_COMPRAS Then
            Bt_AddImpAdic.visible = True
'         End If

         '3284709
        If ItemData(Cb_Estado) = ED_PENDIENTE And CbItemData(Cb_TipoLib) = LIB_VENTAS And GetDiminutivoDoc(CbItemData(Cb_TipoLib), CbItemData(Cb_TipoDoc)) = TDOC_FAVEXENTA And Year(GetTxDate(Tx_FEmision)) < gEmpresa.Ano Then
         Cb_Estado.Enabled = True
         MsgBox1 "Documento proviene del año anterior, favor de dejar estado del documento en centralizado ya que se encuentra en estado pendiente. Favor de ejecutar Recalcular Saldo.", vbInformation
            Call SelItem(Cb_Estado, ED_CENTRALIZADO)
        Else
         'Cb_Estado.Enabled = False
        
        End If
        '3284709
      
      Case Else
         Caption = "Ver Documento"
   End Select
     
 
   Bt_ActivoFijo.visible = gFunciones.ActivoFijo
   Bt_ActivoFijo.Enabled = gFunciones.ActivoFijo
   
   
   lInLoad = False
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
      
   Grid.ColWidth(C_IDMOV) = 0
   Grid.ColWidth(C_ORDEN) = 300
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CODCUENTA) = 1650
   Grid.ColWidth(C_CUENTA) = 2300
   Grid.ColWidth(C_LSTCUENTA) = 250
   Grid.ColWidth(C_DEBE) = 1200
   Grid.ColWidth(C_HABER) = 1200
   Grid.ColWidth(C_GLOSA) = 2900
   Grid.ColWidth(C_TIPOVALLIB) = 3000
   Grid.ColWidth(C_IDTIPOVALLIB) = 0
   Grid.ColWidth(C_TASA) = 600
   Grid.ColWidth(C_ESRECUPERABLE) = 600
   Grid.ColWidth(C_CCOSTO) = 2000
   Grid.ColWidth(C_IDCCOSTO) = 0
   Grid.ColWidth(C_AREANEG) = 2000
   Grid.ColWidth(C_IDAREANEG) = 0
   Grid.ColWidth(C_CODSIIDTE) = 0
   Grid.ColWidth(C_ATRIBUTO) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_IDMOV) = flexAlignRightCenter
   Grid.ColAlignment(C_ORDEN) = flexAlignRightCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_LSTCUENTA) = flexAlignCenterCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPOVALLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_TASA) = flexAlignRightCenter
   Grid.ColAlignment(C_ESRECUPERABLE) = flexAlignCenterCenter
   Grid.ColAlignment(C_CCOSTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_AREANEG) = flexAlignLeftCenter
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   Grid.TextMatrix(0, C_IDMOV) = "."
   Grid.TextMatrix(0, C_ORDEN) = "N°"
   Grid.TextMatrix(0, C_IDCUENTA) = ""
   Grid.TextMatrix(0, C_CODCUENTA) = "Código Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.FlxGrid.Row = 0
   Grid.FlxGrid.Col = C_LSTCUENTA
   Set Grid.CellPicture = Bt_Cuentas.Picture
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   Grid.TextMatrix(0, C_GLOSA) = "Descripción"
   Grid.TextMatrix(0, C_TIPOVALLIB) = "Clasificación"
   Grid.TextMatrix(0, C_TASA) = "Tasa"
   Grid.TextMatrix(0, C_ESRECUPERABLE) = "Recup."
   Grid.TextMatrix(0, C_CCOSTO) = "Centro de Gestión"
   Grid.TextMatrix(0, C_AREANEG) = "Área de Negocio"
   
   Call FGrVRows(Grid)

End Sub
Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, abs(NotValidRut) FROM Entidades"
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY Nombre "
      Call lcbNombre.FillCombo(DbMain, Q1, -1)
   Else
      Call lcbNombre.AddItem(" ", 0)
      Call lcbNombre.AddItem("DIN", -2)
   End If
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim k As Integer
   Dim DimDoc As String, DimDocAsoc As String
   
   If lOper = O_NEW Or lIdDoc <= 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT TipoLib, TipoDoc, NumDoc, NumDocHasta, Documento.IdEntidad, RutEntidad, NombreEntidad, Documento.TipoEntidad, "
   Q1 = Q1 & " FEmision, FVenc, FEmisionOri, Descrip, Afecto, Total, Documento.Estado, DTE, CorrInterno, "
   Q1 = Q1 & " Entidades.Nombre, Entidades.Rut, Entidades.NotValidRut, UrlDTE, "
   Q1 = Q1 & " PorcentRetencion, TipoRetencion, Documento.Giro, IdSucursal, MovEdited, NumCuotas, "
   Q1 = Q1 & " NumFiscImpr, NumInformeZ, CantBoletas, VentasAcumInfZ, IdDocAsoc, Documento.PropIVA, TipoDocAsoc, NumDocAsoc, DTEDocAsoc, CompraBienRaiz "
   
   For k = 0 To MAX_ENTCLASIF
      Q1 = Q1 & ", Clasif" & k
   Next k
   
   '3125512
   Q1 = Q1 & " ,FExported "
   '3125512
   
   
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad"
   Q1 = Q1 & " AND Entidades.IdEmpresa = Documento.IdEmpresa "
   Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Call SelItem(Cb_TipoLib, vFld(Rs("TipoLib")))
      Call SelItem(Cb_TipoDoc, vFld(Rs("TipoDoc")))
      
      lCurTipoLib = vFld(Rs("TipoLib"))
      lCurTipoDoc = vFld(Rs("TipoDoc"))
      
      If vFld(Rs("idEntidad")) > 0 Then
         Call SelItem(Cb_Entidad, vFld(Rs("TipoEntidad")))
      End If
      
      If vFld(Rs("Rut")) <> "" Then
         Tx_Rut = FmtCID(vFld(Rs("Rut")), Not vFld(Rs("NotValidRut")))
      End If
           
      If vFld(Rs("idEntidad")) > 0 Then
         
         If lcbNombre.SelItem(vFld(Rs("idEntidad"))) < 0 Then 'no lo encontró, puede estar en otra clasificación
            
            For k = 0 To MAX_ENTCLASIF
               If vFld(Rs("Clasif" & k)) <> 0 Then
                  Call SelItem(Cb_Entidad, k)
                  lcbNombre.SelItem (vFld(Rs("idEntidad")))
                  Exit For
               End If
            Next k
            
         End If
         
      End If
      
      DimDoc = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
         
      If DimDoc = TDOC_MAQREGISTRADORA Then
         Tx_NumFiscImpr = vFld(Rs("NumFiscImpr"))
         Tx_NumInformeZ = vFld(Rs("NumInformeZ"))
      End If
      
      If vFld(Rs("CantBoletas")) > 0 Then
         Tx_CantBoletas = Format(vFld(Rs("CantBoletas")), NUMFMT)
      End If
      
      If DimDoc <> "VSD" Then      'venta sin documento
         Tx_NumDoc = vFld(Rs("NumDoc"))
         Tx_NumDocHasta = vFld(Rs("NumDocHasta"))
      End If
      
      Tx_CorrInterno = vFld(Rs("CorrInterno"), True)
      Ch_DTE = IIf(vFld(Rs("DTE")) <> 0, 1, 0)
      Ch_DelGiro = IIf(vFld(Rs("Giro")) <> 0, 1, 0)
      Ch_CompraBienRaiz = IIf(vFld(Rs("CompraBienRaiz")) <> 0, 1, 0)
      
      Call SetTxDate(Tx_FEmision, vFld(Rs("FEmision")))
      lOldFEmision = vFld(Rs("FEmision"))
      Call SetTxDate(Tx_FEmisionOri, vFld(Rs("FEmisionOri")))
      Call SetTxDate(Tx_FVenc, vFld(Rs("FVenc")))
      
      '3125512
      Call SetTxDate(Tx_FechaExport, vFld(Rs("FExported")))
      '3125512
      
      Tx_NumCuotas = IIf(vFld(Rs("NumCuotas")) > 0, vFld(Rs("NumCuotas")), "")
      
      Tx_ValTotal = Format(vFld(Rs("Total")), NUMFMT)
      lNetoDoc = vFld(Rs("Afecto"))
      
      If DimDoc = TDOC_MAQREGISTRADORA Then
         Tx_VentasAcumInfZ = Format(vFld(Rs("VentasAcumInfZ")), NUMFMT)
      End If
      
      Call SelItem(Cb_Estado, vFld(Rs("Estado")))
      Call SelItem(Cb_Sucursal, vFld(Rs("IdSucursal")))
      
      lMovEdited = vFld(Rs("MovEdited"))
      lUrlDTE = vFld(Rs("UrlDTE"))
      
      Call SelItem(Cb_PropIVA, vFld(Rs("PropIVA")))

      Tx_Descrip = vFld(Rs("Descrip"), True)
      
      If vFld(Rs("TipoLib")) = LIB_RETEN Then
         Call SelItem(Cb_Impto, vFld(Rs("PorcentRetencion")))
         Call SelItem(Cb_TipoReten, vFld(Rs("TipoRetencion")))
      End If
      
      lIdDocAsoc = vFld(Rs("IdDocAsoc"))
      lTipoDocAsoc = vFld(Rs("TipoDocAsoc"))
      lNumDocAsoc = vFld(Rs("NumDocAsoc"))
      lDTEDocAsoc = vFld(Rs("DTEDocAsoc"))
      lOldIdDocAsoc = lIdDocAsoc
      lOldTipoDocAsoc = lTipoDocAsoc
      lOldNumDocAsoc = lNumDocAsoc
      lOldDTEDocAsoc = lDTEDocAsoc
      
      DimDocAsoc = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDocAsoc")))

      
'      If lCurTipoLib = LIB_VENTAS And DimDoc = TDOC_VALEPAGOELECTR Then
      If lIdDocAsoc <> 0 Then
         Call GetDocAsoc(lIdDocAsoc)
         
      'ElseIf (DimDocAsoc = TDOC_VALEPAGOELECTR Or DimDocAsoc = TDOC_BOLVENTA Or DimDocAsoc = TDOC_BOLEXENTA Or DimDocAsoc = TDOC_BOLVENTAEX) Then   'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX  (FCA 30 oct 2017)
      
      '2814014
      'ElseIf (DimDocAsoc = TDOC_VALEPAGOELECTR Or DimDocAsoc = TDOC_BOLVENTA Or DimDocAsoc = TDOC_BOLEXENTA Or DimDocAsoc = TDOC_BOLVENTAEX) Then   'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX  (FCA 30 oct 2017)
      ElseIf (DimDocAsoc = TDOC_VALEPAGOELECTR Or DimDocAsoc = TDOC_BOLVENTA Or DimDocAsoc = TDOC_BOLEXENTA Or DimDocAsoc = TDOC_BOLVENTAEX Or DimDocAsoc = TDOC_VALVENTAEX) Then   'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX  (FCA 30 oct 2017)
         
         Call CbSelItem(Cb_TipoDocAsoc, lTipoDocAsoc)
         Tx_NumDocAsoc = lNumDocAsoc
         Ch_DTEDocAsoc = IIf(lDTEDocAsoc <> 0, 1, 0)
         
         lEstadoDocAsocPagado = False
      End If
   
   Else ' *** PAM - 5 MAY 2005
      MsgBox1 "No se encontró el documento solicitado. Debe haber sido eliminado por otro usuario.", vbExclamation
      lIdDoc = 0
   End If
   
   
      
   Call CloseRs(Rs)
   
   If lIdDoc <= 0 Then  ' *** PAM - 5 MAY 2005
      Exit Sub
   End If
         
   Call LoadGrMov
   
  
   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - Tx_CurrCell.Height - 550
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 230
   GridTot.Width = Grid.Width - 230
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call bt_Copy_Click
   ElseIf KeyPaste(KeyCode, Shift) Then
      Call Bt_Paste_Click
   End If

End Sub

Private Sub Grid_SelChange()
   Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)
End Sub

Private Sub Tx_FEmision_GotFocus()
   Call DtGotFocus(Tx_FEmision)
End Sub

Private Sub Tx_FEmision_LostFocus()
   
   If Trim$(Tx_FEmision) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FEmision)
   
End Sub

Private Sub Tx_FEmision_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FEmisionOri_GotFocus()
   Call DtGotFocus(Tx_FEmisionOri)
End Sub

Private Sub Tx_FEmisionOri_LostFocus()
   
   If Trim$(Tx_FEmisionOri) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FEmisionOri)
   
End Sub

Private Sub Tx_FEmisionOri_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub


Private Sub Tx_FVenc_GotFocus()
   
   If Tx_FVenc = "" And Tx_FEmisionOri <> "" Then
      Call SetTxDate(Tx_FVenc, DateAdd("d", 30, GetTxDate(Tx_FEmisionOri)))
   End If
   
   Call DtGotFocus(Tx_FVenc)
End Sub

Private Sub Tx_FVenc_LostFocus()
   
   If Trim$(Tx_FVenc) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FVenc)
   
End Sub

Private Sub Tx_FVenc_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim j As Integer
   Dim NumDocVSD As Long
   Dim DimDoc As String
   Dim FldArray(3) As AdvTbAddNew_t
   
   If lOper = O_NEW Then
      
'      Set Rs = DbMain.OpenRecordset("Documento", dbOpenTable)
'      Rs.AddNew
'
'      lIdDoc = vFld(Rs("IdDoc"))
'
'      Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'      Rs.Fields("FechaCreacion") = CLng(Int(Now))
'
'      On Error Resume Next
'
'      For j = 1 To 10
'         Rs.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'         Rs.Update
'
'         If Err = 0 Then
'            Exit For
'         End If
'      Next j
'
'      Rs.Close
'
'      Set Rs = Nothing


      FldArray(0).FldName = "IdUsuario"
      FldArray(0).FldValue = gUsuario.IdUsuario
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "FechaCreacion"
      FldArray(1).FldValue = CLng(Int(Now))
      FldArray(1).FldIsNum = True
            
      FldArray(2).FldName = "IdEmpresa"
      FldArray(2).FldValue = gEmpresa.id
      FldArray(2).FldIsNum = True
                  
      FldArray(3).FldName = "Ano"
      FldArray(3).FldValue = gEmpresa.Ano
      FldArray(3).FldIsNum = True
      
      lIdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
            
   End If

   Q1 = "UPDATE Documento SET "
   Q1 = Q1 & "  TipoLib =" & lCurTipoLib
   Q1 = Q1 & ", TipoDoc =" & lCurTipoDoc
   Q1 = Q1 & ", CorrInterno = " & vFmt(Tx_CorrInterno)
   Q1 = Q1 & ", DTE = " & IIf(Ch_DTE <> 0, -1, 0)
   Q1 = Q1 & ", Giro = " & IIf(Ch_DelGiro <> 0, -1, 0)
   Q1 = Q1 & ", CompraBienRaiz = " & IIf(Ch_CompraBienRaiz <> 0, -1, 0)
   Q1 = Q1 & ", NumFiscImpr ='" & ParaSQL(Tx_NumFiscImpr) & "'"
   Q1 = Q1 & ", NumInformeZ ='" & ParaSQL(Tx_NumInformeZ) & "'"
   
   If DimDoc <> "VSD" Then     'venta sin documento
      Q1 = Q1 & ", NumDoc ='" & ParaSQL(Tx_NumDoc) & "'"
      Q1 = Q1 & ", NumDocHasta ='" & ParaSQL(Tx_NumDocHasta) & "'"
   Else
      NumDocVSD = GetNumDocVSD(lCurTipoLib, lCurTipoDoc)
      Q1 = Q1 & ", NumDoc = '" & NumDocVSD & "'"
      Q1 = Q1 & ", NumDocHasta = '0'"
   End If
   
   Q1 = Q1 & ", CantBoletas =" & vFmt(Tx_CantBoletas)
   
   If ItemData(Cb_Entidad) >= 0 Then
      Q1 = Q1 & ", TipoEntidad =" & ItemData(Cb_Entidad)
   Else
      Q1 = Q1 & ", TipoEntidad =0"
   End If
   If lcbNombre.ItemData >= 0 Then
      Q1 = Q1 & ", idEntidad =" & lcbNombre.ItemData
   Else
      Q1 = Q1 & ", idEntidad =0"
   End If
      
   Q1 = Q1 & ", FEmision =" & GetTxDate(Tx_FEmision)
   Q1 = Q1 & ", PropIVA =" & CbItemData(Cb_PropIVA)
   Q1 = Q1 & ", FEmisionOri =" & GetTxDate(Tx_FEmisionOri)
   Q1 = Q1 & ", FVenc =" & GetTxDate(Tx_FVenc)
   Q1 = Q1 & ", Estado =" & CbItemData(Cb_Estado)
   Q1 = Q1 & ", IdSucursal =" & ItemData(Cb_Sucursal)
   Q1 = Q1 & ", Descrip ='" & Left(ParaSQL(RemoveNoPrtChars(Tx_Descrip, True)), 100) & "'"
   If lCurTipoLib = LIB_RETEN Then
      Q1 = Q1 & ", PorcentRetencion = " & ItemData(Cb_Impto)
      Q1 = Q1 & ", TipoRetencion = " & ItemData(Cb_TipoReten)
   End If
   Q1 = Q1 & ", Total =" & vFmt(Tx_ValTotal)
   Q1 = Q1 & ", VentasAcumInfZ =" & vFmt(Tx_VentasAcumInfZ)
   Q1 = Q1 & ", SaldoDoc = NULL"
   Q1 = Q1 & "  WHERE IdDoc =" & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Tracking 3227543
    If lOper = O_NEW Then
        Call SeguimientoDocumento(lIdDoc, gEmpresa.id, gEmpresa.Ano, "FrmDocLib.SaveAll", Q1, 1, "", gUsuario.IdUsuario, 1, 1)
    Else
        Call SeguimientoDocumento(lIdDoc, gEmpresa.id, gEmpresa.Ano, "FrmDocLib.SaveAll", Q1, 1, "", gUsuario.IdUsuario, 1, 2)
    End If
    ' fin 3227543
   
   Call SaveDocAsoc
      
   Call SaveGrMov
      
End Sub
Private Sub SaveDocAsoc()
   Dim Q1 As String
   
   Q1 = "UPDATE Documento SET IdDocAsoc = " & lIdDocAsoc & ", TipoDocAsoc = " & lTipoDocAsoc & ", NumDocAsoc = '" & lNumDocAsoc & "', DTEDocAsoc = " & IIf(Ch_DTEDocAsoc <> 0, 1, 0)
   Q1 = Q1 & ", SaldoDoc = NULL WHERE IdDoc =" & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Tracking 3227543
    Call SeguimientoDocumento(lIdDoc, gEmpresa.id, gEmpresa.Ano, "FrmDocLib.SaveDocAsoc", "", 1, "", gUsuario.IdUsuario, 1, 2)
    ' fin 3227543
      
   If lIdDocAsoc <> 0 Then          'para que se recalcule el saldo en el documento asociado (se deja con SaldoDoc=0 en RecalcSaldos)
      Q1 = "UPDATE Documento SET SaldoDoc = NULL"
      Q1 = Q1 & "  WHERE IdDoc =" & lIdDocAsoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id    'no le ponemos el año porque podría ser del año anterior
      Call ExecSQL(DbMain, Q1)
      
      'Tracking 3227543
      Call SeguimientoDocumento(lIdDocAsoc, gEmpresa.id, gEmpresa.Ano, "FrmDocLib.SaveDocAsoc", "", 1, "", gUsuario.IdUsuario, 1, 2)
      ' fin 3227543
   End If
   
End Sub
Private Sub LoadGrMov()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim id As Long
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Q1 = "SELECT IdMovDoc, Orden, MovDocumento.IdCuenta, Cuentas.Codigo as CodCta, Nombre, Cuentas.Descripcion, "
   Q1 = Q1 & " MovDocumento.Debe, MovDocumento.Haber, Glosa, MovDocumento.IdTipoValLib, "
   Q1 = Q1 & " MovDocumento.IdCCosto, CentroCosto.Descripcion as DescCCosto, "
   Q1 = Q1 & " MovDocumento.IdAreaNeg, AreaNegocio.Descripcion as DescAreaNeg, Tasa, EsRecuperable "
   Q1 = Q1 & " FROM ((MovDocumento INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "Cuentas") & " )"
   Q1 = Q1 & " LEFT JOIN CentroCosto ON MovDocumento.IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "CentroCosto", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN AreaNegocio ON MovDocumento.IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "AreaNegocio", True, True)
   Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Orden"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDMOV) = vFld(Rs("IdMovDoc"))
      Grid.TextMatrix(i, C_ORDEN) = vFld(Rs("Orden"))
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(vFld(Rs("CodCta")))
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Descripcion"))
      Call FGrSetPicture(Grid, i, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
      Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"), True)
      Grid.TextMatrix(i, C_IDTIPOVALLIB) = vFld(Rs("IdTipoValLib"), True)
      Grid.TextMatrix(i, C_TIPOVALLIB) = GetNombreTipoValLib(CbItemData(Cb_TipoLib), vFld(Rs("IdTipoValLib")))
      Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("DescCCosto"), True)
      Grid.TextMatrix(i, C_IDCCOSTO) = vFld(Rs("IdCCosto"), True)
      Grid.TextMatrix(i, C_AREANEG) = vFld(Rs("DescAreaNeg"), True)
      Grid.TextMatrix(i, C_IDAREANEG) = vFld(Rs("IdAreaNeg"), True)
      id = GetTipoValLib(CbItemData(Cb_TipoLib), vFld(Rs("IdTipoValLib")))
      If id >= 0 Then
         Grid.TextMatrix(i, C_ATRIBUTO) = gTipoValLib(id).Atributo
         Grid.TextMatrix(i, C_CODSIIDTE) = gTipoValLib(id).CodSIIDTE
      End If
      If Val(Grid.TextMatrix(i, C_CODSIIDTE)) > 0 And Grid.TextMatrix(i, C_ATRIBUTO) <> "IVAIRREC" Then
         Grid.TextMatrix(i, C_TASA) = IIf(vFld(Rs("Tasa")) = 0, "", Format(vFld(Rs("Tasa")), DBLFMT2))
         Grid.TextMatrix(i, C_ESRECUPERABLE) = FmtSiNo(Abs(vFld(Rs("EsRecuperable"))))
      End If
      
      
      '2861733 tema 2

      Dim vIdArea As Long
      Dim vIdCentro As Long
      Dim vDescArea As String
      Dim vDescCentro As String

      Dim vTieneArea As Boolean
      Dim vTieneCentro As Boolean


      If ValidarCuentaAFijo(vFld(Rs("IdCuenta")), vTieneArea, vTieneCentro) = True Then

'       '2861733 tema 2
'         If ValidarCuentaAFijo(Val(Grid.TextMatrix(i, C_IDCUENTA)), False, False) Then
'           Call SaveAreaCentroActFijo(Val(Grid.TextMatrix(i, C_IDCUENTA)), vFmt(Grid.TextMatrix(i, C_IDAREANEG)), vFmt(Grid.TextMatrix(i, C_IDCCOSTO)))
'         End If
'         '2861733 tema 2

      If ObtenerAreaCentro(lIdDoc, vFld(Rs("IdCuenta")), vIdArea, vIdCentro, vDescArea, vDescCentro) Then

       If Grid.TextMatrix(i, C_IDCCOSTO) <> vIdCentro Then
         If vTieneCentro Then
            Grid.TextMatrix(i, C_IDCCOSTO) = vIdCentro
            Grid.TextMatrix(i, C_CCOSTO) = vDescCentro
         End If
       End If
       If Grid.TextMatrix(i, C_IDAREANEG) <> vIdArea Then
        If vTieneArea Then
        Grid.TextMatrix(i, C_IDAREANEG) = vIdArea
        Grid.TextMatrix(i, C_AREANEG) = vDescArea
        End If
      End If

       If vTieneCentro Then
         Grid.TextMatrix(i, C_IDCCOSTO) = vIdCentro
         Grid.TextMatrix(i, C_CCOSTO) = vDescCentro
       End If

        If vTieneArea Then
         Grid.TextMatrix(i, C_IDAREANEG) = vIdArea
         Grid.TextMatrix(i, C_AREANEG) = vDescArea
         End If

      End If
      Grid.TextMatrix(i, C_UPDATE) = FGR_U

      End If
      
'2861733       tema 2
      
      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 2)
   
   Grid.Row = Grid.FixedRows
   Grid.Col = C_CODCUENTA
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col
   
   Call CalcTot
End Sub
Private Sub SaveGrMov()
   Dim i As Integer, j As Integer
   Dim Lin As Integer
   Dim Rs As Recordset
   Dim Q1 As String, Q2 As String
   Dim IdMovDoc As Long
   Dim TipoVal() As NTipoVal_t
   Dim NombreCampo As String
   Dim IdxOtrosVal As String
   Dim IdTipoValLib As Integer
   Dim EsRebaja As Boolean
   Dim Valor As Double
   Dim RetParcial As Integer
   Dim IVAInmueble As Integer
   Dim ValIVA As Double
   Dim IVAIrrec As Integer    'tipo de IVA Irrecuperable
   Dim ValIVAIrrec As Double
   Dim ValIVAActFijo As Double
   Dim DimDoc As String
   Dim ValOtros As Double
   Dim TipoLib As Integer
   Dim CodSIIDTEIvaIrrec As Integer, Idx As Integer
   Dim FldArray(2) As AdvTbAddNew_t
   Dim IdCuentaAfecto As Long, IdCuentaExento As Long, IdCuentaTotal As Long
   
   If lIdDoc <= 0 Then
      Exit Sub
   End If
   
   TipoLib = CbItemData(Cb_TipoLib)
   

   Lin = Grid.FixedRows
      
   ReDim TipoVal(UBound(gTipoValLib))   'de más, porque gTipoValLib incluye todos los libros, no sólo el actual
   
   IdxOtrosVal = -1
   EsRebaja = gTipoDoc(GetTipoDoc(TipoLib, CbItemData(Cb_TipoDoc))).EsRebaja

   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If
      
      IdTipoValLib = Val(Grid.TextMatrix(i, C_IDTIPOVALLIB))
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
         

         FldArray(0).FldName = "IdDoc"
         FldArray(0).FldValue = lIdDoc
         FldArray(0).FldIsNum = True
         
         FldArray(1).FldName = "IdEmpresa"
         FldArray(1).FldValue = gEmpresa.id
         FldArray(1).FldIsNum = True
                     
         FldArray(2).FldName = "Ano"
         FldArray(2).FldValue = gEmpresa.Ano
         FldArray(2).FldIsNum = True
                  
         IdMovDoc = AdvTbAddNewMult(DbMain, "MovDocumento", "IdMovDoc", FldArray)
         
         
         Grid.TextMatrix(i, C_IDMOV) = IdMovDoc
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
         
         lMovEdited = True
         
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM MovDocumento WHERE IdMovDoc = " & Grid.TextMatrix(i, C_IDMOV)
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdMovDoc = " & Grid.TextMatrix(i, C_IDMOV)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "MovDocumento", Q1)
         
         lMovEdited = True
         
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Or Val(Grid.TextMatrix(i, C_ORDEN)) <> Lin Or (TipoLib = LIB_RETEN And lMovEdited) Then  'Se agrega la condición del libro de retenciones por error que no se pudo reproducir, que marcaba dos registgrs como Total del Documento (FCA 21 10 21)
         Q1 = "UPDATE MovDocumento SET "
         Q1 = Q1 & "  IdDoc = " & lIdDoc
         Q1 = Q1 & ", Orden = " & Lin
         Q1 = Q1 & ", IdCuenta = " & Grid.TextMatrix(i, C_IDCUENTA)
         Q1 = Q1 & ", Debe = " & vFmt(Grid.TextMatrix(i, C_DEBE))
         Q1 = Q1 & ", Haber = " & vFmt(Grid.TextMatrix(i, C_HABER))
         '636269
         'Q1 = Q1 & ", Glosa = '" & ParaSQL(Grid.TextMatrix(i, C_GLOSA)) & "'"
         Q1 = Q1 & ", Glosa = '" & Left(ParaSQL(Grid.TextMatrix(i, C_GLOSA)), 50) & "'"
         '636269
         Q1 = Q1 & ", IdTipoValLib = " & vFmt(Grid.TextMatrix(i, C_IDTIPOVALLIB))
         Q1 = Q1 & ", Tasa = " & IIf(Grid.TextMatrix(i, C_TASA) <> "", str(vFmt(Grid.TextMatrix(i, C_TASA))), 0)
         Q1 = Q1 & ", EsRecuperable = " & IIf(Grid.TextMatrix(i, C_ESRECUPERABLE) = "", 0, ValSiNo(Grid.TextMatrix(i, C_ESRECUPERABLE)))
         Q1 = Q1 & ", CodSIIDTE = '" & Grid.TextMatrix(i, C_CODSIIDTE) & "'"
         If (TipoLib = LIB_COMPRAS And IdTipoValLib = LIBCOMPRAS_TOTAL) Or (TipoLib = LIB_VENTAS And IdTipoValLib = LIBVENTAS_TOTAL) Or (TipoLib = LIB_RETEN And IdTipoValLib = LIBRETEN_NETO) Or (TipoLib = LIB_REMU And IdTipoValLib = LIBREMU_TOTAL) Or (TipoLib = LIB_OTROS And IdTipoValLib = LIBOTROS_TOTAL) Then
            Q1 = Q1 & ", EsTotalDoc = 1"
            IdCuentaTotal = Val(Grid.TextMatrix(i, C_IDCUENTA))
         Else
            Q1 = Q1 & ", EsTotalDoc = 0"
         End If
         Q1 = Q1 & ", IdCCosto = " & vFmt(Grid.TextMatrix(i, C_IDCCOSTO))
         Q1 = Q1 & ", IdAreaNeg = " & vFmt(Grid.TextMatrix(i, C_IDAREANEG))
         Q1 = Q1 & " WHERE IdMovDoc = " & Grid.TextMatrix(i, C_IDMOV)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         '2861733 tema 2
         If ValidarCuentaAFijo(Val(Grid.TextMatrix(i, C_IDCUENTA)), False, False) Then
           Call SaveAreaCentroActFijo(Val(Grid.TextMatrix(i, C_IDCUENTA)), vFmt(Grid.TextMatrix(i, C_IDAREANEG)), vFmt(Grid.TextMatrix(i, C_IDCCOSTO)))
         End If
         '2861733 tema 2
         
         lMovEdited = True
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
      
         Lin = Lin + 1
                    
         If (TipoLib = LIB_COMPRAS And IdTipoValLib = LIBCOMPRAS_AFECTO) Or (TipoLib = LIB_VENTAS And IdTipoValLib = LIBVENTAS_AFECTO) Or (TipoLib = LIB_RETEN And IdTipoValLib = LIBRETEN_BRUTO) Then
            If IdCuentaAfecto = 0 Then
               IdCuentaAfecto = Val(Grid.TextMatrix(i, C_IDCUENTA))
            End If
         End If
         
         If (TipoLib = LIB_COMPRAS And IdTipoValLib = LIBCOMPRAS_EXENTO) Or (TipoLib = LIB_VENTAS And IdTipoValLib = LIBVENTAS_EXENTO) Or (TipoLib = LIB_RETEN And IdTipoValLib = LIBRETEN_HONORSINRET) Then
            If IdCuentaExento = 0 Then
               IdCuentaExento = Val(Grid.TextMatrix(i, C_IDCUENTA))
            End If
         End If
         
         '2937156
          If (TipoLib = LIB_COMPRAS And IdTipoValLib = LIBCOMPRAS_TOTAL) Or (TipoLib = LIB_VENTAS And IdTipoValLib = LIBVENTAS_TOTAL) Or (TipoLib = LIB_RETEN And IdTipoValLib = LIBRETEN_NETO) Or (TipoLib = LIB_REMU And IdTipoValLib = LIBREMU_TOTAL) Or (TipoLib = LIB_OTROS And IdTipoValLib = LIBOTROS_TOTAL) Then
               IdCuentaTotal = Val(Grid.TextMatrix(i, C_IDCUENTA))
         End If
         '2937156
         
         NombreCampo = GetNombreCampo(IdTipoValLib)
         Valor = CalcValor(TipoLib, NombreCampo, EsRebaja, vFmt(Grid.TextMatrix(i, C_DEBE)), vFmt(Grid.TextMatrix(i, C_HABER)))
              
         'sumamos montos por tipo de valor
            
         For j = 0 To UBound(TipoVal)
                        
            If TipoVal(j).NombreCampo = "" Then
               'Valor = CalcValor(TipoLib, NombreCampo, EsRebaja, vFmt(Grid.TextMatrix(i, C_DEBE)), vFmt(Grid.TextMatrix(i, C_HABER)))
               TipoVal(j).IdTipoValLib = IdTipoValLib
               TipoVal(j).NombreCampo = NombreCampo
               TipoVal(j).Valor = Valor
               Exit For
               
            ElseIf TipoVal(j).NombreCampo = NombreCampo Then
               'Valor = CalcValor(TipoLib, TipoVal(j).NombreCampo, EsRebaja, vFmt(Grid.TextMatrix(i, C_DEBE)), vFmt(Grid.TextMatrix(i, C_HABER)))
               TipoVal(j).Valor = TipoVal(j).Valor + Valor
               Exit For
            End If
            
         Next j
         
         DimDoc = GetDiminutivoDoc(TipoLib, ItemData(Cb_TipoDoc))
         
         If TipoLib = LIB_COMPRAS Then
         
            If DimDoc = "FCC" Or DimDoc = "NCF" Or DimDoc = "NDF" Then
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVARETPARC Then
                  RetParcial = 1
               End If
            End If
            
            If DimDoc = "FAC" Or DimDoc = "NDC" Or DimDoc = "NCC" Or DimDoc = "IMP" Or DimDoc = "FCE" Or DimDoc = "FCC" Or DimDoc = "FIC" Then
               
               'es iva irrecuperable?
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVAIRREC Or (Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) >= LIBCOMPRAS_IVAIRREC1 And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) <= LIBCOMPRAS_IVAIRREC9) Then
                  
                  ValIVAIrrec = ValIVAIrrec + Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
                  
                  If ValIVAIrrec > 0 And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) >= LIBCOMPRAS_IVAIRREC1 And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) <= LIBCOMPRAS_IVAIRREC9 Then
                     Idx = GetTipoValLib(TipoLib, Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)))
                     CodSIIDTEIvaIrrec = gTipoValLib(Idx).CodSIIDTE
                  Else
                     CodSIIDTEIvaIrrec = 0
                  End If
                  
               End If
                'es IVA Crédito?
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVACREDFISC Then
                  ValIVA = Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
               End If
               
               'Es IVA Activo Fijo?
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVAACTFIJO Then
                  ValIVAActFijo = Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
               End If
               
            End If
            
            If DimDoc = "FAC" Or DimDoc = "NCC" Or DimDoc = "NDC" Then
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVAADQCONSTINMUEBLES Then
                  IVAInmueble = 1
               End If
            End If
            
                        
         ElseIf TipoLib = LIB_VENTAS Then
            If Cb_TipoDoc = "Factura de Compra" Then
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBVENTAS_IVARETPARC Then
                  RetParcial = 1
               End If
            End If
            
            If DimDoc = "FAV" Or DimDoc = "NCV" Or DimDoc = "NDV" Then
               If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBVENTAS_IVAADQCONSTINMUEBLES Then
                  IVAInmueble = 1
               End If
            End If
         End If
      
      End If
      
   Next i
   
   If lMovEdited Then
   
      Q1 = ""
      
      For j = 0 To UBound(TipoVal)
         
         If TipoVal(j).IdTipoValLib = 0 Then
            Exit For
         End If
      
         If TipoVal(j).IdTipoValLib <> 0 Then
            If TipoLib = LIB_VENTAS Or TipoLib = LIB_COMPRAS Or TipoLib = LIB_RETEN Then
               Q1 = Q1 & ", " & TipoVal(j).NombreCampo & "=" & TipoVal(j).Valor
            
            ElseIf TipoLib = LIB_OTROS Or TipoLib = LIB_REMU Then
               If TipoVal(j).IdTipoValLib = LIBOTROS_VALOR Then
                  If TipoVal(j).Valor < 0 Then
                     Q1 = Q1 & ",  Afecto= " & Abs(TipoVal(j).Valor)    'Haber
                  Else
                     Q1 = Q1 & ",  Exento= " & Abs(TipoVal(j).Valor)    'Debe
                  End If
               Else   'total
                  Q1 = Q1 & ", " & TipoVal(j).NombreCampo & "=" & Abs(TipoVal(j).Valor)
               End If
            End If
            
         End If
         
      Next j
                           
      If ValIVAIrrec > 0 Then
                     
         If ValIVAIrrec = ValIVA Then
            IVAIrrec = IVAIRREC_TOTAL
         ElseIf ValIVAIrrec < ValIVA Then
            IVAIrrec = IVAIRREC_PARCIAL
         ElseIf ValIVA = 0 Then
            IVAIrrec = IVAIRREC_TOTAL
         End If
         
      Else
         IVAIrrec = IVAIRREC_CERO
         
      End If
      
   
      'limpiamos los campos
      Q2 = "UPDATE Documento SET Afecto=0, Exento=0, IVA=0, OtroImp=0, OtrosVal=0, Total=0, IdANegCCosto='' "
      Q2 = Q2 & " WHERE IdDoc=" & lIdDoc & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q2)
      
      'OJO: En el campo IVA va IVA Crédito/Débito + IVA Act. Fijo + IVA Irrecuperable en todos sus tipos
      
      'actualizamos los campos, incluyendo los que limpiamos
      Q2 = "UPDATE Documento SET MovEdited=-1" & Q1
      Q2 = Q2 & ", FacCompraRetParcial = " & RetParcial & ", IVAIrrecuperable = " & IVAIrrec & ", ValIVAIrrec = " & ValIVAIrrec
      Q2 = Q2 & ", CodSIIDTEIVAIrrec = " & CodSIIDTEIvaIrrec & ", IVAInmueble = " & IVAInmueble & ", IVAActFijo = " & ValIVAActFijo
      Q2 = Q2 & ", IdCuentaAfecto = " & IdCuentaAfecto
      Q2 = Q2 & ", IdCuentaExento = " & IdCuentaExento
      Q2 = Q2 & ", IdCuentaTotal = " & IdCuentaTotal
      Q2 = Q2 & "  WHERE IdDoc=" & lIdDoc & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q2)
   
      'actualizamos OtrosImp calculándolo exactamente de la misma manera como se calcula en el libro de ComprasVentas
      If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Then
         
         Q1 = "SELECT Total, Afecto, Exento, IVA FROM Documento WHERE IdDoc=" & lIdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            'exactamente igual como se calcula en el Libro de ComprasVentas
            ValOtros = Abs(vFld(Rs("Total")) - (vFld(Rs("Exento")) + vFld(Rs("Afecto")) + vFld(Rs("IVA"))))
            Q1 = "UPDATE Documento SET OtroImp = " & ValOtros & "  WHERE IdDoc=" & lIdDoc & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
         End If
         Call CloseRs(Rs)
      End If
      
   End If
   
   'Tracking 3227543
        Call SeguimientoMovDocumento(lIdDoc, gEmpresa.id, gEmpresa.Ano, "FrmDocLib.SaveGrMov", "", 1, "", 1, 1)
    ' fin 3227543
   
End Sub
Private Sub FrmEnable()
   Dim bool As Integer
   Dim i As Integer
   Dim EdEstado As Boolean
   Dim DimDoc As String
   Dim TipoLib As Integer, TipoDoc As Integer
   
   EdEstado = (ItemData(Cb_Estado) = ED_PENDIENTE Or (Year(GetTxDate(Tx_FEmision)) < gEmpresa.Ano And ItemData(Cb_Estado) = ED_CENTRALIZADO))
   
   bool = ((lOper = O_EDIT And lIdDoc > 0 And EdEstado) Or lOper = O_NEW)
   
   If lOper = O_EDIT Or lOper = O_NEW Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         bool = False
      End If
   End If
   
   lEditEnable = bool
   
   Bt_Duplicate.Enabled = bool
   Bt_Del.Enabled = bool
   Bt_MoveUp.Enabled = bool
   Bt_MoveDown.Enabled = bool
   Bt_Duplicate.Enabled = bool
   Bt_Cut.Enabled = bool
   Bt_Copy.Enabled = bool
   Bt_Paste.Enabled = bool
      
   Cb_TipoLib.Enabled = bool
   Cb_TipoDoc.Enabled = bool
   
   'bool = bool And (ItemData(Cb_TipoDoc) <> -1)
     
   Call SetTxRO(Tx_Rut, Not bool)
   
   Ch_Rut.Enabled = bool
   
   Call SetTxRO(Tx_NumDoc, Not bool)
   Call SetTxRO(Tx_FEmision, Not bool)
   Call SetTxRO(Tx_FEmisionOri, Not bool)
   Call SetTxRO(Tx_FVenc, Not bool)
'   Call SettxRO(Tx_NumDocHasta, Not bool)
      
   If Tx_NumDocHasta.Locked = False Then
      Call SetTxRO(Tx_NumDocHasta, Not bool)
   End If

   
   Call SetTxRO(Tx_ValTotal, Not bool)
   Call SetTxRO(Tx_Descrip, Not bool)
   Call SetTxRO(Tx_CorrInterno, Not bool)
   
   Ch_DTE.Enabled = bool
   
   If CbItemData(Cb_TipoLib) > 0 Then
      TipoLib = CbItemData(Cb_TipoLib)
   Else
      Exit Sub
   End If
   
   If CbItemData(Cb_TipoDoc) > 0 Then
      TipoDoc = CbItemData(Cb_TipoDoc)
   Else
      Exit Sub
   End If

   DimDoc = GetDiminutivoDoc(TipoLib, TipoDoc)
   
   If DimDoc = "FAV" Or DimDoc = "FVE" Or DimDoc = "NCV" Or DimDoc = "NDV" Then
      Ch_DelGiro.Enabled = bool
   Else
      Ch_DelGiro = 0
      Ch_DelGiro.Enabled = False
   End If
   
   If DimDoc = "FAC" Or DimDoc = "FCE" Or DimDoc = "NCC" Or DimDoc = "NDC" Then
      Ch_CompraBienRaiz.Enabled = bool
   Else
      Ch_CompraBienRaiz = 0
      Ch_CompraBienRaiz.Enabled = False
   End If
   
   Cb_Entidad.Enabled = bool
   Cb_Nombre.Enabled = bool
   Cb_Estado.Enabled = bool
   Cb_Impto.Enabled = bool
   Cb_TipoReten.Enabled = bool
   Cb_Sucursal.Enabled = bool
   
   For i = 0 To 2
      Bt_Fecha(i).Enabled = bool
   Next i
   
   Grid.Locked = Not bool
   
   Bt_NewEnt.Enabled = bool
   
   Bt_ClearMov.Enabled = bool
   
   If bool = False Then
      Bt_OK.visible = False
      Bt_Cancel.Caption = "Cerrar"
      Bt_Cancel.Top = Bt_OK.Top
   End If
   
   If lOper = O_EDIT And (CbItemData(Cb_TipoLib) = LIB_COMPRAS Or ItemData(Cb_TipoLib) = LIB_VENTAS Or ItemData(Cb_TipoLib) = LIB_RETEN) Then
      Cb_TipoLib.Enabled = False
      'Cb_TipoDoc.Enabled = False
      Cb_Estado.Enabled = False
   End If
   
     '3284709
'   If ItemData(Cb_Estado) = ED_PENDIENTE And CbItemData(Cb_TipoLib) = LIB_VENTAS And DimDoc = TDOC_FAVEXENTA And Year(GetTxDate(Tx_FEmision)) < gEmpresa.Ano Then
'    Cb_Estado.Enabled = True
'    MsgBox1 "Documento proviene del año anterior, favor de dejar estado del documento en centralizado ya que se encuentra en estado pendiente.", vbInformation
'   Else
'    Cb_Estado.Enabled = False
'   End If
'   '3284709
   

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim IdCuenta As Long
   Dim Cod As String
   Dim DescCta As String
   Dim NombCta As String
   Dim UltimoNivel As Boolean
   Dim CtaActiva As Boolean
   Dim i As Integer
   
   Action = vbOK
   
   Select Case Col
   
      Case C_CODCUENTA
      
         Cod = Trim(ReplaceStr(Value, "-", ""))
         If Len(Cod) < Len(VFmtCodigoCta(gFmtCodigoCta)) Then   'asumimos que está usando nombre corto
            NombCta = UCase(Trim(Value))
            Cod = ""
         Else
            NombCta = ""
         End If
         
         IdCuenta = GetIdCuenta(NombCta, Cod, DescCta, UltimoNivel)
         
         If IdCuenta = 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
         
         ElseIf UltimoNivel = False Then
            MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
            Action = vbCancel
         
         Else
            
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta   'se asigna porque se usa en GridActivoFijo
            Value = Format(Cod, gFmtCodigoCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta
                                       
            If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_AREANEG) = 0 Then
               Grid.TextMatrix(Grid.Row, C_IDAREANEG) = 0
               Grid.TextMatrix(Grid.Row, C_AREANEG) = ""
            End If
               
            If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CCOSTO) = 0 Then
               Grid.TextMatrix(Grid.Row, C_IDCCOSTO) = 0
               Grid.TextMatrix(Grid.Row, C_CCOSTO) = ""
            End If
               
            If EsCuentaActFijo(Grid.TextMatrix(Row, C_IDCUENTA)) Then
               MsgBox1 "Recuerde asignar el valor de IVA Activo Fijo, utilizando la columna Clasificación, para el caso en que desee exportar a HR-IVA.", vbInformation + vbOKOnly
            End If
               
         End If
         
         If Grid.Row = Grid.rows - 1 Then
            Grid.rows = Grid.rows + 1
         End If
         
         
      Case C_DEBE
      
         If vFmt(Value) < 0 And Grid.CbList(C_TIPOVALLIB) <> "Rebaja 65%" Then
            MsgBeep vbExclamation
            Action = vbCancel
            
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_HABER)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_HABER) = ""
            End If
            Call CalcTot
            
         End If
         
      Case C_HABER
   
         If vFmt(Value) < 0 And Grid.CbList(C_TIPOVALLIB) <> "Rebaja 65%" Then
            MsgBeep vbExclamation
            Action = vbCancel
            
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_DEBE)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_DEBE) = ""
            End If
            Call CalcTot
            
         End If
         
      'Case C_GLOSA
      
      Case C_TIPOVALLIB
'         Grid.TextMatrix(Row, C_IDTIPOVALLIB) = ItemData(Grid.CbList(C_TIPOVALLIB))
         Grid.TextMatrix(Row, C_IDTIPOVALLIB) = lcbTipoValLib.ItemData
         If Grid.TextMatrix(Row, C_IDTIPOVALLIB) > 0 Then
            Grid.TextMatrix(Row, C_ATRIBUTO) = gTipoValLib(GetTipoValLib(lCurTipoLib, Grid.TextMatrix(Row, C_IDTIPOVALLIB))).Atributo
            Grid.TextMatrix(Row, C_CODSIIDTE) = gTipoValLib(GetTipoValLib(lCurTipoLib, Grid.TextMatrix(Row, C_IDTIPOVALLIB))).CodSIIDTE
            Grid.TextMatrix(Row, C_TASA) = ""
            Grid.TextMatrix(Row, C_ESRECUPERABLE) = ""
          Else
            Grid.TextMatrix(Row, C_ATRIBUTO) = ""
            Grid.TextMatrix(Row, C_CODSIIDTE) = ""
            Grid.TextMatrix(Row, C_TASA) = ""
            Grid.TextMatrix(Row, C_ESRECUPERABLE) = ""
         End If
         
         'vemos si hay dos afectos
         If CbItemData(Cb_TipoLib) = LIB_COMPRAS And Val(Grid.TextMatrix(Row, C_IDTIPOVALLIB)) = LIBCOMPRAS_AFECTO Then
            lNetoDoc = 0
            For i = Grid.FixedRows To Grid.rows - 1
               If Not lMsgAfecto And i <> Row And Val(Grid.TextMatrix(Row, C_IDTIPOVALLIB)) = LIBCOMPRAS_AFECTO Then
                  MsgBox1 "Recuerde que si desglosa el monto Afecto en dos cuentas contables, deberá desglosar el IVA de las compras con su monto, descripción y clasificación.", vbInformation
                  lMsgAfecto = True
               End If
               If Val(Grid.TextMatrix(Row, C_IDTIPOVALLIB)) = LIBCOMPRAS_AFECTO Then   'o LIBVENTAS_AFECTO (son iguales)
                  lNetoDoc = lNetoDoc + Abs(vFmt(Grid.TextMatrix(Row, C_DEBE)) - vFmt(Grid.TextMatrix(Row, C_HABER)))
               End If
            Next i
         End If
         
         If CbItemData(Cb_TipoLib) = LIB_COMPRAS And Grid.TextMatrix(Row, C_IDTIPOVALLIB) >= LIBCOMPRAS_OTROSIMP Then
            'veamos si es imp. adicional descontinuado (Libro de Compras)
            If CbItemData(Cb_TipoLib) = LIB_COMPRAS And Val(lcbTipoValLib.Matrix(M_CODSIIDTE)) = 0 And Grid.TextMatrix(Row, C_ATRIBUTO) <> "IVAIRREC" And Grid.TextMatrix(Row, C_ATRIBUTO) <> "IVAACTFIJO" Then   'está descontinuado
'               MsgBox1 "Este impuesto está descontinuado y no podrá ser utilizado en el Libro Electrónico de Compras.", vbExclamation
               MsgBox1 "ADVERTENCIA: Este impuesto no podrá ser utilizado en el Libro Electrónico de Compras.", vbInformation
'               Action = vbCancel
            End If
            
            'faltan datos de tasa y código SII DTE
            If CbItemData(Cb_TipoLib) = LIB_COMPRAS And Val(lcbTipoValLib.Matrix(M_CODSIIDTE)) <> 0 And Grid.TextMatrix(Row, C_ATRIBUTO) <> "IVAIRREC" And Grid.TextMatrix(Row, C_ATRIBUTO) <> "IVAACTFIJO" Then
               MsgBox1 "Utilice el botón 'Agregar Imp. Adicionales' para asignar automáticamente la Tasa, de acuerdo a la configuración de su empresa." & vbCrLf & vbCrLf & "Sin este dato, no podrá emitir el 'Libro Electrónico de Compras'", vbExclamation
            End If
         End If
  
      Case C_CCOSTO
         Grid.TextMatrix(Row, C_IDCCOSTO) = ItemData(Grid.CbList(C_CCOSTO))
      
      Case C_AREANEG
         Grid.TextMatrix(Row, C_IDAREANEG) = ItemData(Grid.CbList(C_AREANEG))
         
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
   Dim idMov As Long
   Dim orden As Integer
   Dim IdCuenta As Long
   Dim DescCta As String
   Dim NombCuenta As String
   Dim FrmPlan As FrmPlanCuentas
   Dim ValPrevLine As Boolean
   Dim Msg As String
   Dim DetMov As DetMovim_t
   Dim Q1 As String
   Dim Rs As Recordset
   Dim CodCta As String
   
   'no permitimos editar
   'Exit Sub
   
   'si no está habilitado para modificar, nos vamos.
   If Tx_NumDoc.Locked Or Row < Grid.FixedRows Then
      Exit Sub
   End If
               
   If ItemData(Cb_TipoDoc) <= 0 Then
      MsgBox1 "Falta definir el tipo de documento.", vbExclamation + vbOKOnly
      Exit Sub
   End If
               
   idMov = Val(Grid.TextMatrix(Row, C_IDMOV))
   orden = Val(Grid.TextMatrix(Row, C_ORDEN))
      
   'Linea anterior tiene valor o está eliminada?
   ValPrevLine = (Row > Grid.FixedRows And Val(Grid.TextMatrix(Row - 1, C_ORDEN)) > 0 And Trim(Grid.TextMatrix(Row - 1, C_CUENTA)) <> "" And (Val(Grid.TextMatrix(Row - 1, C_DEBE)) > 0 Or Val(Grid.TextMatrix(Row - 1, C_HABER)) > 0)) Or Grid.RowHeight(Row - 1) = 0
   
   If Not (Row = Grid.FixedRows Or (Row > Grid.FixedRows And orden > 0) Or ValPrevLine) Then
      Exit Sub
   End If
      
   'sólo pueden ingresar valores en debe, haber, glosa, etc., si seleccionó una cuenta
   If Col <> C_CODCUENTA And Col <> C_LSTCUENTA And Grid.TextMatrix(Row, C_CUENTA) = "" Then
      Exit Sub
   End If
      
   If idMov = 0 Then    'nuevo
      Grid.TextMatrix(Row, C_ORDEN) = Row
      Call FGrSetPicture(Grid, Row, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
      
      If Row >= Grid.rows - 2 Then
         Grid.rows = Grid.rows + 1
      End If
     
   End If
   
   
   Select Case Col
   
      Case C_CODCUENTA
            
         Grid.TxBox.MaxLength = 20
         EdType = FEG_Edit
         
'      Case C_CUENTA, C_LSTCUENTA
'
'         Set FrmPlan = New FrmPlanCuentas
'
'         If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta) = vbOK Then
'            If DescCta <> "" Then
'               Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
'               Grid.TextMatrix(Row, C_CODCUENTA) = FmtCodCuenta(CodCta)
'               Grid.TextMatrix(Row, C_CUENTA) = DescCta
'
'               Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
'            End If
'
'         End If
'         Set FrmPlan = Nothing
                     
      Case C_DEBE
         EdType = FEG_Edit
         
      Case C_HABER
         EdType = FEG_Edit
         
      Case C_GLOSA
         Grid.TxBox.MaxLength = 50
         EdType = FEG_Edit
         
      Case C_TIPOVALLIB, C_CCOSTO, C_AREANEG
         EdType = FEG_List
      
   End Select
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   Dim IdCuenta As Long
   Dim DescCta As String
   Dim NombCuenta As String
   Dim FrmPlan As FrmPlanCuentas
   Dim CodCta As String
   
   Col = Grid.MouseCol
   Row = Grid.MouseRow
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Locked = True Then
      Exit Sub
   End If
      
   If Col = C_LSTCUENTA Then

      Set FrmPlan = New FrmPlanCuentas

      If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta) = vbOK Then
         If DescCta <> "" Then
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
            Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta

            If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_AREANEG) = 0 Then
               Grid.TextMatrix(Grid.Row, C_IDAREANEG) = 0
               Grid.TextMatrix(Grid.Row, C_AREANEG) = ""
            End If
               
            If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CCOSTO) = 0 Then
               Grid.TextMatrix(Grid.Row, C_IDCCOSTO) = 0
               Grid.TextMatrix(Grid.Row, C_CCOSTO) = ""
            End If
            
            If EsCuentaActFijo(Grid.TextMatrix(Row, C_IDCUENTA)) Then
               MsgBox1 "Recuerde asignar el valor de IVA Activo Fijo, utilizando la columna Clasificación, para el caso en que desee exportar a HR-IVA.", vbInformation + vbOKOnly
            End If
            
            Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)

        End If

      End If
      Set FrmPlan = Nothing
      
   ElseIf Col = C_TASA Or Col = C_ESRECUPERABLE Then
      If CbItemData(Cb_TipoLib) = LIB_COMPRAS Then
         MsgBox1 "Utilice el botón 'Agregar Imp. Adicionales' para seleccionar el impuesto adicional y la tasa correspondiente.", vbInformation
      End If
      
   End If
   
End Sub

Private Sub CalcTot()
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   
   TotDebe = 0
   TotHaber = 0
   lNetoDoc = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.RowHeight(i) > 0 Then     ' no está borrado
         TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
         TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      End If
      If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_AFECTO Then   'o Libventas_afecto (es el mismo)
         lNetoDoc = lNetoDoc + Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
      End If
   Next i
         
   GridTot.TextMatrix(0, C_DEBE) = Format(TotDebe, BL_NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(TotHaber, BL_NUMFMT)
   
End Sub

Private Sub Bt_MoveDown_Click()
   Dim Row As Integer
   Dim Aux As String
   Dim i As Integer
      
   If Grid.Row = Grid.rows - 1 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_ORDEN) = "" Or Grid.TextMatrix(Grid.Row + 1, C_ORDEN) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
  
   Row = Grid.Row
   
   For i = 0 To Grid.Cols - 1
      If i <> C_ORDEN Then
         Aux = Grid.TextMatrix(Row + 1, i)
         Grid.TextMatrix(Row + 1, i) = Grid.TextMatrix(Row, i)
         Grid.TextMatrix(Row, i) = Aux
      End If
   Next i

   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   Call FGrModRow(Grid, Row + 1, FGR_U, C_IDMOV, C_UPDATE)
   
   Grid.Row = Row + 1
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = 0
   Grid.ColSel = Grid.Cols - 1
      
End Sub

Private Sub Bt_MoveUp_Click()
   Dim Row As Integer
   Dim Aux As String
   Dim i As Integer
      
   If Grid.Row = Grid.FixedRows Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_ORDEN) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Row = Grid.Row
   
   For i = 0 To Grid.Cols - 1
      If i <> C_ORDEN Then
         Aux = Grid.TextMatrix(Row - 1, i)
         Grid.TextMatrix(Row - 1, i) = Grid.TextMatrix(Row, i)
         Grid.TextMatrix(Row, i) = Aux
      End If
   Next i
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   Call FGrModRow(Grid, Row - 1, FGR_U, C_IDMOV, C_UPDATE)

   Grid.Row = Row - 1
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = 0
   Grid.ColSel = Grid.Cols - 1
      
End Sub

Private Sub Bt_Cut_Click()
   
   If Grid.Col = C_CUENTA Then
      Exit Sub
   End If
   
   Clipboard.Clear
   Call Clipboard.SetText(Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col))
   
   Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = ""
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDMOV, C_UPDATE)
   
   Call CalcTot
   
End Sub

Private Sub bt_Copy_Click()
   
   Clipboard.Clear
   Clipboard.SetText Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col)
   
End Sub

Private Sub Bt_Paste_Click()
   Dim Fmt As Integer
   Dim DVal As Double
      
   If Grid.Col = C_CUENTA Then
      Exit Sub
   End If
   
   If Clipboard.GetFormat(vbCFText) = False Then
      Exit Sub
   End If
   
   DVal = Val(vFmt(Clipboard.GetText))
   
   If (Grid.FlxGrid.Col = C_DEBE Or Grid.FlxGrid.Col = C_HABER) Then
      If DVal <> 0 Then
         Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = Format(Abs(DVal), BL_NUMFMT)
         Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDMOV, C_UPDATE)
         Call CalcTot
      End If
   Else
      Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = Clipboard.GetText
      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDMOV, C_UPDATE)
   End If
      
End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBE, C_HABER, Grid.FlxGrid.Row, Grid.FlxGrid.RowSel)
   
   Set Frm = Nothing
End Sub

Private Sub bt_Cuadrar_Click()
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim Diff As Double
   Dim Row As Integer
   Dim i As Integer
   
   Row = Grid.Row
   
   If Trim(Grid.TextMatrix(Row, C_CODCUENTA)) = "" Then
      Exit Sub
   End If
   
   'limpiamos la fila actual para que no entre en la suma
   Grid.TextMatrix(Row, C_DEBE) = ""
   Grid.TextMatrix(Row, C_HABER) = ""
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      If Grid.RowHeight(i) > 0 Then
         SumDebe = vFmt(Grid.TextMatrix(i, C_DEBE)) + SumDebe
         SumHaber = vFmt(Grid.TextMatrix(i, C_HABER)) + SumHaber
      End If
   Next i
   
   Diff = SumDebe - SumHaber
   If Diff = 0 Then
      MsgBox1 "Los movimientos ya están cuadrados. No se puede sugerir un valor.", vbExclamation
      Exit Sub
   End If
   
   'Debo cuadrar el debe
   If Diff < 0 Then
      Grid.TextMatrix(Row, C_DEBE) = Format(Abs(Diff), BL_NUMFMT)
   Else
      Grid.TextMatrix(Row, C_HABER) = Format(Diff, BL_NUMFMT)
   End If
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   
   Call CalcTot
   
End Sub

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Col As Integer
   Dim Row As Integer
   Dim Valor As Double
   
   Col = Grid.Col
   Row = Grid.Row
   
   'If Col <> C_DEBE And Col <> C_HABER Then
   '   MsgBox1 "Esta opción se utiliza sólo en las columnas Debe y Haber de los movimientos", vbExclamation
   '   Exit Sub
   'End If
   
   'If Trim(Grid.TextMatrix(Row, C_CUENTA)) = "" Then
   '   Exit Sub
   'End If
   
   If Col = C_DEBE Or Col = C_HABER Then
      Valor = vFmt(Grid.TextMatrix(Row, Col))
   End If
   
   Set Frm = New FrmConverMoneda
   If Frm.FSelect(Valor) = vbOK Then
      
      If Tx_Descrip.Enabled And (Col = C_DEBE Or Col = C_HABER) Then
         
         Grid.TextMatrix(Row, Col) = Format(Valor, BL_NUMFMT)
         
         If Col = C_DEBE And Trim(Grid.TextMatrix(Row, C_HABER)) <> "" Then
            Grid.TextMatrix(Row, C_HABER) = ""
         ElseIf Col = C_HABER And Trim(Grid.TextMatrix(Row, C_DEBE)) <> "" Then
            Grid.TextMatrix(Row, C_DEBE) = ""
         End If
         
         Call CalcTot
      End If
   End If
   Set Frm = Nothing
   
End Sub
Private Sub Bt_Cuentas_Click()
   Dim Frm As FrmPlanCuentas
   
   Set Frm = New FrmPlanCuentas

   Call Frm.FEdit(False)
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Print_Click()
   
   'Call SetUpPrtGrid
   
   Call gPrtLibros.PrtFlexGrid(Printer)
   
End Sub
Private Sub Bt_Calc_Click()
   Call Calculadora
End Sub

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, " Tipo: " & Cb_TipoDoc & vbTab & " Número Documento: " & Tx_NumDoc & vbTab & " " & Cb_Entidad & ": " & Cb_Nombre & vbTab & " Fecha: " & Tx_FEmision)
End Sub
Private Sub ClearForm()
   Dim Col As Integer
   Dim MesActual As Integer
   Dim Mes As Integer
   Dim F1 As Long, F2 As Long
   Dim i As Integer
   
   Call FGrClear(Grid)
   For Col = 0 To GridTot.Cols - 1
      GridTot.TextMatrix(0, Col) = ""
   Next Col
   
   Cb_TipoLib.ListIndex = -1
   Cb_Entidad.ListIndex = -1
   Call SelItem(Cb_Estado, ED_PENDIENTE)
   
   Ch_DTE = 0
   Tx_Rut = ""
   Tx_NumDoc = ""
   Tx_NumDocHasta = ""
   Tx_CorrInterno = ""
   Tx_Descrip = ""
   Tx_ValTotal = ""
   
   MesActual = GetMesActual()
   
   If lMes > 0 Then
      Mes = lMes
   ElseIf MesActual > 0 Then
      Mes = MesActual
   Else
      Mes = GetUltimoMesConMovs()
   End If

   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, Mes, 1), F1, F2)
   
   If month(Now) <> Mes Then
      Call SetTxDate(Tx_FEmision, F2)
      Call SetTxDate(Tx_FEmisionOri, F2)
   Else
      Call SetTxDate(Tx_FEmision, Now)
      Call SetTxDate(Tx_FEmisionOri, Now)
   End If

   Tx_FVenc = ""
   
   Grid.rows = Grid.FixedRows
   Call FGrVRows(Grid)
   
End Sub

Private Sub Tx_NumDocHasta_LostFocus()

   Tx_CantBoletas = ""
   If vFmt(Tx_NumDoc) > 0 Then
      If vFmt(Tx_NumDocHasta) > 0 And vFmt(Tx_NumDocHasta) >= vFmt(Tx_NumDoc) Then
         Tx_CantBoletas = Format(vFmt(Tx_NumDocHasta) - vFmt(Tx_NumDoc), NUMFMT) + 1
      End If
   End If

End Sub

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Exit Sub
   End If
   
'   If Not MsgValidCID(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
      
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut, Ch_Rut <> 0) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEnt = 0
   
   If Rs.EOF = False Then   'existe
      IdEnt = vFld(Rs("IdEntidad"))
            
      'seleccionamos el tipo de entidad y esto llena la lista de nombres de entidades
      For i = 1 To MAX_ENTCLASIF   'el cero tiene blanco
         If vFld(Rs("Clasif" & Cb_Entidad.ItemData(i))) <> 0 Then
            Cb_Entidad.ListIndex = i
            Exit For
         End If
      Next i
   
      'ahora seleccionamos la entidad
      For i = 0 To lcbNombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
   Else
      MsgBox1 "Este RUT no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Cb_Entidad.ListIndex = -1
      
   End If
      
   Call CloseRs(Rs)
   
   If Ch_Rut <> 0 Then
      AuxRut = FmtCID(vFmtCID(Tx_Rut))
      If AuxRut <> "0-0" Then
         Tx_Rut = AuxRut
      End If
   End If
   
End Sub
Private Function valida() As Boolean
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim IdxTipoDoc As Integer
   Dim i As Integer, j As Integer
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim RowCuadra As Integer
   Dim cont As Integer
   Dim Diff As Double
   Dim SinClasif As Integer
   Dim TipoVal() As NTipoVal_t
   Dim ValidTot As Boolean
   Dim RowRetParcial As Integer
   Dim RowRetTotal As Integer
   Dim RowIVA As Integer
   Dim DimDoc As String
   Dim RowIVAIrrec As Integer, TipoIVAIrrec As Integer, TotIVAIrrec As Double
   Dim RowIVAActFijo As Integer
   Dim IVACredDeb As String
   Dim ValLibIVA As Integer, ValLibIVAIrrec As Integer, ValLibIVA_AF As Integer
   Dim TotIVA As Double, TotAfecto As Double, Aux As Double
   Dim Linea As Integer
   Dim IdTipoValLib As Integer
   Dim TipoIVARetenido As Integer
   Dim Idx As Integer
   
   valida = False
   
   If Not ValidaIngresoDoc() Then
      Exit Function
   End If
               
   If Cb_TipoDoc.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un tipo de documento.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If

   If Tx_Rut <> "" And lcbNombre.ItemData <= 0 Then
      MsgBox1 "El RUT ingresado no tiene una entidad asociada.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If
   
   IdxTipoDoc = GetTipoDoc(lCurTipoLib, lCurTipoDoc)
   
   If IdxTipoDoc >= 0 Then
   
      If gTipoDoc(IdxTipoDoc).ExigeRUT And Tx_Rut = "" Then
         MsgBox1 "Debe ingresar una entidad.", vbExclamation
         Tx_Rut.SetFocus
         Exit Function
      End If
      
      If Not gTipoDoc(IdxTipoDoc).DocImpExp And Tx_Rut <> "" And Ch_Rut = 0 Then
         MsgBox1 "RUT inválido para este tipo de documento.", vbExclamation
         Tx_Rut.SetFocus
         Exit Function
      End If
   End If
   
   If Trim(Tx_NumDoc) = "" Or Val(Tx_NumDoc) = 0 Then
      MsgBox1 "Debe ingresar un número de documento.", vbExclamation
      Tx_NumDoc.SetFocus
      Exit Function
   End If
   
   'veamos si este documento ya ha sido ingresado
   If lOper = O_NEW Then
   
      Q1 = "SELECT IdDoc FROM Documento "
      Q1 = Q1 & " WHERE TipoLib=" & lCurTipoLib
      Q1 = Q1 & " AND TipoDoc=" & lCurTipoDoc
      Q1 = Q1 & " AND NumDoc='" & Trim(Tx_NumDoc) & "'"
      If lcbNombre.ItemData >= 0 Then
         Q1 = Q1 & " AND IdEntidad =" & lcbNombre.ItemData
      Else
         Q1 = Q1 & " AND IdEntidad =0"
      End If
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then   'ya existe
      
         If lcbNombre.ItemData < 0 Then
            If MsgBox1("Este documento ya ha sido ingresado al sistema, sin una entidad asociada. Es posible que esté duplicado." & vbNewLine & vbNewLine & "¿Desea verificar los datos antes de grabar?", vbQuestion + vbYesNo) = vbYes Then
               Call CloseRs(Rs)
               Exit Function
            End If
         Else
            MsgBox1 "Este documento ya ha sido ingresado al sistema.", vbExclamation + vbOKOnly
            Call CloseRs(Rs)
            Exit Function
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If Trim(Tx_FEmision) = "" Then
      MsgBox1 "Debe ingresar fecha de recepción del documento o de ingreso al libro.", vbExclamation
      Tx_FEmision.SetFocus
      Exit Function
   End If
   
   If lOper = O_EDIT And lOldFEmision > 0 Then
      If Year(GetTxDate(Tx_FEmision)) <> Year(lOldFEmision) Or month(GetTxDate(Tx_FEmision)) <> month(lOldFEmision) Then
         Call MsgBox1("No es posible cambiar el mes de recepción del documento.", vbOKOnly + vbExclamation)
         Exit Function
      End If
   End If
   
   If Trim(Tx_FEmisionOri) = "" Then
      MsgBox1 "Debe ingresar fecha de emisión.", vbExclamation
      Tx_FEmision.SetFocus
      Exit Function
   End If
   
   If Year(GetTxDate(Tx_FEmisionOri)) > gEmpresa.Ano Then
      Call MsgBox1("La fecha de emisión del documento es posterior al año en que está trabajando.", vbOKOnly + vbExclamation)
      Exit Function
   End If
   
'   If GetTxDate(Tx_FEmision) < GetTxDate(Tx_FEmisionOri) Then
'      MsgBox1 "La fecha de recepción del documento o de ingreso al libro, es anterior a la fecha de emisión del mismo.", vbExclamation
'      Tx_FVenc.SetFocus
'      Exit Function
'   End If
   
   
   If GetTxDate(Tx_FVenc) > 0 And GetTxDate(Tx_FEmisionOri) > GetTxDate(Tx_FVenc) Then
      MsgBox1 "Fecha de emisión mayor a la fecha de vencimiento.", vbExclamation
      Tx_FVenc.SetFocus
      Exit Function
   End If
   
   If lCurTipoLib = LIB_RETEN Then
      If ItemData(Cb_Impto) <= 0 Then
         MsgBox1 "Debe seleccionar el procentaje de impuesto.", vbExclamation
         Exit Function
      End If
      If ItemData(Cb_TipoReten) <= 0 Then
         MsgBox1 "Debe seleccionar el tipo de retención.", vbExclamation
         Exit Function
      End If
   End If
   
   If gEmpresa.Franq14Ter And Trim(Tx_Descrip) = "" Then
      MsgBox1 "Falta ingresar la descripción del documento." & vbCrLf & vbCrLf & "Ésta es obligatoria para el Libro de Caja, en empresas acogidas a 14TER.", vbExclamation
      Exit Function
   End If

   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) <> vFmt(vFmt(GridTot.TextMatrix(0, C_HABER))) Then
      Call MsgBox1("Los totales de las columnas DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      Exit Function
   End If
   
   ReDim TipoVal(UBound(gTipoValLib))   'de más, porque gTipoValLib incluye todos los libros, no sólo el actual
   
   DimDoc = GetDiminutivoDoc(lCurTipoLib, lCurTipoDoc)
   If lCurTipoLib = LIB_COMPRAS Then
      IVACredDeb = "Crédito"
   Else
      IVACredDeb = "Débito"
   End If
   
   'recorremos las líneas para validar Debe y Haber
   TotAfecto = 0
   Linea = 1
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      
      Linea = Val(Grid.TextMatrix(i, C_ORDEN))
       
      If Grid.RowHeight(i) = 0 Then
         GoTo NextRec
      End If
      
      If vFmt(Grid.TextMatrix(i, C_DEBE)) <> 0 And vFmt(Grid.TextMatrix(i, C_HABER)) <> 0 Then
         MsgBox1 "En la línea " & Linea & " los valores en las columnas DEBE y HABER son ambos mayores que 0.", vbExclamation + vbOKOnly
         Exit Function
         
      End If
      
      If Trim(Grid.TextMatrix(i, C_TIPOVALLIB)) = "" Then
         'MsgBox1 "En la línea " & i & " falta definir la clasificación del valor (Neto, Impuesto, Bruto).", vbExclamation + vbOKOnly
         'Exit Function
         SinClasif = SinClasif + 1
         
      Else   'sumamos aparición de cada tipo de valor
      
         If (lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_TOTAL) Or (lCurTipoLib = LIB_VENTAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBVENTAS_TOTAL) Or (lCurTipoLib = LIB_REMU And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBREMU_TOTAL) Or (lCurTipoLib = LIB_OTROS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBOTROS_TOTAL) Then
            
            ValidTot = (vFmt(Grid.TextMatrix(i, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(i, C_HABER)) = 0 And vFmt(Tx_ValTotal) = 0)
            ValidTot = ValidTot Or (vFmt(Grid.TextMatrix(i, C_DEBE)) <> 0 And vFmt(Tx_ValTotal) = vFmt(Grid.TextMatrix(i, C_DEBE)))
            ValidTot = ValidTot Or (vFmt(Grid.TextMatrix(i, C_HABER)) <> 0 And vFmt(Tx_ValTotal) = vFmt(Grid.TextMatrix(i, C_HABER)))
            
            'If vFmt(Tx_ValTotal) <> vFmt(Grid.TextMatrix(i, C_DEBE)) And vFmt(Tx_ValTotal) <> vFmt(Grid.TextMatrix(i, C_HABER)) Then
            If Not ValidTot Then
               MsgBox1 "En la línea " & i & " el valor no coincide con total del documento.", vbExclamation + vbOKOnly
               Exit Function
            End If
            
         End If
         
         For j = 0 To UBound(TipoVal)
            If TipoVal(j).IdTipoValLib = 0 Then
               TipoVal(j).IdTipoValLib = Val(Grid.TextMatrix(i, C_IDTIPOVALLIB))
               TipoVal(j).Count = 1
               Exit For
               
            ElseIf TipoVal(j).IdTipoValLib = Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) Then
               TipoVal(j).Count = TipoVal(j).Count + 1
               Exit For
               
            End If
         Next j
                   
         'sumamos el afecto para validad el valor del IVA
         If lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_AFECTO Then
            TotAfecto = TotAfecto + Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
         End If
         
      End If
      
      
      SumDebe = SumDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
      SumHaber = SumHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      
      If lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_OTROSIMP Then
         If MsgBox1("Todos aquellos Impuestos que estén clasificados como Otros Impuestos no serán traspasados a HR IVA." & vbCrLf & vbCrLf & "¿Esta seguro de mantener esta clasificación?", vbYesNo + vbQuestion) = vbNo Then
            MsgBox1 "Modifique la clasificación Otros Impuestos para poder grabar el documento.", vbInformation
            Exit Function
         End If
         
      End If
      
      If (lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVACREDFISC) Or (lCurTipoLib = LIB_VENTAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBVENTAS_IVADEBFISC) Then
         If RowIVA > 0 Then
            MsgBox1 "El detalle del documento incluye más de un IVA " & IVACredDeb & " Fiscal.", vbExclamation
            Exit Function
         End If
         RowIVA = i
      End If
      
      If lCurTipoLib = LIB_COMPRAS And Grid.TextMatrix(i, C_ATRIBUTO) = "IVAIRREC" Then   'IVA Irrec o IVA Irrec 1 a 9
         
         If Grid.TextMatrix(i, C_IDTIPOVALLIB) = LIBCOMPRAS_IVAIRREC And gEmpresa.Ano >= 2016 Then
            If MsgBox1("En la linea " & Linea & " debe seleccionar el tipo de IVA Irrecuperable para poder emitir el Libro Electrónico de Compras." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
               Exit Function
            End If
         End If
            
         If RowIVAIrrec > 0 And TipoIVAIrrec > 0 Then    'no se permite más de un tipo de IVA irrecuperable en el detalle. Puede ir más de un IVA Irrecuperable pero deben ser todos del mismo tipo
            If Val(Grid.TextMatrix(i, C_CODSIIDTE)) <> TipoIVAIrrec Then
               MsgBox1 "El detalle del documento incluye más de un tipo de IVA Irrecuperable.", vbExclamation
               Exit Function
            End If
         End If
         RowIVAIrrec = i
         TipoIVAIrrec = Val(Grid.TextMatrix(i, C_CODSIIDTE))
         TotIVAIrrec = TotIVAIrrec + Abs(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)))
      End If
      
      If lCurTipoLib = LIB_COMPRAS And Grid.TextMatrix(i, C_ATRIBUTO) = "IVAACTFIJO" Then
         If RowIVAActFijo > 0 Then
            MsgBox1 "El detalle del documento incluye más de un IVA Activo Fijo.", vbExclamation
            Exit Function
         End If
         RowIVAActFijo = i
      End If
      
      If lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) > LIBCOMPRAS_OTROSIMP And Grid.TextMatrix(i, C_ATRIBUTO) <> "IVAACTFIJO" And Grid.TextMatrix(i, C_ATRIBUTO) <> "IVAIRREC" Then
         If Grid.TextMatrix(i, C_CODSIIDTE) = "" Then
            If MsgBox1("En la línea " & Linea & " el impuesto seleccionado está descontinuado. No podrá emitir el Libro Electrónico de Compras." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
               Exit Function
            End If
         End If
         If vFmt(Grid.TextMatrix(i, C_TASA)) = 0 Then
            If MsgBox1("En la línea " & Linea & " falta definir la tasa correspondiente al impuesto. Utilice el botón 'Agregar Imp. Adicionales' para seleccionar los impuestos configurados para esta empresa." & vbCrLf & vbCrLf & "Sin este dato no podrá emitir el Libro Electrónico de Compras." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
               Exit Function
            End If
         End If
      End If
      
      IdTipoValLib = Val(Grid.TextMatrix(i, C_IDTIPOVALLIB))
      Idx = FindTipoValLib(lCurTipoLib, IdTipoValLib)
      TipoIVARetenido = gTipoValLib(Idx).TipoIVARetenido
            
      If lCurTipoLib = LIB_COMPRAS Then
         If DimDoc = "FCC" Or DimDoc = "NCF" Or DimDoc = "NDF" Then
            
'            If Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) >= LIBCOMPRAS_IVARETPARCTRIGO And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) <= LIBCOMPRAS_IVARETPARCFAMBPASAS Then
            If TipoIVARetenido = IVARET_PARCIAL Then
               RowRetParcial = i
'            ElseIf Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_IVARETTOT Or (Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) >= LIBCOMPRAS_IVARETTOTCHATARRA And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) <= LIBCOMPRAS_IVARETTOTCARTONES) Then
            ElseIf TipoIVARetenido = IVARET_TOTAL Then
               RowRetTotal = i
            End If
         End If
         
      ElseIf lCurTipoLib = LIB_VENTAS Then
         If DimDoc = "FCV" Then
            If TipoIVARetenido = IVARET_PARCIAL Then
               RowRetParcial = i
            ElseIf TipoIVARetenido = IVARET_TOTAL Then
               RowRetTotal = i
            End If
         End If
      End If
      
NextRec:
   Next i
   
   'If Abs(DateDiff("d", GetTxDate(Tx_FEmision), GetTxDate(Tx_FEmisionOri))) > 90 Then
   'If MsgBox1("Hay más de 90 días entre la fecha de ingreso o recepción del documento y la fecha de emisión del mismo." & vbNewLine & vbNewLine & "¿Está seguro que desea almacenar esta información?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   If lCurTipoLib = LIB_COMPRAS Then
      If Abs(DateDiff("m", GetTxDate(Tx_FEmision), GetTxDate(Tx_FEmisionOri))) > 2 And TotIVAIrrec = 0 Then
         If MsgBox1("De acuerdo a lo establecido en el art. 24 del D.L. 825 solo es posible recuperar el IVA Crédito Fiscal dentro de los dos períodos tributarios siguientes." & vbNewLine & vbNewLine & "¿Está seguro que desea almacenar esta información?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
   End If
   
   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) <> vFmt(vFmt(GridTot.TextMatrix(0, C_HABER))) Then   'no se propone valor que cuadra porque es complicado calzar
      Call MsgBox1("Los totales de las columnas DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      Exit Function
      
   End If
   
'   If vFmt(GridTot.TextMatrix(0, C_DEBE)) = 0 Then
'      Call MsgBox1("El valor total del documento es 0.", vbOKOnly + vbExclamation)
'      Exit Function
'   End If
      
   If SinClasif > 0 Then

      MsgBox1 "En algunas líneas de detalle, falta definir la clasificación del valor (Neto, Exento, Impuesto).", vbExclamation + vbOKOnly
      Exit Function

   End If
   
      
   'Validamos tipos de IVA en facturas de compra de libro de Compras o Ventas
   If DimDoc = "FCC" Or DimDoc = "FCV" Then
      
      If RowRetParcial = 0 And RowRetTotal = 0 Then
         MsgBox1 "Falta ingresar el valor para IVA Retenido Total o Parcial.", vbExclamation
         Exit Function
         
      ElseIf RowRetParcial > 0 And RowRetTotal > 0 Then
         MsgBox1 "Sólo puede ingresar un valor para IVA Retenido Total o Parcial.", vbExclamation
         Exit Function
         
      ElseIf RowIVA = 0 Then
         MsgBox1 "Falta ingresar el valor para IVA " & IVACredDeb & " Fiscal.", vbExclamation
         Exit Function
         
      ElseIf RowRetParcial > 0 Then
         If Abs(vFmt(Grid.TextMatrix(RowRetParcial, C_DEBE)) - vFmt(Grid.TextMatrix(RowRetParcial, C_HABER))) >= Abs(vFmt(Grid.TextMatrix(RowIVA, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVA, C_HABER))) Then
            MsgBox1 "El valor del IVA Retenido Parcial debe ser inferior al valor del IVA " & IVACredDeb & " Fiscal.", vbExclamation
            Exit Function
         End If
         
      Else     'RetTotal > 0
         If Abs(vFmt(Grid.TextMatrix(RowRetTotal, C_DEBE)) - vFmt(Grid.TextMatrix(RowRetTotal, C_HABER))) <> Abs(vFmt(Grid.TextMatrix(RowIVA, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVA, C_HABER))) Then
            MsgBox1 "El valor del IVA Retenido Total debe igualar el valor del IVA " & IVACredDeb & " Fiscal.", vbExclamation
            Exit Function
         End If
         
      End If
   End If
   
   'Vemos si la suma de IVA Irrec + IVA Créd. Fiscal + IVA Activo Fijo es igual al Valor Neto * 19%
   TotIVA = 0
   If lCurTipoLib = LIB_COMPRAS Then
   
      If RowIVAIrrec > 0 Then
         TotIVA = TotIVA + TotIVAIrrec
      End If
         
      If RowIVA > 0 Then
         TotIVA = TotIVA + Abs(vFmt(Grid.TextMatrix(RowIVA, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVA, C_HABER)))
      End If
         
      If RowIVAActFijo > 0 Then
         TotIVA = TotIVA + Abs(vFmt(Grid.TextMatrix(RowIVAActFijo, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVAActFijo, C_HABER)))
      End If
         
      Aux = TotAfecto * gIVA
      If Abs(TotIVA - Aux) > 2 Then
         If MsgBox1("El total de IVA difiere en más de dos unidades del " & gIVA * 100 & "% del total del Afecto." & vbCrLf & vbCrLf & "¿Está seguro que desea dejar este valor?", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
         End If
      End If
      
   End If
   
   
   
   
'   Se elimina esta validación a solicitud de Victor Morales en reporte 64, 19 Oct. 2011
'   If RowIVA > 0 And RowIVAIrrec > 0 Then
'      If Abs(vFmt(Grid.TextMatrix(RowIVAIrrec, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVAIrrec, C_HABER))) > Abs(vFmt(Grid.TextMatrix(RowIVA, C_DEBE)) - vFmt(Grid.TextMatrix(RowIVA, C_HABER))) Then
'         MsgBox1 "El valor del IVA Irrecuperable no puede superar el valor del IVA " & TipoIVA & " Fiscal.", vbExclamation
'         Exit Function
'      End If
'   End If
   
   
   'Vemos si hay valores repetidos que no pueden ser repetidos
   For j = 0 To UBound(TipoVal)
      If TipoVal(j).IdTipoValLib = 0 Then
         Exit For
      End If
      
      If TipoVal(j).Count > 1 And gTipoValLib(GetTipoValLib(lCurTipoLib, TipoVal(j).IdTipoValLib)).Multiple = 0 Then
         MsgBox "El detalle del documento incluye más de un valor de tipo " & gTipoValLib(GetTipoValLib(lCurTipoLib, TipoVal(j).IdTipoValLib)).Nombre & ".", vbExclamation + vbOKOnly
         Exit Function
      End If
   Next j
   
   'para cálculo de proporcionalidad de IVA no puede haber IVA Cred Fisc e IVA Act. Fijo en un mismo documento
   If lCurTipoLib = LIB_COMPRAS Then
      
      If CbItemData(Cb_PropIVA) <> 0 Then
         
         For j = 0 To UBound(TipoVal)
            If TipoVal(j).IdTipoValLib = 0 Then
               Exit For
            End If
            
            If TipoVal(j).Count > 1 Then
'               If TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVACREDFISC Or TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVAACTFIJO Or TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVAIRREC Then
               If TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVACREDFISC Or TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVAACTFIJO Then
                  MsgBox "El detalle del documento incluye más de un valor de tipo " & gTipoValLib(GetTipoValLib(lCurTipoLib, TipoVal(j).IdTipoValLib)).Nombre & "." & vbCrLf & vbCrLf & "No es posible calcular la Proporcionalidad del IVA en este caso.", vbExclamation + vbOKOnly
                  Exit Function
               End If
            End If
            
            If TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVACREDFISC Then
               ValLibIVA = ValLibIVA + 1
            End If
            
            If TipoVal(j).IdTipoValLib = LIBCOMPRAS_IVAACTFIJO Then
               ValLibIVA_AF = ValLibIVA_AF + 1
            End If
            
         Next j
   
         If ValLibIVA >= 1 And ValLibIVA_AF >= 1 Then
            MsgBox "El detalle del documento incluye IVA Crédito Fiscal e IVA Activo Fijo." & vbCrLf & vbCrLf & "No es posible calcular la Proporcionalidad del IVA en este caso.", vbExclamation + vbOKOnly
            Exit Function
         End If
         
      End If
               
   End If
         
   If Not ValidaDocAsoc() Then
      Exit Function
   End If
   
   valida = True
   
End Function
Private Function ValidaDocAsoc() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TipoLib As Long
   Dim TipoDocAsoc As Integer
   Dim DimDoc As String
   
   ValidaDocAsoc = False
   lIdDocAsoc = 0
   lTipoDocAsoc = 0

   If CbItemData(Cb_TipoLib) > 0 Then
      TipoLib = CbItemData(Cb_TipoLib)
   Else
      Exit Function
   End If
      
   TipoDocAsoc = CbItemData(Cb_TipoDocAsoc)
   
   'veamos si el documento asociado existe
   
   DimDoc = GetDiminutivoDoc(TipoLib, TipoDocAsoc)
   
   
   '2814014
   'If (DimDoc = TDOC_VALEPAGOELECTR Or DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_BOLVENTAEX) And Trim(Tx_NumDocAsoc) = "" Then      'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX  (FCA 30 oct 2017)
   If (DimDoc = TDOC_VALEPAGOELECTR Or DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_BOLVENTAEX Or DimDoc = TDOC_VALVENTAEX) And Trim(Tx_NumDocAsoc) = "" Then       'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX,VPEE
      lIdDocAsoc = 0
      lTipoDocAsoc = TipoDocAsoc
      lNumDocAsoc = ""
      lDTEDocAsoc = IIf(Ch_DTEDocAsoc <> 0, 1, 0)
      ValidaDocAsoc = True
      MsgBox1 "Recuerde verificar  que el documento asociado esté efectivamente registrado en el Sistema.", vbInformation
      Exit Function
   End If
      
   If (CbItemData(Cb_TipoDocAsoc) <= 0 And Trim(Tx_NumDocAsoc) <> "") Or (CbItemData(Cb_TipoDocAsoc) > 0 And Trim(Tx_NumDocAsoc) = "") Then
      MsgBox1 "Informacion de documento asociado incompleta.", vbExclamation
      Exit Function
            
   ElseIf CbItemData(Cb_TipoDocAsoc) > 0 And Trim(Tx_NumDocAsoc) <> "" Then
   
      Q1 = "SELECT IdDoc, TipoDoc, NumDoc, DTE FROM Documento "
      Q1 = Q1 & " WHERE TipoLib=" & CbItemData(Cb_TipoLib)
      Q1 = Q1 & " AND TipoDoc=" & CbItemData(Cb_TipoDocAsoc)
      Q1 = Q1 & " AND NumDoc='" & Trim(Tx_NumDocAsoc) & "'"
      
      'pipe2807009
      If gDbType = SQL_ACCESS Then
      Q1 = Q1 & " AND DTE=" & IIf(Ch_DTEDocAsoc <> 0, -1, 0)
      Else
      Q1 = Q1 & " AND DTE=" & IIf(Ch_DTEDocAsoc <> 0, 1, 0)
      End If
      'fin 2807009
      
      Q1 = Q1 & " AND IdEntidad =" & lcbNombre.ItemData
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF Then   ' no existe
      '2814014
         'If (DimDoc = TDOC_VALEPAGOELECTR Or DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_BOLVENTAEX) Then       'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX  (FCA 30 oct 2017)
         If (DimDoc = TDOC_VALEPAGOELECTR Or DimDoc = TDOC_BOLVENTA Or DimDoc = TDOC_BOLEXENTA Or DimDoc = TDOC_BOLVENTAEX Or DimDoc = TDOC_VALVENTAEX) Then       'de acuerdo a lo solicitado por TR (Nicolás Cartrin) no se debe validar número de VPE, BOV, BOE, BEX , VPEE
            lIdDocAsoc = 0
            lTipoDocAsoc = TipoDocAsoc
            lNumDocAsoc = Trim(Tx_NumDocAsoc)
            lDTEDocAsoc = IIf(Ch_DTEDocAsoc <> 0, 1, 0)
           MsgBox1 "Recuerde verificar  que el documento asociado esté efectivamente registrado en el Sistema.", vbInformation
         Else
            MsgBox1 "El documento asociado no ha sido ingresado al sistema. Verifique los datos y la entidad asociada.", vbExclamation + vbOKOnly
            Call CloseRs(Rs)
            Exit Function
         End If
                  
      Else
         lIdDocAsoc = vFld(Rs("IdDoc"))
         lTipoDocAsoc = vFld(Rs("TipoDoc"))
         lNumDocAsoc = vFld(Rs("NumDoc"))
         lDTEDocAsoc = vFld(Rs("DTE"))
      End If
      
      Call CloseRs(Rs)

   End If

   ValidaDocAsoc = True
   
End Function
Private Function Valida_Old() As Boolean
   Dim i As Integer
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim RowCuadra As Integer
   Dim cont As Integer
   Dim Diff As Double
   Dim SinClasif As Integer
   
   Valida_Old = False
  
   If Cb_TipoDoc.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un tipo de documento.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If

   If Tx_Rut = "" Then
      MsgBox1 "Debe ingresar una entidad.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If
   
   If Trim(Tx_NumDoc) = "" Then
      MsgBox1 "Debe ingresar un número de documento.", vbExclamation
      Tx_NumDoc.SetFocus
      Exit Function
   End If
      
   If Trim(Tx_FEmision) = "" Then
      MsgBox1 "Debe ingresar fecha de emisión.", vbExclamation
      Tx_FEmision.SetFocus
      Exit Function
   End If
   
   'If Trim(Tx_FVenc) = "" Then
   '   MsgBox1 "Debe ingresar fecha de vecimiento, ya que usted seleccionó un tipo de documento.", vbExclamation
   '   Tx_FVenc.SetFocus
   '   Exit Function
   'End If
   
   If GetTxDate(Tx_FVenc) > 0 And GetTxDate(Tx_FEmision) > GetTxDate(Tx_FVenc) Then
      MsgBox1 "Fecha de emisión mayor a la fecha de vencimiento.", vbExclamation
      Tx_FVenc.SetFocus
      Exit Function
   End If
      
   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) <> vFmt(vFmt(GridTot.TextMatrix(0, C_HABER))) Then
      Call MsgBox1("Los totales de las columnas DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      Exit Function
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      
      If vFmt(Grid.TextMatrix(i, C_DEBE)) <> 0 And vFmt(Grid.TextMatrix(i, C_HABER)) <> 0 Then
         MsgBox1 "En la línea " & i & " los valores en las columnas DEBE y HABER son ambos mayores que 0.", vbExclamation + vbOKOnly
         Exit Function
         
      End If
      
      If Trim(Grid.TextMatrix(i, C_TIPOVALLIB)) = "" Then
         'MsgBox1 "En la línea " & i & " falta definir la clasificación del documento (Neto, Impuesto, Bruto).", vbExclamation + vbOKOnly
         'Exit Function
         SinClasif = SinClasif + 1
     
      End If
      
      If vFmt(Grid.TextMatrix(i, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(i, C_HABER)) = 0 Then
         RowCuadra = i
         cont = cont + 1
         
      End If
      
      SumDebe = SumDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
      SumHaber = SumHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      
   Next i
   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) <> vFmt(vFmt(GridTot.TextMatrix(0, C_HABER))) Then
      Call MsgBox1("Los totales de las columnas DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      
      'Propongo valor que cuadra, si debe y haber es =0 y es solo una fila que no tiene debe ni haber
      Diff = SumDebe - SumHaber
      If cont = 1 Then
         If Diff < 0 Then
            MsgBox1 "Para cuadrar el documento se sugiere valor en el DEBE", vbExclamation
            Grid.TextMatrix(RowCuadra, C_DEBE) = Format(Abs(Diff), BL_NUMFMT)
            
         Else
            MsgBox1 "Para cuadrar el documento se sugiere valor en el HABER", vbExclamation
            Grid.TextMatrix(RowCuadra, C_HABER) = Format(Diff, BL_NUMFMT)
            
         End If
      End If
      Call CalcTot
      Exit Function
      
   End If
   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) = 0 Then
      Call MsgBox1("El valor total del documento es 0.", vbOKOnly + vbExclamation)
      Exit Function
   End If
      
   If SinClasif > 0 Then
      If MsgBox1("¡Atención!" & vbNewLine & vbNewLine & "En algunas líneas de detalle, falta definir la clasificación del valor (Neto, Exento, Impuesto)." & vbNewLine & "¿Desea continuar sin actualizar esta información ?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
         Exit Function
      End If
   End If

      
   Valida_Old = True
   
End Function
Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_Rut_LostFocus
      KeyAscii = 0
   ElseIf Ch_Rut <> 0 Then
      Call KeyCID(KeyAscii)
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
   
   
End Sub

Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

Private Function CalcValor(ByVal TipoLib As Integer, ByVal NombreCampo As String, ByVal EsRebaja As Boolean, ByVal Debe As Double, ByVal Haber As Double) As Double
    
   If NombreCampo = "OtrosVal" Or NombreCampo = "OtroImp" Then
      If TipoLib = LIB_VENTAS Then
         If EsRebaja Then
            CalcValor = Debe - Haber
         Else
            CalcValor = Haber - Debe
         End If
      ElseIf TipoLib = LIB_COMPRAS Then
         If EsRebaja Then
            CalcValor = Haber - Debe
         Else
            CalcValor = Debe - Haber
         End If
      ElseIf TipoLib = LIB_RETEN Then
         If EsRebaja Then
            CalcValor = Debe - Haber
         Else
            CalcValor = Haber - Debe
         End If
      End If
      
   Else
      If TipoLib = LIB_OTROS Then
         CalcValor = Debe - Haber
      Else
         CalcValor = Abs(Debe - Haber)
      End If
   End If

End Function



Private Function GetNombreCampo(ByVal IdTipoValLib As Integer) As String
   Dim NombreCampo As String

   If ItemData(Cb_TipoLib) = LIB_COMPRAS Then
    
      Select Case IdTipoValLib
       
         Case LIBCOMPRAS_AFECTO
            NombreCampo = "Afecto"
         
         Case LIBCOMPRAS_EXENTO
            NombreCampo = "Exento"
      
         Case LIBCOMPRAS_IVACREDFISC, LIBCOMPRAS_IVAACTFIJO, LIBCOMPRAS_IVAIRREC, LIBCOMPRAS_IVAIRREC1, LIBCOMPRAS_IVAIRREC2, LIBCOMPRAS_IVAIRREC3, LIBCOMPRAS_IVAIRREC4, LIBCOMPRAS_IVAIRREC9
            NombreCampo = "IVA"
   
         Case LIBCOMPRAS_OTROSIMP
            NombreCampo = "OtroImp"
            
          Case LIBCOMPRAS_TOTAL
            NombreCampo = "Total"
            
        Case Else
            NombreCampo = "OtrosVal"
         
      End Select
       
   ElseIf ItemData(Cb_TipoLib) = LIB_VENTAS Then
   
      Select Case IdTipoValLib
       
         Case LIBVENTAS_AFECTO
            NombreCampo = "Afecto"
         
         Case LIBVENTAS_EXENTO
            NombreCampo = "Exento"
      
         Case LIBVENTAS_IVADEBFISC
            NombreCampo = "IVA"
   
         Case LIBVENTAS_OTROSIMP
            NombreCampo = "OtroImp"
   
         Case LIBVENTAS_TOTAL
            NombreCampo = "Total"
            
         Case Else
            NombreCampo = "OtrosVal"
         
      End Select
       
   ElseIf ItemData(Cb_TipoLib) = LIB_RETEN Then
   
      Select Case IdTipoValLib
   
         Case LIBRETEN_BRUTO
            NombreCampo = "Afecto"
         
         Case LIBRETEN_HONORSINRET
            NombreCampo = "Exento"
      
         Case LIBRETEN_IMPUESTO
            NombreCampo = "OtroImp"
            
         Case LIBRETEN_NETO
            NombreCampo = "Total"
            
         Case Else
            NombreCampo = "OtrosVal"
            
      End Select
            
   ElseIf ItemData(Cb_TipoLib) = LIB_REMU Then
         
      Select Case IdTipoValLib
         
         Case LIBREMU_VALOR
            NombreCampo = "Exento"
         
         Case LIBREMU_TOTAL
            NombreCampo = "Total"
    
      End Select
      
   ElseIf ItemData(Cb_TipoLib) = LIB_OTROS Then
         
      Select Case IdTipoValLib
         
         Case LIBOTROS_VALOR
            NombreCampo = "Exento"
         
         Case LIBOTROS_TOTAL
            NombreCampo = "Total"
    
      End Select
      
   End If
   
   GetNombreCampo = NombreCampo
   
End Function

Private Sub Bt_ActivoFijo_Click()
   Dim Frm As FrmLstActFijo
   Dim i As Integer
   Dim CtaActFijo As Boolean
   Dim vAreaN As Long, vCentroG As Long
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_IDCUENTA)) = 0 Then
         Exit For
      End If
   
      If EsCuentaActFijo(Val(Grid.TextMatrix(i, C_IDCUENTA))) Then
         CtaActFijo = True
        vAreaN = Val(Grid.TextMatrix(i, C_IDAREANEG))
        vCentroG = Val(Grid.TextMatrix(i, C_IDCCOSTO))
      End If
      
   Next i
   
   If Not CtaActFijo Then
      If MsgBox1("No hay cuentas de Activo Fijo en este documento." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         Exit Sub
      End If
   End If
   
   Set Frm = New FrmLstActFijo
   If lOper <> O_EDIT Then
      Call Frm.FViewFromDoc(lIdDoc, CbItemData(Cb_TipoLib))
   Else
      '2861733
      'Call Frm.FEditFromDoc(lIdDoc, CbItemData(Cb_TipoLib), GetTxDate(Tx_FEmision))
      Call Frm.FEditFromDocActiFijo(lIdDoc, CbItemData(Cb_TipoLib), GetTxDate(Tx_FEmision), vAreaN, vCentroG)
      '2861733
   End If
   
   Call LoadGrMov

   
   Set Frm = Nothing
   
End Sub
Private Sub FillCb()
   Dim i As Integer
   Dim Q1 As String

   Cb_Entidad.AddItem ""
   Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = -1
   Call FillCbClasifEnt(Cb_Entidad)
   Cb_Entidad.ListIndex = -1     'para no seleccionar ninguno al partir
   
   'Cb_TipoLib.AddItem ""
   'Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = 0
   For i = 1 To UBound(gTipoLibNew)
      Cb_TipoLib.AddItem ReplaceStr(gTipoLibNew(i).Nombre, "Libro de ", "")
      Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).id 'i
   Next i
   Cb_TipoLib.ListIndex = -1
   
   For i = 1 To UBound(gEstadoDoc)
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   Call SelItem(Cb_Estado, ED_PENDIENTE)
      
   Call AddItem(Grid.CbList(C_CCOSTO), " ", 0)
   Q1 = "SELECT Descripcion,idCCosto FROM CentroCosto WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call FillCombo(Grid.CbList(C_CCOSTO), DbMain, Q1, -1)
   
   Call AddItem(Grid.CbList(C_AREANEG), " ", 0)
   Q1 = "SELECT Descripcion, idAreaNegocio FROM AreaNegocio WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call FillCombo(Grid.CbList(C_AREANEG), DbMain, Q1, -1)

   Cb_Impto.Clear
   Call AddItem(Cb_Impto, gImpRet(IMPRET_NAC) * 100 & "%", IMPRET_NAC)
   Call AddItem(Cb_Impto, gImpRet(IMPRET_EXT) * 100 & "%", IMPRET_EXT)
   Call AddItem(Cb_Impto, "Otro", IMPRET_OTRO)
   
   Cb_TipoReten.Clear
   Call AddItem(Cb_TipoReten, "Honorarios", TR_HONORARIOS)
   Call AddItem(Cb_TipoReten, "Dieta", TR_DIETA)
   Call AddItem(Cb_TipoReten, "Otro", TR_OTRO)
      
   Call AddItem(Cb_Sucursal, " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)
   
   Cb_PropIVA.Clear
   Call CbAddItem(Cb_PropIVA, gStrPropIVA(PIVA_SINPROP), PIVA_SINPROP)
   Call CbAddItem(Cb_PropIVA, gStrPropIVA(PIVA_TOTAL), PIVA_TOTAL)
   Call CbAddItem(Cb_PropIVA, gStrPropIVA(PIVA_NULO), PIVA_NULO)
   Call CbAddItem(Cb_PropIVA, gStrPropIVA(PIVA_PROP), PIVA_PROP)
   Cb_PropIVA.ListIndex = 0
  
End Sub

Private Function GetDocAsoc(ByVal IdDoc As Long, Optional ByVal Msg As Boolean = True) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   If IdDoc <= 0 Then
      Exit Function
   End If
   
   Cb_TipoDocAsoc.ListIndex = -1
   Tx_NumDocAsoc = ""
   Ch_DTEDocAsoc = 0
   lEstadoDocAsocPagado = False
   
   Q1 = "SELECT TipoDoc, NumDoc, Estado, DTE FROM Documento WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If Cb_TipoDocAsoc.ListCount > 0 Then
         For i = 0 To Cb_TipoDocAsoc.ListCount - 1
            If vFld(Rs("TipoDoc")) = Cb_TipoDocAsoc.ItemData(i) Then
               
               If vFld(Rs("Estado")) = ED_PAGADO Then
                  lEstadoDocAsocPagado = True
                  Cb_TipoDocAsoc.Locked = True
                  Call SetTxRO(Tx_NumDocAsoc, True)
                  Ch_DTEDocAsoc.Enabled = False
               End If
               
               Cb_TipoDocAsoc.ListIndex = i
               Tx_NumDocAsoc = vFld(Rs("NumDoc"))
               Ch_DTEDocAsoc = IIf(vFld(Rs("DTE")) <> 0, 1, 0)
               
               Exit For
            End If
         Next i
      ElseIf Msg Then
         MsgBox1 "Documento asociado inválido.", vbExclamation
      End If
      
   ElseIf Msg Then
      MsgBox1 "Documento asociado inválido.", vbExclamation
   End If
   
   Call CloseRs(Rs)

End Function
Friend Sub AppendImpAdic(ImpAdic() As ImpAdic_t, ByVal Glosa As String)
   Dim i As Integer
   Dim FirstEmptyRow As Integer, Row As Integer
   Dim n As Integer, RowOtrosImp As Integer
   Dim EsRebaja As Boolean, TipoDoc As Integer, Diminutivo As String, EsIVARetenido As Boolean
   Dim Tasa As Single
   Dim IdTipoValLib As Integer
   Dim TipoIVARetenido As Integer
   Dim Idx As Integer
    
    
   EsIVARetenido = 0
  
   If ImpAdic(0).CodTipoValor = 0 Then
      Exit Sub
   End If
   
   Row = Grid.Row
   TipoDoc = CbItemData(Cb_TipoDoc)
   EsRebaja = gTipoDoc(GetTipoDoc(lCurTipoLib, TipoDoc)).EsRebaja
   Diminutivo = gTipoDoc(GetTipoDoc(lCurTipoLib, TipoDoc)).Diminutivo
   
   RowOtrosImp = 0
   
   'buscamos la primera fila vacía
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDCUENTA)) = 0 Then
         FirstEmptyRow = i
         Exit For
      End If
      If (lCurTipoLib = LIB_COMPRAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBCOMPRAS_OTROSIMP) Or (lCurTipoLib = LIB_VENTAS And Val(Grid.TextMatrix(i, C_IDTIPOVALLIB)) = LIBVENTAS_OTROSIMP) Then
         RowOtrosImp = i
      End If
   Next i
   
   Row = FirstEmptyRow
   n = 0
   
   'agregamos cada impuesto adicional
   For i = 0 To UBound(ImpAdic) - 1
   
      If ImpAdic(i).CodTipoValor = 0 Then
         Exit For
      End If
     
      IdTipoValLib = ImpAdic(i).CodTipoValor
      Idx = FindTipoValLib(lCurTipoLib, IdTipoValLib)
      TipoIVARetenido = gTipoValLib(Idx).TipoIVARetenido
                 
      If lCurTipoLib = LIB_COMPRAS Then
         EsIVARetenido = (ImpAdic(i).CodTipoValor = LIBCOMPRAS_IVARETTOT)
         EsIVARetenido = EsIVARetenido Or (ImpAdic(i).CodTipoValor >= LIBCOMPRAS_IVARETPARCTRIGO And ImpAdic(i).CodTipoValor <= LIBCOMPRAS_IVARETPARCFAMBPASAS)
         EsIVARetenido = EsIVARetenido Or (ImpAdic(i).CodTipoValor >= LIBCOMPRAS_IVARETTOTCHATARRA And ImpAdic(i).CodTipoValor <= LIBCOMPRAS_IVARETTOTCARTONES)
         EsIVARetenido = EsIVARetenido Or (ImpAdic(i).CodTipoValor = LIBCOMPRAS_IVARETORO)
      
      ElseIf lCurTipoLib = LIB_VENTAS Then
         EsIVARetenido = IIf(TipoIVARetenido <> 0, True, False)
      End If
      
      Grid.TextMatrix(Row, C_IDCUENTA) = ImpAdic(i).IdCuenta
      Grid.TextMatrix(Row, C_CODCUENTA) = ImpAdic(i).CodCuenta
      Grid.TextMatrix(Row, C_CUENTA) = ImpAdic(i).Cuenta
      
      If ImpAdic(i).Tasa = 100 Then
         Tasa = gIVA * 100
      Else
         Tasa = ImpAdic(i).Tasa
      End If
      
      If lCurTipoLib = LIB_COMPRAS Then
         If EsRebaja Then
'            If Diminutivo = "NCF" And EsIVARetenido Then
            If (Diminutivo = "NCC" Or Diminutivo = "NCF") And EsIVARetenido Then   'Cambio solicitado por Nicolás Catrin el 21 ago 2018
               Grid.TextMatrix(Row, C_DEBE) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_HABER) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            End If
         Else
            If (Diminutivo = "FCC" Or Diminutivo = "NDF") And EsIVARetenido Then
               Grid.TextMatrix(Row, C_HABER) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_DEBE) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            End If
         End If
         
      ElseIf lCurTipoLib = LIB_VENTAS Then
         If EsRebaja Then
            If (Diminutivo = "NCV") And EsIVARetenido Then    'Cambio solicitado por Nicolás Catrin el 21 ago 2018
               Grid.TextMatrix(Row, C_HABER) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_DEBE) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            End If
         Else
            If (Diminutivo = "FCV") And EsIVARetenido Then
               Grid.TextMatrix(Row, C_DEBE) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_HABER) = Format(lNetoDoc * Tasa / 100, NUMFMT)
            End If
         End If
      End If
      
      Grid.TextMatrix(Row, C_TIPOVALLIB) = ImpAdic(i).TipoValor
      Grid.TextMatrix(Row, C_IDTIPOVALLIB) = ImpAdic(i).CodTipoValor
      Grid.TextMatrix(Row, C_TASA) = Format(ImpAdic(i).Tasa, DBLFMT2)
      Grid.TextMatrix(Row, C_ESRECUPERABLE) = FmtSiNo(Abs(ImpAdic(i).EsRecuperable))
      Grid.TextMatrix(Row, C_CODSIIDTE) = gTipoValLib(GetTipoValLib(lCurTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPOVALLIB)))).CodSIIDTE
      Grid.TextMatrix(Row, C_GLOSA) = Glosa
      Grid.TextMatrix(Row, C_ORDEN) = Row
      Call FGrSetPicture(Grid, Row, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
      
      If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_AREANEG) = 0 Then
         Grid.TextMatrix(Grid.Row, C_IDAREANEG) = 0
         Grid.TextMatrix(Grid.Row, C_AREANEG) = ""
      End If
         
      If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CCOSTO) = 0 Then
         Grid.TextMatrix(Grid.Row, C_IDCCOSTO) = 0
         Grid.TextMatrix(Grid.Row, C_CCOSTO) = ""
      End If
         
      If Row = Grid.rows - 1 Then
         Grid.rows = Grid.rows + 1
      End If

      Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)

      Row = Row + 1
      
      n = n + 1
      
   Next i
   
   If n > 0 Then
      If RowOtrosImp <> 0 Then
         If MsgBox1("¿Desea eliminar el movimiento asociado a la Clasificación 'Otros Impuestos'?", vbYesNo + vbQuestion) = vbYes Then
            Grid.Row = RowOtrosImp
'            Grid.RowSel = RowOtrosImp
'            Grid.Col = C_DEBE
'            Grid.ColSel = C_DEBE
            Call Bt_Del_Click
         End If
      End If
   End If
   
   Call CalcTot
   
End Sub


Private Sub Tx_ValTotal_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)

End Sub

Private Sub Tx_ValTotal_LostFocus()
   Tx_ValTotal = Format(vFmt(Tx_ValTotal), NUMFMT)
End Sub


'2861733 tema 2

Private Function ValidarCuentaAFijo(ByVal vIdCuenta As Long, ByRef vTieneArea As Boolean, ByRef vTieneCentro As Boolean) As Boolean
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim Rs2 As Recordset

 ValidarCuentaAFijo = False
 vTieneArea = False
 vTieneCentro = False
 
    Q1 = ""
    Q1 = " SELECT count(*) as existe  from cuentas "
    Q1 = Q1 & " where idcuenta  = " & vIdCuenta
    Q1 = Q1 & " and atrib3 = 1 "
    
    
   Set Rs = OpenRs(DbMain, Q1)
    If Not Rs.EOF Then
        If vFld(Rs("existe")) = 1 Then
         ValidarCuentaAFijo = True
        End If
    End If
    
   Call CloseRs(Rs)
   
   If ValidarCuentaAFijo = True Then
   
    Q1 = ""
    Q1 = " SELECT atrib5,atrib6 from cuentas "
    Q1 = Q1 & " where idcuenta  = " & vIdCuenta
    Q1 = Q1 & " and atrib3 = 1 "
    
    
   Set Rs = OpenRs(DbMain, Q1)
    If Not Rs.EOF Then
        If vFld(Rs("atrib6")) = 1 Then
         vTieneArea = True
        End If
        
        If vFld(Rs("atrib5")) = 1 Then
         vTieneCentro = True
        End If
    End If
    
   Call CloseRs(Rs)
  End If
End Function

Private Function ObtenerAreaCentro(ByVal vIdDoc As Long, ByVal vIdCuenta As Long, ByRef vIdArea As Long, ByRef vIdCentro As Long, ByRef vDescArea As String, ByRef vDescCentro As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Double
    
    ObtenerAreaCentro = False
    
    vIdArea = 0
    vIdCentro = 0
    i = 0
    
   Q1 = ""
   Q1 = " SELECT IdActFijo,IdCuenta,movactivofijo.idCCosto,movactivofijo.IdAreaNeg,CentroCosto.Descripcion as DescCCosto,AreaNegocio.Descripcion as DescAreaNeg from (movactivofijo "
   Q1 = Q1 & " LEFT JOIN CentroCosto ON movactivofijo.IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & JoinEmpAno(gDbType, "movactivofijo", "CentroCosto", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN AreaNegocio ON movactivofijo.IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & JoinEmpAno(gDbType, "movactivofijo", "AreaNegocio", True, True)
   Q1 = Q1 & " Where iddoc = " & vIdDoc
   Q1 = Q1 & " And IdCuenta = " & vIdCuenta
   Q1 = Q1 & " And movactivofijo.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " And Ano = " & gEmpresa.Ano
   Q1 = Q1 & " Order by IdActFijo asc"
   
   Set Rs = OpenRs(DbMain, Q1)
    Do While Rs.EOF = False
    i = i + 1
    
        If i = 1 Then
        
                If vFld(Rs("IdAreaNeg")) > 0 Then
                 vIdArea = vFld(Rs("IdAreaNeg"))
                 vDescArea = vFld(Rs("DescAreaNeg"))
                End If
                
                If vFld(Rs("idCCosto")) > 0 Then
                 vIdCentro = vFld(Rs("idCCosto"))
                 vDescCentro = vFld(Rs("DescCCosto"))
                End If
                
           ObtenerAreaCentro = True
        
        End If
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

End Function

Private Sub SaveAreaCentroActFijo(ByVal vIdCuenta As Long, ByVal vIdArea As Long, ByVal vIdCentro As Long)
   Dim Q1 As String
   
   Q1 = "UPDATE movactivofijo SET IdAreaNeg = " & vIdArea & ", idCCosto = " & vIdCentro
   Q1 = Q1 & " WHERE IdDoc =" & lIdDoc
   Q1 = Q1 & " And IdCuenta = " & vIdCuenta & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
         
End Sub

'2861733 tema 2

