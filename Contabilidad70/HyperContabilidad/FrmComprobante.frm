VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmComprobante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Comprobante"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "FrmComprobante.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   12225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr_Paginacion 
      Caption         =   "Paginacion"
      Height          =   735
      Left            =   120
      TabIndex        =   66
      Top             =   7680
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton Bt_ToLeft 
         Height          =   375
         Left            =   120
         Picture         =   "FrmComprobante.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Anterior conjunto de registros"
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton Bt_ToRight 
         Height          =   375
         Left            =   540
         Picture         =   "FrmComprobante.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Siguente conjunto de registros"
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.CommandButton Bt_PrtCheque 
      Caption         =   "Imprimir Cheque..."
      Height          =   315
      Left            =   9960
      TabIndex        =   64
      Top             =   1980
      Width           =   1965
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   4275
      Left            =   60
      TabIndex        =   11
      Top             =   2520
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7541
      Cols            =   24
      Rows            =   23
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
   Begin VB.Frame Frame 
      Height          =   615
      Left            =   60
      TabIndex        =   36
      Top             =   0
      Width           =   11895
      Begin VB.Frame Fr_Doc 
         BorderStyle     =   0  'None
         Height          =   470
         Left            =   3660
         TabIndex        =   53
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton Bt_BuscarDoc 
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
            Left            =   420
            Picture         =   "FrmComprobante.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Buscar un documento para asociarlo al movimiento seleccionado"
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton Bt_DelDoc 
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
            Left            =   840
            Picture         =   "FrmComprobante.frx":0A5E
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Eliminar enlace con documento definido en movimiento seleccionado"
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton Bt_GenPago 
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
            Left            =   1260
            Picture         =   "FrmComprobante.frx":0EE2
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Seleccionar documentos de compra, venta o retención para generar automáticamente comprobante de pago "
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton Bt_NewDoc 
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
            Left            =   0
            Picture         =   "FrmComprobante.frx":12DB
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Crear un nuevo documento para asociarlo al movimiento seleccionado"
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton Bt_DetMov 
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
            Left            =   600
            Picture         =   "FrmComprobante.frx":1746
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Detalle del documento asociado al movimiento seleccionado"
            Top             =   -120
            Visible         =   0   'False
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
            Left            =   1680
            Picture         =   "FrmComprobante.frx":1A50
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Ver/agregar detalle de Activo Fijo asociado al documento seleccionado"
            Top             =   60
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.CommandButton Bt_Salir 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10920
         TabIndex        =   35
         Top             =   180
         Width           =   885
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
         Left            =   7200
         Picture         =   "FrmComprobante.frx":1E4E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Copiar Excel"
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
         Left            =   3180
         Picture         =   "FrmComprobante.frx":2293
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cuadrar comprobante"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Preview 
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
         Left            =   6360
         Picture         =   "FrmComprobante.frx":2618
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Print 
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
         Left            =   6780
         Picture         =   "FrmComprobante.frx":2ABF
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprimir"
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
         Left            =   1860
         Picture         =   "FrmComprobante.frx":2F79
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Copiar dato"
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
         Left            =   2280
         Picture         =   "FrmComprobante.frx":3359
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Pegar dato copiado"
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
         Left            =   1380
         Picture         =   "FrmComprobante.frx":3742
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Mover hacia abajo"
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
         Left            =   960
         Picture         =   "FrmComprobante.frx":37C3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Mover hacia arriba"
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
         Picture         =   "FrmComprobante.frx":3844
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar movimiento seleccionado"
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
         Picture         =   "FrmComprobante.frx":3C40
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Duplicar movimiento seleccionado"
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
         Left            =   5880
         Picture         =   "FrmComprobante.frx":4092
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Plan de Cuentas"
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
         Left            =   8100
         Picture         =   "FrmComprobante.frx":4453
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Calculadora"
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
         Left            =   7680
         Picture         =   "FrmComprobante.frx":47B4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Convertir moneda"
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
         Left            =   2760
         Picture         =   "FrmComprobante.frx":4B52
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   9000
         TabIndex        =   33
         Top             =   180
         Width           =   885
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
         Left            =   8520
         Picture         =   "FrmComprobante.frx":4BF6
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   9960
         TabIndex        =   34
         Top             =   180
         Width           =   885
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6780
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Frame Fr_SelCompTipo 
      Caption         =   "Comprobante Tipo"
      Height          =   1155
      Left            =   9960
      TabIndex        =   50
      Top             =   660
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton Bt_CompTipo 
         Height          =   315
         Left            =   1560
         Picture         =   "FrmComprobante.frx":501F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Mostrar lista de  Comprobantes Tipo para seleccionar uno"
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton Bt_NewCompTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         Picture         =   "FrmComprobante.frx":5329
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Crear Comprobante Tipo a partir de este Comprobante"
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton Bt_SelCompTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         Picture         =   "FrmComprobante.frx":5789
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Buscar Comprobante Tipo por nombre corto"
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox Tx_NombCompTipo 
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Fr_Header 
      Caption         =   "Encabezado Comprobante"
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Index           =   0
      Left            =   60
      TabIndex        =   37
      Top             =   660
      Width           =   9795
      Begin VB.CheckBox Ch_OtrosIngEg14TER 
         Caption         =   "Otros Ingresos/Egresos 14TER"
         Height          =   195
         Left            =   3480
         TabIndex        =   60
         ToolTipText     =   "Ingresos distintos a Compras, Ventas y Retenciones"
         Top             =   1200
         Width           =   2595
      End
      Begin VB.CheckBox Ch_ImpRes 
         Caption         =   "Imprimir en forma Resumida"
         Height          =   195
         Left            =   840
         TabIndex        =   59
         ToolTipText     =   "Imprimir en forma resumida desde listado de comprobantes y libro diario"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Tx_Usuario 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_Correlativo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton Bt_Glosas 
         Height          =   315
         Index           =   0
         Left            =   9360
         Picture         =   "FrmComprobante.frx":5BC7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Left            =   3600
         Picture         =   "FrmComprobante.frx":5ED1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox Tx_Glosa 
         Height          =   315
         Left            =   840
         MaxLength       =   100
         TabIndex        =   4
         Top             =   720
         Width           =   8535
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   2460
         TabIndex        =   0
         Top             =   300
         Width           =   1155
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Tx_IdComp 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
      End
      Begin VB.Frame Fr_TAjuste 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   6420
         TabIndex        =   65
         Top             =   1080
         Width           =   3315
         Begin VB.OptionButton Op_TAjuste 
            Caption         =   "Financiero"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   61
            Top             =   120
            Width           =   1035
         End
         Begin VB.OptionButton Op_TAjuste 
            Caption         =   "Tributario"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   62
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton Op_TAjuste 
            Caption         =   "Ambos"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   63
            Top             =   120
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usr:"
         Height          =   195
         Index           =   5
         Left            =   8160
         TabIndex        =   52
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   780
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   42
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   40
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° comp.:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Fr_Header 
      Caption         =   "Encabezado Comprobante Tipo"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   44
      Top             =   660
      Width           =   9795
      Begin VB.TextBox Tx_Descrip 
         Height          =   315
         Left            =   3600
         MaxLength       =   40
         TabIndex        =   7
         Top             =   300
         Width           =   3555
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   8340
         MaxLength       =   15
         TabIndex        =   8
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Glosas 
         Height          =   315
         Index           =   1
         Left            =   9360
         Picture         =   "FrmComprobante.frx":61DB
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   780
         Width           =   255
      End
      Begin VB.ComboBox Cb_TipoCompTipo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox Tx_GlosaCompTipo 
         Height          =   315
         Left            =   840
         MaxLength       =   100
         TabIndex        =   9
         Top             =   780
         Width           =   8535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrip.:"
         Height          =   195
         Index           =   9
         Left            =   2880
         TabIndex        =   58
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   8
         Left            =   7680
         TabIndex        =   47
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTotFull 
      Height          =   315
      Left            =   60
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Menu M_Opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu M_SelCuenta 
         Caption         =   "Seleccionar Cuenta..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu M_Copy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu M_Paste 
         Caption         =   "Pegar"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu M_ViewNote 
         Caption         =   "Ver Nota..."
      End
      Begin VB.Menu M_AddNote 
         Caption         =   "Agregar Nota..."
      End
      Begin VB.Menu M_EditNote 
         Caption         =   "Editar Nota..."
      End
      Begin VB.Menu M_DelNote 
         Caption         =   "Eliminar Nota..."
      End
   End
End
Attribute VB_Name = "FrmComprobante"
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
Const C_GLOSA = 8
Const C_TIPODOC = 9
Const C_NUMDOC = 10
Const C_ENTIDAD = 11
Const C_DETALLE = 12
Const C_AREANEG = 13
Const C_IDAREANEG = 14
Const C_CCOSTO = 15
Const C_IDCCOSTO = 16
Const C_DETACTFIJO = 17
Const C_IDDOC = 18
Const C_IDDOCCUOTA = 19
Const C_DECENTRALIZ = 20
Const C_DEPAGO = 21
Const C_ATRIB_CONCIL = 22
Const C_NOTA = 23
Const C_UPDATE = 24

Const NCOLS = C_UPDATE

Const NEW_DEPAGO = 2
Const NEW_DESELDOC = 3

Const TX_ACTFIJO = "AF >>"

Const COMP_FOOTER = 2000
Const COMP_VBOX = 1000

Dim lRc As Integer
Dim lOper As Integer
Dim lidComp As Long
Dim lCorrelativo As Long

Dim lEstado As Integer
Dim lTblComprobante As String
Dim lTblMovComprobante As String
Dim lTblDocumento As String
Dim lTblDocumentoFull As String
Dim lCompTipo As Boolean
Dim lFromCentraliz As Boolean

Dim lTieneMovCentraliz As Boolean
Dim lTieneMovPago As Boolean           'indica que tiene doc de Pago pero no es nuevo

Dim lGenPago As Boolean                'Indica que se generó un pago de documentos en este comprobante

Dim lTipoLib As Integer

Dim lPregIngCompMesCerrado As Boolean

Dim lGrDobleClick As Boolean

Dim lViewResumido As Boolean
Dim lPrtResumido As Boolean

Dim lDelDocLst As String      'lista de documentos que fueron eliminados de este
                              'comprobante, para setear SaldoDoc en NULL y obligar
                              'que se recalcule el Saldo
                              
Dim lEnImpresionMasiva As Boolean

Dim lOldFechaEmision As Long         'almacena la fecha de emisión del comprobante al entrar a la edición (si es new, vale 0)

Dim lHayDocLibros As Boolean         'indica si hay pagos de documentos de compras, ventas o retenciones

Dim lCuentasDisponible As Boolean    'indica si se usan cuentas de ajuste extra contable disponible

Dim lOtrosIngEg14TER As Boolean      'Indica el valor de este campo en la base de datos (se carga en el LoadAll y se asigna a la check en el FormActivate

Dim lMesAnoCentraliz As Long           'Indica el año-mes que se está centralizando en formato yyyymm (201706)

'14523812
Dim EsCompTipo As Boolean
'14523812

'3238739
Dim lClsPaging As ClsPaging   'Clase de paginamiento
Dim lCentrFull As Integer

Dim ltotalDebeFull As Double
Dim ltotalHaberFull As Double

'3238739

Private Sub Bt_ActivoFijo_Click()
   Dim Frm As FrmLstActFijo
   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lCompTipo Then
      Exit Sub
   End If
   
   If EsCuentaActFijo(Val(Grid.TextMatrix(Grid.Row, C_IDCUENTA))) Then
      
      If (lOper = O_EDIT Or lOper = O_NEW) And Grid.Locked = False Then
         Call ActivoFijo(Grid.Row, Grid.Col)
      
      Else
         Set Frm = New FrmLstActFijo
         Call FrmLstActFijo.FViewFromComp(lidComp, Grid.TextMatrix(Grid.Row, C_IDMOV))
      
      End If
   Else
      MsgBox1 "La cuenta del movimiento seleccionado no es de Activo Fijo.", vbExclamation
      
   End If

End Sub
Private Sub Bt_GenPago_Click()
   Dim LstIdDoc() As LstDoc_t
   Dim Frm As FrmLstDoc
   Dim Rc As Integer
   Dim TipoLib As Integer
      
   Set Frm = New FrmLstDoc
   
   If InStr(LCase(Tx_Glosa), "ventas") Then
      TipoLib = LIB_VENTAS
   ElseIf InStr(LCase(Tx_Glosa), "compras") Then
      TipoLib = LIB_COMPRAS
   End If
      
   Rc = Frm.FSelect(LstIdDoc, TipoLib, ED_CENTPAG, True, False)
   Set Frm = Nothing
      
   If Rc = vbOK Then
      
      If LstIdDoc(0).IdDoc <> 0 Then  'eligió al menos un doc
         Call AsignarLstDoc(LstIdDoc)
                     
      Else                       'eligió el botón Nuevo Doc
         Call Bt_NewDoc_Click
      End If
      
   End If
   
   lTipoLib = 0
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_BuscarDoc_Click()
   Dim LstIdDoc() As LstDoc_t
   Dim Frm As FrmLstDoc
   Dim Rc As Integer
   Dim TipoLib As Integer
      
   If Val(Grid.TextMatrix(Grid.Row, C_DECENTRALIZ)) <> 0 Then  'viene de centralización, no puede seleccionar un doc
      MsgBox1 "El movimiento seleccionado proviene de una centralización de documentos. No es posible cambiar la identificación del documento asociado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_DEPAGO)) = NEW_DEPAGO Then  'viene de pago automático, no puede seleccionar un doc
      MsgBox1 "El movimiento seleccionado proviene de una generación de pago automático. No es posible cambiar la identificación del documento asociado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
    
   If InStr(LCase(Tx_Glosa), "recauda") Then
      TipoLib = LIB_VENTAS
   ElseIf InStr(LCase(Tx_Glosa), "pago") Then
      TipoLib = LIB_COMPRAS
   End If
   
   Set Frm = New FrmLstDoc
   Rc = Frm.FSelect(LstIdDoc, TipoLib, ED_CENTRALIZADO, True, True)
   Set Frm = Nothing
      
   If Rc = vbOK Then
      
      'marcamos el doc anterior si había, para que recalcule su saldo
      If Val(Grid.TextMatrix(Grid.Row, C_IDDOC)) <> 0 Then
         lDelDocLst = lDelDocLst & "," & Grid.TextMatrix(Grid.Row, C_IDDOC)
      End If
      
      If LstIdDoc(0).IdDoc <> 0 Then      'seleccionó al menos un doc
      
         If LstIdDoc(1).IdDoc = 0 Then    'seleccionó un solo doc => lo asignamos al mov. actual
         
            'If Grid.TextMatrix(Grid.Row, C_IDCUENTA) = "" Then  'no puede seleccionar un doc
            '   MsgBox1 "El movimiento seleccionado no tiene cuenta asignada.", vbExclamation + vbOKOnly
            '   Exit Sub
            'End If
            
            Call AsignarDoc(LstIdDoc(0).IdDoc, LstIdDoc(0).IdDocCuota)
            
         Else    'seleccionó varios Docs, los asignamos sin generar movs automáticamente
         
            Call AsignarLstDoc(LstIdDoc, False)
            
         End If
         
      Else                       'eligió el botón Nuevo Doc
         Call Bt_NewDoc_Click
      End If
      
   End If
   
   
   lTipoLib = 0
   
   Set Frm = Nothing
   
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

Private Sub Bt_Cancel_Click()
   Dim Q1 As String
   If lidComp <> 0 Then
   
      If lOper = O_NEW Then      'por si grabó mientras estaba editando
         Call RemoveComprobante
      
      ElseIf lOper = O_EDIT And lFromCentraliz = True Then   'viene de una operación de centralización y se arrepintió
         Call DeleteComprobante(lidComp, False)
      
      End If
   
   End If
   '*** 2699584
    

    
    Call RemovePercNull
    Call AjustarPercepciones
   
   ' fin 2699584
   
   If lCompTipo = False And lOper = O_NEW And gEmpresa.FCierre = 0 Then
      Call ClearForm
   Else
      lRc = vbCancel
      Unload Me
   End If
   
End Sub

Private Sub bt_CompTipo_Click()
   Dim Frm As FrmSelCompTipo
   Dim IdCompTipo As Long
   Dim Row As Integer
   
   For Row = 1 To Grid.rows - 1
      If Grid.TextMatrix(Row, C_CODCUENTA) <> "" Then
         If MsgBox1("¡ATENCION!, existen movimientos asociados a este Comprobante. Al seleccionar un Comprobante Tipo, éstos se perderán. ¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
            Exit Sub
         Else
            Row = Grid.rows - 1
         End If
      End If
      
   Next Row
   
   Set Frm = New FrmSelCompTipo
   If Frm.FSelect(IdCompTipo) <> vbOK Then
      Exit Sub
   End If
   Set Frm = Nothing
   
   'Cargar Comprobante tipo
   If IdCompTipo > 0 Then
      Call FillCompTipo(IdCompTipo)
   End If
   
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
      
      If Tx_Glosa.Enabled And (Col = C_DEBE Or Col = C_HABER) Then
         
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

Private Sub bt_Copy_Click()
   
   Clipboard.Clear
   Clipboard.SetText Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col)
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, "Número comprobante: " & Tx_Correlativo & " Tipo: " & Cb_Tipo & " Estado: " & Cb_Estado & " Fecha: " & Tx_Fecha)
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
   
   For i = 1 To Grid.rows - 1
      If i <> Row And Grid.RowHeight(i) <> 0 Then
         SumDebe = vFmt(Grid.TextMatrix(i, C_DEBE)) + SumDebe
         SumHaber = vFmt(Grid.TextMatrix(i, C_HABER)) + SumHaber
      End If
   Next i
   
   Diff = SumDebe - SumHaber
   If Diff = 0 Then
      MsgBox1 "El comprobante ya esta cuadrado. No se puede sugerir un valor", vbExclamation
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

Private Sub Bt_DelDoc_Click()
   Dim Q1 As String
   
   If Val(Grid.TextMatrix(Grid.Row, C_IDDOC)) = 0 Then  'no hay doc asignado
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_DECENTRALIZ)) <> 0 Then  'viene de centralización, no puede seleccionar un doc
      MsgBox1 "El movimiento seleccionado proviene de una centralización de documentos. No es posible eliminar el enlace con el documento definido en este movimiento.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_DEPAGO)) = NEW_DEPAGO Then  'viene de pago automático, no puede seleccionar un doc
      MsgBox1 "El movimiento seleccionado proviene de una generación de pago automático. No es posible cambiar la identificación del documento asociado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If MsgBox1("¿Está seguro que desea eliminar el enlace con el documento definido en este movimiento?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   lDelDocLst = lDelDocLst & "," & Grid.TextMatrix(Grid.Row, C_IDDOC)
   
   Call AsignarDoc(0, 0)

End Sub

Private Sub Bt_DetMov_Click()
   Dim EdType As FlexEdGrid2.FEG2_EdType
   Dim Frm As FrmDoc
   Dim IdDoc As Long
   Dim IdDocCuota As Long
   Dim Rc As Integer
   Dim TipoDoc As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   IdDocCuota = Val(Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA))
   
   If IdDoc > 0 Then
      TipoDoc = 0
      If InStr(Grid.TextMatrix(Grid.Row, C_TIPODOC), "ODF") > 0 Then
        TipoDoc = 8
      End If
      Set Frm = New FrmDoc
      Rc = Frm.FEdit(IdDoc, TipoDoc)
      Set Frm = Nothing
      
      'llamamos a asignar doc para que actualice los datos, si algo cambió en el doc
      Call AsignarDoc(IdDoc, IdDocCuota)
      Call CalcTot  'por si acaso

   Else
      MsgBeep vbExclamation
   End If
   
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
   
   If Row >= Grid.rows - 2 Then
      Grid.rows = Grid.rows + 1
   End If
   
   Grid.TextMatrix(Row, C_ORDEN) = Val(Grid.TextMatrix(Row - 1, C_ORDEN)) + 1
   Grid.TextMatrix(Row, C_IDCUENTA) = Grid.TextMatrix(Grid.Row, C_IDCUENTA)
   Grid.TextMatrix(Row, C_CODCUENTA) = Grid.TextMatrix(Grid.Row, C_CODCUENTA)
   Grid.TextMatrix(Row, C_CUENTA) = Grid.TextMatrix(Grid.Row, C_CUENTA)
   Grid.TextMatrix(Row, C_DEBE) = Grid.TextMatrix(Grid.Row, C_DEBE)
   Grid.TextMatrix(Row, C_HABER) = Grid.TextMatrix(Grid.Row, C_HABER)
   Grid.TextMatrix(Row, C_GLOSA) = Grid.TextMatrix(Grid.Row, C_GLOSA)
   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   
'   Call FGrSetPicture(Grid, Row, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
'   Call FGrSetPicture(Grid, Row, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)
   
   Grid.TextMatrix(Row, C_LSTCUENTA) = ">>"
   Grid.TextMatrix(Row, C_DETALLE) = ">>"
   
   Call CalcTot
      
   Grid.Row = Row
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = 0
   Grid.ColSel = Grid.Cols - 1
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Grid.RowSel <> Grid.Row Then
      MsgBox1 "Sólo se puede eliminar un movimiento a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_ORDEN) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If MsgBox1("¿Está seguro que desea borrar este movimiento?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Sub
   End If
      
   
   Call DelMov(Row)
   
   
   Call CalcTot
   
End Sub

Private Sub Bt_Glosas_Click(Index As Integer)
   Dim Frm As FrmGlosas
   Dim Glosa As String
   
   Set Frm = New FrmGlosas
   If Index = 0 Then
      Glosa = FrmGlosas.FSelect(Tx_Glosa)
      If Glosa <> "" Then
         Tx_Glosa = Glosa
      End If
   Else
      Glosa = FrmGlosas.FSelect(Tx_GlosaCompTipo)
      If Glosa <> "" Then
         Tx_GlosaCompTipo = Glosa
      End If
   End If
   
   Set Frm = Nothing
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

Private Sub Bt_NewCompTipo_Click()
   Dim ConValores As Boolean
   Dim i As Integer, Row As Integer
   
   'buscamos la primera línea no borrada
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      If Grid.RowHeight(i) > 0 Then
         Exit For
      End If
   Next i
   
   Row = i
   
   If Trim(Tx_Glosa) = "" Or Val(Grid.TextMatrix(Row, C_IDCUENTA)) = 0 Then
      MsgBox1 "Agregue datos al comprobante antes de crear un Comprobante Tipo.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea crear un Comprobante Tipo a partir de este Comprobante?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
      Exit Sub
   End If
   
   ConValores = False
   
   If vFmt(GridTot.TextMatrix(0, C_DEBE)) <> 0 Or vFmt(vFmt(GridTot.TextMatrix(0, C_HABER))) <> 0 Then
      If MsgBox1("¿Desea incluir los valores del Debe y Haber de este Comprobante en el nuevo Comprobante Tipo?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         ConValores = True
      End If
   End If
   
   Call GenComprobanteTipo(ConValores)
   
End Sub

Private Sub Bt_NewDoc_Click()
   Dim Frm As Form
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim Valor As Double
   Dim i As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Q2 As String
   Dim IdEnt As Long
   Dim TipoLibEnt As Long
   
   If Not ValidaIngresoDoc() Then
      Exit Sub
   End If
   
   'obtenemos la entidad del documento idicaco en la fila inmediatamente anterior a la fila donde se quiere crear el documento
   'pensando que es un cheque de pago del documento anterior
   i = Grid.Row - 1
   
   If i >= Grid.FixedRows Then
   
      If Grid.TextMatrix(i, C_TIPODOC) <> "" And Grid.TextMatrix(i, C_TIPODOC) <> "CHE" Or Grid.TextMatrix(i, C_TIPODOC) = "CHF" Then
         
         'obtenemos el Id de la Entidad y el TipoLib asociada al doc que vamos a pagar
         Q1 = "SELECT IdEntidad, TipoLib FROM Documento WHERE IdDoc=" & Grid.TextMatrix(i, C_IDDOC)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         'Q2 = Replace(Q1, "Documento", "DocumentoFull")
         'Q1 = Q1 = " UNION ALL " & Q2
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdEnt = vFld(Rs("IdEntidad"))
            TipoLibEnt = vFld(Rs("TipoLib"))
         End If
         Call CloseRs(Rs)
                  
      End If
      
   End If

   
'   If lTipoLib = 0 Then
'
'      Set Frm = New FrmSelLibDocs
'
'      Rc = Frm.FSelect(lTipoLib, True)
'      Set Frm = Nothing
'
'      If Rc <> vbOK Then
'         Exit Sub
'      End If
'
'   End If
'
'   If lTipoLib = LIB_COMPRAS Or lTipoLib = LIB_VENTAS Then
'      Set Frm = New FrmCompraVenta
'      Rc = Frm.FEdit(lTipoLib, GetMesActual(), gEmpresa.Ano, IdDoc)
'   ElseIf lTipoLib = LIB_RETEN Then
'      Set Frm = New FrmLibRetenciones
'      Rc = Frm.FEdit(GetMesActual(), gEmpresa.Ano, IdDoc)
'   Else
      Set Frm = New FrmDoc
      Rc = Frm.FNew(0, IdDoc, False, GetMesActual(), gEmpresa.Ano, Abs(vFmt(Grid.TextMatrix(Grid.Row, C_DEBE)) - vFmt(Grid.TextMatrix(Grid.Row, C_HABER))), IdEnt, TipoLibEnt)
'   End If
   
   Set Frm = Nothing
   lTipoLib = 0
   
   If Rc = vbOK And IdDoc > 0 Then
   
'      If Grid.TextMatrix(Grid.Row, C_IDCUENTA) = "" Then  'puede seleccionar un doc
'         MsgBox1 "El movimiento seleccionado no tiene cuenta asignada.", vbExclamation + vbOKOnly
'         Exit Sub
'      End If
      
      If Val(Grid.TextMatrix(Grid.Row, C_DECENTRALIZ)) <> 0 Then  'viene de centralización, no puede seleccionar un doc
         MsgBox1 "El movimiento seleccionado proviene de una centralización de documentos. No es posible cambiar la identificación del documento asociado.", vbExclamation + vbOKOnly
         Exit Sub
      End If
      
      If Val(Grid.TextMatrix(Grid.Row, C_DEPAGO)) = NEW_DEPAGO Then 'viene de pago automático, no puede seleccionar un doc
         MsgBox1 "El movimiento seleccionado proviene de una generación de pago automático. No es posible cambiar la identificación del documento asociado.", vbExclamation + vbOKOnly
         Exit Sub
      End If
      
      Call AsignarDoc(IdDoc, 0)
      Call CalcTot  'PS Se agregó porque sumaba al total el nuevo valor en caso de ponerlo, y se grababa sin reclamar nada, ya que el DEBE y el HABER quedaba en cero
      
   End If

End Sub

Private Sub Bt_Paste_Click()
   Dim Fmt As Integer
   Dim DVal As Double
   Dim Action As Integer
   
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
      
   ElseIf Grid.Col = C_CODCUENTA Then
   
      If Val(Grid.TextMatrix(Grid.Row, C_ORDEN)) = 0 Then    'nuevo
         Call AddNewMov(Grid.Row)
      End If
      
      Call Grid_AcceptValue(Grid.Row, Grid.Col, Clipboard.GetText, Action)
      If Action = vbOK Then
         Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = Clipboard.GetText
      End If
      
   Else
      Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = Clipboard.GetText
      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDMOV, C_UPDATE)
   End If
      
End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   
   If Val(Grid.TextMatrix(Grid.FixedRows, C_IDCUENTA)) = 0 Then
      MsgBox1 "Este comprobante no tiene movimientos.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If Ch_ImpRes <> 0 And Not lViewResumido Then
      If MsgBox1("Este comprobante se imprimirá tal como se ve en pantalla." & vbLf & vbLf & "Si desea imprimirlo en forma  resumida, utilice el botón: " & vbLf & vbLf & "      'Ver comprobante seleccionado en forma resumida'" & vbLf & vbLf & "disponible en el listado de comprobantes.", vbInformation + vbOKCancel) = vbCancel Then
         Exit Sub
      End If
   End If
   
   If SetUpPrtGrid = False Then
      Exit Sub
   End If
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   
   '2860036
   'Pag = gPrtReportes.PrtFlexGrid(Frm)
   Pag = gPrtReportes.PrtFlexGridMembrete(Frm)
   
   'fin 2860036
   
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call ResetPrtBas(gPrtReportes)

   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   Dim Pag As Integer
   
   If Val(Grid.TextMatrix(Grid.FixedRows, C_IDCUENTA)) = 0 Then
      MsgBox1 "Este comprobante no tiene movimientos.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If vFmt(Tx_Correlativo) = 0 Then
      If MsgBox1("Se debe grabar el comprobante antes de imprimirlo." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton3) = vbNo Then
         Exit Sub
      End If
      
      If valida() Then
         Call SaveAll
      Else
         Exit Sub
      End If
   End If
   
   If Ch_ImpRes <> 0 And Not lViewResumido Then
      If MsgBox1("Este comprobante se imprimirá tal como se ve en pantalla." & vbLf & vbLf & "Si desea imprimirlo en forma  resumida, utilice el botón: " & vbLf & vbLf & "'Ver comprobante seleccionado en forma resumida'" & vbLf & vbLf & "disponible en el listado de comprobantes.", vbInformation + vbOKCancel) = vbCancel Then
         Exit Sub
      End If
   End If
   
   OldOrientation = Printer.Orientation
   
   If SetUpPrtGrid() = False Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   gPrtReportes.CallEndDoc = False
   
   '2860036
   Pag = gPrtReportes.PrtFlexGrid(Printer)
   'Pag = gPrtReportes.PrtFlexGridMembrete(Printer)
   
   'fin 2860036
   
   Call PrtVisacion(Printer, Pag, gPrtReportes.GrLeft, gPrtReportes.GrRight)
   
   gPrtReportes.CallEndDoc = True
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
   Call ResetPrtBas(gPrtReportes)

End Sub
Private Function PrtVisacion(PrtObj As Object, ByVal Pag As Integer, ByVal LeftX As Integer, ByVal RightX As Integer) As Integer
   Dim PrtPage As Object
   Dim CurY As Integer
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   PrtPage.Print
   PrtPage.Print
   
   CurY = PrtPage.CurrentY
   
   If PrtPage.CurrentY >= PrtPage.Height - COMP_FOOTER - COMP_VBOX Then
      Call gPrtReportes.PrtFooter(PrtPage, "Continua >>>", RightX)
      Set PrtPage = NewPage(PrtObj)
      Pag = Pag + 1
      
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      CurY = PrtPage.CurrentY

   End If
   
   PrtPage.Line (LeftX, CurY)-(RightX, CurY)
   PrtPage.Line (LeftX, CurY)-(LeftX, CurY + COMP_VBOX)
   PrtPage.Line (RightX, CurY)-(RightX, CurY + COMP_VBOX)
   PrtPage.Line (LeftX, CurY + COMP_VBOX)-(RightX, CurY + COMP_VBOX)
   
   PrtPage.Line (LeftX + 1000, CurY)-(LeftX + 1000, CurY + COMP_VBOX)
   PrtPage.Line (LeftX + 2000, CurY)-(LeftX + 2000, CurY + COMP_VBOX)
   
   PrtPage.CurrentY = CurY + COMP_VBOX - Grid.RowHeight(0)
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "V°B°1";
   PrtPage.CurrentX = LeftX + 1300
   PrtPage.Print "V°B°2";
   PrtPage.CurrentX = LeftX + 2400
   PrtPage.Print "Fecha";
   PrtPage.CurrentX = LeftX + 4000
   PrtPage.Print "Nombre";
   PrtPage.CurrentX = LeftX + 6500
   PrtPage.Print "RUT";
   PrtPage.CurrentX = LeftX + 8500
   PrtPage.Print "Firma"
   
   RightX = PrtPage.Width - 2000
   '2860036
   
   If gMembrete.TxtTitMembrete2 <> "" Then
   
    Call PrtFooterMembreteRigth(PrtPage, gMembrete.TxtTitMembrete2, gMembrete.TxtTexto2, RightX)
    
    End If
    
     If gMembrete.TxtTitMembrete1 <> "" Then
    
    Call PrtFooterMembreteLeft(PrtPage, gMembrete.TxtTitMembrete1, gMembrete.TxtTexto1, LeftX)
    
    End If
   'fin 2860036
   PrtPage.EndDoc
   

End Function

Private Sub Bt_PrtCheque_Click()
   Dim i As Integer
   Dim IdCheque As Long
   Dim Ref As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Frm As FrmPrtCheque
   Dim NumEgreso As Long
   Dim FechaCheque As Long
   Dim NumCheque As Long
   Dim Nombre As String
   Dim Banco As String
   Dim Valor As String
   Dim IdCuenta As Long, IdCuentaCheque As Long
   Dim NombCuenta As String
   Dim ColWi(NCOLS) As Integer
   Dim NombreEnt As String
   
   
   If Bt_OK.visible Then
   
      If MsgBox1("Se grabarán los datos antes de imprimir el cheque." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If valida() Then
         Call SaveAll
      Else
         Exit Sub
      End If
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1

      If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If

      'ubicamos el documento que se va a pagar (factura u otro)
      If Grid.TextMatrix(i, C_TIPODOC) <> "" And Grid.TextMatrix(i, C_TIPODOC) <> "CHE" Or Grid.TextMatrix(i, C_TIPODOC) = "CHF" Then
         Ref = Ref & Grid.TextMatrix(i, C_TIPODOC) & " " & Grid.TextMatrix(i, C_NUMDOC) & ", "
         NombreEnt = Grid.TextMatrix(i, C_ENTIDAD)
         
      End If

      'cheque o cheque a fecha
      If Grid.TextMatrix(i, C_TIPODOC) = "CHE" Or Grid.TextMatrix(i, C_TIPODOC) = "CHF" Then
         IdCheque = Val(Grid.TextMatrix(i, C_IDDOC))
         IdCuenta = Val(Grid.TextMatrix(i, C_IDCUENTA))
         NombCuenta = Grid.TextMatrix(i, C_CUENTA)
         
         Exit For
      End If
      
   Next i
   
   If IdCheque = 0 Then
      MsgBox1 "No se encontró cheque para imprimir.", vbExclamation
      Exit Sub
   End If
   
   If Ref <> "" Then
      Ref = "Pago " & Left(Ref, Len(Ref) - 2)
   End If
   NumEgreso = Val(Tx_Correlativo)
   
   Q1 = "SELECT NumDoc, FEmision, Entidades.Nombre, Cuentas.Descripcion, Documento.IdCtaBanco, Documento.Total"
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True) & ")"
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCtaBanco = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento")
   Q1 = Q1 & " WHERE IdDoc = " & IdCheque
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      NumCheque = vFld(Rs("NumDoc"))
      FechaCheque = vFld(Rs("FEmision"))
      Nombre = vFld(Rs("Nombre"))
      Banco = vFld(Rs("Descripcion"))
      IdCuentaCheque = vFld(Rs("IdCtaBanco"))
      Valor = vFld(Rs("Total"))
   End If
   
   Call CloseRs(Rs)
   
   If Nombre = "" Then
      MsgBox1 "No es posible imprimir el cheque. Falta ingresar el nombre de la entidad a la orden de la cual se emite el cheque.", vbExclamation
      Exit Sub
   End If
   
   If Nombre <> NombreEnt Then
      If MsgBox1("El nombre de la entidad a la orden de la cual se emite el cheque difiere de la entidad asociada al documento a pagar." & vbCrLf & vbCrLf & "¿Desea continar con la impresión del cheque?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
      
'   If IdCuentaCheque = 0 Then
'      MsgBox1 "No es posible imprimir el cheque. Falta ingresar la cuenta bancaria en el detalle del documento.", vbExclamation
'      Exit Sub
'   End If
   
'  Esto se valida en la función Valida
'   If IdCuenta <> IdCuentaCheque And IdCuentaCheque <> 0 Then
'      If MsgBox1("La cuenta bancaria especificada en el detalle del cheque no coincide con la cuenta bancaria indicada en el comprobante." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'         Exit Sub
'      End If
'   End If

   Call SetUpPrtGrid(False)
      
   Set Frm = New FrmPrtCheque
   Call Frm.FPrint(False, gPrtReportes, NumCheque, FechaCheque, Nombre, Ref, NombCuenta, Valor, NumEgreso)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Salir_Click()
   Dim Rc As Integer
   Dim Q1 As String

   If (Tx_Glosa <> "" Or Grid.TextMatrix(Grid.FixedRows, C_ORDEN) <> "") And lOper = O_NEW Then
      Rc = MsgBox1("¿Desea guardar el comprobante actual?", vbYesNoCancel + vbDefaultButton1 + vbQuestion)
      
      If Rc = vbYes Then
         If valida() Then
            Call SaveAll
            lRc = vbOK
            Unload Me
         End If
      ElseIf Rc = vbNo Then
         Call RemovePercNull
         Call RemoveComprobante
         lRc = vbCancel
         Unload Me
      Else
         Exit Sub
      End If
   Else
      lRc = vbCancel
      Unload Me
   End If
   
   
End Sub

Private Function SetUpPrtGrid(Optional ByVal PreguntaColumnas As Boolean = True) As Boolean
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim Rc As Integer
   
   SetUpPrtGrid = True
   
   Set gPrtReportes.Grid = Grid.FlxGrid
   
   Printer.Orientation = ORIENT_VER
   
   If gTituloTipoComp Then
      Titulos(0) = "Comprobante de " & Cb_Tipo & " N° " & lCorrelativo
   Else
      Titulos(0) = "Comprobante Contable N° " & lCorrelativo
   End If
   If CbItemData(Cb_Estado) = EC_ANULADO Then
      Titulos(1) = "(Anulado)"
   End If
   gPrtReportes.Titulos = Titulos
   
   Encabezados(0) = "N° Comp.:" & vbTab & lCorrelativo
   Encabezados(1) = "Fecha:" & vbTab & Tx_Fecha
   Encabezados(2) = "Tipo:" & vbTab & Cb_Tipo
   Encabezados(3) = "Glosa:" & vbTab & Tx_Glosa
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To C_NUMDOC
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   ColWi(C_CODCUENTA) = ColWi(C_CODCUENTA) - 200
   ColWi(C_CUENTA) = ColWi(C_CUENTA) - 100
   ColWi(C_CODCUENTA) = ColWi(C_CODCUENTA) - 200
   ColWi(C_GLOSA) = ColWi(C_GLOSA) - 200
      
   If gPrtMovDetOpt = 0 Then
   
      If PreguntaColumnas Then
      
         Rc = MsgBox1("¿Desea imprimir la columna Entidad en vez de la columna Descripción?", vbQuestion + vbYesNoCancel + vbDefaultButton2)
         
         If Rc = vbCancel Then
            SetUpPrtGrid = False
            Exit Function
         
         ElseIf Rc = vbYes Then
            gPrtMovDetOpt = PRTMOV_ENTIDAD
            Call GrabarPrtMovDet
         
         End If
      
      End If
   
   End If
   
   Select Case gPrtMovDetOpt        'Default: PRTMOV_DESC (descripción)
      
      Case PRTMOV_ENTIDAD
         ColWi(C_ENTIDAD) = ColWi(C_GLOSA)
         ColWi(C_GLOSA) = 0
            
      Case PRTMOV_CCOSTO
         ColWi(C_CCOSTO) = ColWi(C_GLOSA)
         ColWi(C_GLOSA) = 0
         
       Case PRTMOV_AREANEG
         ColWi(C_AREANEG) = ColWi(C_GLOSA)
         ColWi(C_GLOSA) = 0
        
   End Select
   
   ColWi(C_LSTCUENTA) = 0
   ColWi(C_DETALLE) = 0
   
   Total(C_CODCUENTA) = "Total"
   Total(C_DEBE) = GridTot.TextMatrix(0, C_DEBE)
   Total(C_HABER) = GridTot.TextMatrix(0, C_HABER)
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.NTotLines = 1
   gPrtReportes.ColObligatoria = C_ORDEN

   gPrtReportes.PrintFecha = False

End Function

Private Sub Bt_SelCompTipo_Click()
   Dim IdCompTipo As Long
   Dim Row As Integer
   Dim Nombre As String
   
   For Row = 1 To Grid.rows - 1
      If Grid.TextMatrix(Row, C_CODCUENTA) <> "" Then
         If MsgBox1("¡ATENCION!, existen movimientos asociados a este Comprobante. Al seleccionar un Comprobante Tipo, éstos se perderán. ¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
            Exit Sub
         Else
            Row = Grid.rows - 1
         End If
      End If
      
   Next Row
   Nombre = UCase(Trim(Tx_NombCompTipo))
   If Nombre <> "" Then
      IdCompTipo = GetCompTipo(Nombre)
      If IdCompTipo > 0 Then
         Call FillCompTipo(IdCompTipo)
      Else
         MsgBox1 "No existe un Comprobante Tipo con este nombre.", vbExclamation + vbOKOnly
      End If
   End If
   

End Sub

Private Sub Bt_SelFecha_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
  
   Call Frm.TxSelDate(Tx_Fecha)
   
'   Call Frm.SelDate(Fecha)
'   Call SetTxDate(Tx_Fecha, Fecha)
   
   Set Frm = Nothing
End Sub

Private Sub Bt_Cuentas_Click()
   Dim Frm As FrmPlanCuentas
   
   Set Frm = New FrmPlanCuentas

   Call Frm.FEdit(False)
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()
   Call SetTblName(lCompTipo)
   If valida() Then
      Call SaveAll
        
       '3126513
      #If DATACON <> 1 Then
       
        Call CorrigeCuentaCompTipo
      
      #End If
      '3126513
      
      If lCompTipo = False And lOper = O_NEW Then
         'Limpio la grilla
         Call ClearForm
      Else
         'Para comprobante Tipo sirve así
         lRc = vbOK
         Unload Me
      End If
   End If

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBE, C_HABER, Grid.FlxGrid.Row, Grid.FlxGrid.RowSel)
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_ToLeft_Click()
   
   If lClsPaging.CurReg <= 1 Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call SaveMovs
   
   lClsPaging.ToLeft
   
   Call LoadAll
   
   'Call CalcTotFull
   
   Me.MousePointer = vbDefault
   
   
End Sub

Private Sub Bt_ToRight_Click()
    Me.MousePointer = vbHourglass
   Dim Modif As Boolean
   Dim i As Integer

   If Grid.TextMatrix(Grid.FixedRows, C_ORDEN) = "" Then
      Exit Sub
   End If
   
  ' If Not ((lOper <> O_EDIT And lOper <> O_NEW)) Then
            
'      For i = Grid.FixedRows To Grid.rows - 1
'         If Grid.TextMatrix(i, C_FECHA) = "" Then
'            Exit For
'         End If
'         If Grid.TextMatrix(i, C_UPDATE) <> "" Then
'            Modif = True
'            Exit For
'         End If
'
'      Next i
      
      'If Modif Then

         If MsgBox1("Antes de pasar al siguiente conjunto de registros se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Me.MousePointer = vbDefault
            Exit Sub
         Else
         Call SaveMovs
   
         Call lClsPaging.ToRight
        
         Call LoadAll
         End If
         
      'End If
   
   
   
   'Call CalcTotFull
   
   Me.MousePointer = vbDefault
   
   lClsPaging.ToRightPressed = False
 'End If
End Sub

Private Sub Cb_Estado_Click()
   
   If ItemData(Cb_Estado) = EC_PENDIENTE And lEstado = EC_APROBADO Then
      Call EnabForm(True)
      
      If lOper = O_EDIT And gTipoCorrComp = TCC_TIPOCOMP Then
         'si ya fue creado y el correlativo es por tipo, no es posible cambiar el tipo
         Cb_Tipo.Enabled = False
      End If
   
   'anula doc.
   ElseIf ItemData(Cb_Estado) = EC_ANULADO And lEstado <> EC_ANULADO Then
      
      If lTieneMovCentraliz Then
         
         If MsgBox1("Recuerde que este comprobante contiene movimientos de centralización de documentos ¿Está seguro que desea anularlo?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Cb_Estado.ListIndex = FindItem(Cb_Estado, lEstado)
         
         ElseIf lTieneMovPago Then
            If MsgBox1("Recuerde que este comprobante contiene movimientos de pago de documentos ¿Está seguro que desea anularlo?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Cb_Estado.ListIndex = FindItem(Cb_Estado, lEstado)
            End If
         End If
         
      Else
         If MsgBox1("¿Está seguro que desea anular este comprobante?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Cb_Estado.ListIndex = FindItem(Cb_Estado, lEstado)
         End If
      End If
      
      If ItemData(Cb_Estado) = EC_ANULADO Then
         If MsgBox1("¿Desea dejar el valor de todos los movimientos en cero?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            Call DejarEnCeroMovs
         End If
      End If
      
   End If
   
   
End Sub


Private Sub Cb_Tipo_Click()

   If CbItemData(Cb_Tipo) = TC_TRASPASO Then
      Ch_OtrosIngEg14TER.Value = 0
      Ch_OtrosIngEg14TER.Enabled = False
   ElseIf Not lGenPago Then
      Ch_OtrosIngEg14TER.Enabled = True
   End If
      
End Sub

Private Sub Form_Activate()

   If Not Ch_OtrosIngEg14TER.visible Then
      Ch_OtrosIngEg14TER.Value = 0
      Ch_OtrosIngEg14TER.Enabled = False
      
   ElseIf Ch_OtrosIngEg14TER.Enabled Then
      Ch_OtrosIngEg14TER = IIf(lOtrosIngEg14TER <> 0, 1, 0)
   End If
  
  
   
End Sub

Private Sub Form_Load()
   Dim i As Integer, ActionLock As Integer
   Dim FrmEnable As Boolean
   Dim MesActual As Integer
   Dim EditEnable As Boolean
            
   lRc = 0
   
   lGenPago = 0
   
   If lOper = O_NEW Then
      Caption = "Nuevo Comprobante"
   ElseIf lOper = O_EDIT Then
      Caption = "Modificar Comprobante"
   ElseIf lOper = O_VIEW Then
      Caption = "Ver Comprobante"
   End If
   
   Bt_PrtCheque.visible = gFunciones.PrtCheque

   If lCompTipo Then
      Caption = Caption & " Tipo"
   End If
   
   Ch_ImpRes.visible = gFunciones.ComprobanteResumido
   
   If gEmpresa.Ano >= 2020 Then
      Ch_OtrosIngEg14TER.Caption = "Otros Ingresos/Egresos 14D"
   End If
   
   Ch_OtrosIngEg14TER.Value = 1
   Ch_OtrosIngEg14TER.Enabled = True

   Ch_OtrosIngEg14TER.visible = gEmpresa.Franq14Ter Or gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp
   
   
   lCuentasDisponible = False
   lHayDocLibros = False
   
   For i = 1 To N_TIPOCOMP
   
      If Not (lOper = O_NEW And i = TC_APERTURA) Then
         Cb_Tipo.AddItem gTipoComp(i)
         Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
      End If
         
      If i <> TC_APERTURA Then
         Cb_TipoCompTipo.AddItem gTipoComp(i)
         Cb_TipoCompTipo.ItemData(Cb_TipoCompTipo.NewIndex) = i
      End If
         
   Next i
   
   Cb_Tipo.ListIndex = -1
   Cb_TipoCompTipo.ListIndex = -1
   
   For i = 1 To N_ESTADOCOMP
      Cb_Estado.AddItem gEstadoComp(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   Call SelItem(Cb_Estado, gEstadoNewComp)
   
   'Tipo de ajuste comprobante
   Op_TAjuste(TAJUSTE_AMBOS).Value = 1 'valor por omisón
      
   Call SetUpGrid
   
   Bt_ActivoFijo.visible = gFunciones.ActivoFijo
   Bt_ActivoFijo.Enabled = gFunciones.ActivoFijo
   
   MesActual = GetMesActual()
   
   If MesActual = 0 Then
      MesActual = GetUltimoMesConComps()
   End If
   
   If MesActual = month(Now) Then
      Call SetTxDate(Tx_Fecha, DateSerial(gEmpresa.Ano, month(Now), Day(Now)))
   Else
      Call SetTxDate(Tx_Fecha, DateSerial(gEmpresa.Ano, MesActual, 1))
   End If
   
    '3188805
   
   Set lClsPaging = New ClsPaging

    Call lClsPaging.Init(Bt_ToLeft, Bt_ToRight)

   If gDbType = SQL_SERVER Then

    fr_Paginacion.visible = True

        If lOper = O_NEW Or lOper = O_EDIT Then
         If lCentrFull = 1 Then
          GridTotFull.visible = True

         Call CalcTotFull(0, 0, 0, 0)
         End If
        End If

       If ValidaCompFull(lidComp) Then
         lCentrFull = 1

        GridTotFull.visible = True

        Call CalcTotFull(0, 0, 0, 0)
       End If

   End If
'
   '3188805
   
   Call LoadAll(lViewResumido)
   
   'se habilita el form si el periodo no está cerrado y (si es New y hay un  mes actual abierto) o (si es edit y el mes del comprobante está abierto)
   FrmEnable = (gEmpresa.FCierre = 0) And ((lOper = O_NEW And MesActual > 0) Or (lOper = O_EDIT And lidComp > 0 And GetEstadoMes(month(GetTxDate(Tx_Fecha))) = EM_ABIERTO))
   
   Call EnabForm(FrmEnable)
      

   Bt_CompTipo.Enabled = (lOper = O_NEW) And FrmEnable
   Tx_NombCompTipo.Enabled = Bt_CompTipo.Enabled
   
   Fr_Header(0).visible = (lCompTipo = False)
   Fr_Header(1).visible = (lCompTipo = True)
   Fr_SelCompTipo.visible = (lCompTipo = False)

   Bt_SelCompTipo.Enabled = False
   
   If lCompTipo Or lOper <> O_NEW Then
      Bt_OK.Left = Bt_Cancel.Left
      Bt_Cancel.Left = Bt_Salir.Left
      Bt_Salir.visible = False
   
   ElseIf Not lCompTipo Then
      Bt_OK.Caption = "Grabar"
      
   End If
   
   If lCompTipo Then
      Fr_Doc.visible = False
      Fr_Doc.Enabled = False
   End If
   
   If lFromCentraliz = True And FrmEnable Then
      'Bt_Cancel.Enabled = False
      Bt_NewDoc.Enabled = False
      Bt_DetMov.Enabled = False
      Fr_SelCompTipo.visible = False
   End If
      
   
   If lCompTipo = False And FrmEnable Then
      
      'Sólo se habilita para New y Edit
      Call EnabForm(lOper = O_NEW Or lOper = O_EDIT)
      
      'si el estado es anulado o aprobado, no se puede modificar el comprobante, pero se puede modificar si es New y parte con estado apropbado
      If (CbItemData(Cb_Estado) = EC_APROBADO And lOper = O_EDIT) Or CbItemData(Cb_Estado) = EC_ANULADO Then
         
         Call EnabForm(False)
         
         'Permitimos cambiar el estado sólo si está aprobado y es del mes abierto. Si está anulado no se puede cambiar.
         If Cb_Estado.ItemData(Cb_Estado.ListIndex) = EC_APROBADO And lOper = O_EDIT Then
            Cb_Estado.Enabled = True
            Bt_OK.visible = True
            Bt_Cancel.Caption = "Cancelar"
         End If
      End If
                  
      If lOper = O_EDIT And gTipoCorrComp = TCC_TIPOCOMP Then
         'si ya fue creado y el correlativo es por tipo, no es posible cambiar el tipo
         Cb_Tipo.Enabled = False
      End If
         
   End If
   
   'vemos si podemos bloquear el comprobante, para que nadie más lo edite
   If lidComp <> 0 Then
      If lCompTipo Then
         ActionLock = LK_COMPTIPO
      Else
         ActionLock = LK_COMPROBANTE
      End If
      
      EditEnable = LockAction(DbMain, ActionLock, lidComp)
      
      If EditEnable = False Then    'alguien más lo está editando, no podemos editarlo
         Call EnabForm(False)
         MsgBox1 "Este comprobante se está editando en el equipo '" & IsLockedAction(DbMain, ActionLock, lidComp) & "'. Sólo se abrirá de lectura.", vbInformation
      End If
   End If

   Call SetTxRO(Tx_IdComp, True)
   Call SetTxRO(Tx_Correlativo, True)
   
   'no se puede cambiar el tipo de ajuste una vez que ha sido creado el comprobante
   If Fr_TAjuste.Enabled And lOper = O_EDIT Then
      Fr_TAjuste.Enabled = False
   End If
   If lCompTipo Then
      Fr_TAjuste.visible = False
   End If


   If lEnImpresionMasiva Then
      Call Bt_Print_Click
   End If

End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   Dim Q1 As String
   
   Grid.Cols = NCOLS + 1
      
   Grid.ColWidth(C_IDMOV) = 0
   Grid.ColWidth(C_ORDEN) = 450
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CODCUENTA) = 1650
   Grid.ColWidth(C_CUENTA) = 2200
   Grid.ColWidth(C_LSTCUENTA) = 250
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_IDDOCCUOTA) = 0
   Grid.ColWidth(C_DECENTRALIZ) = 0
   Grid.ColWidth(C_DEPAGO) = 0
   Grid.ColWidth(C_ATRIB_CONCIL) = 0
   Grid.ColWidth(C_NOTA) = 0
   Grid.ColWidth(C_GLOSA) = 3000
   
   If lCompTipo = False Then
      Grid.ColWidth(C_DETALLE) = 250
      Grid.ColWidth(C_TIPODOC) = 550
      Grid.ColWidth(C_NUMDOC) = 930
      Grid.ColWidth(C_ENTIDAD) = 1500
      Grid.ColWidth(C_AREANEG) = 2100
      Grid.ColWidth(C_IDAREANEG) = 0
      Grid.ColWidth(C_CCOSTO) = 2000
      Grid.ColWidth(C_IDCCOSTO) = 0
      If gFunciones.ActivoFijo Then
         Grid.ColWidth(C_DETACTFIJO) = 800
      Else
         Grid.ColWidth(C_DETACTFIJO) = 0  'se oculta por ahora
      End If
      
      'If lFromCentraliz Then     'de centralización -> ocultamos la columna Orden para que no se vea desordenada y confunda al usuario. Al grabar se ordena.
      '   Grid.ColWidth(C_GLOSA) = Grid.ColWidth(C_GLOSA) + Grid.ColWidth(C_ORDEN)
      '   Grid.ColWidth(C_ORDEN) = 0
      'End If
      
   Else
      Grid.ColWidth(C_DETALLE) = 0
      Grid.ColWidth(C_TIPODOC) = 0
      Grid.ColWidth(C_NUMDOC) = 0
'      Grid.ColWidth(C_GLOSA) = 4500
      Grid.ColWidth(C_ENTIDAD) = 0
      Grid.ColWidth(C_AREANEG) = 2100
      Grid.ColWidth(C_IDAREANEG) = 0
      Grid.ColWidth(C_CCOSTO) = 2000
      Grid.ColWidth(C_IDCCOSTO) = 0
      Grid.ColWidth(C_DETACTFIJO) = 0
   End If
     
   Grid.ColAlignment(C_IDMOV) = flexAlignRightCenter
   Grid.ColAlignment(C_ORDEN) = flexAlignRightCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_LSTCUENTA) = flexAlignCenterCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_DETALLE) = flexAlignCenterCenter
   Grid.ColAlignment(C_DETACTFIJO) = flexAlignCenterCenter
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   '3269719
   Call FGrTotales(Grid, GridTotFull)
   GridTotFull.TextMatrix(0, C_CUENTA) = "TOTAL COMP.:"
   '3269719
   
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
   Grid.TextMatrix(0, C_AREANEG) = "Área Negocio"
   Grid.TextMatrix(0, C_CCOSTO) = "Centro Gestión"
   
   If lCompTipo = False Then
      Grid.TextMatrix(0, C_TIPODOC) = "TD"
      Grid.TextMatrix(0, C_NUMDOC) = "Nº Doc."
      Grid.TextMatrix(0, C_ENTIDAD) = "Entidad"
'      Grid.TextMatrix(0, C_AREANEG) = "Área Negocio"
'      Grid.TextMatrix(0, C_CCOSTO) = "Centro Gestión"
   End If
   If gFunciones.ActivoFijo Then
      Grid.TextMatrix(0, C_DETACTFIJO) = "Act. Fijo"
   End If
   Grid.TextMatrix(0, C_DETALLE) = ""
   Grid.FlxGrid.Row = 0
   Grid.FlxGrid.Col = C_DETALLE
   Set Grid.CellPicture = FrmMain.Pc_Lupa
   
   GridTot.TextMatrix(0, C_CUENTA) = "TOTAL:"

   Call AddItem(Grid.CbList(C_AREANEG), " ", 0)
   Q1 = "SELECT Descripcion, IdAreaNegocio FROM AreaNegocio "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 "
   Q1 = Q1 & " ORDER BY Descripcion "
   Call FillCombo(Grid.CbList(C_AREANEG), DbMain, Q1, 0, True)
   
   Call AddItem(Grid.CbList(C_CCOSTO), " ", 0)
   Q1 = "SELECT Descripcion, IdCCosto FROM CentroCosto "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 "
   Q1 = Q1 & " ORDER BY Descripcion "
   Call FillCombo(Grid.CbList(C_CCOSTO), DbMain, Q1, 0, True)
  
   Call FGrVRows(Grid)
   
End Sub
Public Function FNew(CompTipo As Boolean) As Integer
   
   lOper = O_NEW
   Call SetTblName(CompTipo)
   
   lCompTipo = CompTipo
   lidComp = 0
   lCorrelativo = 0
   
   Me.Show vbModal
   
   FNew = lRc
End Function
Public Function FEdit(idcomp As Long, CompTipo As Boolean) As Integer
   
   lOper = O_EDIT
   Call SetTblName(CompTipo)
   
   lCompTipo = CompTipo
   lidComp = idcomp
   
   Me.Show vbModal
   
   FEdit = lRc
End Function
Public Function FEditCentraliz(ByVal idcomp As Long, ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal vCentrFull As Integer = 0) As Integer
   
   lOper = O_EDIT
   lidComp = idcomp
   
   Call SetTblName(False)
   lFromCentraliz = True
   If Mes > 0 Then
      lMesAnoCentraliz = DateSerial(Ano, Mes, 1)
   Else
      lMesAnoCentraliz = 0
   End If
   
   '3238739
   lCentrFull = vCentrFull
   '3238739
   Me.Show vbModal
   
   FEditCentraliz = lRc
End Function

Public Function FView(idcomp As Long, CompTipo As Boolean) As Integer
   lOper = O_VIEW
   lidComp = idcomp
   'lCorrelativo = GetCorrelativoComp(idComp)
   
   Call SetTblName(CompTipo)
      
   lCompTipo = CompTipo
   
   Me.Show vbModal
   
   FView = lRc
End Function

Private Sub SetTblName(CompTipo As Boolean)

   lTblDocumento = "Documento"
   If CompTipo = True Then
      lTblComprobante = "CT_Comprobante"
      lTblMovComprobante = "CT_MovComprobante"
   Else
      'If CodTipoLib <> 8 Then
        lTblComprobante = "Comprobante"
        lTblMovComprobante = "MovComprobante"
        lTblDocumento = "Documento"
'      Else
'        lTblComprobante = "ComprobanteFull"
'        lTblMovComprobante = "MovComprobanteFull"
'        lTblDocumento = "DocumentoFull"
'      End If
   End If

End Sub
Private Sub LoadAll(Optional ByVal Resumido As Boolean = False, Optional ByVal Row As Integer = 0)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Tipo As Integer
   Dim n As Integer

   If lidComp = 0 Then
      Tx_Usuario = gUsuario.Nombre
      Exit Sub
   End If
   
   Q1 = "SELECT Correlativo, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber, IdUsuario  "
   If lCompTipo = True Then
      Q1 = Q1 & ", Nombre, Descrip"
   Else
      Q1 = Q1 & ", OtrosIngEg14TER, ImpResumido, TipoAjuste "
   End If
   Q1 = Q1 & " FROM " & lTblComprobante & " WHERE IdComp = " & lidComp
   
   If lCompTipo = False Then
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Else
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   End If
   
'   If lOper <> O_EDIT Then
'        Q2 = Replace(Q1, "Comprobante", "ComprobanteFull")
'        Q1 = Q1 & " UNION ALL " & Q2
'   End If

  
    
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      lCorrelativo = vFld(Rs("Correlativo"))
      
      If lCompTipo = False Then
         Call SetTxDate(Tx_Fecha, vFld(Rs("Fecha")))
         lOldFechaEmision = vFld(Rs("Fecha"))
         Tx_IdComp = lidComp
         If vFld(Rs("Correlativo")) > 0 Then
            Tx_Correlativo = lCorrelativo
         End If
         lEstado = vFld(Rs("Estado"))
         Call SelItem(Cb_Estado, vFld(Rs("Estado")))
         Tx_Glosa = vFld(Rs("Glosa"), True)
         Tx_Usuario = GetNombreUsuario(vFld(Rs("IdUsuario")))
         Ch_ImpRes = IIf(vFld(Rs("ImpResumido")) <> 0, 1, 0)
         lOtrosIngEg14TER = IIf(vFld(Rs("OtrosIngEg14TER")) <> 0, 1, 0)
         
         
         If vFld(Rs("TipoAjuste")) > 0 Then
            Op_TAjuste(vFld(Rs("TipoAjuste"))).Value = 1
         Else
            Op_TAjuste(TAJUSTE_AMBOS).Value = 1 'valor por omisón
         End If
      Else
         Tx_GlosaCompTipo = DeSQL(Rs("Glosa"))
         Tx_Nombre = vFld(Rs("Nombre"), True)
         Tx_Descrip = vFld(Rs("Descrip"), True)
      End If
      
      
      Call SelItem(Cb_Tipo, vFld(Rs("Tipo")))
      Call SelItem(Cb_TipoCompTipo, vFld(Rs("Tipo")))
   
   Else ' *** PAM - 5 MAY 2005
      MsgBox1 "No se encontró el comprobante solicitado. Debe haber sido eliminado por otro usuario.", vbExclamation
      lidComp = 0
   End If
      
   Call CloseRs(Rs)
   
   If lidComp <= 0 Then  ' *** PAM - 5 MAY 2005
      Tx_Usuario = gUsuario.Nombre
      Exit Sub
   End If
   
   If lPrtResumido And Ch_ImpRes <> 0 Then
      Resumido = True
      lViewResumido = True
   End If
   
   Grid.FlxGrid.Redraw = False
   
   Q1 = "SELECT "
   If Resumido Then
      Q1 = Q1 & " 0, 0,"
   Else
      Q1 = Q1 & " IdMov, Orden, "
   End If
   Q1 = Q1 & lTblMovComprobante & ".IdCuenta, Cuentas.Codigo As CodCta, Cuentas.Nombre, Cuentas.Atrib" & ATRIB_ACTIVOFIJO & ",Cuentas.Atrib" & ATRIB_CONCILIACION & ","
   Q1 = Q1 & "Cuentas.Descripcion As DescCta, "
   If Resumido Then
      Q1 = Q1 & "Sum(" & lTblMovComprobante & ".Debe) as Debe , " & "Sum(" & lTblMovComprobante & ".Haber) as Haber,"
      Q1 = Q1 & " ' ', 0, 0, "
      Q1 = Q1 & " ' ', ' ' "
   Else
      Q1 = Q1 & lTblMovComprobante & ".Debe, " & lTblMovComprobante & ".Haber,"
      Q1 = Q1 & " Glosa," & lTblMovComprobante & ".IdAreaNeg," & lTblMovComprobante & ".IdCCosto, "
      Q1 = Q1 & " AreaNegocio.Descripcion As DescAreaNeg, CentroCosto.Descripcion As DescCCosto "
   End If
   If lCompTipo = False Then
      If Resumido Then
         Q1 = Q1 & ", ' ' "
         Q1 = Q1 & ", 0, 0, 0, 0, 0,"
         Q1 = Q1 & " 0, 0, ' ' "
      Else
         Q1 = Q1 & ", Entidades.Nombre as NombEnt "
         Q1 = Q1 & ", NumDoc, TipoLib, TipoDoc, NumDocHasta, " & lTblMovComprobante & ".IdDoc, "
         Q1 = Q1 & lTblMovComprobante & ".DeCentraliz, " & lTblMovComprobante & ".DePago, " & lTblMovComprobante & ".Nota "
      End If
   End If
   
   If lCompTipo = False Then
      Q1 = Q1 & " FROM ((((" & lTblMovComprobante
   Else
      Q1 = Q1 & " FROM (((" & lTblMovComprobante
   End If
   
   Q1 = Q1 & " INNER JOIN Cuentas ON " & lTblMovComprobante & ".IdCuenta = Cuentas.IdCuenta "
   If lCompTipo = False Then
      If gDbType = SQL_ACCESS Then
        Q1 = Q1 & "  AND " & lTblMovComprobante & ".IdEmpresa = Cuentas.IdEmpresa AND " & lTblMovComprobante & ".Ano = Cuentas.Ano)"
      Else
        Q1 = Q1 & "  AND " & lTblMovComprobante & ".IdEmpresa = Cuentas.IdEmpresa )"
      End If
   Else
      Q1 = Q1 & "  AND " & lTblMovComprobante & ".IdEmpresa = Cuentas.IdEmpresa )"
   End If
   Q1 = Q1 & " LEFT JOIN AreaNegocio ON " & lTblMovComprobante & ".IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & "  AND " & lTblMovComprobante & ".IdEmpresa = AreaNegocio.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN CentroCosto ON " & lTblMovComprobante & ".IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & "  AND " & lTblMovComprobante & ".IdEmpresa = CentroCosto.IdEmpresa )"
   If lCompTipo = False Then
      Q1 = Q1 & " LEFT JOIN Documento ON " & lTblMovComprobante & ".IdDoc=Documento.IdDoc "
      If lCompTipo = False Then
         Q1 = Q1 & " AND " & lTblMovComprobante & ".IdEmpresa = Documento.IdEmpresa AND " & lTblMovComprobante & ".Ano = Documento.Ano)"
      Else
         Q1 = Q1 & " AND " & lTblMovComprobante & ".IdEmpresa = Documento.IdEmpresa )"
      End If
      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " AND Entidades.IdEmpresa = Documento.IdEmpresa "
   End If
   
   Q1 = Q1 & " WHERE " & lTblMovComprobante & ".IdComp = " & lidComp
   If lCompTipo = False Then
      Q1 = Q1 & " AND " & lTblMovComprobante & ".IdEmpresa = " & gEmpresa.id & " AND " & lTblMovComprobante & ".Ano = " & gEmpresa.Ano
   Else
      Q1 = Q1 & " AND " & lTblMovComprobante & ".IdEmpresa = " & gEmpresa.id
   End If
   'If lOper <> O_EDIT Then
    'Q2 = Replace(Replace(Replace(Replace(Q1, "MovComprobante", "MovComprobanteFull"), " Documento ", " DocumentoFull "), "Documento.", " DocumentoFull."), "FullFull", "Full")
    'Q2 = Replace(Replace(Q1, "MovComprobante", "MovComprobanteFull"), " Documento ", " DocumentoFull ")
    'Q1 = Q1 & " UNION ALL " & Q2
   'End If
   If Resumido Then
      Q1 = Q1 & " GROUP BY "
      Q1 = Q1 & lTblMovComprobante & ".IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Atrib" & ATRIB_ACTIVOFIJO & ",Cuentas.Atrib" & ATRIB_CONCILIACION & ","
      Q1 = Q1 & "Cuentas.Descripcion "
   
   End If
   
   If Not Resumido Then
      Q1 = Q1 & " ORDER BY Orden, IdMov"
   End If
   
'   Q1 = "SELECT IdMov, Orden," & lTblMovComprobante & ".IdCuenta, Cuentas.Codigo As CodCta, Cuentas.Nombre, Cuentas.Atrib" & ATRIB_ACTIVOFIJO & ",Cuentas.Atrib" & ATRIB_CONCILIACION & ","
'   Q1 = Q1 & "Cuentas.Descripcion As DescCta, " & lTblMovComprobante & ".Debe, " & lTblMovComprobante & ".Haber,"
'   Q1 = Q1 & " Glosa," & lTblMovComprobante & ".IdAreaNeg," & lTblMovComprobante & ".IdCCosto, "
'   Q1 = Q1 & " AreaNegocio.Descripcion As DescAreaNeg, CentroCosto.Descripcion As DescCCosto "
'   If lCompTipo = False Then
'      Q1 = Q1 & ", Entidades.Nombre as NombEnt "
'      Q1 = Q1 & ", NumDoc, TipoLib, TipoDoc, NumDocHasta, " & lTblMovComprobante & ".IdDoc, "
'      Q1 = Q1 & lTblMovComprobante & ".DeCentraliz, " & lTblMovComprobante & ".DePago"
'   End If
'
'   If lCompTipo = False Then
'      Q1 = Q1 & " FROM ((((" & lTblMovComprobante
'   Else
'      Q1 = Q1 & " FROM (((" & lTblMovComprobante
'   End If
'
'   Q1 = Q1 & " INNER JOIN Cuentas ON " & lTblMovComprobante & ".IdCuenta = Cuentas.IdCuenta) "
'   Q1 = Q1 & " LEFT JOIN AreaNegocio ON " & lTblMovComprobante & ".IdAreaNeg = AreaNegocio.IdAreaNegocio) "
'   Q1 = Q1 & " LEFT JOIN CentroCosto ON " & lTblMovComprobante & ".IdCCosto = CentroCosto.IdCCosto) "
'   If lCompTipo = False Then
'      Q1 = Q1 & " LEFT JOIN Documento ON " & lTblMovComprobante & ".IdDoc=Documento.IdDoc) "
'      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
'   End If
'
'   Q1 = Q1 & " WHERE " & lTblMovComprobante & ".IdComp = " & lidComp
'   Q1 = Q1 & " ORDER BY Orden, IdMov"
 
 
'   If lTblMovComprobante = "MovComprobanteFull" Then
'    Q1 = Replace(Q1, "Documento", "DocumentoFull")
'   End If
   
   If Row = 0 Then
     If lCentrFull = 1 Then
      Q1 = Q1 & SqlPaging(gDbType, lClsPaging.CurReg - 1, gPageNumReg)
     End If
   End If

   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   n = 1
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If
          

   '3238739 pipe
      If Not Resumido Then
         Grid.TextMatrix(i, C_IDMOV) = vFld(Rs("IdMov"))
         Grid.TextMatrix(i, C_ORDEN) = vFld(Rs("Orden"))
         
         Grid.TextMatrix(i, C_ORDEN) = i - Grid.FixedRows + 1 + (lClsPaging.CurReg - 1)
            If Row = 0 Then
               lClsPaging.NumReg = vFmt(Grid.TextMatrix(i, C_ORDEN)) - (lClsPaging.CurReg - 1)
            End If
         
      End If
      
      'If lFromCentraliz Then
      '3387904
       If Resumido Then
         Grid.TextMatrix(i, C_ORDEN) = n
       End If
       '3387904
         Call FGrModRow(Grid, i, FGR_U, C_IDMOV, C_UPDATE)
         n = n + 1
      'End If
      
      
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(vFld(Rs("CodCta")))
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("DescCta"))
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), BL_NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), BL_NUMFMT)
      
      If InStr(gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas, "," & vFld(Rs("IdCuenta")) & ",") > 0 Then
         lCuentasDisponible = True
      End If
      
      If Not Resumido Then
         Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"), True)
         Grid.TextMatrix(i, C_AREANEG) = vFld(Rs("DescAreaNeg"), True)
         Grid.TextMatrix(i, C_IDAREANEG) = vFld(Rs("IdAreaNeg"))
         Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("DescCCosto"), True)
         Grid.TextMatrix(i, C_IDCCOSTO) = vFld(Rs("IdCCosto"))
      End If
      
      'Call FGrSetPicture(Grid, i, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
      Grid.TextMatrix(i, C_LSTCUENTA) = ">>"
      
      
      If Not Resumido Then
         
         If Not lCompTipo Then
            'Call FGrSetPicture(Grid, i, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)
            Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
            Grid.TextMatrix(i, C_DETALLE) = ">>"
            If vFld(Rs("IdDoc")) > 0 Then
               Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
               Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
               Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("NombEnt"), True)
               If Grid.TextMatrix(i, C_TIPODOC) = "VSD" Then
                  Grid.TextMatrix(i, C_NUMDOC) = ""
               End If
               If vFld(Rs("TipoLib")) = LIB_COMPRAS Or vFld(Rs("TipoLib")) = LIB_VENTAS Or vFld(Rs("TipoLib")) = LIB_RETEN Then
                  lHayDocLibros = True
               End If
            End If
            
            If vFld(Rs("Atrib" & ATRIB_ACTIVOFIJO)) <> 0 Then ' es cuenta de activo fijo
               Grid.TextMatrix(i, C_DETACTFIJO) = TX_ACTFIJO
            End If
            
            If vFld(Rs("Atrib" & ATRIB_CONCILIACION)) <> 0 Then ' es cuenta banco
               Grid.TextMatrix(i, C_ATRIB_CONCIL) = 1
            End If
            
            Grid.TextMatrix(i, C_DECENTRALIZ) = Abs(vFld(Rs("DeCentraliz")))
            Grid.TextMatrix(i, C_DEPAGO) = Abs(vFld(Rs("DePago")))
            
            If (vFld(Rs("DeCentraliz"))) <> 0 Then
               lTieneMovCentraliz = True
            End If
            
            If (vFld(Rs("DePago"))) <> 0 Then
               lTieneMovPago = True
            End If
         
            Grid.TextMatrix(i, C_NOTA) = vFld(Rs("Nota"))
            If vFld(Rs("Nota")) <> "" Then
               Grid.Row = i
               Grid.Col = C_CODCUENTA
               Grid.CellPictureAlignment = flexAlignRightTop
               Set Grid.CellPicture = FrmMain.Pc_Nota
            End If
         
         End If
         
      End If
     
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
   
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = 0
   Grid.ColSel = Grid.Col
   
   Grid.FlxGrid.Redraw = True

   Call CalcTot
   
   '3269719
   Call CalcTotFull(0, 0, 0, 0)
   '3269719
   
   If lHayDocLibros Then
      Ch_OtrosIngEg14TER.Value = 0
      Ch_OtrosIngEg14TER.Enabled = False
   End If
      
End Sub

Private Sub Form_Resize()
   Dim d As Long
   Dim H As Long

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - 2 * (Grid.Left + W.xFrame)
   If d > 1000 Then
      Grid.Width = d
   End If
   
   If Not gFunciones.ComprobanteResumido Then
      
      H = Fr_Header(0).Height - (Ch_ImpRes.Top + 60)
      
      Fr_Header(0).Height = Ch_ImpRes.Top + 60
      Fr_Header(1).Height = Fr_Header(0).Height
      Fr_SelCompTipo.Height = Fr_Header(0).Height
      
      Grid.Top = Grid.Top - H
   End If
   
   d = Me.Height - (Grid.Top - 100) - W.YCaption * 2 - GridTot.Height + 50
   If d > 1000 Then
      '3269719
      Grid.Height = d - 1600
      'Grid.Height = d
      '3269719
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   GridTot.Top = Grid.Height + Grid.Top
   GridTot.Width = Grid.Width - 300
   
   '3269719
   GridTotFull.Top = GridTot.Height + GridTot.Top + 60
   GridTotFull.Width = Grid.Width - 300
   
   fr_Paginacion.Top = GridTotFull.Top + GridTotFull.Height + 60
   fr_Paginacion.Width = Grid.Width - 300
   '3269719
   
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub Form_Unload(Cancel As Integer)

   If lRc <> vbOK And Bt_Cancel.Enabled = False Then
      
      Call Bt_OK_Click
      If lRc <> vbOK Then
         Cancel = True
      End If
   End If

   If lCompTipo Then
      Call UnLockAction(DbMain, LK_COMPTIPO, lidComp)
   Else
      Call UnLockAction(DbMain, LK_COMPROBANTE, lidComp)
   End If

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim IdCuenta As Long
   Dim Cod As String
   Dim DescCta As String
   Dim NombCta As String
   Dim UltimoNivel As Boolean
   Dim OldIdCuenta As Long
   Dim CodiF2214Ter As String
   
   '3269719
   Dim vValorAnterior As Double
   '3269719
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
         
'   ***** Ado 2699584 Tema 1 del 3.4
            If Col = 3 Or Col = C_LSTCUENTA Then
            CodiF2214Ter = ""
            Call CodF2214Ter(Cod, CodiF2214Ter)

            If CodiF2214Ter = "1" Then 'Or CodiF2214Ter = "629" Then
                If MsgBox1("¡Atención! ¿Partida Ingresada corresponde a una Participación?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    Exit Sub
                Else
                    Dim Frm As FrmPercepciones
                    Set Frm = New FrmPercepciones
                    Frm.CodCta = IdCuenta
                    Frm.GIdPerc = 0
                    Frm.Fecha = Tx_Fecha
                    Frm.orden = Grid.TextMatrix(Row, C_ORDEN)
                    Frm.idcomp = lidComp
                    Frm.Show vbModal
                    Set Frm = Nothing
                End If
            End If
            End If
'   ***** Fin Ado 2699584 Tema 1 del 3.4
            
            
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta   'se asigna porque se usa en GridActivoFijo
            Value = Format(Cod, gFmtCodigoCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta
            
            Call GridActivoFijo(IdCuenta, Row, Col)
                           
            Call GridAtribCuenta(Row)
            
         End If
         
         If Grid.Row = Grid.rows - 1 Then
            Grid.rows = Grid.rows + 1
         End If
         
      Case C_DEBE
   
         If vFmt(Value) < 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            If Grid.TextMatrix(Row, Col) <> "" Then
            vValorAnterior = Grid.TextMatrix(Row, Col)
            End If
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_HABER)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_HABER) = ""
            End If
            Call CalcTot
            '3269719
            Call CalcTotFull(Row, Col, vValorAnterior, 1)
            '3269719
            
         End If
         
      Case C_HABER
   
         If vFmt(Value) < 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            If Grid.TextMatrix(Row, Col) <> "" Then
              vValorAnterior = Grid.TextMatrix(Row, Col)
            End If
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_DEBE)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_DEBE) = ""
            End If
            Call CalcTot
            '3269719
             Call CalcTotFull(Row, Col, vValorAnterior, 2)
            '3269719
         End If
         
      Case C_GLOSA
         Value = Trim(Value)
         Value = ReplaceStr(Value, vbCr, "")
         Value = ReplaceStr(Value, vbLf, "")
      
      Case C_AREANEG
         If ItemData(Grid.CbList(C_AREANEG)) > 0 Then
            Grid.TextMatrix(Row, C_IDAREANEG) = ItemData(Grid.CbList(C_AREANEG))
         Else
            Grid.TextMatrix(Row, C_IDAREANEG) = 0
         End If
         
      Case C_CCOSTO
         If ItemData(Grid.CbList(C_CCOSTO)) > 0 Then
            Grid.TextMatrix(Row, C_IDCCOSTO) = ItemData(Grid.CbList(C_CCOSTO))
         Else
            Grid.TextMatrix(Row, C_IDCCOSTO) = 0
         End If
         
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
   End If

End Sub
Private Function CodF2214Ter(Cod As String, CodF22 As String) As Boolean
Dim Rs As Recordset
   Dim Q1 As String
    
    CodF2214Ter = False
    
    Q1 = "SELECT Percepcion FROM Cuentas "
    Q1 = Q1 & " WHERE Codigo = '" & ReplaceStr(Cod, "-", "") & "'"
    Q1 = Q1 & " AND Percepcion IS NOT NULL"
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
    Q1 = Q1 & " AND Ano = " & gEmpresa.Ano
    
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   'Do While Rs.EOF = False
      CodF2214Ter = True
      CodF22 = vFld(Rs("Percepcion"))
      
      If Trim(CodF22) = "0" Then
         Q1 = "SELECT CodF22_14Ter FROM Cuentas "
         Q1 = Q1 & " WHERE Codigo = '" & ReplaceStr(Cod, "-", "") & "'"
         Q1 = Q1 & " AND CodF22_14Ter IS NOT NULL"
        
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
           CodF2214Ter = True
           CodF22 = vFld(Rs("CodF22_14Ter"))
         End If
         Call CloseRs(Rs)
      End If
    'Rs.MoveNext
   'Loop
   Else
      Q1 = "SELECT CodF22_14Ter FROM Cuentas "
      Q1 = Q1 & " WHERE Codigo = '" & ReplaceStr(Cod, "-", "") & "'"
      Q1 = Q1 & " AND CodF22_14Ter IS NOT NULL"
     
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
        CodF2214Ter = True
        CodF22 = vFld(Rs("CodF22_14Ter"))
      End If
      Call CloseRs(Rs)
   End If
Call CloseRs(Rs)
End Function


Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
   Dim orden As Integer
   Dim IdCuenta As Long
   Dim DescCta As String
   Dim NombCuenta As String
   Dim FrmPlan As FrmPlanCuentas
   Dim ValPrevLine As Boolean
   Dim Msg As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdDoc As Long
   Dim idMov As Long
   Dim CodCta As String
   Dim GrDobleClick As Boolean
   
   GrDobleClick = lGrDobleClick
   lGrDobleClick = False
   
   'si no está habilitado para modificar, nos vamos. Dejamos sólo pasar a ver el detalle del movimiento
   If (Not Tx_Glosa.Enabled And Col <> C_DETALLE) Or Row = 0 Then
      Exit Sub
   End If
               
   idMov = Val(Grid.TextMatrix(Row, C_IDMOV))
   orden = Val(Grid.TextMatrix(Row, C_ORDEN))
      
   'Linea anterior tiene valor o está eliminada?
   'If lCompTipo Then
      ValPrevLine = (Row > Grid.FixedRows And Val(Grid.TextMatrix(Row - 1, C_ORDEN)) > 0 And Trim(Grid.TextMatrix(Row - 1, C_CUENTA)) <> "") Or Grid.RowHeight(Row - 1) = 0
   'Else
   '   ValPrevLine = (Row > Grid.FixedRows And Val(Grid.TextMatrix(Row - 1, C_ORDEN)) > 0 And Trim(Grid.TextMatrix(Row - 1, C_CUENTA)) <> "" And (Val(Grid.TextMatrix(Row - 1, C_DEBE)) > 0 Or Val(Grid.TextMatrix(Row - 1, C_HABER)) > 0)) Or Grid.RowHeight(Row - 1) = 0
   'End If
   
   If Not (Row = Grid.FixedRows Or (Row > Grid.FixedRows And orden > 0) Or ValPrevLine) Then
      Exit Sub
   End If
   
   'If Col = C_DETALLE And vFmt(Grid.TextMatrix(Row, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(Row, C_HABER)) = 0 Then
   '   Exit Sub
   'End If
   
   'si no hay cuenta seleccionada, sólo puede ingresar cuenta o asignar un doc o ver detalle de doc asignado
   If Col <> C_DETALLE And Col <> C_NUMDOC And Col <> C_ENTIDAD And Col <> C_TIPODOC And Col <> C_CODCUENTA And Col <> C_LSTCUENTA And Col <> C_CUENTA And Grid.TextMatrix(Row, C_CUENTA) = "" Then
      Exit Sub
   End If
   
   If orden = 0 Then    'nuevo
      
      '3269719
      If lCentrFull = 1 Then
      Call AddNewMov(Row - 1)
      Else
      Call AddNewMov(Row)
      End If
      
      'Call AddNewMov(Row)
      '3269719
      
      
      
     
      
   End If
   
   Select Case Col
   
      Case C_CODCUENTA
            
         Grid.TxBox.MaxLength = 20
         EdType = FEG_Edit
         
      'Case C_CUENTA, C_LSTCUENTA
      
      Case C_LSTCUENTA
      
         If GrDobleClick Then
         
            Call AsignaCuenta(Row, Col)
            
         End If
         
      Case C_DEBE
         EdType = FEG_Edit
         
      Case C_HABER
         EdType = FEG_Edit
         
      Case C_GLOSA
         Grid.TxBox.MaxLength = 50
         EdType = FEG_Edit
      
      Case C_DETALLE, C_NUMDOC, C_ENTIDAD, C_TIPODOC
         If lCompTipo = False Then
            
            IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
            
            If IdDoc = 0 Then
            
               Call Bt_BuscarDoc_Click
            Else
            
               Call Bt_DetMov_Click
               
            End If
         End If
         
      Case C_AREANEG
         If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_AREANEG) <> 0 Then
            EdType = FEG_List
         End If
            
      Case C_CCOSTO
         If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CCOSTO) <> 0 Then
            EdType = FEG_List
         End If
         
      Case C_DETACTFIJO
         If GrDobleClick And Grid.TextMatrix(Row, C_DETACTFIJO) <> "" Then
            Call PostClick(Bt_ActivoFijo)
         End If

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
   Dim Frm As FrmDoc
   Dim IdDoc As Long
   Dim TipoDoc As Integer
   
   Col = Grid.MouseCol
   Row = Grid.MouseRow
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   lGrDobleClick = True
   
   If Grid.Locked Then   'para poder ver el detalle de un documento aún cuando no se puedan hacer modificaciones en el comprobante
   
      If lCompTipo = True Then
         Exit Sub
      End If
      
      If Grid.Col = C_DETALLE Or Grid.Col = C_NUMDOC Or Grid.Col = C_ENTIDAD Or Grid.Col = C_TIPODOC Then

         IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
         
         If IdDoc = 0 Then
            Exit Sub
         Else
            TipoDoc = 0
            If InStr(Grid.TextMatrix(Grid.Row, C_TIPODOC), "ODF") > 0 Then
                TipoDoc = 8
            End If
            Set Frm = New FrmDoc
            Call Frm.FView(IdDoc, TipoDoc)
            Set Frm = Nothing
         End If
      End If
      
   End If
   

'   If Col = C_DETACTFIJO And Grid.TextMatrix(Row, C_DETACTFIJO) <> "" Then
'      Call PostClick(Bt_ActivoFijo)
'
'   ElseIf Col = C_LSTCUENTA Then
'
'      Set FrmPlan = New FrmPlanCuentas
'
'      If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta) = vbOK Then
'         If DescCta <> "" Then
'            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
'            Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
'            Grid.TextMatrix(Row, C_CUENTA) = DescCta
'
'            Call GridActivoFijo(IdCuenta, Row, Col)
'
'            Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
'
'        End If
'
'      End If
'      Set FrmPlan = Nothing
'
'   End If
   
End Sub

Private Sub SaveAll()
   Dim Rs As Recordset
   Dim Rc As Long, RNum As Long
   Dim Q1 As String
   Dim sWhere As String, WhConWhere As String, WhConAnd As String
   Dim AddUniqueRecord As Boolean
   Dim MesActual As Integer
   Dim Estado As Integer
   Dim TipoAjuste As Integer, WhTAjuste As String
   Dim i As Integer
   Dim EsNuevo As Boolean
   Dim Frm As FrmMsgConBreak
   Dim Msg As String
   Dim FldArray(7) As AdvTbAddNew_t
   Dim FldArrayT(1) As AdvTbAddNew_t
    
   For i = 1 To N_TIPOAJUSTE
      If Op_TAjuste(i) <> 0 Then
         TipoAjuste = i
         Exit For
      End If
   Next i

   If lidComp = 0 Then      'nuevo comprobante, lo agregamos
'      Set Rs = DbMain.OpenRecordset(lTblComprobante)
'      Rs.AddNew
'
'      lidComp = Rs("IdComp")
'      If lCompTipo Then
'         Rs.Fields("Tipo") = Cb_Tipo.ItemData(Cb_TipoCompTipo.ListIndex)
'      Else
'         Rs.Fields("Tipo") = Cb_Tipo.ItemData(Cb_Tipo.ListIndex)
'      End If

'      If lCompTipo Then
'         lidComp = AdvTbAddNew(DbMain, lTblComprobante, "IdComp", "Tipo", Cb_Tipo.ItemData(Cb_TipoCompTipo.ListIndex))
'      Else
'         lidComp = AdvTbAddNew(DbMain, lTblComprobante, "IdComp", "Tipo", Cb_Tipo.ItemData(Cb_Tipo.ListIndex))
'      End If
      
'      If Not lCompTipo Then
'         Rs.Fields("Fecha") = GetTxDate(Tx_Fecha)
'         Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'         Rs.Fields("FechaCreacion") = CLng(Int(Now))
'         Rs.Fields("TipoAjuste") = TipoAjuste
'         Rs.Fields("Correlativo") = -1
'      End If
'
'      Rs.Update
'      Rs.Close
'      Set Rs = Nothing
      
'      Q1 = "UPDATE Comprobante SET "
'
'      If Not lCompTipo Then
'         Q1 = Q1 & "  Fecha = " & GetTxDate(Tx_Fecha)
'         Q1 = Q1 & ", IdUsuario = " & gUsuario.IdUsuario
'         Q1 = Q1 & ", FechaCreacion = " & CLng(Int(Now))
'         Q1 = Q1 & ", TipoAjuste = " & TipoAjuste
'         Q1 = Q1 & ", Correlativo = -1 "
'         Q1 = Q1 & ", IdEmpresa =  " & gEmpresa.id
'         Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'      Else
'         Q1 = Q1 & "  IdEmpresa =  " & gEmpresa.id
'      End If
'
'      Q1 = Q1 & " WHERE IdComp = " & lidComp
'      Call ExecSQL(DbMain, Q1)
      
      
      If lCompTipo Then
         
         FldArrayT(0).FldName = "IdEmpresa"
         FldArrayT(0).FldValue = gEmpresa.id
         FldArrayT(0).FldIsNum = True
                     
         FldArrayT(1).FldName = "Tipo"
         FldArrayT(1).FldValue = Cb_Tipo.ItemData(Cb_TipoCompTipo.ListIndex)
         FldArrayT(1).FldIsNum = True
        
         lidComp = AdvTbAddNewMult(DbMain, lTblComprobante, "IdComp", FldArrayT)
      
      Else
      
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
         
         FldArray(4).FldName = "Fecha"
         FldArray(4).FldValue = GetTxDate(Tx_Fecha)
         FldArray(4).FldIsNum = True
               
         FldArray(5).FldName = "TipoAjuste"
         FldArray(5).FldValue = TipoAjuste
         FldArray(5).FldIsNum = True
               
         FldArray(6).FldName = "Tipo"
         FldArray(6).FldValue = Cb_Tipo.ItemData(Cb_Tipo.ListIndex)
         FldArray(6).FldIsNum = True
               
         FldArray(7).FldName = "Correlativo"
         FldArray(7).FldValue = -1
         FldArray(7).FldIsNum = True
            
         lidComp = AdvTbAddNewMult(DbMain, lTblComprobante, "IdComp", FldArray)
      End If
      
         
      lCorrelativo = 0
      EsNuevo = True
   End If
   
   MesActual = month(GetTxDate(Tx_Fecha))
   
   If lCorrelativo = 0 Then
   
      If lCompTipo Then
      'If (gTipoCorrComp = TCC_UNICO And (gPerCorrComp = TCC_ANUAL Or gPerCorrComp = TCC_CONTINUO)) Or lCompTipo Then
         lCorrelativo = lidComp
         
         Q1 = "UPDATE " & lTblComprobante & " SET Correlativo=" & lCorrelativo
         Q1 = Q1 & " WHERE IdComp=" & lidComp
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Rc = ExecSQL(DbMain, Q1)
         
      Else
      
         If gTipoCorrComp = TCC_UNICO Then
                     
            If gPerCorrComp = TCC_MENSUAL Then   'si es anual o continuo sWhere = ""
               sWhere = SqlMonthLng("Fecha") & " = " & MesActual
            End If
            
         ElseIf gTipoCorrComp = TCC_TIPOCOMP Then
            sWhere = " Tipo = " & ItemData(Cb_Tipo)
            
            If gPerCorrComp = TCC_MENSUAL Then
               sWhere = sWhere & " AND " & SqlMonthLng("Fecha") & " = " & MesActual    'SQL Server tiene los días desplazados en dos
            End If
            
         End If
         
         'agregamos el tipo de ajuste
         If TipoAjuste = TAJUSTE_TRIBUTARIO Then
            WhTAjuste = " TipoAjuste = " & TAJUSTE_TRIBUTARIO
         Else
            WhTAjuste = " TipoAjuste IN ( " & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
         End If
         
         If sWhere <> "" Then
            sWhere = sWhere & " AND " & WhTAjuste
         Else
            sWhere = WhTAjuste
         End If
                  
         If sWhere <> "" Then
            WhConWhere = " WHERE " & sWhere & " AND Correlativo > 0"
            WhConAnd = " AND " & sWhere  ' sin > 0
         Else
            WhConWhere = " WHERE Correlativo > 0"
            
         End If
         
         Do
            Q1 = "SELECT Max(Correlativo) as N FROM " & lTblComprobante & WhConWhere
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then
               lCorrelativo = vFld(Rs("N")) + 1
            Else
               lCorrelativo = 1
            End If
            Call CloseRs(Rs)
                     
            Q1 = "UPDATE " & lTblComprobante & " SET Correlativo=" & lCorrelativo
            Q1 = Q1 & " WHERE IdComp=" & lidComp
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Rc = ExecSQL(DbMain, Q1)
         
            DoEvents    'produce cosas raras
            
            Q1 = "SELECT Correlativo, idComp FROM " & lTblComprobante
            Q1 = Q1 & " WHERE Correlativo = " & lCorrelativo & WhConAnd
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            
            Set Rs = OpenRs(DbMain, Q1)
            If Rs.EOF = False Then
               AddUniqueRecord = True
            
               Do Until Rs.EOF
                  If vFld(Rs("idComp")) <> lidComp Then
                     AddUniqueRecord = False
                     Exit Do
                  End If
                  Rs.MoveNext
               Loop
               Call CloseRs(Rs)
               
               If AddUniqueRecord = False Then
                  Q1 = "UPDATE " & lTblComprobante & " SET Correlativo=-1"
                  Q1 = Q1 & " WHERE IdComp=" & lidComp
                  Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                  Rc = ExecSQL(DbMain, Q1)
               Else
                  Exit Do ' tenemos el correlativo
               End If
               
            Else
               Call CloseRs(Rs)
            End If
                           
         Loop
               
      End If
      
      Tx_Correlativo = lCorrelativo
      
   End If
   
   If lidComp <> 0 Then
   
      If lCompTipo = False Then
         'actualizamos el encabezado
         
         Estado = IIf(CbItemData(Cb_Estado) >= 0, CbItemData(Cb_Estado), EC_PENDIENTE)
         '3238739
         If lCentrFull = 1 Then
            Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante WHERE Idcomp = " & lidComp
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            Dim TotDebe As Long
            Dim TotHaber As Long
            
            If Rs.EOF = False Then
               TotDebe = vFld(Rs("TotDebe"))
               TotHaber = vFld(Rs("TotHaber"))
            End If
            
            Call CloseRs(Rs)
        End If
        '3238739
         
         Q1 = "UPDATE " & lTblComprobante & " SET "
         Q1 = Q1 & "  Fecha = " & GetTxDate(Tx_Fecha)
         Q1 = Q1 & ", Tipo = " & CbItemData(Cb_Tipo)
         Q1 = Q1 & ", Estado = " & Estado
          '2971346
         'Q1 = Q1 & ", Glosa = '" & ParaSQL(RemoveNoPrtChars(Tx_Glosa)) & "'"
         Q1 = Q1 & ", Glosa = '" & Trim(ParaSQL(RemoveNoPrtChars(Tx_Glosa))) & "'"
         '2971346
         Q1 = Q1 & ", ImpResumido = " & IIf(Ch_ImpRes <> 0, 1, 0)
         
         '3238739
         If lCentrFull = 1 Then
         Q1 = Q1 & ", TotalDebe = " & vFmt(TotDebe)
         Q1 = Q1 & ", TotalHaber = " & vFmt(TotHaber)
         Else
         Q1 = Q1 & ", TotalDebe = " & vFmt(GridTot.TextMatrix(0, C_DEBE))
         Q1 = Q1 & ", TotalHaber = " & vFmt(GridTot.TextMatrix(0, C_HABER))
         End If
         '3238739
         Q1 = Q1 & ", OtrosIngEg14TER = " & IIf(Ch_OtrosIngEg14TER <> 0, 1, 0)
         Q1 = Q1 & " WHERE IdComp = " & lidComp
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
      Else
         'actualizamos el Encabezado Comprobante Tipo
         Q1 = "UPDATE " & lTblComprobante & " SET "
         Q1 = Q1 & " Tipo = " & ItemData(Cb_TipoCompTipo)
         Q1 = Q1 & ", Descrip = '" & IIf(Trim(Tx_Descrip) <> "", ParaSQL(Tx_Descrip), ParaSQL(Tx_GlosaCompTipo)) & "'"
         Q1 = Q1 & ", Glosa = '" & IIf(Trim(Tx_GlosaCompTipo) <> "", ParaSQL(RemoveNoPrtChars(Tx_GlosaCompTipo)), ParaSQL(RemoveNoPrtChars(Tx_Glosa))) & "'"
         Q1 = Q1 & ", TotalDebe = " & vFmt(GridTot.TextMatrix(0, C_DEBE))
         Q1 = Q1 & ", TotalHaber = " & vFmt(GridTot.TextMatrix(0, C_HABER))
         Q1 = Q1 & ", Nombre ='" & ParaSQL(Tx_Nombre) & "'"
         Q1 = Q1 & " WHERE IdComp = " & lidComp
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      End If
      
      Call ExecSQL(DbMain, Q1)
      
      'Modificamos el IdComp en percepciones
      Q1 = "UPDATE  percepciones "
      Q1 = Q1 & " SET idcomp = " & lidComp
      Q1 = Q1 & " where idcomp is null"
      Q1 = Q1 & " AND IDEMPRESA = " & gEmpresa.id
      Q1 = Q1 & " AND ANO = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)
      
      'generamos log
      If lCompTipo = False Then
         Call AddLogComprobantes(lidComp, gUsuario.IdUsuario, lOper, Now, Estado, lCorrelativo, GetTxDate(Tx_Fecha), CbItemData(Cb_Tipo), Estado, TipoAjuste)
         If EsNuevo Then
            Set Frm = New FrmMsgConBreak
            Msg = "Se ha creado un Nuevo Comprobante con Correlativo " & Left(Cb_Tipo, 1) & "-" & Tx_Correlativo
            Call Frm.FView(Msg, "NoDispMsgNewComp")
            Set Frm = Nothing
         End If
      End If
      
   End If
   
   If lCentrFull = 1 Then
    Call InsertComprCentraFull(lidComp)
   End If
      
   Call SaveMovs
      
End Sub
Private Sub SaveMovs()
   Dim i As Integer
   Dim Lin As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim idMov As Long
   Dim StrIdDoc As String
   Dim StrIdDocSel As String
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim FldArray(2) As AdvTbAddNew_t
   Dim FldArrayT(1) As AdvTbAddNew_t
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim Rc As Long
   
   SumDebe = vFmt(GridTot.TextMatrix(0, C_DEBE))
   SumHaber = vFmt(GridTot.TextMatrix(0, C_HABER))


   If lidComp <= 0 Then
      Exit Sub
   End If

   Lin = Grid.FixedRows
   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If
            
      If lCompTipo = False And vFmt(Grid.TextMatrix(i, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(i, C_HABER)) = 0 And Grid.TextMatrix(i, C_DETACTFIJO) = "" And SumDebe <> 0 Then   'Ya se validó que SumDebe = SumHaber
         'eliminamos movs. sin valor, sabiendo que el total no es cero
         If Grid.TextMatrix(i, C_IDMOV) <> "" Then
'            Q1 = "DELETE FROM " & lTblMovComprobante & " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV)
'            Call ExecSQL(DbMain, Q1)
            Call DeleteSQL(DbMain, lTblMovComprobante, " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
         End If
      
      Else
      
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
'            Set Rs = DbMain.OpenRecordset(lTblMovComprobante)
'            Rs.AddNew
'
'            idMov = Rs("IdMov")
'
'            Rs.Update
'            Rs.Close
'            Set Rs = Nothing
            

            FldArray(0).FldName = "IdComp"
            FldArray(0).FldValue = lidComp
            FldArray(0).FldIsNum = True
                              
            FldArray(1).FldName = "IdEmpresa"
            FldArray(1).FldValue = gEmpresa.id
            FldArray(1).FldIsNum = True
                        
            If lCompTipo = False Then
               FldArray(2).FldName = "Ano"
               FldArray(2).FldValue = gEmpresa.Ano
               FldArray(2).FldIsNum = True
            Else
               FldArray(2).FldName = "CodCuenta"
               FldArray(2).FldValue = VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA))
               FldArray(2).FldIsNum = False
            End If
      
            idMov = AdvTbAddNewMult(DbMain, lTblMovComprobante, "IdMov", FldArray)
            
            Grid.TextMatrix(i, C_IDMOV) = idMov
            Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
            
'            If lCompTipo = False Then
'               Q1 = "UPDATE " & lTblMovComprobante & " SET "
'               Q1 = Q1 & "  IdEmpresa = " & gEmpresa.id
'               Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'               Q1 = Q1 & " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV)
'            Else
'               Q1 = "UPDATE " & lTblMovComprobante & " SET "
'               Q1 = Q1 & " IdEmpresa = " & gEmpresa.id
'               Q1 = Q1 & " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV)
'            End If
'
'            Call ExecSQL(DbMain, Q1)
            
            
         End If
         
         If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'            Q1 = "DELETE FROM " & lTblMovComprobante & " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV)
'            Call ExecSQL(DbMain, Q1)
            If lCompTipo Then
               Call DeleteSQL(DbMain, lTblMovComprobante, " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV) & " AND IdEmpresa = " & gEmpresa.id)
            Else
               Call DeleteSQL(DbMain, lTblMovComprobante, " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
            End If
            
            'soltamos docs si es centralización o pago
            If vFmt(Grid.TextMatrix(i, C_IDDOC)) <> 0 And (Val(Grid.TextMatrix(i, C_DECENTRALIZ)) <> 0 Or Val(Grid.TextMatrix(i, C_DEPAGO)) <> 0) Then
               Call SoltarDocs(vFmt(Grid.TextMatrix(i, C_IDDOC)))
            End If
               
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Or Val(Grid.TextMatrix(i, C_ORDEN)) <> Lin Then 'Update
            
            Call DesConciliarMov(Val(Grid.TextMatrix(i, C_IDMOV))) ' pam - 18 oct 2010
            
            Q1 = "UPDATE " & lTblMovComprobante & " SET "
            Q1 = Q1 & "  IdComp = " & lidComp
            '3238739
            'Q1 = Q1 & ", Orden = " & Lin
             Q1 = Q1 & ", Orden = " & i - Grid.FixedRows + 1 + (lClsPaging.CurReg - 1)
            '3238739
            Q1 = Q1 & ", IdCuenta = " & Grid.TextMatrix(i, C_IDCUENTA)
            Q1 = Q1 & ", Debe = " & vFmt(Grid.TextMatrix(i, C_DEBE))
            Q1 = Q1 & ", Haber = " & vFmt(Grid.TextMatrix(i, C_HABER))
            Q1 = Q1 & ", Glosa = '" & Left(ParaSQL(RemoveNoPrtChars(Grid.TextMatrix(i, C_GLOSA), True)), 50) & "'"
            Q1 = Q1 & ", IdAreaNeg = " & vFmt(Grid.TextMatrix(i, C_IDAREANEG))
            Q1 = Q1 & ", IdCCosto = " & vFmt(Grid.TextMatrix(i, C_IDCCOSTO))
            
            If lCompTipo Then
            Q1 = Q1 & ", CodCuenta = " & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA))
            End If
            
            If lCompTipo = False Then
               Q1 = Q1 & ", DePago = " & Int(Val(Grid.TextMatrix(i, C_DEPAGO)) <> 0)
               Q1 = Q1 & ", IdDoc =" & vFmt(Grid.TextMatrix(i, C_IDDOC))
               Q1 = Q1 & ", IdDocCuota =" & vFmt(Grid.TextMatrix(i, C_IDDOCCUOTA))
               Q1 = Q1 & ", Nota = '" & ParaSQL(Left(Grid.TextMatrix(i, C_NOTA), 120)) & "'"
            End If
            
            Q1 = Q1 & " WHERE IdMov = " & Grid.TextMatrix(i, C_IDMOV)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            
            If lCompTipo = False Then
               Q1 = Q1 & " AND Ano = " & gEmpresa.Ano
            End If
            
            Call ExecSQL(DbMain, Q1)
            
            If vFmt(Grid.TextMatrix(i, C_IDDOCCUOTA)) > 0 Then
               Q1 = "UPDATE DocCuotas SET Estado = " & ED_PAGADO & ", IdCompPago = " & lidComp
               Q1 = Q1 & " WHERE IdDocCuota = " & vFmt(Grid.TextMatrix(i, C_IDDOCCUOTA))
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
               Call ExecSQL(DbMain, Q1)
            End If
            
         End If
         
         If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
            Lin = Lin + 1
         End If
      
      End If
      
      
        If InStr(Grid.TextMatrix(i, C_TIPODOC), "ODF") > 0 Then
        
            Q1 = "Select IdCtaBanco,Tratamiento From Documento "
            Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
            Q1 = Q1 & " And Idempresa= " & gEmpresa.id
            Q1 = Q1 & " And Ano = " & gEmpresa.Ano
            
            Set Rs = OpenRs(DbMain, Q1)
            If Rs.EOF = False Then
                If vFld(Rs("IdCtaBanco")) = 0 Then
                
                    Q1 = "Update Documento "
                    Q1 = Q1 & " Set IdCtaBanco = " & Grid.TextMatrix(i, C_IDCUENTA)
                    Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
                    Rc = ExecSQL(DbMain, Q1)
                
                End If
            
            End If
            
            '616437 ffv
'            If Cb_Tipo.ItemData(Cb_Tipo.ListIndex) = TC_TRASPASO Then
'
'                    'lTieneMovCentraliz = True
'
'                    Q1 = "Update Documento "
'                    Q1 = Q1 & " Set estado = " & ED_CENTRALIZADO
'                    Q1 = Q1 & " ,IdCompCent = " & lidComp
'                    Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
'                    Q1 = Q1 & " And ano =" & gEmpresa.Ano
'                    Q1 = Q1 & " and idempresa = " & gEmpresa.id
'                    Rc = ExecSQL(DbMain, Q1)
'
'                    Q1 = "UPDATE " & lTblMovComprobante & " SET DeCentraliz = 1"
'                    Q1 = Q1 & " ,DePago = 0"
'                    Q1 = Q1 & " WHERE IdComp = " & lidComp & " AND DeCentraliz is null"
'                    Q1 = Q1 & " AND DePago = 1 "
'                    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'                    Q1 = Q1 & " AND idDoc = " & Grid.TextMatrix(i, C_IDDOC)
'                    Call ExecSQL(DbMain, Q1)
'
'             ElseIf (Cb_Tipo.ItemData(Cb_Tipo.ListIndex) = TC_INGRESO And vFld(Rs("tratamiento")) = 1) Then
'
'                    'lTieneMovPago = True
'
'                    If Val(Grid.TextMatrix(i, C_HABER)) > 0 Then
'
'                        Q1 = "Update Documento "
'                        Q1 = Q1 & " Set estado = " & ED_PAGADO
'                        Q1 = Q1 & " ,IdCompPago = " & lidComp
'                        Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
'                        Q1 = Q1 & " And ano =" & gEmpresa.Ano
'                        Q1 = Q1 & " and idempresa = " & gEmpresa.id
'                        Rc = ExecSQL(DbMain, Q1)
'
'                        Q1 = "UPDATE " & lTblMovComprobante & " SET DePago = 1, IdDoc =" & Grid.TextMatrix(i, C_IDDOC) & " , IdCartola = 0 "
'                        Q1 = Q1 & " WHERE IdComp = " & lidComp
'                        Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'                        Call ExecSQL(DbMain, Q1)
'                    Else
'                       ' MsgBox1 "Documento ODF tiene Tratamiento Activo, monto Haber debe ser mayor a 0 para realizar pago", vbExclamation
'                        'Grid.TextMatrix(i, C_DEBE) = 0
'                        'Exit Sub
'                    End If
'
'             ElseIf (Cb_Tipo.ItemData(Cb_Tipo.ListIndex) = TC_EGRESO And vFld(Rs("tratamiento")) = 2) Then
'
'                    'lTieneMovPago = True
'
'                    If Val(Grid.TextMatrix(i, C_DEBE)) > 0 Then
'
'                        Q1 = "Update Documento "
'                        Q1 = Q1 & " Set estado = " & ED_PAGADO
'                        Q1 = Q1 & " ,IdCompPago = " & lidComp
'                        Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
'                        Q1 = Q1 & " And ano =" & gEmpresa.Ano
'                        Q1 = Q1 & " and idempresa = " & gEmpresa.id
'                        Rc = ExecSQL(DbMain, Q1)
'
'                        Q1 = "UPDATE " & lTblMovComprobante & " SET DePago = 1, IdDoc =" & Grid.TextMatrix(i, C_IDDOC) & " , IdCartola = 0 "
'                        Q1 = Q1 & " WHERE IdComp = " & lidComp & " AND DePago = 0"
'                        Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'                        Call ExecSQL(DbMain, Q1)
'                    Else
'                       ' MsgBox1 "Documento ODF tiene Tratamiento Activo, monto Haber debe ser mayor a 0 para realizar pago", vbExclamation
'                        'Grid.TextMatrix(i, C_Haber) = 0
'                        'Exit Sub
'                    End If

             
            'End If
            '616437
           Call CloseRs(Rs)
           
        End If
      
   Next i

   If Not lCompTipo And (lTieneMovCentraliz Or lTieneMovPago) Then  'es de centralización o pago y es antiguo, no recién creado
      
      If ItemData(Cb_Estado) = EC_ANULADO And lEstado <> EC_ANULADO Then
         
         'acaban de anular el comprobante => debemos soltar los documentos asociados al comprobante
      
         'soltamos docs de centralización y pago, y cartola asociada (si hubiera)
         Q1 = "UPDATE " & lTblMovComprobante & " SET DeCentraliz = 0, DePago = 0, IdDoc = 0, IdCartola = 0 "
         Q1 = Q1 & " WHERE IdComp = " & lidComp & " AND (DeCentraliz <> 0 or DePago <> 0)"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         'soltamos cuotas de pago si corresponde
         Q1 = "UPDATE DocCuotas SET Estado = " & ED_PENDIENTE & ", IdCompPago = 0"
         Q1 = Q1 & " WHERE IdCompPago = " & lidComp & " AND Estado = " & ED_PAGADO
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         'desconciliamos los movimientos
         Tbl = " DetCartola "
         sFrom = " DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
         sFrom = sFrom & JoinEmpAno(gDbType, "DetCartola", "MovComprobante")
         sSet = " DetCartola.IdMov = 0 "
         sWhere = " WHERE MovComprobante.IdComp = " & lidComp
         sWhere = sWhere & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
         Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
         
                  
'         'los docs. que vienen de centralización los dejamos pendientes
'         Q1 = "UPDATE Documento SET IdCompCent = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE IdCompCent = " & lidComp
'         Call ExecSQL(DbMain, Q1)
'
'         'los docs. que vienen de pago autom.: dejamos en estado ED_CENTRALIZADO si tiene IdCompCent <> 0 o son del año anterior
'         Q1 = "UPDATE Documento SET IdCompPago = 0, Estado = " & ED_CENTRALIZADO & ", SaldoDoc = NULL WHERE IdCompPago = " & lidComp & " AND (IdCompCent <> 0 OR Year(FEmision) < " & gEmpresa.Ano & ")"
'         Call ExecSQL(DbMain, Q1)
'
'         'los docs. que vienen de pago autom.: dejamos pendientes si tiene IdCompCent = 0 (esto no debiera ocurrir nunca, pero por si las moscas)
'         Q1 = "UPDATE Documento SET IdCompPago = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE IdCompPago = " & lidComp & " AND IdCompCent = 0 "
'         Call ExecSQL(DbMain, Q1)
'
'         'ahora borramos referencia de mov docs
'         Q1 = "UPDATE MovDocumento SET IdCompCent = 0, IdCompPago = 0 WHERE IdCompCent = " & lidComp & " OR IdCompPago = " & lidComp
'         Call ExecSQL(DbMain, Q1)

         Call SoltarDocs
         
      End If
   End If
   
   If Not lCompTipo Then
   
      'movimientos de pago recién creados
      For i = Grid.FixedRows To Grid.rows - 1
      
         If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
            Exit For
         End If
      
         If Val(Grid.TextMatrix(i, C_DEPAGO)) = NEW_DEPAGO Then
            StrIdDoc = StrIdDoc & "," & Val(Grid.TextMatrix(i, C_IDDOC))
         
         ElseIf Val(Grid.TextMatrix(i, C_DEPAGO)) = NEW_DESELDOC Then
            StrIdDocSel = StrIdDocSel & "," & Val(Grid.TextMatrix(i, C_IDDOC))
         
         End If
      Next i
      
      If Trim(StrIdDoc) <> "" Then
         StrIdDoc = Mid(StrIdDoc, 2)
           
         'marcamos docs como pagados y que se recaulcule el saldo
         Q1 = "UPDATE Documento SET IdCompPago = " & lidComp & ", Estado = " & ED_PAGADO & ", SaldoDoc = NULL "
         Q1 = Q1 & " WHERE IdDoc IN(" & StrIdDoc & ")"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE MovDocumento SET IdCompPago = " & lidComp
         Q1 = Q1 & " WHERE IdDoc IN(" & StrIdDoc & ")"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
      
      If Trim(StrIdDocSel) <> "" Then
         StrIdDocSel = Mid(StrIdDocSel, 2)
           
         'marcamos docs para que recalcule saldo
         Q1 = "UPDATE Documento SET SaldoDoc = NULL WHERE IdDoc IN(" & StrIdDocSel & ")"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      If Trim(lDelDocLst) <> "" Then
         lDelDocLst = Mid(lDelDocLst, 2)
           
         'marcamos docs para que recalcule saldo
         Q1 = "UPDATE Documento SET SaldoDoc = NULL WHERE IdDoc IN(" & lDelDocLst & ")"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      'actualizamos el IdCtaBanco de documento para cada movimiento con cuenta banco
'      For i = Grid.FixedRows To Grid.rows - 1
'
'         If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
'            Exit For
'         End If
'
'         If Grid.TextMatrix(i, C_UPDATE) = FGR_U And Val(Grid.TextMatrix(i, C_ATRIB_CONCIL)) <> 0 Then    'es cuenta banco
'            Q1 = "UPDATE Documento SET IdCtaBanco=" & Val(Grid.TextMatrix(i, C_IDCUENTA)) & " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'            Call ExecSQL(DbMain, Q1)
'         End If
'      Next i
'
   End If
   
   Call AjustarPercepciones
   
End Sub

Private Sub AjustarPercepciones()
   Dim existe As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   existe = True
    Q1 = "SELECT ORDEN, IDCUENTA "
    Q1 = Q1 & " FROM  PERCEPCIONES "
    Q1 = Q1 & " WHERE IDCOMP = " & lidComp
    Q1 = Q1 & " AND IDEMPRESA = " & gEmpresa.id
    Q1 = Q1 & " AND ANO = " & gEmpresa.Ano
    Q1 = Q1 & " ORDER BY ORDEN "
    Set Rs = OpenRs(DbMain, Q1)
   Do While Rs.EOF = False
        For i = Grid.FixedRows To Grid.rows - 1
        
            If vFld(Rs("ORDEN")) = Grid.TextMatrix(i, C_ORDEN) And vFld(Rs("IDCUENTA")) = Grid.TextMatrix(i, C_IDCUENTA) Then
                existe = False
                Exit For
            End If
        
        Next i
        If existe Then
            Call DelPercepciones(vFld(Rs("ORDEN")), vFld(Rs("IDCUENTA")))
        End If
        existe = True
    Rs.MoveNext
    Loop
    Call CloseRs(Rs)


End Sub

Private Sub SoltarDocs(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim WhIdDoc As String
   
   If Not lCompTipo And (lTieneMovCentraliz Or lTieneMovPago) Then  'es de centralización o pago y es antiguo, no recién creado
   
      If IdDoc <> 0 Then
         WhIdDoc = " IdDoc = " & IdDoc & " AND "
      End If
   
      'los docs. que vienen de centralización los dejamos pendientes
      Q1 = "UPDATE Documento SET IdCompCent = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE " & WhIdDoc & " IdCompCent = " & lidComp
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      'los docs. que vienen de pago autom.: dejamos en estado ED_CENTRALIZADO si tiene IdCompCent <> 0 o son del año anterior
      Q1 = "UPDATE Documento SET IdCompPago = 0, Estado = " & ED_CENTRALIZADO & ", SaldoDoc = NULL WHERE " & WhIdDoc & "  IdCompPago = " & lidComp & " AND (IdCompCent <> 0 OR " & SqlYearLng("FEmision") & " < " & gEmpresa.Ano & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      'los docs. que vienen de pago autom.: dejamos pendientes si tiene IdCompCent = 0 (esto no debiera ocurrir nunca, pero por si las moscas)
      Q1 = "UPDATE Documento SET IdCompPago = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE " & WhIdDoc & "  IdCompPago = " & lidComp & " AND IdCompCent = 0 "
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      'ahora borramos referencia de mov docs
      Q1 = "UPDATE MovDocumento SET IdCompCent = 0, IdCompPago = 0 WHERE " & WhIdDoc & "(IdCompCent = " & lidComp & " OR IdCompPago = " & lidComp & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
   
   End If
   
End Sub



Private Function valida(Optional ByVal ValidaParaGenCompTipo As Boolean = False) As Boolean
   Dim i As Integer, j As Integer
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim RowCuadra As Integer
   Dim RegValCero As Integer
   Dim Diff As Double
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Q2 As String
   Dim Fecha As Long

   valida = False
   
   Fecha = GetTxDate(Tx_Fecha)
      
   If Not lCompTipo Then
   
      If Not ValidaIngresoComp(lOper = O_EDIT) Then
         Exit Function
      End If
      
      If gAppCode.Demo Then
         If lOper = O_NEW Then
            Q1 = "SELECT Count(*) FROM Comprobante "
            Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
         
            If Not Rs.EOF Then
               If vFld(Rs(0)) >= MAX_COMPDEMO Then
                  MsgBox1 "Ha superado la cantidad de comprobantes de la versión DEMO.", vbExclamation
                  Call CloseRs(Rs)
                  Exit Function
               End If
            End If
            Call CloseRs(Rs)
         End If
      End If

   
      If Not ValidaParaGenCompTipo Then
            
         If Trim(Tx_Fecha) = "" Then
            Call MsgBox1("Falta ingresar la fecha del comprobante.", vbOKOnly + vbExclamation)
            Exit Function
         End If
         
         If Year(Fecha) <> gEmpresa.Ano Then
            MsgBox1 "El año de la fecha del comprobante no corresponde al periodo actual.", vbExclamation
            Exit Function
         End If
         
'         6 oct 2017 - pam - se comenta
'         If gPerCorrComp = TCC_MENSUAL Then
'            If lOldFechaEmision > 0 And Month(Fecha) <> Month(lOldFechaEmision) Then
'               MsgBox1 "Dado que la numeración de los comprobantes es mensual, no puede cambiar el mes de emisión de este comprobante.", vbExclamation
'               Exit Function
'            End If
'         End If
         
         If GetEstadoMes(month(Fecha)) <> EM_ABIERTO Then
            MsgBox1 "El mes correspondiente a la fecha del comprobante no está abierto.", vbExclamation
            Exit Function
         End If
         
         If ItemData(Cb_Tipo) <> TC_APERTURA Then
            'Q1 = "SELECT Fecha FROM Comprobante WHERE Tipo=" & TC_APERTURA
            Q1 = "SELECT Fecha FROM " & lTblComprobante & " WHERE Tipo=" & TC_APERTURA
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
           
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then
               If vFld(Rs("Fecha")) > Fecha Then
                  Call MsgBox1("La fecha del comprobante es anterior a la fecha del Comprobante de Apertura.", vbOKOnly + vbExclamation)
                  Call CloseRs(Rs)
                  Exit Function
               End If
            End If
            
            Call CloseRs(Rs)
         End If
         
      End If
      
      If CbItemData(Cb_Estado) <= 0 Then
         MsgBox1 "Falta seleccionar el estado del comprobante.", vbExclamation + vbOKOnly
         Exit Function
      End If
      
      If ItemData(Cb_Tipo) <= 0 Then
         Call MsgBox1("Falta seleccionar el tipo de comprobante.", vbOKOnly + vbExclamation)
         Exit Function
      End If
   
      If Trim(Tx_Glosa) = "" Then
         Call MsgBox1("Falta ingresar la glosa del comprobante.", vbOKOnly + vbExclamation)
         Exit Function
      End If
            
      If ItemData(Cb_Estado) = EC_ANULADO And ChkPriv(PRV_ADM_COMP) = False Then
         MsgBox1 "Este usuario no tiene el perfil requerido para anular un comprobante.", vbExclamation + vbOKOnly
         Exit Function
      End If
         
   Else
   
      If ItemData(Cb_TipoCompTipo) <= 0 Then
         Call MsgBox1("Falta seleccionar el tipo de comprobante tipo.", vbOKOnly + vbExclamation)
         Exit Function
      End If
   
      If Trim(Tx_GlosaCompTipo) = "" Then
         Call MsgBox1("Falta ingresar la glosa del comprobante tipo.", vbOKOnly + vbExclamation)
         Exit Function
      End If
      
      If Trim(Tx_Nombre) = "" Then
         MsgBox1 "Ingrese nombre al comprobante tipo.", vbExclamation
         Exit Function
         
      ElseIf Trim(Tx_Descrip) = "" Then
         MsgBox1 "Ingrese descripción para comprobante tipo.", vbExclamation
         Exit Function
         
      ElseIf lOper <> O_EDIT Then
      
         Q1 = "SELECT IdComp FROM CT_Comprobante WHERE Nombre = '" & Trim(Tx_Nombre) & "'"
         If Not lCompTipo Then
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Else
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         End If
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then   'ya existe uno con este nombre
            MsgBox1 "Ya existe un Comprobante Tipo con este nombre.", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
         
         Call CloseRs(Rs)
         
      End If
            
   End If
   
   SumDebe = vFmt(GridTot.TextMatrix(0, C_DEBE))
   SumHaber = vFmt(GridTot.TextMatrix(0, C_HABER))
      
   If Not ValidaParaGenCompTipo Then
      If Not lCompTipo And ItemData(Cb_Tipo) <> TC_APERTURA And SumDebe = 0 And SumHaber = 0 Then
         If MsgBox1("¡Atención! El valor total del comprobante es cero." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
      End If
      
      'validamos comprobante de apertura nuevo
      If Not lCompTipo And ItemData(Cb_Tipo) = TC_APERTURA And lidComp = 0 Then
         
         'verificamos que no cree más de un comp. de apertura
         Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo=" & TC_APERTURA
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         'Feña
         'Q2 = Replace(Q1, "Comprobante", "ComprobanteFull")
         'Q1 = Q1 & " UNION ALL " & Q2
         'FIN Feña
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then                          'ya hay un comp de apertura
            Call MsgBox1("Ya existe un comprobante de apertura. No es posible ingresar otro.", vbOKOnly + vbExclamation)
            Call CloseRs(Rs)
            Exit Function
         End If
         
         'verificamos que no haya un comprobante anterior a este comp. de apertura
         Q1 = "SELECT IdComp, Fecha FROM Comprobante WHERE Tipo<>" & TC_APERTURA
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Q1 = Q1 & " ORDER BY Fecha, IdComp"
         'Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then                          'ya hay un comp de apertura
            Call MsgBox1("Ya existe un comprobante cuya fecha es anterior a la fecha de este comprobante de apertura.", vbOKOnly + vbExclamation)
            'Call CloseRs(Rs)
            Exit Function
         End If
         Call CloseRs(Rs)
      End If
      
   End If
   
   If lFromCentraliz And lMesAnoCentraliz > 0 Then
   
      If month(Fecha) <> month(lMesAnoCentraliz) Or Year(Fecha) <> Year(lMesAnoCentraliz) Then
         If MsgBox1("Los documentos serán centralizados en un mes distinto al del ingreso de éstos." & vbCrLf & vbCrLf & "¿Está seguro que desea continuar?", vbYesNo + vbQuestion) = vbNo Then
            Exit Function
         End If
      End If
   End If
      
   lCuentasDisponible = False
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then
      
         If Val(Grid.TextMatrix(i, C_IDCUENTA)) = 0 Then
            MsgBox1 "En la línea " & i & " falta ingresar la cuenta.", vbExclamation + vbOKOnly
            Exit Function
         End If
         
         If vFmt(Grid.TextMatrix(i, C_DEBE)) <> 0 And vFmt(Grid.TextMatrix(i, C_HABER)) <> 0 Then
            MsgBox1 "En la línea " & i & " los valores en las columnas DEBE y HABER son ambos mayores que 0.", vbExclamation + vbOKOnly
            Exit Function
         End If
         
         If Not ValidaParaGenCompTipo Then
         
            If vFmt(Grid.TextMatrix(i, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(i, C_HABER)) = 0 Then
               RowCuadra = i
               RegValCero = RegValCero + 1
            End If
            
            If Not lCompTipo And ItemData(Cb_Tipo) <> TC_APERTURA And (vFmt(Grid.TextMatrix(i, C_DEBE)) <> 0 Or vFmt(Grid.TextMatrix(i, C_HABER)) <> 0) Then
               If GetAtribCuenta(Val(Grid.TextMatrix(i, C_IDCUENTA)), ATRIB_RUT) <> 0 And Trim(Grid.TextMatrix(i, C_ENTIDAD)) = "" Then 'And Val(Grid.TextMatrix(Grid.Row, C_DECENTRALIZ)) = 0 Then
                  'debe tener doc asociado y no proviene de centralización ni proviene de pago automático
                  MsgBox1 "En la línea " & i & " falta ingresar el detalle de documento asociado y entidad correspondiente.", vbExclamation + vbOKOnly
                  Exit Function
               End If
                  
               If GetAtribCuenta(Val(Grid.TextMatrix(i, C_IDCUENTA)), ATRIB_AREANEG) <> 0 And Val(Grid.TextMatrix(i, C_IDAREANEG)) = 0 Then
                  MsgBox1 "En la línea " & i & " falta ingresar el Área de Negocio.", vbExclamation + vbOKOnly
                  Exit Function
               End If
                  
               If GetAtribCuenta(Val(Grid.TextMatrix(i, C_IDCUENTA)), ATRIB_CCOSTO) <> 0 And Val(Grid.TextMatrix(i, C_IDCCOSTO)) = 0 Then
                  MsgBox1 "En la línea " & i & " falta ingresar el Centro de Gestión.", vbExclamation + vbOKOnly
                  Exit Function
               End If
               
               If Val(Grid.TextMatrix(i, C_ATRIB_CONCIL)) <> 0 And Val(Grid.TextMatrix(i, C_IDDOC)) <> 0 Then
                  If Not ValidIdCtaBanco(i) Then
                     MsgBox1 "En la línea " & i & " la cuenta no es igual a la cuenta bancaria definida en el documento.", vbExclamation + vbOKOnly
                     Exit Function
                  End If
               End If
               
            End If
         End If
         
         If Grid.TextMatrix(i, C_GLOSA) = "" Then
            Grid.TextMatrix(i, C_GLOSA) = Tx_Glosa
         End If
         
         If Ch_OtrosIngEg14TER <> 0 Then
            If InStr(gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas, "," & Grid.TextMatrix(i, C_IDCUENTA) & ",") > 0 Then
               lCuentasDisponible = True
            End If
         End If
         
         '616437 ffv
'          If InStr(Grid.TextMatrix(i, C_TIPODOC), "ODF") > 0 Then
'
'            Q1 = "Select Tratamiento From Documento "
'            Q1 = Q1 & " Where TipoLib = " & LIB_OTROFULL & " And IdDoc = " & Grid.TextMatrix(i, C_IDDOC) & "  And NumDoc = '" & Grid.TextMatrix(i, C_NUMDOC) & "' "
'            Q1 = Q1 & " And Idempresa= " & gEmpresa.id
'            Q1 = Q1 & " And Ano = " & gEmpresa.Ano
'            Set Rs = OpenRs(DbMain, Q1)
'
'            Dim TratamientoODF As Integer
'
'            If Rs.EOF = False Then
'             TratamientoODF = vFld(Rs("tratamiento"))
'            End If
'
'
'             If (Cb_Tipo.ItemData(Cb_Tipo.ListIndex) = TC_INGRESO And TratamientoODF = 1) Then
'
'                    'lTieneMovPago = True
'
'                    If Val(Grid.TextMatrix(i, C_DEBE)) > 0 Then
'
'                        MsgBox1 "Documento ODF tiene Tratamiento Activo, monto Haber debe ser mayor a 0 para realizar pago", vbExclamation
'
'                        Grid.TextMatrix(i, C_DEBE) = ""
'                         Call CalcTot
'
'                        Exit Function
'                    End If
'
'             ElseIf (Cb_Tipo.ItemData(Cb_Tipo.ListIndex) = TC_EGRESO And TratamientoODF = 2) Then
'
'                    'lTieneMovPago = True
'
'                    If Val(Grid.TextMatrix(i, C_HABER)) > 0 Then
'
'                        MsgBox1 "Documento ODF tiene Tratamiento Pasivo, monto Debe debe ser mayor a 0 para realizar pago", vbExclamation
'                        Grid.TextMatrix(i, C_HABER) = ""
'                        Call CalcTot
'                        Exit Function
'                    End If
'             ElseIf (Cb_Tipo.ItemData(Cb_Tipo.ListIndex)) = TC_TRASPASO Then
'
'             Else
'               MsgBox1 "Tipo Comprobante no corresponde a Tratamiento de documento ODF, Valide Tipo Comprobante para realizar pago", vbExclamation
'               Cb_Tipo.SetFocus
'
'               Exit Function
'
'             End If
'
'
'          End If
'          '616437 ffv

      End If
   Next i
   
   'Vemos si hay pago de cuotas repetidas
   For i = Grid.FixedRows To Grid.rows - 1
      For j = Grid.FixedRows To Grid.rows - 1
      
         If i <> j Then
            If Val(Grid.TextMatrix(i, C_IDDOCCUOTA)) <> 0 Then
               If Val(Grid.TextMatrix(j, C_IDDOCCUOTA)) = Val(Grid.TextMatrix(i, C_IDDOCCUOTA)) Then
                  MsgBox1 "Hay cuotas que están repetidas.", vbOKOnly + vbExclamation
                  Exit Function
               End If
            End If
         End If
      Next j
   Next i
   '3269719
   If lCentrFull = 1 Then
   
    If ltotalDebeFull <> ltotalHaberFull Then
        Call MsgBox1("Los totales de las columnas total Comprobante DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      
             If Not ValidaParaGenCompTipo Then
            
                'Propongo valor que cuadra, si debe y haber es =0 y es solo una fila que no tiene debe ni haber
                Diff = ltotalDebeFull - ltotalHaberFull
                       If RegValCero = 1 Then
                          If Diff < 0 Then
                             MsgBox1 "Para cuadrar el comprobante se sugiere valor en el DEBE", vbExclamation
                             Grid.TextMatrix(RowCuadra, C_DEBE) = Format(Abs(Diff), BL_NUMFMT)
                             
                          Else
                             MsgBox1 "Para cuadrar el comprobante se sugiere valor en el HABER", vbExclamation
                             Grid.TextMatrix(RowCuadra, C_HABER) = Format(Diff, BL_NUMFMT)
                             
                          End If
                          
                          Call CalcTot
                       End If
                
             End If
      
        Exit Function
      End If
   
   Else
   
   'Si el comprobante no cuadra y hay un sólo movimiento sin debe ni haber, se sugiere valor de cuadre
   If SumDebe <> SumHaber Then
      Call MsgBox1("Los totales de las columnas DEBE y HABER no son iguales.", vbOKOnly + vbExclamation)
      
      If Not ValidaParaGenCompTipo Then
     
         'Propongo valor que cuadra, si debe y haber es =0 y es solo una fila que no tiene debe ni haber
         Diff = SumDebe - SumHaber
         If RegValCero = 1 Then
            If Diff < 0 Then
               MsgBox1 "Para cuadrar el comprobante se sugiere valor en el DEBE", vbExclamation
               Grid.TextMatrix(RowCuadra, C_DEBE) = Format(Abs(Diff), BL_NUMFMT)
               
            Else
               MsgBox1 "Para cuadrar el comprobante se sugiere valor en el HABER", vbExclamation
               Grid.TextMatrix(RowCuadra, C_HABER) = Format(Diff, BL_NUMFMT)
               
            End If
            
            Call CalcTot
         End If
         
      End If
      
      Exit Function
   
     End If
   End If
   '3269719
   
   If Ch_OtrosIngEg14TER <> 0 And Not lCuentasDisponible Then
      MsgBox1 "Para marcar un comprobante como " & Ch_OtrosIngEg14TER.Caption & " debe incluir al menos una cuenta configurada en Cuentas de Ajustes Extra-contables >> Disponibles >> Disponible Ingresos-Egresos", vbExclamation
      Exit Function
   End If
   
   valida = True
   
End Function
Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   If Grid.Col = C_DEBE Or Grid.Col = C_HABER Then
      Call KeyNum(KeyAscii)
   ElseIf Grid.Col = C_CODCUENTA Then
      Call KeyUpper(KeyAscii)
   End If
      
End Sub

Private Sub CalcTot()
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   
   TotDebe = 0
   TotHaber = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      If Grid.RowHeight(i) > 0 Then     ' no está borrado
         TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
         TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      End If
   Next i
         
   GridTot.TextMatrix(0, C_DEBE) = Format(TotDebe, BL_NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(TotHaber, BL_NUMFMT)
   
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call bt_Copy_Click
   ElseIf KeyPaste(KeyCode, Shift) Then
      Call Bt_Paste_Click
   End If


End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim CodiF2214Ter As String
   Dim Cod As String
   Grid.Row = Grid.MouseRow
   Grid.Col = Grid.MouseCol
   
   If lOper <> O_EDIT And lOper <> O_NEW Then
      Exit Sub
   End If

   If Button = vbRightButton Then
      Call PopupMenu(M_Opciones)
   End If
   
    If Grid.Row = Grid.rows - 1 Then
       Grid.rows = Grid.rows + 1
    End If
   
'   '   ***** Ado 2699584 Tema 1 del 3.4
'            If Grid.Col = 3 Or Grid.Col = C_LSTCUENTA Then
'                CodiF2214Ter = ""
'                cod = Grid.TextMatrix(Grid.Row, C_CODCUENTA)
'                Call CodF2214Ter(cod, CodiF2214Ter)
'
'                If CodiF2214Ter = "1" Then 'Or CodiF2214Ter = "629" Then
'                    If MsgBox1("¡Atención! ¿Partida Ingresada corresponde a una Participación?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
'                        Exit Sub
'                    Else
'                        Dim Frm As FrmPercepciones
'                        Set Frm = New FrmPercepciones
'                        Frm.CodCta = CodCuentaSelec
'                        Frm.GIdPerc = 0
'                        Frm.Fecha = Tx_Fecha
'                        Frm.Show vbModal
'                        Set Frm = Nothing
'                    End If
'                End If
'            End If
'
''   ***** Fin Ado 2699584 Tema 1 del 3.4

End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim Txt As String
   
   If Grid.MouseCol <> C_CODCUENTA And Grid.MouseCol <> C_CUENTA Then
      Grid.ToolTipText = ""
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.MouseRow, C_NOTA) <> "" Then
      Txt = ReplaceStr(Grid.TextMatrix(Grid.MouseRow, C_NOTA), vbCr, " / ")
      Txt = ReplaceStr(Txt, vbLf, "")
      Grid.ToolTipText = Txt
   Else
      Grid.ToolTipText = ""
   End If

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub




Private Sub M_AddNote_Click()
   Dim Frm As FrmNote
   Dim Row As Integer
   Dim Txt As String
   
   Row = Grid.Row
   
   If Not (Bt_OK.visible) Then
      Exit Sub
   End If
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_ORDEN)) = 0 Then
      Exit Sub
   End If
   
   Txt = Grid.TextMatrix(Row, C_NOTA)
   
   Set Frm = New FrmNote
   Call Frm.FEdit(Txt)
   Set Frm = Nothing
   
   Grid.TextMatrix(Row, C_NOTA) = Txt
   If Txt <> "" Then
      Grid.Row = Row
      Grid.Col = C_CODCUENTA
      Grid.CellPictureAlignment = flexAlignRightTop
      Set Grid.CellPicture = FrmMain.Pc_Nota
   End If

   Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
  
End Sub

Private Sub M_Copy_Click()
   Call bt_Copy_Click
End Sub

Private Sub M_DelNote_Click()

   If Grid.TextMatrix(Grid.Row, C_NOTA) = "" Then
      MsgBox1 "Este registro no tiene nota asociada.", vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar la nota asociada a este registro?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Grid.TextMatrix(Grid.Row, C_NOTA) = ""
   Grid.Row = Grid.Row
   Grid.Col = C_CODCUENTA
   Set Grid.CellPicture = LoadPicture()

   Call FGrModRow(Grid, Grid.Row, FGR_U, C_IDMOV, C_UPDATE)
   
End Sub

Private Sub M_EditNote_Click()
   Call M_AddNote_Click
End Sub

Private Sub M_Paste_Click()
   Call Bt_Paste_Click
End Sub

Private Sub M_SelCuenta_Click()
   
   Call AsignaCuenta(Grid.Row, Grid.Col)
   
End Sub

Private Sub M_ViewNote_Click()
   Dim Frm As FrmNote
   Dim Row As Integer
   Dim Txt As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_ORDEN)) = 0 Then
      Exit Sub
   End If
   
   Txt = Grid.TextMatrix(Row, C_NOTA)
   
   Set Frm = New FrmNote
   Call Frm.FView(Txt)
   Set Frm = Nothing
   

End Sub

Private Sub Tx_Fecha_GotFocus()
   Call DtGotFocus(Tx_Fecha)
End Sub

Private Sub Tx_Fecha_LostFocus()
   Call DtLostFocus(Tx_Fecha)
End Sub
Private Sub EnabForm(bool As Boolean)

   If Not lCompTipo Then
      If lOper = O_EDIT Or lOper = O_NEW Then
         If Not ChkPriv(PRV_ING_COMP) Then
            bool = False
         End If
      End If
   Else
      If lOper = O_EDIT Or lOper = O_NEW Then
         If Not ChkPriv(PRV_CFG_EMP) Then
            bool = False
         End If
      End If
   End If
   
   Bt_Duplicate.Enabled = bool
   'Bt_Cut.Enabled = bool
   Bt_Del.Enabled = bool
   Bt_Cuadrar.Enabled = bool
   Bt_Glosas(0).Enabled = bool
   Bt_Glosas(1).Enabled = bool
   Bt_MoveDown.Enabled = bool
   Bt_MoveUp.Enabled = bool
   Bt_Paste.Enabled = bool
   Bt_NewDoc.Enabled = bool
   Bt_BuscarDoc.Enabled = bool
   Bt_DelDoc.Enabled = bool
   Bt_GenPago.Enabled = bool
   Bt_SelFecha.Enabled = bool
   Tx_Fecha.Enabled = bool
   Tx_Glosa.Enabled = bool
   Ch_ImpRes.Enabled = bool
   Tx_GlosaCompTipo.Enabled = bool
   Tx_Descrip.Enabled = bool
   Tx_Nombre.Enabled = bool
   Cb_Tipo.Enabled = bool
   Cb_TipoCompTipo.Enabled = bool
   Cb_Estado.Enabled = bool
   Grid.Locked = Not bool
   Fr_TAjuste.Enabled = bool
   
   If Fr_TAjuste.Enabled And lOper = O_EDIT Then
      Fr_TAjuste.Enabled = False
   End If

      
   Bt_OK.visible = bool
   If Not bool Then
      If lOper <> O_NEW Then
         Bt_Cancel.Caption = "Cerrar"
      Else
         Bt_Cancel.visible = False
      End If
      M_AddNote.Enabled = False
      M_EditNote.Enabled = False
      M_DelNote.Enabled = False
   Else
      Bt_Cancel.Caption = "Cancelar"
   End If
   
   If Not ChkPriv(PRV_ADM_DOCS) Then
      Bt_GenPago.Enabled = False
   End If
               
   If Not ChkPriv(PRV_ING_DOCS) And Not ChkPriv(PRV_ING_COMP) Then
      Bt_ActivoFijo.Enabled = False
   End If
   
End Sub
Private Sub RemoveComprobante()
   Dim Q1 As String
   'En algún momento grabó el movimiento para ingresar detalles
   'Por lo tanto debo eliminarlo junto a sus posibles detalles
   
   If lidComp = 0 Then
      Exit Sub
   End If
   
   If lCompTipo Then
'      Q1 = "DELETE * FROM " & lTblComprobante & " WHERE idComp=" & lidComp
'      Call ExecSQL(DbMain, Q1)
      Call DeleteSQL(DbMain, lTblComprobante, " WHERE idComp=" & lidComp & " AND IdEmpresa = " & gEmpresa.id)
      
'      Q1 = "DELETE * FROM " & lTblMovComprobante & " WHERE idComp=" & lidComp
'      Call ExecSQL(DbMain, Q1)
      Call DeleteSQL(DbMain, lTblMovComprobante, " WHERE idComp=" & lidComp & " AND IdEmpresa = " & gEmpresa.id)
   
   Else
      Call DeleteSQL(DbMain, lTblComprobante, " WHERE idComp=" & lidComp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Call DeleteSQL(DbMain, lTblMovComprobante, " WHERE idComp=" & lidComp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   End If
      
End Sub
Private Sub RemovePercNull()
Dim Q1 As String
Dim Rs As Recordset

    Q1 = "SELECT IDPerc "
    Q1 = Q1 & "From Percepciones "
    Q1 = Q1 & " WHERE idcomp IS NULL "
    Q1 = Q1 & " AND idempresa = " & gEmpresa.id
    Q1 = Q1 & " AND ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
    Q1 = "DELETE FROM DETPercepciones "
    Q1 = Q1 & " WHERE IDPERC = " & vFld(Rs("IDPerc"))
    Call ExecSQL(DbMain, Q1)
    
   Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

    Q1 = "DELETE FROM  PERCEPCIONES "
    Q1 = Q1 & " WHERE idcomp IS NULL "
    Q1 = Q1 & " AND idempresa = " & gEmpresa.id
    Q1 = Q1 & " AND ano = " & gEmpresa.Ano
    Call ExecSQL(DbMain, Q1)
End Sub

Private Sub FillCompTipo(IdCompTipo As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Tipo As Integer
   
   Q1 = "SELECT Tipo, Glosa, TotalDebe, TotalHaber, Nombre "
   Q1 = Q1 & " FROM CT_Comprobante WHERE IdComp = " & IdCompTipo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Tx_Glosa = DeSQL(Rs("Glosa"))
      Tx_NombCompTipo = DeSQL(Rs("Nombre"))
      Call SelItem(Cb_Tipo, vFld(Rs("Tipo")))
      
   End If
   Call CloseRs(Rs)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows

   Q1 = "SELECT IdMov, Orden, CT_MovComprobante.IdCuenta, Cuentas.Codigo, Nombre, "
   Q1 = Q1 & " Cuentas.Descripcion, CT_MovComprobante.Debe, CT_MovComprobante.Haber, Glosa, "
   Q1 = Q1 & " CT_MovComprobante.IdCCosto, CT_MovComprobante.IdAreaNeg, "
   Q1 = Q1 & " AreaNegocio.Descripcion As DescAreaNeg, CentroCosto.Descripcion As DescCCosto "
   Q1 = Q1 & " FROM ((CT_MovComprobante"
   Q1 = Q1 & " INNER JOIN Cuentas ON CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = CT_MovComprobante.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN AreaNegocio ON CT_MovComprobante.IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & " AND AreaNegocio.IdEmpresa = CT_MovComprobante.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN CentroCosto ON CT_MovComprobante.IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & " AND CentroCosto.IdEmpresa = CT_MovComprobante.IdEmpresa "
   Q1 = Q1 & " WHERE IdComp = " & IdCompTipo & " AND Cuentas.Nivel =" & gLastNivel
   Q1 = Q1 & " AND CT_MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY Orden "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
'     Grid.TextMatrix(i, C_IDMOV) = vFld(Rs("IdMov"))
      Grid.TextMatrix(i, C_ORDEN) = vFld(Rs("Orden"))
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(i, C_CUENTA) = FCase(vFld(Rs("Descripcion")))
      
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), BL_NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), BL_NUMFMT)
      Grid.TextMatrix(i, C_GLOSA) = DeSQL(Rs("Glosa"))
      
      Grid.TextMatrix(i, C_AREANEG) = vFld(Rs("DescAreaNeg"), True)
      Grid.TextMatrix(i, C_IDAREANEG) = vFld(Rs("IdAreaNeg"))
      Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("DescCCosto"), True)
      Grid.TextMatrix(i, C_IDCCOSTO) = vFld(Rs("IdCCosto"))
      
'      Call FGrSetPicture(Grid, i, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
'      Call FGrSetPicture(Grid, i, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)

      Grid.TextMatrix(i, C_LSTCUENTA) = ">>"
      Grid.TextMatrix(i, C_DETALLE) = ">>"
      
      Grid.TextMatrix(i, C_UPDATE) = FGR_I
           
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
   Call CalcTot
End Sub
Private Sub ClearForm()
   Dim Col As Integer
   
   Call FGrClear(Grid)
   For Col = 0 To GridTot.Cols - 1
      GridTot.TextMatrix(0, Col) = ""
   Next Col
   
   Tx_Glosa = ""
   Tx_IdComp = ""
   Tx_Correlativo = ""
   lidComp = 0
   lOldFechaEmision = 0
   Cb_Tipo.ListIndex = -1
   
   Ch_OtrosIngEg14TER.Value = 0
   Ch_OtrosIngEg14TER.Enabled = True
   
   'Call SetTxDate(Tx_Fecha, Now)
   'Cb_Tipo.ListIndex = 0
   'Cb_Estado.ListIndex = EC_PENDIENTE - 1
   Call SelItem(Cb_Estado, gEstadoNewComp)
   
End Sub

Private Function GetCompTipo(ByVal Nombre As String) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT IdComp FROM CT_Comprobante WHERE Nombre = '" & Nombre & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      GetCompTipo = vFld(Rs("IdComp"))
   Else
      GetCompTipo = 0
   End If
   
   Call CloseRs(Rs)
End Function

Private Sub GenComprobanteTipo(ByVal ConValores As Boolean)
   Dim Rs As Recordset
   Dim Rc As Long, RNum As Long
   Dim Q1 As String
   Dim i As Integer
   Dim IdCompTipo As Long
   Dim Lin As Integer
   Dim idMov As Long
   Dim Nombre As String
    
   If Not valida(True) Then
      Exit Sub
   End If
   
   Nombre = InputBox("Ingrese un NOMBRE CORTO para el nuevo" & vbNewLine & "Comprobante Tipo (máx. 10 caracteres).", "Nuevo Comprobante Tipo:")
   Nombre = Trim(UCase(Nombre))
   
   If Nombre = "" Then
      MsgBox1 "Debe ingresar un nombre para el nuevo Comprobante Tipo.", vbExclamation + vbOKOnly
      Exit Sub
   
   Else
      Q1 = "SELECT IdComp FROM CT_Comprobante WHERE Nombre = '" & Nombre & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then   'ya existe uno con este nombre
         MsgBox1 "Ya existe un Comprobante Tipo con el nombre '" & Nombre & "'.", vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
      
      Call CloseRs(Rs)
   End If
   
   If ConValores = True And GridTot.TextMatrix(0, C_DEBE) <> GridTot.TextMatrix(0, C_HABER) Then
      MsgBox1 "El Comprobante Tipo no cuadra.", vbExclamation
      Exit Sub
   End If

   'todo bien, generamos comprobante tipo
'   Set Rs = DbMain.OpenRecordset("CT_Comprobante")
'   Rs.AddNew
'
'   IdCompTipo = Rs("IdComp")
'   Rs.Fields("Tipo") = Cb_Tipo.ItemData(Cb_Tipo.ListIndex)
'   Rs.Fields("Correlativo") = Rs("IdComp")
'   Rs.Update
'   Rs.Close
'   Set Rs = Nothing

   IdCompTipo = AdvTbAddNew(DbMain, "CT_Comprobante", "IdComp", "IdEmpresa", gEmpresa.id)
   
'   Q1 = "UPDATE CT_Comprobante SET( Correlativo ) VALUES(" & Rs("IdComp") & ") WHERE IdCompTipo = " & IdCompTipo
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
'   Call ExecSQL(DbMain, Q1)
   
   'actualizamos el Encabezado Comprobante Tipo
   Q1 = "UPDATE CT_Comprobante SET "
   Q1 = Q1 & "  Tipo =" & Cb_Tipo.ItemData(Cb_Tipo.ListIndex)
   Q1 = Q1 & ", Correlativo =" & IdCompTipo
   '3225619 FPG Se le agrego el Left para que haga el limite de caracteres antes de ejecutar la Query
   Q1 = Q1 & ", Glosa = '" & ParaSQL(Left(Tx_Glosa, 100)) & "'"
   Q1 = Q1 & ", Descrip = '" & ParaSQL(Left(Tx_Glosa, 40)) & "'"
   Q1 = Q1 & ", Nombre ='" & ParaSQL(Left(Nombre, 10)) & "'"
   ' FIN 3225619
'   Q1 = Q1 & ", Glosa = '" & ParaSQL(Tx_Glosa) & "'"
'   Q1 = Q1 & ", Descrip = '" & ParaSQL(Tx_Glosa) & "'"
'   Q1 = Q1 & ", Nombre ='" & ParaSQL(Nombre) & "'"
   If ConValores = True Then
      Q1 = Q1 & ", TotalDebe = " & vFmt(GridTot.TextMatrix(0, C_DEBE))
      Q1 = Q1 & ", TotalHaber = " & vFmt(GridTot.TextMatrix(0, C_HABER))
   End If
   Q1 = Q1 & " WHERE IdComp = " & IdCompTipo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
    
   Call ExecSQL(DbMain, Q1)
            
   Tx_NombCompTipo = Nombre
            
   'ahora los movimientos
   
   Lin = Grid.FixedRows
   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_ORDEN) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
      
'         Set Rs = DbMain.OpenRecordset("CT_MovComprobante")
'         Rs.AddNew
'
'         idMov = Rs("IdMov")
'
'         Rs.Update
'         Rs.Close
'         Set Rs = Nothing

         idMov = AdvTbAddNew(DbMain, "CT_MovComprobante", "IdMov", "IdEmpresa", gEmpresa.id)
                  
         Q1 = "UPDATE CT_MovComprobante SET "
         Q1 = Q1 & "  IdComp = " & IdCompTipo
         Q1 = Q1 & ", Orden = " & Lin
         Q1 = Q1 & ", IdCuenta = " & Grid.TextMatrix(i, C_IDCUENTA)
         
         If ConValores = True Then
            Q1 = Q1 & ", Debe = " & vFmt(Grid.TextMatrix(i, C_DEBE))
            Q1 = Q1 & ", Haber = " & vFmt(Grid.TextMatrix(i, C_HABER))
         End If
         
         Q1 = Q1 & ", Glosa = '" & ParaSQL(Grid.TextMatrix(i, C_GLOSA)) & "'"
         Q1 = Q1 & ", IdAreaNeg = " & vFmt(Grid.TextMatrix(i, C_IDAREANEG))
         Q1 = Q1 & ", IdCCosto = " & vFmt(Grid.TextMatrix(i, C_IDCCOSTO))
            
         Q1 = Q1 & " WHERE IdMov = " & idMov
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id

         Call ExecSQL(DbMain, Q1)
      
         Lin = Lin + 1
      End If
      
   Next i

End Sub

Private Sub Tx_NombCompTipo_Change()

   If Trim(Tx_NombCompTipo) <> "" Then
      Bt_SelCompTipo.Enabled = True
   Else
      Bt_SelCompTipo.Enabled = False
   End If

End Sub

Private Sub Tx_NombCompTipo_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
   Call KeyUpper(KeyAscii)
End Sub

Private Sub Tx_Nombre_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
   Call KeyUpper(KeyAscii)
End Sub

Private Sub AsignarDoc(ByVal IdDoc As Long, ByVal IdDocCuota As Long)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Q2 As String
   
   If IdDoc = 0 Then   'ningún doc
      Grid.TextMatrix(Grid.Row, C_IDDOC) = ""
      Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA) = ""
      Grid.TextMatrix(Grid.Row, C_NUMDOC) = ""
      Grid.TextMatrix(Grid.Row, C_TIPODOC) = ""
      Grid.TextMatrix(Grid.Row, C_ENTIDAD) = ""
      
   Else
               
      Q1 = "SELECT Documento.IdDoc, NumDoc, TipoLib, TipoDoc, Entidades.Nombre, Total, SaldoDoc, NumCuotas, tratamiento "
      If IdDocCuota <> 0 Then
         Q1 = Q1 & ", DocCuotas.IdDocCuota, DocCuotas.MontoCuota, DocCuotas.NumCuota "
         Q1 = Q1 & " FROM (Documento LEFT JOIN DocCuotas ON Documento.IdDoc = DocCuotas.IdDoc "
         Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "DocCuotas") & " )"
     Else
         Q1 = Q1 & ", 0 as IdDocCuota, 0 as MontoCuota, 0 as NumCuota"
         Q1 = Q1 & " FROM Documento "
      End If
      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " AND Documento.IdEmpresa = Entidades.IdEmpresa "
      Q1 = Q1 & " WHERE Documento.IdDoc = " & IdDoc
      
      If IdDocCuota <> 0 Then
         Q1 = Q1 & " AND  (IdDocCuota IS NULL OR IdDocCuota = " & IdDocCuota & ")"
      End If
      
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
'      If CodTipoLib = LIB_OTROFULL Then
'        Q1 = Replace(Replace(Q1, "Documento", "DocumentoFull"), ",0", ",DocumentoFull.tratamiento")
'      End If
      'Q1 = Q1 & " UNION ALL " & Q2
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
               
         If Val(Grid.TextMatrix(Grid.Row, C_ORDEN)) = 0 Then   'nuevo
            Call AddNewMov(Grid.Row)
         End If
                   
         Grid.TextMatrix(Grid.Row, C_IDDOC) = vFld(Rs("IdDoc"))
         Grid.TextMatrix(Grid.Row, C_NUMDOC) = vFld(Rs("NumDoc"))
         Grid.TextMatrix(Grid.Row, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
         Grid.TextMatrix(Grid.Row, C_ENTIDAD) = vFld(Rs("Nombre"), True)
         If Grid.TextMatrix(Grid.Row, C_TIPODOC) = "VSD" Then
            Grid.TextMatrix(Grid.Row, C_NUMDOC) = ""
         End If
         
         'si no se genera el mov. en forma automática, por lo menos le traemos el monto y se lo asignamos al DEBE o HABEr, dependiendo del saldo.
         'Esto sólo si DEBE = 0 y HABER = 0
         
         If vFmt(Grid.TextMatrix(Grid.Row, C_DEBE)) = 0 And vFmt(Grid.TextMatrix(Grid.Row, C_HABER)) = 0 Then
         
            If vFld(Rs("SaldoDoc")) > 0 Then
               If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                  
                  'Grid.TextMatrix(Grid.row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                  'feña
                  '3044643
                  If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                        
                        If vFld(Rs("tratamiento")) <> 1 Then
                          Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                        Else
                          Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                        End If
                  Else
                  Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                  End If
                  '3044643
                        
                        'fin feña
                  Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                  Grid.TextMatrix(Grid.Row, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
               Else
                  'Grid.TextMatrix(Grid.row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  'feña
                  '3044643
                 If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                  If vFld(Rs("tratamiento")) <> 1 Then
                    Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  Else
                    Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  End If
                 Else
                 Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                 
                 End If
                  '3044643
                  'fin feña
                  
               End If
            ElseIf vFld(Rs("SaldoDoc")) < 0 Then
               If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                  'Grid.TextMatrix(Grid.row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                  'feña
                  '3044643
                If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                    If vFld(Rs("tratamiento")) <> 1 Then
                      Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                    Else
                      Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                    End If
                Else
                Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                    
                End If
                '3044643
                    'fin feña
                    Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                  Grid.TextMatrix(Grid.Row, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
               Else
                  'Grid.TextMatrix(Grid.row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  'feña
                  '3044643
                 If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                  If vFld(Rs("tratamiento")) <> 1 Then
                    Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  Else
                    Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                  End If
                 Else
                 Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                 End If
                  '3044643
                  'fin feña
               End If
            Else
               If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then          'esto no debería ocurrir nunca si saldo = 0
                  'Grid.TextMatrix(Grid.row, C_DEBE) = Format(vFld(Rs("MontoCuota")), NUMFMT)
                  'feña
                   '3044643
                If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                  If vFld(Rs("tratamiento")) <> 1 Then
                    Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                  Else
                    Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                  End If
                 Else
                 Grid.TextMatrix(Grid.Row, C_DEBE) = Format(vFld(Rs("MontoCuota")), NUMFMT)
                 End If
                  '3044643
                  'fin feña
                  Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                  Grid.TextMatrix(Grid.Row, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
               Else
                  'Grid.TextMatrix(Grid.row, C_DEBE) = Format(vFld(Rs("Total")), NUMFMT)
                  
                  'feña
                  '3044643
                 If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                    If vFld(Rs("tratamiento")) <> 1 Then
                      Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("Total"))), NUMFMT)
                    Else
                      Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("Total"))), NUMFMT)
                    End If
                    
                 Else
                   Grid.TextMatrix(Grid.Row, C_DEBE) = Format(vFld(Rs("Total")), NUMFMT)
                 End If
                 '3044643
                  'fin feña
               End If
            End If
            
            
'            If vFld(Rs("SaldoDoc")) < 0 Then
'               Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
'            ElseIf vFld(Rs("SaldoDoc")) > 0 Then
'               Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
'            Else
'               Grid.TextMatrix(Grid.Row, C_DEBE) = Format(vFld(Rs("Total")), NUMFMT)
'            End If
         
         End If
         
         Grid.TextMatrix(Grid.Row, C_DEPAGO) = NEW_DESELDOC
         
         If vFld(Rs("TipoLib")) = LIB_COMPRAS Or vFld(Rs("TipoLib")) = LIB_VENTAS Or vFld(Rs("TipoLib")) = LIB_RETEN Then
            Ch_OtrosIngEg14TER.Value = 0
            Ch_OtrosIngEg14TER.Enabled = False
         End If
         
         lGenPago = True
         
      End If
      
      Call CloseRs(Rs)
   
   End If
   
   Call FGrModRow(Grid, Grid.Row, FGR_U, C_IDMOV, C_UPDATE)

End Sub

Private Sub AsignarLstDoc(LstIdDoc() As LstDoc_t, Optional ByVal AskGenMov As Boolean = True)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdDocStr As String
   Dim IdDocCuotaStr As String
   Dim i As Integer
   Dim GenMov As Integer
   Dim AsigDoc As Integer
   Dim j As Integer
   Dim NewRow As Integer
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim TipoValLib As String
   Dim TipoLib As Integer
   Dim MsgNotMov As Boolean
   Dim Tratamiento As Long
   
   If LstIdDoc(0).IdDoc = 0 Then   'ningún doc
      Grid.TextMatrix(Grid.Row, C_IDDOC) = ""
      Grid.TextMatrix(Grid.Row, C_NUMDOC) = ""
      Grid.TextMatrix(Grid.Row, C_TIPODOC) = ""
      Grid.TextMatrix(Grid.Row, C_ENTIDAD) = ""
  
   Else
         
      For i = 0 To UBound(LstIdDoc)
         If LstIdDoc(i).IdDoc = 0 Then
            Exit For
         End If
         IdDocStr = IdDocStr & "," & LstIdDoc(i).IdDoc
         
         If LstIdDoc(i).IdDocCuota <> 0 Then
            IdDocCuotaStr = IdDocCuotaStr & "," & LstIdDoc(i).IdDocCuota
         End If
         
      Next i
      
      If IdDocStr = "" Then
         Exit Sub
      End If
      
      IdDocStr = Mid(IdDocStr, 2)
     
      If IdDocCuotaStr <> "" Then
         IdDocCuotaStr = Mid(IdDocCuotaStr, 2)
      End If
      
     
      'obtenemos el primer registro en blanco, para generación automática de movimientos
      For j = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(j, C_ORDEN) = "" Then
            NewRow = j
            Exit For
         End If
      Next j
                  
      TipoValLib = " (EsTotalDoc <> 0) OR (MovDocumento.IdTipoValLib IS NULL OR MovDocumento.IdTipoValLib = 0)"
      
'      Q1 = "SELECT Documento.IdDoc, NumDoc, TipoLib, TipoDoc, Entidades.Nombre, Documento.Total, Documento.SaldoDoc, "
'      Q1 = Q1 & " MovDocumento.IdCuenta, MovDocumento.Debe, MovDocumento.Haber, "
'      Q1 = Q1 & " Cuentas.Codigo, Cuentas.Descripcion, MovDocumento.EsTotalDoc, Documento.SaldoDoc "
'      Q1 = Q1 & " FROM ((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad) "
'      Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc )"
'      Q1 = Q1 & " LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'      Q1 = Q1 & " WHERE Documento.IdDoc IN (" & IdDocStr & ")"
'      Q1 = Q1 & " AND (" & TipoValLib & ")"
      
      
      Q1 = "SELECT Documento.IdDoc, NumDoc, TipoLib, TipoDoc, Entidades.Nombre, Documento.Total, Documento.SaldoDoc"
      'Q1 = Q1 & ", MovDocumento.IdCuenta, MovDocumento.Debe, MovDocumento.Haber"
            Q1 = Q1 & ", iif(TipoLib = 8, Documento.IdCtaBanco, MovDocumento.IdCuenta) as IdCuenta, MovDocumento.Debe, MovDocumento.Haber"
      Q1 = Q1 & ", Cuentas.Codigo, Cuentas.Descripcion, MovDocumento.EsTotalDoc, Documento.SaldoDoc, Documento.NumCuotas, Documento.Tratamiento "
      If IdDocCuotaStr <> "" Then
         Q1 = Q1 & ", DocCuotas.IdDocCuota, DocCuotas.MontoCuota, DocCuotas.NumCuota "
      Else
         Q1 = Q1 & ", 0 as IdDocCuota, 0 as MontoCuota, 0 as NumCuota "
      End If
      
      If IdDocCuotaStr <> "" Then
         Q1 = Q1 & " FROM (((Documento LEFT JOIN DocCuotas ON Documento.IdDoc = DocCuotas.IdDoc "
         Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "DocCuotas") & " )"
      Else
         Q1 = Q1 & " FROM ((Documento "
      End If
      
      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
      Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
      Q1 = Q1 & " LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
      Q1 = Q1 & " WHERE Documento.IdDoc IN(" & IdDocStr & ")"
      
      If IdDocCuotaStr <> "" Then
         Q1 = Q1 & " AND  (IdDocCuota IS NULL OR IdDocCuota IN (" & IdDocCuotaStr & "))"
      End If
      
      Q1 = Q1 & " AND ( (EsTotalDoc <> 0) OR (MovDocumento.IdTipoValLib IS NULL OR MovDocumento.IdTipoValLib = 0))"
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoLib, TipoDoc, NumDoc"
      If IdDocCuotaStr <> "" Then
         Q1 = Q1 & ", NumCuota"
      End If
      Set Rs = OpenRs(DbMain, Q1)
      
      If Grid.Row < Grid.FixedRows Then
         Grid.Row = Grid.FixedRows
      End If
      
      i = Grid.Row   'registro seleccionado, por si selecciona un sólo documento
         
      GenMov = 0
      AsigDoc = 0
       
      Do While Rs.EOF = False
      
         TipoLib = vFld(Rs("TipoLib"))
         Tratamiento = vFld(Rs("Tratamiento"))
      
         'If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then
         
         If vFld(Rs("EsTotalDoc")) <> 0 Or TipoLib = LIB_OTROFULL Then   'ventas, compras o retenciones
         
            If GenMov = 0 Then
               
               If AskGenMov Then
                  If gCtasBas.IdCtaPagoFacturas = 0 Then
                     MsgBox1 "No es posible generar automáticamente los movimientos de pago de los documentos de compra o venta seleccionados, debido a que no se ha seleccionado la cuenta de pago de facturas." & vbNewLine & vbNewLine & "Para tal efecto, utilice la opción 'Definición de Cuentas Básicas' incluída en la Configuración del Sistema.", vbExclamation + vbOKOnly
                     GenMov = vbNo
                  Else
                     GenMov = MsgBox1("¿Desea que el sistema genere automáticamente los movimientos de pago para los documentos de compra, venta u honorarios seleccionados?", vbQuestion + vbYesNo)
                  End If
               Else
                  GenMov = vbNo
                  
               End If
               
               If GenMov = vbYes Then
                  i = NewRow     'desde el primer registro en blanco
                  If Tx_Glosa = "" Then
                     'proponemos glosa
                     Tx_Glosa = "Contabiliza Pago de Documentos de " & gTipoLibNew(IIf(TipoLib = 8, 6, TipoLib)).Nombre
                  End If

               Else
                  AsigDoc = MsgBox1("¿Desea asignar los documentos seleccionados a los movimientos a partir del registro " & Grid.Row & "?", vbQuestion + vbYesNo)
               End If
            End If
            
            If GenMov = vbYes Then    'generamos un movimiento por cada doc, a partir del primer registro en blanco
               
               If vFld(Rs("IdCuenta")) > 0 Then
                  
                  Call AddNewMov(i)
                  Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
                  Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(GetCodCuenta(vFld(Rs("IdCuenta")))) 'FmtCodCuenta(vFld(Rs("Codigo")))
                  Grid.TextMatrix(i, C_CUENTA) = FCase(GetDescCuenta(vFld(Rs("IdCuenta"), True))) 'FCase(vFld(Rs("Descripcion"), True))
                  
                  If vFld(Rs("SaldoDoc")) = 0 Then
                  
                     If vFld(Rs("Debe")) <> 0 Then

                        Grid.TextMatrix(i, C_HABER) = Format(0, NUMFMT)
                        'Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Debe")), NUMFMT)
                        TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
                     Else
                        Grid.TextMatrix(i, C_DEBE) = Format(0, NUMFMT)
                        'Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Haber")), NUMFMT)
                        TotDebe = TotDebe + Grid.TextMatrix(i, C_DEBE)
                     End If
                  
                  Else
                  
                     If vFld(Rs("SaldoDoc")) > 0 Then
                        If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                            ' Feña
                            If vFld(Rs("tratamiento")) <> 1 Then
                              Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                            Else
                              Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                            End If
                            'Fin Feña
                           'Grid.TextMatrix(i, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                           Grid.TextMatrix(i, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                           Grid.TextMatrix(i, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
                        Else
                            ' Feña
                            If TipoLib = LIB_OTROFULL Then
                                If vFld(Rs("tratamiento")) <> 1 Then
                                  Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                                Else
                                  Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                                End If
                            Else
                                Grid.TextMatrix(i, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                            End If
                           'Fin Feña
                           'Grid.TextMatrix(i, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                        End If
                        'TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
                        TotHaber = TotHaber + IIf(Grid.TextMatrix(i, C_HABER) = "", Grid.TextMatrix(i, C_DEBE), Grid.TextMatrix(i, C_HABER))
                     Else
                        If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                            ' Feña
                            If TipoLib = LIB_OTROFULL Then
                                If vFld(Rs("tratamiento")) <> 1 Then
                                  Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                                Else
                                  Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                                End If
                            Else
                                Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                            End If
                            'Fin Feña
                           'Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                           Grid.TextMatrix(i, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                           Grid.TextMatrix(i, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
                        Else
                            ' Feña
                            If TipoLib = LIB_OTROFULL Then
                                If vFld(Rs("tratamiento")) <> 1 Then
                                  Grid.TextMatrix(Grid.Row, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                                Else
                                  Grid.TextMatrix(Grid.Row, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                                End If
                            Else
                                Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                            End If
                            'Fin Feña
                           'Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                        End If
                        'TotDebe = TotDebe + Grid.TextMatrix(i, C_DEBE)
                        TotDebe = TotDebe + IIf(Grid.TextMatrix(i, C_DEBE) = "", Grid.TextMatrix(i, C_HABER), Grid.TextMatrix(i, C_DEBE))
                     End If
                  
                  End If
                  
                  'marcamos que lo generamos automáticamente por pago, para actualizar campos Documento.Estado y MovComprobante.DePago en Save
                  Grid.TextMatrix(i, C_DEPAGO) = NEW_DEPAGO
                  
               End If
               
            End If
         
         ElseIf AsigDoc = 0 Then
            AsigDoc = MsgBox1("¿Desea asignar los documentos seleccionados a los movimientos a partir del registro " & Grid.Row & "?", vbQuestion + vbYesNo)
         
         End If
         
         If GenMov = vbYes Or AsigDoc = vbYes Then
            If i >= Grid.rows Then
               If Not MsgNotMov Then
                  MsgBox1 "Uno o más documentos no podrán ser asignados ya que no hay suficientes movimientos creados.", vbInformation + vbOKOnly
                  MsgNotMov = True
               End If
            Else
               Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
               Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
               Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
               Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Nombre"), True)
               If Grid.TextMatrix(i, C_TIPODOC) = "VSD" Then
                  Grid.TextMatrix(i, C_NUMDOC) = ""
               End If
               
               If AsigDoc = vbYes Then
                  Grid.TextMatrix(i, C_DEPAGO) = NEW_DESELDOC
               End If
            
               If Not (GenMov = vbYes) Then
                  'si no se genera el mov. en forma automática, por lo menos le traemos el monto y se lo asignamos al DEBE, el usuario verá si lo mueve al Haber
                  
                  If vFld(Rs("SaldoDoc")) < 0 Then
                     If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                        
                        Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                        Grid.TextMatrix(i, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                        Grid.TextMatrix(i, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
                     Else
                        Grid.TextMatrix(i, C_DEBE) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                     End If
                  ElseIf vFld(Rs("SaldoDoc")) > 0 Then
                     If vFld(Rs("IdDocCuota")) > 0 And vFld(Rs("MontoCuota")) > 0 Then
                        Grid.TextMatrix(i, C_HABER) = Format(Abs(vFld(Rs("MontoCuota"))), NUMFMT)
                        Grid.TextMatrix(i, C_IDDOCCUOTA) = Format(Abs(vFld(Rs("IdDocCuota"))), NUMFMT)
                        Grid.TextMatrix(i, C_GLOSA) = "Cuota " & vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas"))
                     Else
                        Grid.TextMatrix(i, C_HABER) = Format(Abs(vFld(Rs("SaldoDoc"))), NUMFMT)
                     End If
                  Else
                     Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Total")), NUMFMT)
                  End If
               
               End If
            
               Call FGrModRow(Grid, i, FGR_U, C_IDMOV, C_UPDATE)
            End If
         End If
         
         Rs.MoveNext
         i = i + 1
         'Grid.Rows = Grid.Rows + 1
         
      Loop
      
      Call CloseRs(Rs)
   
   End If
   
   If GenMov = vbYes And (TotDebe - TotHaber) <> 0 Then
                  
      'generamos mov de pago
      Call AddNewMov(i)
      
      If TotDebe - TotHaber > 0 Then   'es al revés para que cuadre el comprobante
         'Grid.TextMatrix(i, C_HABER) = Format(TotDebe - TotHaber, NUMFMT)
        ' Feña
        If TipoLib = LIB_OTROFULL Then
            If Tratamiento = 1 Then
              Grid.TextMatrix(i, C_DEBE) = Format(Abs(TotDebe - TotHaber), NUMFMT)
            Else
              Grid.TextMatrix(i, C_HABER) = Format(Abs(TotDebe - TotHaber), NUMFMT)
            End If
        Else
            Grid.TextMatrix(i, C_HABER) = Format(Abs(TotDebe - TotHaber), NUMFMT)
        End If
        'Fin Feña
         
         Grid.TextMatrix(i, C_IDCUENTA) = gCtasBas.IdCtaPagoFacturas
         GetCodCuenta (gCtasBas.IdCtaPagoFacturas)
         Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(GetCodCuenta(gCtasBas.IdCtaPagoFacturas))
         Grid.TextMatrix(i, C_CUENTA) = GetDescCuenta(gCtasBas.IdCtaPagoFacturas)
         
      Else
         
        'feña
        If TipoLib = LIB_OTROFULL Then
            If Tratamiento = 1 Then
              Grid.TextMatrix(i, C_DEBE) = Format(Abs(TotDebe - TotHaber), NUMFMT)
            Else
              Grid.TextMatrix(i, C_HABER) = Format(Abs(TotDebe - TotHaber), NUMFMT)
            End If
        Else
            Grid.TextMatrix(i, C_DEBE) = Format(Abs(TotDebe - TotHaber), NUMFMT)
        End If
        'Fin Feña
         
         'Grid.TextMatrix(i, C_DEBE) = Format(Abs(TotDebe - TotHaber), NUMFMT)
         Grid.TextMatrix(i, C_IDCUENTA) = gCtasBas.IdCtaCobFacturas
         GetCodCuenta (gCtasBas.IdCtaCobFacturas)
         Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(GetCodCuenta(gCtasBas.IdCtaCobFacturas))
         Grid.TextMatrix(i, C_CUENTA) = GetDescCuenta(gCtasBas.IdCtaCobFacturas)
      
      End If
      
      Call GridAtribCuenta(i)
      Call FGrModRow(Grid, i, FGR_U, C_IDMOV, C_UPDATE)

      Ch_OtrosIngEg14TER.Value = 0
      Ch_OtrosIngEg14TER.Enabled = False
        
      lGenPago = True
  
   End If
      
   Call CalcTot
   
End Sub

Private Sub AddNewMov(ByVal Row As Integer)

   If Row >= Grid.rows Then
      Grid.rows = Grid.rows + (Row - Grid.rows) + 1
   End If
   
   

   Grid.TextMatrix(Row, C_ORDEN) = Row
'   Call FGrSetPicture(Grid, Row, C_LSTCUENTA, FrmMain.Pc_Flecha, vbButtonFace)
   Grid.TextMatrix(Row, C_LSTCUENTA) = ">>"
   
   If Not lCompTipo Then
'      Call FGrSetPicture(Grid, Row, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)
      Grid.TextMatrix(Row, C_DETALLE) = ">>"
   End If
   
   'If Row - 1 >= Grid.FixedRows And Grid.TextMatrix(Row - 1, C_GLOSA) = "" Then
   '   Grid.TextMatrix(Row - 1, C_GLOSA) = Tx_Glosa
   'End If
   
   Grid.TextMatrix(Row, C_GLOSA) = Tx_Glosa
    
End Sub
Private Sub DelMov(ByVal Row As Integer)
  
'   Grid.Row = Row
'   Grid.FlxGrid.Col = C_LSTCUENTA
'   Set Grid.CellPicture = LoadPicture()
   Grid.TextMatrix(Row, C_LSTCUENTA) = ""
   
'
'   Grid.FlxGrid.Col = C_DETALLE
'   Set Grid.CellPicture = LoadPicture()
   Grid.TextMatrix(Row, C_DETALLE) = ""
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDMOV, C_UPDATE)
   Grid.rows = Grid.rows + 1
   Grid.Row = Grid.Row - 1
   '3238739

   '3238739
End Sub

Private Sub DelPercepciones(Row As Integer, Cuenta As Long)
Dim Q1 As String

Q1 = "DELETE FROM  PERCEPCIONES "
Q1 = Q1 & " WHERE idcomp = " & lidComp
Q1 = Q1 & " AND ORDEN =  " & Row
Q1 = Q1 & " AND IDEMPRESA = " & gEmpresa.id
Q1 = Q1 & " AND ANO = " & gEmpresa.Ano
Q1 = Q1 & " AND IDCUENTA = " & Cuenta
Call ExecSQL(DbMain, Q1)

End Sub

Private Sub ActivoFijo(ByVal Row As Integer, ByVal Col As Integer)
   Dim idMov As Long
   Dim n As Integer
   Dim ValNeto As Double
   Dim ValIVA As Double
   Dim Frm As Form
   Dim Fecha As Long
   Dim IdCuenta As Long
   
   If lCompTipo Then
      Exit Sub
   End If
   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
      
   'si no hay IdMov en esta fila, hay que grabar primero
   
   If Val(Grid.TextMatrix(Row, C_IDMOV)) = 0 Or lidComp = 0 Then
      If MsgBox1("Para ingresar movimientos de Activo Fijo, es necesario grabar este comprobante." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If valida() = False Then
         Exit Sub
      End If

      Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO   'para que grabe el movimiento aún cuando los valores de debe y haber sean cero
      
      Call SaveAll
      
      If vFmt(Grid.TextMatrix(Row, C_IDMOV)) = 0 Or lidComp = 0 Then   'algo falló
         Exit Sub
      End If
   End If
   
   idMov = vFmt(Grid.TextMatrix(Row, C_IDMOV))
   Fecha = GetTxDate(Tx_Fecha)

   'vemos si hay algún mov. de activo fijo asociado a este doc
   n = CountActFijo(idMov)
   
   If n = 0 Then           'no hay ninguno aún, llamamos al form de activo fijo que permite crear un mov directamente, sin pasar por la lista
      
      Set Frm = New FrmActivoFijo
      
      If vFmt(Grid.TextMatrix(Row, C_DEBE)) <> 0 Then
         ValNeto = vFmt(Grid.TextMatrix(Row, C_DEBE))
      Else
         ValNeto = vFmt(Grid.TextMatrix(Row, C_HABER))
      End If
      
      IdCuenta = Val(Grid.TextMatrix(Row, C_IDCUENTA))
      If Not EsCuentaActFijo(IdCuenta) Then
         IdCuenta = 0
      End If
        
      If Frm.FNewFromComp(lidComp, idMov, Fecha, ValNeto, 0, Grid.TextMatrix(Row, C_GLOSA), IdCuenta) = vbOK Then
         Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO
      Else
         Grid.TextMatrix(Row, C_DETACTFIJO) = ""
      End If
      
      Set Frm = Nothing
   
   Else              'hay uno o más => llamamos a la lista de activos fijos para este mov
      Set Frm = New FrmLstActFijo
      Call Frm.FEditFromComp(lidComp, Val(Grid.TextMatrix(Row, C_IDMOV)), Fecha)
      Set Frm = Nothing
 
   End If

End Sub

Private Function CountActFijo(ByVal idMov As Long) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   CountActFijo = 0
   
   If lidComp = 0 Or idMov = 0 Then
      Exit Function
   End If
      
   Q1 = "SELECT Count(*) FROM MovActivoFijo WHERE IdComp = " & lidComp & " AND IdMovComp = " & idMov
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      CountActFijo = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)

End Function

Private Sub GridActivoFijo(ByVal IdCuenta As Long, ByVal Row As Integer, Col As Integer)
   Dim HayCtaActFijo As Boolean
      
   If Col <> C_CODCUENTA And Col <> C_LSTCUENTA Then
      Exit Sub
   End If
   
   'vemos si es cuenta de activo fijo
   If EsCuentaActFijo(IdCuenta) Then
               
      Call ActivoFijo(Row, Col)
   
   Else     'no es de activo fijo, vemos si tenía un activo fijo definido, para borrarlo
      If CountActFijo(Val(Grid.TextMatrix(Row, C_IDMOV))) > 0 Then
         
         If MsgBox1("Esta cuenta NO es de Activo Fijo. Se eliminarán los detalles de Activo Fijo asociados a este movimiento." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
'            Call ExecSQL(DbMain, "DELETE * FROM MovActivoFijo WHERE IdComp = " & lidComp & " AND IdMovComp = " & Val(Grid.TextMatrix(Row, C_IDMOV)))
            Call DeleteSQL(DbMain, "MovActivoFijo", " WHERE IdComp = " & lidComp & " AND IdMovComp = " & Val(Grid.TextMatrix(Row, C_IDMOV)) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
         End If
              
      End If
      
      Grid.TextMatrix(Row, C_DETACTFIJO) = ""
      
   End If
End Sub

Public Function FPrtComp(ByVal idcomp As Long)
   
   
   lEnImpresionMasiva = True
   
   lOper = O_VIEW
   lidComp = idcomp
   lCompTipo = False
   If gFunciones.ComprobanteResumido Then
      lPrtResumido = True   'indica que si el comprobante tiene Ch_ImpRes en True, debe imprimir resumido, si no, extendido
   End If
   
   Call SetTblName(lCompTipo)
   
   Load Me
   
   'Call Bt_Print_Click    'ahora se hace en Form_Load
   
   lEnImpresionMasiva = False
   
   DoEvents
   
   Unload Me
   
End Function
Public Function FViewResumido(ByVal idcomp As Long) As Integer
   
   lOper = O_VIEW
   lidComp = idcomp
   lCompTipo = False
   
   lViewResumido = True
   
  ' lCompTipo = False
   Call SetTblName(lCompTipo)
   
   Me.Show vbModal
   
   FViewResumido = lRc
   
   
   
End Function

Private Sub GridAtribCuenta(ByVal Row As Integer)

   If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CONCILIACION) <> 0 Then
      Grid.TextMatrix(Row, C_ATRIB_CONCIL) = 1
   Else
      Grid.TextMatrix(Row, C_ATRIB_CONCIL) = 0
   End If
   
   If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_AREANEG) = 0 Then
      Grid.TextMatrix(Row, C_IDAREANEG) = 0
      Grid.TextMatrix(Row, C_AREANEG) = ""
   End If
      
   If GetAtribCuenta(Grid.TextMatrix(Row, C_IDCUENTA), ATRIB_CCOSTO) = 0 Then
      Grid.TextMatrix(Row, C_IDCCOSTO) = 0
      Grid.TextMatrix(Row, C_CCOSTO) = ""
   End If

End Sub

Private Function ValidIdCtaBanco(ByVal Row As Integer) As Boolean
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   
   ValidIdCtaBanco = True
   
   Q1 = "SELECT IdCtaBanco FROM Documento WHERE IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC))
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
'   Q2 = Replace(Q1, "Documento", "DocumentoFull")
'   Q1 = Q1 & " UNION ALL " & Q2
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      If vFld(Rs("IdCtaBanco")) <> 0 And vFld(Rs("IdCtaBanco")) <> Val(Grid.TextMatrix(Row, C_IDCUENTA)) Then
         ValidIdCtaBanco = False
         'borramos correlativo cheque por si las moscas
         Q1 = "UPDATE Cuentas SET CorrelativoCheque = 0 WHERE IdCuenta = " & vFld(Rs("IdCtaBanco"))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   Call CloseRs(Rs)

End Function
Private Sub GrabarPrtMovDet()
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PRTMOVDET'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      'actualizamos
      Q1 = "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & gPrtMovDetOpt & "' WHERE Tipo = 'PRTMOVDET'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
   Else
      'insertamos
      Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('PRTMOVDET', 0, '" & gPrtMovDetOpt & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
   End If

   Call CloseRs(Rs)
   

End Sub
Private Sub DejarEnCeroMovs()
   Dim i As Integer

   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ORDEN) = "" Then
         Exit For
      End If
      If Grid.RowHeight(i) > 0 Then     ' no está borrado
         If vFmt(Grid.TextMatrix(i, C_DEBE)) > 0 Then
            Grid.TextMatrix(i, C_DEBE) = 0
         End If
         If vFmt(Grid.TextMatrix(i, C_HABER)) > 0 Then
            Grid.TextMatrix(i, C_HABER) = 0
         End If
      End If
      
      Call FGrModRow(Grid, i, FGR_U, C_IDMOV, C_UPDATE)
   Next i
         
   GridTot.TextMatrix(0, C_DEBE) = 0
   GridTot.TextMatrix(0, C_HABER) = 0
End Sub

Private Sub AsignaCuenta(ByVal Row As Integer, ByVal Col As Integer)
   Dim FrmPlan As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim DescCta As String
   Dim NombCuenta As String
   Dim CodCta As String
   Dim CodiF2214Ter As String

   Set FrmPlan = New FrmPlanCuentas

   If FrmPlan.FSelEdit(IdCuenta, CodCta, DescCta, NombCuenta, False) = vbOK Then
      If DescCta <> "" Then
         
         If Val(Grid.TextMatrix(Grid.Row, C_ORDEN)) = 0 Then    'nuevo
            Call AddNewMov(Grid.Row)
         End If

         Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
         Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
         Grid.TextMatrix(Row, C_CUENTA) = DescCta

         Call FGrModRow(Grid, Row, FGR_U, C_IDMOV, C_UPDATE)
         
         Call GridActivoFijo(IdCuenta, Row, Col)

         Call GridAtribCuenta(Row)

     End If

   End If
   Set FrmPlan = Nothing
   
   
      '   ***** Ado 2699584 Tema 1 del 3.4
            If Grid.ColSel = C_CODCUENTA Or Grid.ColSel = C_LSTCUENTA Then
            CodiF2214Ter = ""
            Call CodF2214Ter(Grid.TextMatrix(Grid.RowSel, C_CODCUENTA), CodiF2214Ter)

            If CodiF2214Ter = "1" Then 'Or CodiF2214Ter = "629" Then
                If MsgBox1("¡Atención! ¿Partida Ingresada corresponde a una Participación?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                    Exit Sub
                Else
                    Dim Frm As FrmPercepciones
                    Set Frm = New FrmPercepciones
                    Frm.CodCta = Grid.TextMatrix(Grid.RowSel, C_IDCUENTA)
                    Frm.GIdPerc = 0
                    Frm.Fecha = Tx_Fecha
                    Frm.orden = Grid.TextMatrix(Grid.RowSel, C_ORDEN)
                    Frm.idcomp = lidComp
                    Frm.Show vbModal
                    Set Frm = Nothing
                End If
            End If
            End If
'   ***** Fin Ado 2699584 Tema 1 del 3.4

End Sub

Private Function ValidaOtrosIngEg() As Boolean

   Dim EsOtrosIngEg As Boolean
   
   EsOtrosIngEg = gEmpresa.Franq14Ter
   EsOtrosIngEg = EsOtrosIngEg And CbItemData(Cb_Tipo) = TC_EGRESO Or CbItemData(Cb_Tipo) = TC_INGRESO
   EsOtrosIngEg = EsOtrosIngEg And Not lHayDocLibros
   EsOtrosIngEg = EsOtrosIngEg And lCuentasDisponible
   
   ValidaOtrosIngEg = EsOtrosIngEg
   
End Function

'2860036
Public Sub PrtFooterMembreteRigth(PrtPage As Object, ByVal StrTitMembrete As String, ByVal StrTextMembrete As String, ByVal RightX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   PrtPage.CurrentY = PrtPage.Height - 1700
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   PrtPage.CurrentX = RightX - 2200
   PrtPage.Print "------------------------------------------"
    PrtPage.CurrentX = RightX - 2200
   PrtPage.Print StrTitMembrete
    PrtPage.CurrentX = RightX - 2200
   PrtPage.Print StrTextMembrete
   
   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub
'fin 2860036
'2860036
Public Sub PrtFooterMembreteLeft(PrtPage As Object, ByVal StrTitMembrete As String, ByVal StrTextMembrete As String, ByVal LeftX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   PrtPage.CurrentY = PrtPage.Height - 1700
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   'PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print "------------------------------------------"
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print StrTitMembrete
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print StrTextMembrete
   
   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

'2861591
Public Function FNewCompActivo(ByVal Nlinea As Integer, ByVal vDebe As Double, ByVal vHaber As Double, ByVal vIdCuenta As String, ByVal vCuenta As String, ByVal vDescCuenta As String) As Integer
    Dim i As Integer
   lOper = O_NEW
   Call SetTblName(False)

   lCompTipo = False
   lidComp = 0
   lCorrelativo = 0
      Grid.TextMatrix(Nlinea, C_ORDEN) = Nlinea
      Grid.TextMatrix(Nlinea, C_IDCUENTA) = vIdCuenta
      Grid.TextMatrix(Nlinea, C_CODCUENTA) = FmtCodCuenta(vCuenta)
      Grid.TextMatrix(Nlinea, C_CUENTA) = vDescCuenta
      If vDebe > 0 Then
        Grid.TextMatrix(Nlinea, C_DEBE) = Format(vDebe, BL_NUMFMT)
        Call Grid_AcceptValue(Nlinea, C_DEBE, Format(vDebe, BL_NUMFMT), 1)
        Call CalcTot
      Else
        Grid.TextMatrix(Nlinea, C_DEBE) = Format(0, BL_NUMFMT)
        Call Grid_AcceptValue(Nlinea, C_DEBE, Format(0, BL_NUMFMT), 1)
        Call CalcTot
      End If

      If vHaber > 0 Then
      Grid.TextMatrix(Nlinea, C_HABER) = Format(vHaber, BL_NUMFMT)
      Call Grid_AcceptValue(Nlinea, C_HABER, Format(vHaber, BL_NUMFMT), 1)
         Call CalcTot
         Else
      Grid.TextMatrix(Nlinea, C_HABER) = Format(0, BL_NUMFMT)
      Call Grid_AcceptValue(Nlinea, C_HABER, Format(0, BL_NUMFMT), 1)
         Call CalcTot
      End If


   FNewCompActivo = lRc
End Function
'2861591

Private Sub CalcTotFull(ByVal vRow As Double, ByVal vCol As Double, ByVal vValor As Double, ByVal vTipo As Integer) 'vtipo valor 1 = debe / 2 = haber / 0 = no toma en consideracion vValor
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   TotDebe = 0
   TotHaber = 0
   
    Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante WHERE IdComp = " & lidComp
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
      
     ' ltotalDebeFull = TotDebe
      'ltotalHaberFull = TotHaber
   End If
   
   Call CloseRs(Rs)
         
    If vTipo = 1 Then
      If vFmt(GridTotFull.TextMatrix(0, C_DEBE)) <> TotDebe Then
        'ltotalDebeFull = (TotDebe - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
        ltotalDebeFull = (vFmt(GridTotFull.TextMatrix(0, C_DEBE)) - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
      Else
        ltotalDebeFull = (TotDebe - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
        'ltotalDebeFull = (vFmt(GridTotFull.TextMatrix(0, C_DEBE)) - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
      End If
        
    ElseIf vTipo = 2 Then
      If vFmt(GridTotFull.TextMatrix(0, C_HABER)) <> TotHaber Then
        'ltotalHaberFull = (TotHaber - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
        ltotalHaberFull = (vFmt(GridTotFull.TextMatrix(0, C_HABER)) - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
      Else
        ltotalHaberFull = (TotHaber - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
        'ltotalHaberFull = (vFmt(GridTotFull.TextMatrix(0, C_HABER)) - vValor) + vFmt(Grid.TextMatrix(vRow, vCol))
      End If
    End If
    
    If ltotalDebeFull = 0 Then
     ltotalDebeFull = TotDebe
    End If
    If ltotalHaberFull = 0 Then
     ltotalHaberFull = TotHaber
    End If
    
         
   GridTotFull.TextMatrix(0, C_DEBE) = Format(ltotalDebeFull, BL_NUMFMT)
   GridTotFull.TextMatrix(0, C_HABER) = Format(ltotalHaberFull, BL_NUMFMT)
   
End Sub

Private Sub InsertComprCentraFull(ByVal vIdComp As Double)

   Dim Q1 As String
   Dim Rs As Recordset
   
   On Error Resume Next

    ERR.Clear
    
   
    Q1 = "INSERT INTO tbl_Comp_Centra_Full(IdComp,IdEmpresa,Tipo,Fecha,ano) "
    Q1 = Q1 & "VALUES "
    Q1 = Q1 & "(" & vIdComp & "," & gEmpresa.id & ",1," & CLng(Int(Now)) & "," & gEmpresa.Ano & ")"
    
    Call ExecSQL(DbMain, Q1)
   
End Sub

Private Function ValidaCompFull(ByVal vIdComp As Double) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   ValidaCompFull = False
   
    Q1 = "SELECT count(*) as cant FROM tbl_Comp_Centra_Full WHERE IdComp = " & vIdComp
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
     If vFld(Rs("cant")) = 1 Then
       ValidaCompFull = True
     End If
   End If
   
   Call CloseRs(Rs)
         
   
End Function

'este metodo evita que se cierre formulario comprobante, esto ayuda a no generar comprobantes con correlativo = 0
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then

Cancel = True

End If

End Sub
