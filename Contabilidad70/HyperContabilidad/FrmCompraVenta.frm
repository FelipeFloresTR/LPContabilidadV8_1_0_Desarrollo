VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmCompraVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros Compras y Ventas"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "FrmCompraVenta.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Bt_Centralizar_Full 
      Caption         =   "&Centralizar"
      Height          =   675
      Left            =   12000
      Picture         =   "FrmCompraVenta.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Generar un comprobante con todos los documentos con tick"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Bt_ToLeft 
      Height          =   375
      Left            =   7020
      Picture         =   "FrmCompraVenta.frx":0521
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Anterior conjunto de registros"
      Top             =   7800
      Width           =   315
   End
   Begin VB.CommandButton Bt_ToRight 
      Height          =   375
      Left            =   7440
      Picture         =   "FrmCompraVenta.frx":082B
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Siguente conjunto de registros"
      Top             =   7800
      Width           =   315
   End
   Begin VB.CommandButton Bt_DelAll 
      Caption         =   "Eliminar Todo..."
      Height          =   315
      Left            =   9360
      TabIndex        =   74
      ToolTipText     =   "Eliminar todos los documentos en estado Pendiente"
      Top             =   7800
      Width           =   1275
   End
   Begin VB.Timer Tm_ColWi 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12900
      Top             =   1080
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Capturar..."
      Enabled         =   0   'False
      Height          =   315
      Left            =   8040
      TabIndex        =   40
      ToolTipText     =   "Capturar documentos desde archivo de texto"
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton Bt_HlpImport 
      Caption         =   "?"
      Enabled         =   0   'False
      Height          =   315
      Left            =   9060
      TabIndex        =   41
      ToolTipText     =   "Formato del archivo de importación"
      Top             =   7800
      Width           =   255
   End
   Begin VB.CommandButton Bt_Opciones 
      Caption         =   "Opciones de Vista/Ediciï¿½n"
      Height          =   315
      Left            =   10800
      TabIndex        =   42
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones de Ediciï¿½n"
      Height          =   4335
      Left            =   10800
      TabIndex        =   61
      Top             =   3420
      Width           =   2295
      Begin VB.CheckBox Ch_ViewPropIVA 
         Caption         =   "Ver Prop. IVA"
         Height          =   195
         Left            =   300
         TabIndex        =   71
         Top             =   3540
         Width           =   1845
      End
      Begin VB.CheckBox Ch_ViewCantBoletas 
         Caption         =   "Ver Cantidad Boletas"
         Height          =   195
         Left            =   300
         TabIndex        =   69
         Top             =   2820
         Width           =   1905
      End
      Begin VB.CheckBox Ch_ViewMaqReg 
         Caption         =   "Ver Mï¿½q. Registradora"
         Height          =   195
         Left            =   300
         TabIndex        =   68
         Top             =   2460
         Width           =   1905
      End
      Begin VB.CommandButton Bt_CerrarOpt 
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   120
         Width           =   195
      End
      Begin VB.CheckBox Ch_ViewOtrosImp 
         Caption         =   "Ver Otros Impuestos"
         Height          =   195
         Left            =   300
         TabIndex        =   67
         Top             =   2100
         Width           =   1845
      End
      Begin VB.CheckBox Ch_ViewNumInterno 
         Caption         =   "Ver Nï¿½ Interno"
         Height          =   195
         Left            =   300
         TabIndex        =   66
         Top             =   1740
         Width           =   1545
      End
      Begin VB.CheckBox Ch_ViewDocHasta 
         Caption         =   "Ver Nï¿½Doc Hasta"
         Height          =   195
         Left            =   300
         TabIndex        =   65
         Top             =   1380
         Width           =   1545
      End
      Begin VB.CheckBox Ch_ViewSucursal 
         Caption         =   "Ver  Sucursal"
         Height          =   195
         Left            =   300
         TabIndex        =   64
         Top             =   1020
         Width           =   1275
      End
      Begin VB.CheckBox Ch_ViewExento 
         Caption         =   "Ver  Exento"
         Height          =   195
         Left            =   300
         TabIndex        =   63
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox Ch_ViewDTE 
         Caption         =   "Ver  DTE"
         Height          =   195
         Left            =   300
         TabIndex        =   62
         Top             =   300
         Width           =   1020
      End
      Begin VB.CheckBox Ch_ViewDetOtrosImp 
         Caption         =   "Ver detalle Otros Imp."
         Height          =   195
         Left            =   300
         TabIndex        =   70
         Top             =   3180
         Width           =   1845
      End
      Begin VB.CheckBox Ch_RepetirGlosa 
         Caption         =   "Repetir Glosa"
         Height          =   195
         Left            =   300
         TabIndex        =   72
         Top             =   3900
         Width           =   1455
      End
   End
   Begin VB.CommandButton Bt_ExitNewDoc 
      Caption         =   "&Nuevo Doc"
      Height          =   675
      Left            =   10320
      Picture         =   "FrmCompraVenta.frx":0B35
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Tx_CurrCell 
      Height          =   315
      Left            =   0
      TabIndex        =   56
      Top             =   7800
      Width           =   6915
   End
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   12960
      Picture         =   "FrmCompraVenta.frx":10AD
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   50
      Top             =   660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Bt_Centralizar 
      Caption         =   "&Centralizar"
      Height          =   675
      Left            =   12000
      Picture         =   "FrmCompraVenta.frx":1124
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Generar un comprobante con todos los documentos con tick"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   13095
      Begin VB.ComboBox Cb_Sucursal 
         Height          =   315
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   38
         ToolTipText     =   "Sucursales"
         Top             =   180
         Width           =   2175
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   8760
         TabIndex        =   39
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Fr_BtEdit 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   60
         TabIndex        =   55
         Top             =   180
         Width           =   3435
         Begin VB.CommandButton Bt_SelEnt 
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
            Picture         =   "FrmCompraVenta.frx":1639
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Seleccionar Entidad"
            Top             =   0
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
            Left            =   420
            Picture         =   "FrmCompraVenta.frx":1AD7
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Plan de Cuentas"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_AnulaDoc 
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
            Left            =   2640
            Picture         =   "FrmCompraVenta.frx":1E98
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anular documento seleccionado"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_TipoDoc 
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
            Left            =   0
            Picture         =   "FrmCompraVenta.frx":2309
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Seleccionar tipo de documento"
            Top             =   0
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
            Left            =   2220
            Picture         =   "FrmCompraVenta.frx":26AE
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Duplicar documento seleccionado"
            Top             =   0
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
            Left            =   3060
            Picture         =   "FrmCompraVenta.frx":2B00
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Eliminar documento seleccionado"
            Top             =   0
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
            Left            =   1320
            Picture         =   "FrmCompraVenta.frx":2EFC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Copiar dato"
            Top             =   0
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
            Left            =   1740
            Picture         =   "FrmCompraVenta.frx":32DC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Pegar dato copiado"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Fr_BtGen 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   3540
         TabIndex        =   54
         Top             =   180
         Width           =   4875
         Begin VB.CommandButton Bt_DocCuotas 
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
            Left            =   840
            Picture         =   "FrmCompraVenta.frx":36C5
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Ver/Editar detalle de Cuotas Documento"
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Bt_DetDoc 
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
            Picture         =   "FrmCompraVenta.frx":3BA2
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Detalle documento seleccionado"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_Resumen 
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
            Left            =   1800
            Picture         =   "FrmCompraVenta.frx":4007
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Resumen IVA Compras - Ventas y Otros Impuestos"
            Top             =   0
            Width           =   375
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
            Left            =   4440
            Picture         =   "FrmCompraVenta.frx":447F
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Calendario"
            Top             =   0
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
            Left            =   3600
            Picture         =   "FrmCompraVenta.frx":48A8
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Convertir moneda"
            Top             =   0
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
            Left            =   4020
            Picture         =   "FrmCompraVenta.frx":4C46
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Calculadora"
            Top             =   0
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
            Left            =   1380
            Picture         =   "FrmCompraVenta.frx":4FA7
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Sumar datos seleccionados"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_CopyExcel 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            Picture         =   "FrmCompraVenta.frx":504B
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Copiar Excel"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_Print 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2700
            Picture         =   "FrmCompraVenta.frx":5490
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Imprimir"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton Bt_Preview 
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
            Picture         =   "FrmCompraVenta.frx":594A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Vista previa de la impresión"
            Top             =   0
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
            Left            =   420
            Picture         =   "FrmCompraVenta.frx":5DF1
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Ver/agregar detalle de Activo Fijo asociado al documento seleccionado"
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   11280
         TabIndex        =   43
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   12180
         TabIndex        =   44
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Lb_Sucursal 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   8000
         TabIndex        =   60
         Top             =   260
         Width           =   660
      End
   End
   Begin VB.PictureBox Pc_Cent 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   12960
      Picture         =   "FrmCompraVenta.frx":61EF
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   46
      Top             =   900
      Visible         =   0   'False
      Width           =   375
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5355
      Left            =   30
      TabIndex        =   18
      Top             =   1980
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   9446
      Cols            =   46
      Rows            =   30
      FixedCols       =   0
      FixedRows       =   2
      ScrollBars      =   3
      AllowUserResizing=   1
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7380
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   11
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Frame Fr_List 
      ForeColor       =   &H00FF0000&
      Height          =   1300
      Left            =   0
      TabIndex        =   47
      Top             =   660
      Width           =   11835
      Begin VB.CheckBox Ch_CentralizacionFull 
         Caption         =   "Centralizacion full"
         Height          =   195
         Left            =   9960
         TabIndex        =   80
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Tx_NumDocAsoc 
         Height          =   315
         Left            =   4260
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   900
         Width           =   1275
      End
      Begin VB.ComboBox Cb_DTE 
         Height          =   315
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   1155
      End
      Begin VB.CheckBox Ch_EsSupermercado 
         Caption         =   "Supermercado"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   3540
         TabIndex        =   2
         Top             =   240
         Width           =   225
      End
      Begin VB.TextBox Tx_Descrip 
         Height          =   315
         Left            =   900
         MaxLength       =   100
         TabIndex        =   10
         Top             =   900
         Width           =   2175
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   8220
         TabIndex        =   13
         Top             =   900
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   5640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   7500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   2595
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   4260
         MaxLength       =   12
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   8220
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   4260
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton Bt_List 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   10320
         Picture         =   "FrmCompraVenta.frx":6653
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   2175
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc. Asoc.:"
         Height          =   195
         Index           =   6
         Left            =   3120
         TabIndex        =   78
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DTE:"
         Height          =   195
         Index           =   5
         Left            =   5640
         TabIndex        =   75
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   3780
         TabIndex        =   59
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrip.:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Afecto:"
         Height          =   195
         Index           =   9
         Left            =   7500
         TabIndex        =   57
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   4
         Left            =   7500
         TabIndex        =   53
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc.:"
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   52
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   675
      Left            =   12000
      Picture         =   "FrmCompraVenta.frx":6A91
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Menu M_TipoDoc 
      Caption         =   "Tipo Documento"
      Visible         =   0   'False
      Begin VB.Menu M_ItTipoDoc 
         Caption         =   "TipoDoc0"
         Index           =   0
      End
   End
   Begin VB.Menu M_Cuenta 
      Caption         =   "Cuenta"
      Visible         =   0   'False
      Begin VB.Menu M_ItCuenta 
         Caption         =   "Cuenta0"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmCompraVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOC = 0
Const C_NUMLIN = 1
Const C_FECHA = 2
Const C_IDTIPODOC = 3
Const C_TIPODOC = 4
Const C_DOCIMPEXP = 5
Const C_GIRO = 6
Const C_DTE = 7
Const C_NUMFISCIMPR = 8
Const C_NUMINFORMEZ = 9
Const C_NUMDOC = 10
Const C_NUMDOCHASTA = 11
Const C_CANTBOLETAS = 12
Const C_IDPROPIVA = 13
Const C_PROPIVA = 14
Const C_FECHAEMIORI = 15
Const C_LNGFECHAEMIORI = 16
Const C_CHECK = 17
Const C_RUT = 18
Const C_NOMBRE = 19
Const C_IDENTIDAD = 20
Const C_DESCRIP = 21
Const C_IDSUCURSAL = 22
Const C_SUCURSAL = 23
Const C_AFECTO = 24
Const C_AF_IDCUENTA = 25
Const C_AF_CODCUENTA = 26
Const C_AF_CUENTA = 27
Const C_EXENTO = 28
Const C_EX_IDCUENTA = 29
Const C_EX_CODCUENTA = 30
Const C_EX_CUENTA = 31
Const C_IVA = 32
Const C_IVA_IDCUENTA = 33
Const C_OTROIMP = 34
Const C_OIMP_IDCUENTA = 35
Const C_IDANEG_CCOSTO = 36
Const C_INIDETOTROIMP = 37
Const C_ENDDETOTROIMP = MAX_COL ' 64 o C_INIDETOTROIMP + MAX_COLOTROIMP
Const C_TOTAL = C_ENDDETOTROIMP + 1
Const C_TOT_IDCUENTA = C_ENDDETOTROIMP + 2
Const C_TOT_CODCUENTA = C_ENDDETOTROIMP + 3
Const C_TOT_CUENTA = C_ENDDETOTROIMP + 4
Const C_VENTASACUM = C_ENDDETOTROIMP + 5
Const C_DETALLE = C_ENDDETOTROIMP + 6
Const C_FECHAVENC = C_ENDDETOTROIMP + 7
Const C_LNGFECHAVENC = C_ENDDETOTROIMP + 8
Const C_CORRINTERNO = C_ENDDETOTROIMP + 9
Const C_DETACTFIJO = C_ENDDETOTROIMP + 10
Const C_ESTADO = C_ENDDETOTROIMP + 11
Const C_IDESTADO = C_ENDDETOTROIMP + 12
Const C_DOCASOC = C_ENDDETOTROIMP + 13
Const C_USUARIO = C_ENDDETOTROIMP + 14
Const C_MOVEDITED = C_ENDDETOTROIMP + 15
Const C_IDCOMPCENT = C_ENDDETOTROIMP + 16
Const C_IDCOMPPAGO = C_ENDDETOTROIMP + 17
Const C_MSGACTFIJO = C_ENDDETOTROIMP + 18
Const C_EXPORTED = C_ENDDETOTROIMP + 19
Const C_UPDATE = C_ENDDETOTROIMP + 20



Public ColNum As Integer

Const NCOLS = C_UPDATE

Const TX_ACTFIJO = "AF >>"
Const TX_DETALLE = ">>"

Const O_VIEWLIBLEGAL = -1

Const MITEM_OTRA = "(Otra)..."

Dim lOper As Integer
Dim lRc As Integer
Dim lTipoLib As Integer
Dim lMes As Integer
Dim lAno As Integer
Dim LDia As Integer
Dim lLibOf As Integer

Dim lIdDoc As Long
'Dim lIdEnt As Long

Const MAX_CUENTAS = 50
Dim lCuentas(MAX_CUENTAS) As Long

'cuentas default
Dim lCtaAfecto As Cuenta_t
Dim lCtaExento As Cuenta_t
Dim lCtaTotal As Cuenta_t

Dim lIdCuentaIVA As Long
Dim lIdCuentaIVAIrrec As Long
Dim lIdCuentaOtrosImp As Long
Dim lIdCuentaOtrosImpFacCompra As Long

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3
Dim lcbNombre As ClsCombo

Dim lOrdenGr(C_UPDATE) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lWCodCuenta As Integer
Dim lWCuenta As Integer

Dim lWhere As String

'Dim lCurReg As Long     'primer registro del ragno de registros de la página actual
'Dim lNumReg As Long     'cantidad de registros en la grilla
'Dim lToRightPressed As Boolean   'para indicar que se presionó Bt_ToRight

Dim lClsPaging As ClsPaging   'Clase de paginamiento


'indican si se puede editar/adm/ingresar en el libro actual, de acuerdo a los privilegios y a si está o no locked por otro usuario
Dim lEditEnabled As Boolean         'editar (todo)
Dim lAdmDocsEnabled As Boolean      'administrar (botón Bt_Centralizar)
Dim lIngDocsEnabled As Boolean      'ingresar (botón Bt_ExitNewDoc)

Dim lSucursal As Long

Dim lMsgPagadoNoCent As Boolean     'para no desplegar más el mensaje de pagado pero no centralizado

Dim lInBt_DetDoc As Boolean         'para indicar que no valide detalles que se ingresan en el detalle de documento

Dim lMsgFechaErr As Boolean         'para indicar si ya se mandó el mensaje "la fecha de ingreso o recepción del documento es posterior a los dos meses siguientes de la fecha de emisión del mismo"

Dim lMsgNotaCred As Boolean         'para indicar que ya no deseo ver mensaje de Nota de Crédito

#If DATACON = 1 Then
Dim lDbAnoAnt As Database           'base de datos año anterior
#End If

Dim lFNameLogImp As String

Dim lCurWhere As String             'Where de la lista actual

Dim lDetOtrosImpFilled As Boolean   'indica si se llenó el detalle de otros impuestos o no

Dim lTitEspecialVentas As String    ' Título especial para libros de venta legales

Dim lHayPropIVA As Boolean          'para ver si ofrecemos recálculo de proporcionalidad de IVA

'3340329
Dim vTotalRemMesAnt As Double
'3340329

'3389677 FP
Dim remMesAnt As Boolean


Private Sub Bt_ActivoFijo_Click()
   Dim Frm As FrmLstActFijo
   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If EsCuentaActFijo(Val(Grid.TextMatrix(Grid.Row, C_EX_IDCUENTA))) Or EsCuentaActFijo(Val(Grid.TextMatrix(Grid.Row, C_AF_IDCUENTA))) Then
      'cuentas de activo en exento o afecto
      
      If lOper <> O_EDIT Then
         Set Frm = New FrmLstActFijo
         Call FrmLstActFijo.FViewFromDoc(Grid.TextMatrix(Grid.Row, C_IDDOC), lTipoLib)
         Set Frm = Nothing
      Else
         Call ActivoFijo(Grid.Row, Grid.Col)
      End If
   Else
      MsgBox1 "Las cuentas de afecto o exento del documento seleccionado no son de Activo Fijo.", vbExclamation
      
   End If
   
End Sub

Private Sub Bt_AnulaDoc_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   'If Grid.TextMatrix(Grid.Row, C_IDDOC) = "" Then
   '   MsgBox1 "Este documento acaba de ser ingresado. Elimínelo directamente con el botón que sigue.", vbExclamation + vbOKOnly
   '   Exit Sub
   'End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_IDESTADO)) = ED_CENTRALIZADO And lAno = gEmpresa.Ano Then
      MsgBox1 "Este documento no se puede anular, ya que ha sido centralizado en un comprobante.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If Val(Grid.TextMatrix(Grid.Row, C_IDESTADO)) = ED_PAGADO Then
      MsgBox1 "Este documento no se puede anular, ya que ha sido pagado en un comprobante.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea anular este documento?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Grid.TextMatrix(Grid.Row, C_ESTADO) = gEstadoDoc(ED_ANULADO)
   Grid.TextMatrix(Grid.Row, C_IDESTADO) = ED_ANULADO
   Grid.TextMatrix(Grid.Row, C_DESCRIP) = "NULO"
   Grid.TextMatrix(Grid.Row, C_RUT) = ""
   Grid.TextMatrix(Grid.Row, C_NOMBRE) = "NULO"
   
   Grid.TextMatrix(Grid.Row, C_AFECTO) = 0
   Grid.TextMatrix(Grid.Row, C_AF_IDCUENTA) = 0
   Grid.TextMatrix(Grid.Row, C_AF_CODCUENTA) = ""
   Grid.TextMatrix(Grid.Row, C_AF_CUENTA) = ""
   
   Grid.TextMatrix(Grid.Row, C_EXENTO) = 0
   Grid.TextMatrix(Grid.Row, C_EX_IDCUENTA) = 0
   Grid.TextMatrix(Grid.Row, C_EX_CODCUENTA) = ""
   Grid.TextMatrix(Grid.Row, C_EX_CUENTA) = ""
   
   Grid.TextMatrix(Grid.Row, C_IVA) = 0
   Grid.TextMatrix(Grid.Row, C_IVA_IDCUENTA) = 0
   
   Grid.TextMatrix(Grid.Row, C_OTROIMP) = 0
   Grid.TextMatrix(Grid.Row, C_OIMP_IDCUENTA) = 0
   
   Grid.TextMatrix(Grid.Row, C_TOTAL) = 0
   Grid.TextMatrix(Grid.Row, C_TOT_IDCUENTA) = 0
   Grid.TextMatrix(Grid.Row, C_TOT_CODCUENTA) = ""
   Grid.TextMatrix(Grid.Row, C_TOT_CUENTA) = ""
   
   Call FGrModRow(Grid, Grid.Row, FGR_U, C_IDDOC, C_UPDATE)
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Centralizar_Click()

   Dim i As Integer
   Dim StrIdDoc As String
   Dim idcomp As Long
   Dim Rc As Integer
   Dim HayPropIVA As Boolean
   
   If Not ValidaIngresoComp() Then
      Exit Sub
   End If
      
   HayPropIVA = False
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_IDDOC) = "" Then
         Exit For
      End If
      
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
         StrIdDoc = StrIdDoc & "," & Grid.TextMatrix(i, C_IDDOC)
         
         If lTipoLib = LIB_COMPRAS And (Val(Grid.TextMatrix(i, C_IDPROPIVA)) = PIVA_NULO Or Val(Grid.TextMatrix(i, C_IDPROPIVA)) = PIVA_PROP) And Not HayPropIVA Then   '(si es N o P)
            HayPropIVA = True
         End If
      End If
      
   Next i
      
   If StrIdDoc = "" Then
      'MsgBox1 "No hay documentos marcados para centralizar.", vbExclamation + vbOKOnly
      
      If MsgBox1("¿Está seguro que desea centralizar todos los documentos pendientes, y los pagados pero no centralizados, en un solo comprobante?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      Else
         'marcamos todos los pendientes
         Grid.FlxGrid.Redraw = False
         For i = Grid.FixedRows To Grid.rows - 1
            If Grid.TextMatrix(i, C_IDDOC) = "" Then
               Exit For
            End If
            
            Grid.Row = i
            Grid.Col = C_CHECK
                        
            If Grid.TextMatrix(i, C_IDESTADO) = ED_PENDIENTE Or (Val(Grid.TextMatrix(i, C_IDESTADO)) = ED_PAGADO And vFmt(Grid.TextMatrix(i, C_IDCOMPCENT)) = 0) Then
               Call FGrSetPicture(Grid, i, C_CHECK, Pc_Check, 0)
               StrIdDoc = StrIdDoc & "," & Grid.TextMatrix(i, C_IDDOC)
               
               If lTipoLib = LIB_COMPRAS And (Val(Grid.TextMatrix(i, C_IDPROPIVA)) = PIVA_NULO Or Val(Grid.TextMatrix(i, C_IDPROPIVA)) = PIVA_PROP) And Not HayPropIVA Then    '(si es N o P)
                  HayPropIVA = True
               End If
               
            End If
            
         Next i
         Grid.FlxGrid.Redraw = True
         
         If StrIdDoc = "" Then
            MsgBox1 "No hay documentos para centralizar.", vbExclamation + vbOKOnly
            Exit Sub
         End If
         
      End If
      
   ElseIf MsgBox1("¿Está seguro que desea centralizar todos los documentos marcados en un solo comprobante?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
       Exit Sub
       
   End If
   
   If HayPropIVA Then
      If MsgBox1("Recuerde verificar  que el monto de IVA Irrecuperable esté clasificado en la cuenta contable que corresponde." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
   End If


   StrIdDoc = Mid(StrIdDoc, 2)

   Me.MousePointer = vbHourglass
   '3217885
   idcomp = GenComprobante(StrIdDoc, lTipoLib, CbItemData(Cb_Mes), Val(Cb_Ano))
   
   If idcomp > 0 Then
   
      If FrmComprobante.FEditCentraliz(idcomp, CbItemData(Cb_Mes), Val(Cb_Ano)) = vbOK Then
      
         'limpiamos los check
         For i = Grid.FixedRows To Grid.rows - 1
            If Grid.TextMatrix(i, C_IDDOC) = "" Then
               Exit For
            End If
            
            Grid.Row = i
            Grid.Col = C_CHECK
               
            If Grid.CellPicture <> 0 Then
               Set Grid.CellPicture = LoadPicture()
               If Val(Grid.TextMatrix(i, C_IDESTADO)) <> ED_PAGADO Then
                  Grid.TextMatrix(i, C_IDESTADO) = ED_CENTRALIZADO
                  Grid.TextMatrix(i, C_ESTADO) = gEstadoDoc(ED_CENTRALIZADO)
                  If Grid.ColWidth(C_CHECK) > 0 Then
                     Grid.TextMatrix(i, C_CHECK) = "C"
                  End If
               End If
               Grid.TextMatrix(i, C_IDCOMPCENT) = idcomp
            End If
            
         Next i
      
      End If
      
   Else
      MsgBox1 "Problemas al generar el comprobante.", vbExclamation + vbOKOnly
   End If
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Centralizar_Full_Click()
  Me.MousePointer = vbHourglass
Call CentralizarFull
Me.MousePointer = vbDefault
End Sub

Private Sub Bt_CerrarOpt_Click()
   Fr_Opciones.visible = False
End Sub

Private Sub bt_Copy_Click()
   Clipboard.Clear
   Clipboard.SetText Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col)

End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   'Call FGr2Clip(Grid, gTipoLib(lTipoLib) & vbTab & Cb_Mes & " " & Val(Cb_Ano))
   Clip = FGr2String(Grid, gTipoLib(lTipoLib) & vbTab & Cb_Mes & " " & Val(Cb_Ano), False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTot)
   
   Clipboard.Clear
   Clipboard.SetText Clip

End Sub

Private Sub Bt_Cuentas_Click()
   Dim IdCuenta As Long
   Dim Descrip As String
   Dim Nombre As String
   Dim Frm As FrmPlanCuentas
   Dim Col As Integer
   Dim Row As Integer
   Dim Codigo As String
   Dim Rc As Integer
   
   Col = Grid.Col
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Set Frm = New FrmPlanCuentas
      Call FrmPlanCuentas.FEdit(False)
      Set Frm = Nothing
   
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Or (Col <> C_EX_CODCUENTA And Col <> C_AF_CODCUENTA And Col <> C_TOT_CODCUENTA) Or Not ValidaEstadoEdit(Row) Then
      ValidaFExported (Row)
      
      Set Frm = New FrmPlanCuentas
      Call FrmPlanCuentas.FEdit(False)
      Set Frm = Nothing
      
      Exit Sub
   End If
   
   If Grid.Col = C_EX_CODCUENTA Or Grid.Col = C_AF_CODCUENTA Or Grid.Col = C_TOT_CODCUENTA Then
      Call LoadCuentasMenu(Grid.Col)
      If M_ItCuenta.Count > 1 Then
         Call PopupMenu(M_Cuenta, , Grid.FlxGrid.ColPos(Grid.Col) + Grid.Left + 200, Grid.FlxGrid.RowPos(Grid.Row) + Grid.Top + 100)
      Else
         MsgBox1 "No hay cuentas definidas para este item del libro. Defínalas en la configuración de la empresa.", vbExclamation + vbOKOnly
      End If
   End If


End Sub

Private Sub Bt_Cut_Click()
   Dim ValidCol As Boolean
   
   ValidCol = (Grid.Col = C_NUMDOC Or Grid.Col = C_NUMDOCHASTA Or Grid.Col = C_EXENTO Or Grid.Col = C_AFECTO Or Grid.Col = C_IVA Or Grid.Col = C_OTROIMP Or Grid.Col = C_DESCRIP)
   
   If Not ValidCol Then
      Exit Sub
   End If
   
   Clipboard.Clear
   Call Clipboard.SetText(Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col))
   
   Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = ""
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
   
   Call CalcTot
End Sub

Private Sub Bt_DocCuotas_Click()
   Dim Frm As FrmDocCuotas
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim Row As Integer
   Dim Msg As String
   Dim FVenc As Long
   Dim NumCuotas As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then    'registro en blanco
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_TIPODOC) = "NCC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCF" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDF" Then
      MsgBox1 "Este tipo de documento no permite pago en cuotas. " & Msg, vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDV" Or Grid.TextMatrix(Row, C_TIPODOC) = "DVB" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCE" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDE" Then
      MsgBox1 "Este tipo de documento no permite pago en cuotas. " & Msg, vbExclamation + vbOKOnly
      Exit Sub
   End If
   
      
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   If IdDoc = 0 Then
      If Not IsValidLine(Row, Msg) Then
         MsgBox1 "Línea " & Row - Grid.FixedRows + 1 & " inválida. " & Msg, vbExclamation + vbOKOnly
         Exit Sub
      End If
   End If
   
   If (lOper <> O_EDIT And lOper <> O_NEW) Or lEditEnabled = False Then
      Set Frm = New FrmDocCuotas
      Call Frm.FView(IdDoc)
      Set Frm = Nothing
      Exit Sub
   End If
      
   If Grid.TextMatrix(Row, C_UPDATE) <> "" Then
   
      If MsgBox1("Antes de ingresar al detalle de cuotas del documento se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If Not valida() Then
         Exit Sub
      End If
      
      Call SaveGrid
      
      IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   
   End If
   
   Set Frm = New FrmDocCuotas
   If Frm.FEdit(IdDoc, FVenc, NumCuotas) = vbOK Then
      Grid.TextMatrix(Row, C_LNGFECHAVENC) = FVenc
      Grid.TextMatrix(Row, C_FECHAVENC) = Format(FVenc, SDATEFMT)
   End If
   Set Frm = Nothing
   
   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un documento a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.RowHeight(Row) = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_CENTRALIZADO And lAno = gEmpresa.Ano Then
       '3292777
      If EliminarDocCentralizadoPagado(Grid.TextMatrix(Row, C_IDDOC)) = False Then
        MsgBox1 "Este documento no se puede eliminar, ya que ha sido centralizado en un comprobante.", vbExclamation + vbOKOnly
        Exit Sub
      End If
      '3292777
   
   
'      MsgBox1 "Este documento no se puede eliminar, ya que ha sido centralizado en un comprobante.", vbExclamation + vbOKOnly
'      Exit Sub
   End If
      
   If Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PAGADO And Val(Grid.TextMatrix(Row, C_IDCOMPPAGO)) > 0 Then
      '3292777
      If EliminarDocCentralizadoPagado(Grid.TextMatrix(Row, C_IDDOC)) = False Then
        MsgBox1 "Este documento no se puede borrar, ya que ha sido pagado en un comprobante.", vbExclamation + vbOKOnly
        Exit Sub
      End If
      '3292777
   
   
'      MsgBox1 "Este documento no se puede borrar, ya que ha sido pagado en un comprobante.", vbExclamation + vbOKOnly
'      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea borrar este documento?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot
End Sub

Private Sub Bt_DelAll_Click()
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   If MsgBox1("¿Está seguro que desea eliminar TODOS los documentos en estado Pendiente o Anulado?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
     
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(Row, C_FECHA) = "" Then
         Exit For
      End If

      If Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PENDIENTE Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO Then
         Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE, False)
         Grid.RowHeight(Row) = 0
      ElseIf Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_CENTRALIZADO Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PAGADO Then
        
        Q1 = " SELECT DISTINCT Comprobante.IdComp, Comprobante.Correlativo, Comprobante.Fecha,  Comprobante.Tipo, Comprobante.Estado, Comprobante.Glosa,  Comprobante.TotalDebe, Comprobante.TotalHaber, Usuarios.Usuario, Comprobante.FechaImport, Comprobante.TipoAjuste "
        Q1 = Q1 & " FROM ((((  Comprobante LEFT JOIN Usuarios ON Comprobante.IdUsuario = Usuarios.IdUsuario ) INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp ) LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc) )"
        Q1 = Q1 & " WHERE  (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (1,3))"
        Q1 = Q1 & " AND Documento.NumDoc = '" & Grid.TextMatrix(Row, C_NUMDOC) & "'"
        Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & ""
        Q1 = Q1 & " AND Comprobante.Ano = " & gEmpresa.Ano
        Q1 = Q1 & " AND Documento.TipoDoc =  " & Grid.TextMatrix(Row, C_IDTIPODOC)
        Q1 = Q1 & " ORDER BY Comprobante.Fecha, Comprobante.Correlativo"
        
        Set Rs = OpenRs(DbMain, Q1)
        If Rs.EOF = True Then
          Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE, False)
          Grid.RowHeight(Row) = 0
        End If
        Call CloseRs(Rs)
        
      End If
      
   Next Row
      
   Call FGrVRows(Grid, 2)
      
   Call CalcTot
   
   Me.MousePointer = vbDefault
   
   MsgBox1 "Si presiona el botón Cancelar, se anulará esta operación.", vbInformation

End Sub

Private Sub Bt_DetDoc_Click()
   Dim Frm As FrmDocLib
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim Row As Integer
   Dim Msg As String
   
   Row = Grid.FlxGrid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then    'registro en blanco
      Exit Sub
   End If
   
   lInBt_DetDoc = True
   
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   If IdDoc = 0 Then
      If Not IsValidLine(Row, Msg) Then
         MsgBox1 "Línea " & Row - Grid.FixedRows + 1 & " inválida. " & Msg, vbExclamation + vbOKOnly
         Exit Sub
      End If
   End If
   
   If (lOper <> O_EDIT And lOper <> O_NEW) Or lEditEnabled = False Then
      Set Frm = New FrmDocLib
      Call Frm.FView(IdDoc)
      Set Frm = Nothing
      Exit Sub
   End If
      
   If Grid.TextMatrix(Row, C_UPDATE) <> "" Then
   
      If MsgBox1("Antes de ingresar al detalle de documento se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      Me.MousePointer = vbHourglass
      
      If Not valida() Then
         Exit Sub
      End If
      
      Call SaveGrid
      
      IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
      
      Me.MousePointer = vbDefault
   
   End If
   
   Set Frm = New FrmDocLib
   Call Frm.FEdit(IdDoc)
   Set Frm = Nothing
   
   lInBt_DetDoc = False
   
   Call LoadGrid(Row)
      
End Sub

Private Sub Bt_Duplicate_Click()
   Dim i As Integer
   Dim Row As Integer
   
   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Row = 0
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
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

   For i = C_FECHA To C_LNGFECHAVENC
      Grid.TextMatrix(Row, i) = Grid.TextMatrix(Grid.Row, i)
   Next i

   If Row > Grid.FixedRows Then
      Grid.TextMatrix(Row, C_NUMLIN) = Grid.TextMatrix(Row - 1, C_NUMLIN) + 1
   Else
      Grid.TextMatrix(Row, C_NUMLIN) = 1
   End If
   
   Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(ED_PENDIENTE)
   Grid.TextMatrix(Row, C_IDESTADO) = ED_PENDIENTE
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
   
   Call CalcTot
      
   Grid.Row = Row
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = C_NUMLIN
   Grid.ColSel = Grid.Cols - 1
End Sub

Private Sub Bt_ExitNewDoc_Click()
   lIdDoc = 0
   lRc = vbOK
   Unload Me
End Sub

Private Sub Bt_HlpImport_Click()
   Dim Frm As FrmFmtImpEnt
    'se comenta segun lo solicitado en req. 2764744
    'pipe 2738156 tema 3
'        If gDbType = SQL_SERVER And lTipoLib = LIB_COMPRAS Or gDbType = SQL_SERVER And lTipoLib = LIB_VENTAS Then
'         MsgBox1 "En version SQL para Compras y Ventas se debe capturar la información mediante registro CSV (Menu Procesos - Importar Registros SII)", vbExclamation + vbOKOnly
'        Bt_Importar.Enabled = False
'         Bt_HlpImport.Enabled = False
'         Exit Sub
'        End If
    'fin
     'fin se comenta segun lo solicitado en req. 2764744
   
   Set Frm = New FrmFmtImpEnt
   If lTipoLib = LIB_COMPRAS Then
      Call Frm.FViewLibCompras
   Else
      Call Frm.FViewLibVentas
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_Importar_Click()
   Call ImportFromFile
End Sub
Private Sub Bt_OK_Click()

   If SaveAll() Then
      lRc = vbOK
   
      Unload Me
   End If

End Sub
Private Function SaveAll() As Boolean
   
   SaveAll = False
      
   Me.MousePointer = vbHourglass
   If valida() = False Then
      Me.MousePointer = vbDefault
      Exit Function
   End If
   
   Call SaveGrid
   Me.MousePointer = vbDefault
   
   If gFunciones.ProporcionalidadIVA Then
      Call CalcPropIVALibro
   End If

   SaveAll = True

End Function
Private Function SaveBeforePrint() As Boolean
   Dim i As Integer
   Dim Modif As Boolean

   SaveBeforePrint = False
   Modif = False
   
   If Bt_OK.visible Then
   
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_FECHA) = "" Then
            Exit For
         End If
         If Grid.TextMatrix(i, C_UPDATE) <> "" Then
            Modif = True
            Exit For
         End If
      
      Next i
      
      If Modif Then
         If MsgBox1("Antes de continuar se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
         
         If valida Then
            Me.MousePointer = vbHourglass
            Call SaveAll
            Me.MousePointer = vbDefault

         Else
            Exit Function
         End If
   
      Else
      
         If gFunciones.ProporcionalidadIVA Then
            Call CalcPropIVALibro
         End If
   
      End If
   End If

   SaveBeforePrint = True

End Function


Private Sub Bt_Opciones_Click()

   If lOper <> O_EDIT Then
      Fr_Opciones.Caption = "Opciones de Vista"
   End If
   
   Fr_Opciones.visible = Not Fr_Opciones.visible
   
End Sub

Private Sub Bt_Paste_Click()
   Dim Fmt As Integer
   Dim DVal As Double
   Dim ValidCol As Boolean
   Dim Value As String
   Dim Action As Integer
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.FlxGrid.Row
   Col = Grid.FlxGrid.Col
   
   ValidCol = (Col = C_NUMDOC Or Col = C_NUMDOCHASTA Or Col = C_RUT Or Col = C_EXENTO Or Col = C_AFECTO Or Col = C_IVA Or Col = C_OTROIMP Or Col = C_FECHAEMIORI Or Col = C_FECHAVENC Or Col = C_DESCRIP)

   If Not ValidCol Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Not ValidaEstadoEdit(Row) Then
      MsgBox1 "Este documento no puede ser modificado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Call ValidaFExported(Row)
      
   If Clipboard.GetFormat(vbCFText) = False Then
      Exit Sub
   End If
   
   If Col = C_RUT Then
      Value = Clipboard.GetText
      Call Grid_AcceptValue(Row, Col, Value, Action)
      If Action = vbOK Then
         Grid.TextMatrix(Row, Col) = Value
      End If
      Exit Sub
   End If
   
   
   DVal = Val(vFmt(Clipboard.GetText))
   
   If (Col >= C_AFECTO And Col <= C_OTROIMP) Then
      
      If DVal <> 0 Then
         DVal = Abs(DVal)
         
         If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).EsRebaja Then
            DVal = DVal * -1
         End If
            
         Grid.TextMatrix(Row, Col) = Format(DVal, NEGNUMFMT)
         
         If Col = C_EXENTO Then
         
            If DVal <> 0 Then
               If Val(Grid.TextMatrix(Row, C_EX_IDCUENTA)) = 0 Then
                  Grid.TextMatrix(Row, C_EX_IDCUENTA) = lCtaExento.id
                  Grid.TextMatrix(Row, C_EX_CODCUENTA) = FmtCodCuenta(lCtaExento.Codigo)
                  Grid.TextMatrix(Row, C_EX_CUENTA) = lCtaExento.Descripcion
               End If
            Else
               Grid.TextMatrix(Row, C_EX_IDCUENTA) = 0
               Grid.TextMatrix(Row, C_EX_CODCUENTA) = ""
               Grid.TextMatrix(Row, C_EX_CUENTA) = ""
            End If
            
         ElseIf Col = C_AFECTO Then
         
            If vFmt(DVal) <> 0 Then
               If Val(Grid.TextMatrix(Row, C_AF_IDCUENTA)) = 0 Then
                  Grid.TextMatrix(Row, C_AF_IDCUENTA) = lCtaAfecto.id
                  Grid.TextMatrix(Row, C_AF_CODCUENTA) = FmtCodCuenta(lCtaAfecto.Codigo)
                  Grid.TextMatrix(Row, C_AF_CUENTA) = lCtaAfecto.Descripcion
               End If
            Else
               Grid.TextMatrix(Row, C_AF_IDCUENTA) = 0
               Grid.TextMatrix(Row, C_AF_CODCUENTA) = 0
               Grid.TextMatrix(Row, C_AF_CUENTA) = 0
            End If
            
         End If
         
         Call CalcTotRow(Row, True)
         
         If vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Then
            Grid.TextMatrix(Row, C_IVA_IDCUENTA) = lIdCuentaIVA
         Else
            Grid.TextMatrix(Row, C_IVA_IDCUENTA) = 0
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Then
            'si es factura de compras, nota de crédito de fac. compras o nota de débito de fac. compras, se pone la cuenta al revés
            If Grid.TextMatrix(Row, C_TIPODOC) = "FCC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCF" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDF" Or Grid.TextMatrix(Row, C_TIPODOC) = "FCV" Then
               Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = lIdCuentaOtrosImpFacCompra
            Else
               Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = lIdCuentaOtrosImp
            End If
         Else
            Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = 0
         End If
         
         If Grid.TextMatrix(Row, C_TOT_IDCUENTA) = "" Then
            Grid.TextMatrix(Row, C_TOT_IDCUENTA) = lCtaTotal.id
            Grid.TextMatrix(Row, C_TOT_CODCUENTA) = FmtCodCuenta(lCtaTotal.Codigo)
            Grid.TextMatrix(Row, C_TOT_CUENTA) = lCtaTotal.Descripcion
         End If

         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         Call CalcTot
         
      End If
   
   Else
      
      If Col = C_FECHAEMIORI Then
         Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = GetDate(Clipboard.GetText, "dmy")
         Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)), SDATEFMT)
      ElseIf Col = C_FECHAVENC Then
         Grid.TextMatrix(Row, C_LNGFECHAVENC) = GetDate(Clipboard.GetText, "dmy")
         Grid.TextMatrix(Row, C_FECHAVENC) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAVENC)), SDATEFMT)
      Else
         Grid.TextMatrix(Row, Col) = Clipboard.GetText
      End If
      
      Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
   
   End If
      
End Sub

Private Sub Bt_List_Click()
   
   If Trim(Tx_Rut) <> "" And Val(lcbNombre.Matrix(M_IDENTIDAD)) = 0 Then
      MsgBox1 "El RUT ingresado no es válido o no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Tx_Rut.SetFocus
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   If ExitDemo() Then
      Unload Me
   End If
   
'   lCurReg = 1
'   lNumReg = 0
   
   lClsPaging.Clear
   
   Call LoadGrid
   Me.MousePointer = vbDefault
   
End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   Dim FrmPrt As FrmPrtSetup
   Dim OldOrientacion As Integer
   
   If Bt_List.Enabled Then
      MsgBox1 "Debe presionar el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If SaveBeforePrint() = False Then
      Exit Sub
   End If
   
   If Grid.ColWidth(C_NUMDOCHASTA) > 0 Then
      If MsgBox1("En la impresión se utilizará letra pequeña para visualizar la mayor cantidad de datos posible." & vbNewLine & vbNewLine & "Si no utiliza la columna ""N° Doc. Hasta"", ocúltela seleccionando las ""Opciones de Vista/Edición""," & vbCrLf & vbCrLf & "en la parte inferior derecha de la ventana. De esta manera, el sistema realizará la impresión" & vbCrLf & vbCrLf & "con tamaño de letra normal." & vbNewLine & vbNewLine & "¿Desea continuar con la impresión?", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If
   
   lPapelFoliado = False
      
   If Ch_ViewMaqReg Then
      lOrientacion = ORIENT_HOR
   End If
   
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar, False) = vbOK Then
      
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      Set FrmPrt = Nothing
      
      Set Frm = New FrmPrintPreview
                  
      gPrtLibros.CallEndDoc = False
      'gPrtLibros.CallEndDoc = -3     'para que imprima el resumen en otra página, si se imprime en más de una franja
      gPrtLibros.PermitirMasDe1Franja = True
      gPrtLibros.FixedCols = C_NUMDOCHASTA + 1
      
      Pag = gPrtLibros.PrtFlexGrid(Frm)
      
      Call PrtResumenIVA(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, gPrtLibros.TieneMasDe1Franja)
            
      Me.MousePointer = vbDefault
      
      Set Frm.PrtControl = Bt_Print
      
      Call Frm.FView(Caption)
      
'      Call PrtResumenIVA(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, gPrtLibros.TieneMasDe1Franja)
'
'      Me.MousePointer = vbDefault
'
'      Set Frm.PrtControl = Bt_Print
'
'      Call Frm.FView(Caption)
      
      Set Frm = Nothing
            
      Call SetPrtNotas 'para reponer las notas que se sacaron en SetupPrtGrid
      
      gPrtLibros.GrFontName = Grid.Font.Name
      gPrtLibros.GrFontSize = Grid.Font.Size
      gPrtLibros.TotFntBold = True
      gPrtLibros.CallEndDoc = True
      gPrtLibros.PermitirMasDe1Franja = False

      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
   End If
      
    Call ResetPrtBas(gPrtLibros)
  
   
   Call SetUpPrtGrid
   
'   Set Frm = Nothing
'   Set Frm = New FrmPrintPreview
'
'   Me.MousePointer = vbHourglass
'
'   gPrtLibros.CallEndDoc = False
'
'   Pag = gPrtLibros.PrtFlexGrid(Frm)
'
'   Call PrtResumenIVA(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
'
'   Call SetPrtNotas 'para reponer las notas que se sacaron en SetupPrtGrid
'   gPrtLibros.GrFontName = Grid.Font.Name
'   gPrtLibros.GrFontSize = Grid.Font.Size
'   gPrtLibros.TotFntBold = True
'
'   Set Frm.PrtControl = Bt_Print
'   Me.MousePointer = vbDefault
'
'   Call Frm.FView(Caption)
'
'   Set Frm = Nothing
'
'   Call SetPrtNotas   'para repone las notas que sacamos en SetupPrtGrid
'
'    Call ResetPrtBas(gPrtLibros)
'
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   Dim Pag As Integer
   
   If Bt_List.Enabled Then
      MsgBox1 "Debe presionar el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If SaveBeforePrint() = False Then
      Exit Sub
   End If
   
   If Grid.ColWidth(C_NUMDOCHASTA) > 0 Then
      If MsgBox1("En la impresión se utilizará letra pequeña para visualizar la mayor cantidad de datos posible." & vbNewLine & vbNewLine & "Si no utiliza la columna ""N° Doc. Hasta"", ocúltela utilizando las ""Opciones de Vista / Edición"" que se preoveen en la parte inferior derecha de la ventana." & vbNewLine & vbNewLine & "Alternativamente, seleccione orientación del papel ""Horizontal"" cuando utilice el botón ""Imprimir""." & vbNewLine & vbNewLine & "De esta manera, el sistema realizará la impresión con de tamaño letra normal." & vbNewLine & vbNewLine & "¿Desea continuar con la impresión?", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If
   
   lPapelFoliado = False
   
   If Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(lLibOf, CbItemData(Cb_Mes), FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(lLibOf) & " Oficial del mes de " & gNomMes(CbItemData(Cb_Mes)) & " ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
      
      lPapelFoliado = True
   End If
   
   If Ch_ViewMaqReg Then
      lOrientacion = ORIENT_HOR
   End If
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Call SetUpPrtGrid
      
      gPrtLibros.CallEndDoc = False
      'gPrtLibros.CallEndDoc = -3       'para que imprima el resumen en otra página, si se imprime en más de una franja
      gPrtLibros.PermitirMasDe1Franja = True
      gPrtLibros.FixedCols = C_NUMDOCHASTA + 1
      
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
      If Pag > 0 Then
         Call PrtResumenIVA(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, gPrtLibros.TieneMasDe1Franja)
      End If
      Printer.EndDoc
      
      If lPapelFoliado And Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(lLibOf, CbItemData(Cb_Mes))
      End If
      
      Call SetPrtNotas 'para reponer las notas que se sacaron en SetupPrtGrid
      
      gPrtLibros.CallEndDoc = True
      gPrtLibros.PermitirMasDe1Franja = False
      
      gPrtLibros.GrFontName = Grid.Font.Name
      gPrtLibros.GrFontSize = Grid.Font.Size
      gPrtLibros.TotFntBold = True

      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
   End If
   
   Call ResetPrtBas(gPrtLibros)
   Me.MousePointer = vbDefault

End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(3) As String
   Dim FontTit(2) As FontDef_t
   Dim OldOrient As Integer
   Dim Mes As String
   Dim Idx As Integer
   
   Set gPrtLibros.Grid = Grid.FlxGrid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = UCase(gTipoLib(lTipoLib))
   If lOper = O_VIEWLIBLEGAL Then
      Idx = InStr(Me.Caption, " - ")
      If Idx > 0 Then
         Titulos(0) = Left(Me.Caption, Idx - 1)
      End If
   End If
   
   FontTit(0).FontBold = True
   
   If CbItemData(Cb_Sucursal) <> 0 Then
      Titulos(1) = "Sucursal: " & Cb_Sucursal & " - "
   End If
   
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Titulos(1) = Titulos(1) & gNomMes(CbItemData(Cb_Mes)) & " " & lAno
   Else
      Titulos(1) = Titulos(1) & Cb_Mes & " " & Val(Cb_Ano)
   End If
   
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
      
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
   gPrtLibros.Encabezados = Encabezados
   
   gPrtLibros.GrFontName = Grid.Font.Name
   gPrtLibros.GrFontSize = Grid.Font.Size
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   'se imprime como en lOper = O_VIEWLIBLEGAL
   
   ColWi(C_NUMLIN) = Grid.ColWidth(C_NUMLIN) - 120
   'ColWi(C_CORRINTERNO) = 0
'   ColWi(C_NOMBRE) = 1900
   'ColWi(C_FECHAEMIORI) = 0
   ColWi(C_FECHA) = 0
   ColWi(C_FECHAVENC) = 0
   ColWi(C_GIRO) = 0
   ColWi(C_CHECK) = 0
   ColWi(C_DESCRIP) = 0
   ColWi(C_SUCURSAL) = 0
   ColWi(C_ESTADO) = 0
   ColWi(C_EX_CODCUENTA) = 0
   ColWi(C_EX_CUENTA) = 0
   ColWi(C_AF_CODCUENTA) = 0
   ColWi(C_AF_CUENTA) = 0
   ColWi(C_TOT_CODCUENTA) = 0
   ColWi(C_TOT_CUENTA) = 0
   ColWi(C_DETALLE) = 0
   ColWi(C_DETACTFIJO) = 0
   ColWi(C_USUARIO) = 0
   ColWi(C_DOCASOC) = 0
   ColWi(C_PROPIVA) = 0
   
   If (Grid.ColWidth(C_NUMDOCHASTA) > 0 Or Ch_ViewMaqReg <> 0) And lOrientacion = ORIENT_VER Then
      ColWi(C_AFECTO) = ColWi(C_AFECTO) - 180
      ColWi(C_EXENTO) = ColWi(C_EXENTO) - 180
      ColWi(C_IVA) = ColWi(C_IVA) - 190
      ColWi(C_OTROIMP) = ColWi(C_OTROIMP) - 190
      ColWi(C_TOTAL) = ColWi(C_TOTAL) - 160
      ColWi(C_RUT) = ColWi(C_RUT) - 220
'      ColWi(C_NOMBRE) = ColWi(C_NOMBRE) - 120
      ColWi(C_NOMBRE) = 1900 - 120
      ColWi(C_NUMDOC) = ColWi(C_NUMDOC) - 160
      ColWi(C_NUMDOCHASTA) = ColWi(C_NUMDOCHASTA) - 160
      ColWi(C_FECHAEMIORI) = ColWi(C_FECHAEMIORI) - 100
      '2814014
      'ColWi(C_TIPODOC) = ColWi(C_TIPODOC) - 30
      ColWi(C_TIPODOC) = ColWi(C_TIPODOC) * 1.2
      'fin 2814014
      If ColWi(C_DTE) > 0 Then
         ColWi(C_DTE) = Grid.ColWidth(C_DTE) - 90
      End If
      If Ch_ViewMaqReg <> 0 Then
         ColWi(C_CORRINTERNO) = 0
      End If
      
      ColWi(C_NUMFISCIMPR) = ColWi(C_NUMFISCIMPR) - 180
      ColWi(C_NUMINFORMEZ) = ColWi(C_NUMINFORMEZ) - 180
      ColWi(C_CANTBOLETAS) = ColWi(C_CANTBOLETAS) - 180
      ColWi(C_VENTASACUM) = ColWi(C_VENTASACUM) - 180
      
      gPrtLibros.GrFontSize = 7
      gPrtLibros.GrFontName = "Arial"
      gPrtLibros.TotFntBold = False
      
   Else
      ColWi(C_AFECTO) = ColWi(C_AFECTO) - 90
      ColWi(C_EXENTO) = ColWi(C_EXENTO) - 90
      ColWi(C_IVA) = ColWi(C_IVA) - 100
      ColWi(C_OTROIMP) = ColWi(C_OTROIMP) - 100
      ColWi(C_TOTAL) = ColWi(C_TOTAL) - 80
      ColWi(C_RUT) = ColWi(C_RUT) - 100
      ColWi(C_NOMBRE) = ColWi(C_NOMBRE) - 120
      ColWi(C_NUMDOC) = ColWi(C_NUMDOC) - 90
      ColWi(C_FECHAEMIORI) = ColWi(C_FECHAEMIORI) - 60
      
       '2814014
      'ColWi(C_TIPODOC) = ColWi(C_TIPODOC) - 30
      ColWi(C_TIPODOC) = ColWi(C_TIPODOC) * 1.2
      'fin 2814014
      
      
      If ColWi(C_DTE) > 0 Then
         ColWi(C_DTE) = Grid.ColWidth(C_DTE) - 60
      End If
      
   End If
      
   If lOrientacion = ORIENT_VER Then
      gPrtLibros.GrFontSize = 8
   Else
      gPrtLibros.GrFontSize = 9
      For i = 0 To Grid.Cols - 1
         ColWi(i) = ColWi(i) * 1.2
      Next i
      
   End If
   
   gPrtLibros.GrFontName = "Arial"
   gPrtLibros.TotFntBold = False
   
   For i = 0 To UBound(ColWi)
      If ColWi(i) < 30 Then
         ColWi(i) = 0
      End If
   Next i

   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.NTotLines = 1
   gPrtLibros.ColObligatoria = C_IDDOC
   
   gPrtLibros.Obs = ""   'para que no ponga las notas
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
End Sub


Private Sub Bt_Resumen_Click()
   Dim Frm As FrmResIVA
   
   If Not SaveBeforePrint Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResIVA
   If CbItemData(Cb_Mes) > 0 Then
      Call Frm.FView(CbItemData(Cb_Mes), Val(Cb_Ano), lTipoLib)
   Else
      Call Frm.FView(0, Val(Cb_Ano), lTipoLib)
   End If
   
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Sel_Click()
   Dim IdDoc As Long
   
   lIdDoc = 0
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   
   If IdDoc > 0 And Val(Grid.TextMatrix(Grid.Row, C_IDESTADO)) <> ED_ANULADO Then
      lIdDoc = IdDoc
      lRc = vbOK
      Unload Me
   Else
      MsgBeep vbExclamation
   End If

End Sub

Private Sub Bt_SelEnt_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   Dim Row As Integer
   Dim TipoEnt As Integer
   Dim Col As Integer
   Dim Rc As Integer
      
   Col = Grid.Col
   Row = Grid.Row
   
   If lTipoLib = LIB_COMPRAS Then
      TipoEnt = ENT_PROVEEDOR
   Else
      TipoEnt = ENT_CLIENTE
   End If
   
   Set Frm = New FrmEntidades
   Rc = Frm.FSelEdit(Entidad, TipoEnt)
   Set Frm = Nothing
   
   If Rc <> vbOK Then
      Exit Sub
   End If
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Or (Grid.Col <> C_RUT And Grid.Col <> C_NOMBRE) Or Not ValidaEstadoEdit(Row) Then
      Exit Sub
   End If
   
   Call ValidaFExported(Row)

         
   If Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And Entidad.NotValidRut <> 0 Then
      MsgBox1 "Rut inválido para este tipo de documento.", vbExclamation
      Exit Sub
   End If
   
   Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
   Grid.TextMatrix(Row, C_RUT) = FmtCID(Entidad.Rut, Entidad.NotValidRut = False)
   Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
   
'   If GetEntRelacionada(Entidad.id) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then  'si es Ent Relacionada y 14TER => pago contado
   If gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then                                      'si es 14TER => pago contado (11 jul 2017 - Claudio Villegas)
      Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))))
      Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
   End If
   
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
      
End Sub

Private Sub Bt_TipoDoc_Click()

   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Or Not ValidaEstadoEdit(Grid.Row) Then
      Exit Sub
   End If
   
   Call ValidaFExported(Grid.Row)
   
'   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      If M_ItTipoDoc.Count > 1 Then
         Call PopupMenu(M_TipoDoc, , Grid.FlxGrid.ColPos(Grid.Col) + Grid.Left + 200, Grid.FlxGrid.RowPos(Grid.Row) + Grid.Top + 100)
      End If
'   End If
   
End Sub

Private Sub Bt_ToLeft_Click()
   Dim Modif As Boolean
   Dim i As Integer

   If lClsPaging.CurReg <= 1 Then
      Exit Sub
   End If
   
   If Not ((lOper <> O_EDIT And lOper <> O_NEW) Or lEditEnabled = False) Then
   
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_FECHA) = "" Then
            Exit For
         End If
         If Grid.TextMatrix(i, C_UPDATE) <> "" Then
            Modif = True
            Exit For
         End If
      
      Next i
      
      If Modif Then

         If MsgBox1("Antes de pasar al siguiente conjunto de registros se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
         
      End If
      
      If Not valida() Then
         Exit Sub
      End If
      
      Me.MousePointer = vbHourglass
      
      Call SaveGrid
                        
   End If
   
   Me.MousePointer = vbHourglass
   
   lClsPaging.ToLeft
   
   Call LoadGrid
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_ToRight_Click()
   Dim Modif As Boolean
   Dim i As Integer

   If Grid.TextMatrix(Grid.FixedRows, C_NUMLIN) = "" Then
      Exit Sub
   End If
   
   If Not ((lOper <> O_EDIT And lOper <> O_NEW) Or lEditEnabled = False) Then
            
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_FECHA) = "" Then
            Exit For
         End If
         If Grid.TextMatrix(i, C_UPDATE) <> "" Then
            Modif = True
            Exit For
         End If
      
      Next i
      
      If Modif Then

         If MsgBox1("Antes de pasar al siguiente conjunto de registros se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
         
      End If
      
      If Not valida() Then
         Exit Sub
      End If
      
      Me.MousePointer = vbHourglass
      
      Call SaveGrid
                              
   End If
   
   
'   lCurReg = lCurReg + lNumReg
'   lToRightPressed = True
   
   Me.MousePointer = vbHourglass
   
   Call lClsPaging.ToRight
   
   Call LoadGrid
   
   Me.MousePointer = vbDefault
   
   lClsPaging.ToRightPressed = False

End Sub

Private Sub Cb_Estado_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub

Private Sub Cb_Mes_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub
Private Sub Cb_Ano_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
   
   If Val(Cb_Ano) <> gEmpresa.Ano Then
      MsgBox1 "Recuerde que para años anteriores al año actual, sólo debieran aparecer los documentos pendientes de pago.", vbInformation
   End If
End Sub


Private Sub Cb_Sucursal_Click()
   Static InClick As Boolean

   If InClick = True Then
      Exit Sub
   End If
   
   InClick = True
   
   If Me.visible Then
   
      If GrabarParaContinuar("cambiar la sucursal seleccionada") = False Then
         Call SelItem(Cb_Sucursal, lSucursal)
         InClick = False
         Exit Sub
      End If
      
      Me.MousePointer = vbHourglass
      
      Call LoadGrid
      
      Me.MousePointer = vbDefault
      
      lSucursal = CbItemData(Cb_Sucursal)
      
   End If
   
   InClick = False
   
End Sub

Private Sub Cb_TipoDoc_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub

Private Sub Cb_DTE_Click()
   Bt_List.Enabled = True
   Bt_Sel.Enabled = False
   Bt_Centralizar.Enabled = False

End Sub

Private Sub Ch_CentralizacionFull_Click()
   If Ch_CentralizacionFull Then
     Bt_Centralizar_Full.visible = True
     Bt_Centralizar.visible = False
   Else
     Bt_Centralizar.visible = True
     Bt_Centralizar_Full.visible = False
   End If
End Sub

Private Sub Ch_EsSupermercado_Click()
   Bt_List.Enabled = True
   Bt_Sel.Enabled = False
   Bt_Centralizar.Enabled = False

End Sub

Private Sub Ch_RepetirGlosa_Click()
   
   Call SetIniString(gIniFile, "Opciones", "RepetirGlosa", Abs(Ch_RepetirGlosa.Value))
   gVarIniFile.RepetirGlosa = Abs(Ch_RepetirGlosa.Value)
   
End Sub

Private Sub Ch_Rut_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

Private Sub Ch_ViewCantBoletas_Click()

   If Ch_ViewCantBoletas = 0 Then
      Grid.ColWidth(C_CANTBOLETAS) = 0
      Grid.TextMatrix(0, C_CANTBOLETAS) = ""
      Grid.TextMatrix(1, C_CANTBOLETAS) = ""
      
   Else
      Grid.ColWidth(C_CANTBOLETAS) = 900
      Grid.TextMatrix(0, C_CANTBOLETAS) = "Cantidad"
      Grid.TextMatrix(1, C_CANTBOLETAS) = "de Boletas"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerCantBoletas", Abs(Ch_ViewCantBoletas.Value))
   gVarIniFile.VerCantBoletas = Abs(Ch_ViewCantBoletas.Value)

End Sub

Private Sub Ch_ViewDetOtrosImp_Click()

   If Not Me.visible Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass

   If Ch_ViewDetOtrosImp = 0 Then
      Call DetOtrosImp(-1)
            
   Else
      Call DetOtrosImp(lTipoLib, lCurWhere)
      
   End If
   
   Call LoadGrid
   
   Me.MousePointer = vbDefault

   Call SetIniString(gIniFile, "Opciones", "VerDetOtrosImp", Abs(Ch_ViewDetOtrosImp.Value))
   gVarIniFile.VerDetOtrosImp = Abs(Ch_ViewDetOtrosImp.Value)

End Sub

Private Sub Ch_ViewDocHasta_Click()
   
   If Ch_ViewDocHasta = 0 Then
      Grid.ColWidth(C_NUMDOCHASTA) = 0
      Grid.TextMatrix(0, C_NUMDOCHASTA) = ""
      Grid.TextMatrix(1, C_NUMDOCHASTA) = ""
      
   Else
      Grid.ColWidth(C_NUMDOCHASTA) = 900
      Grid.TextMatrix(0, C_NUMDOCHASTA) = "N° Doc"
      Grid.TextMatrix(1, C_NUMDOCHASTA) = "Hasta"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerNumDocHasta", Abs(Ch_ViewDocHasta.Value))
   gVarIniFile.VerNumDocHasta = Abs(Ch_ViewDocHasta.Value)
   
End Sub

Private Sub Ch_ViewNumInterno_Click()

   If Ch_ViewNumInterno = 0 Then
      Grid.ColWidth(C_CORRINTERNO) = 0
      Grid.TextMatrix(0, C_CORRINTERNO) = ""
      Grid.TextMatrix(1, C_CORRINTERNO) = ""
      
   Else
      Grid.ColWidth(C_CORRINTERNO) = 600
      Grid.TextMatrix(0, C_CORRINTERNO) = "Nro."
      Grid.TextMatrix(1, C_CORRINTERNO) = "Interno"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerNumInterno", Abs(Ch_ViewNumInterno.Value))
   gVarIniFile.VerNumInterno = Abs(Ch_ViewNumInterno.Value)

End Sub

Private Sub Ch_ViewOtrosImp_Click()

   If Ch_ViewOtrosImp = 0 Then
      Grid.ColWidth(C_OTROIMP) = 0
      Grid.TextMatrix(0, C_OTROIMP) = ""
      Grid.TextMatrix(1, C_OTROIMP) = ""
      
   Else
      Grid.ColWidth(C_OTROIMP) = 1200
      Grid.TextMatrix(0, C_OTROIMP) = "Otros"
      Grid.TextMatrix(1, C_OTROIMP) = "Impuestos"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerNumDocHasta", Abs(Ch_ViewOtrosImp.Value))
   gVarIniFile.VerOtrosImp = Abs(Ch_ViewOtrosImp.Value)

End Sub

Private Sub Ch_ViewPropIVA_Click()
   
   If Ch_ViewPropIVA = 0 Then
      Grid.ColWidth(C_PROPIVA) = 0
      Grid.TextMatrix(0, C_PROPIVA) = ""
      Grid.TextMatrix(1, C_PROPIVA) = ""
      Grid.ColAlignment(C_PROPIVA) = flexAlignLeftCenter
      
   Else
      Grid.ColWidth(C_PROPIVA) = 430
      Grid.TextMatrix(0, C_PROPIVA) = "Prop."
      Grid.TextMatrix(1, C_PROPIVA) = "IVA"
      Grid.ColAlignment(C_PROPIVA) = flexAlignCenterCenter
         
   End If
   
   If lTipoLib = LIB_COMPRAS Then
      Call SetIniString(gIniFile, "Opciones", "VerPropIVA", Abs(Ch_ViewPropIVA.Value))
      gVarIniFile.VerPropIVA = Abs(Ch_ViewPropIVA.Value)
   End If

End Sub

Private Sub Ch_ViewSucursal_Click()
   
   If Ch_ViewSucursal = 0 Then
      Grid.ColWidth(C_SUCURSAL) = 0
      Grid.TextMatrix(1, C_SUCURSAL) = ""
      
   Else
      Grid.ColWidth(C_SUCURSAL) = 1500
      Grid.TextMatrix(1, C_SUCURSAL) = "Sucursal"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerSucursal", Abs(Ch_ViewSucursal.Value))
   gVarIniFile.VerSucursal = Abs(Ch_ViewSucursal.Value)

End Sub

Private Sub Ch_ViewDTE_Click()
   
   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
      Grid.TextMatrix(1, C_DTE) = ""
      Grid.ColAlignment(C_DTE) = flexAlignLeftCenter
      
   Else
      Grid.ColWidth(C_DTE) = 400
      Grid.TextMatrix(1, C_DTE) = "DTE"
      Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerDTE", Abs(Ch_ViewDTE.Value))
   gVarIniFile.VerDTE = Abs(Ch_ViewDTE.Value)

End Sub

Private Sub Ch_ViewMaqReg_Click()
   
   If Ch_ViewMaqReg.Value = 0 Then
      Grid.ColWidth(C_NUMFISCIMPR) = 0
      Grid.ColWidth(C_NUMINFORMEZ) = 0
      Grid.ColWidth(C_VENTASACUM) = 0
      Grid.TextMatrix(0, C_NUMFISCIMPR) = ""
      Grid.TextMatrix(1, C_NUMFISCIMPR) = ""
      Grid.TextMatrix(0, C_NUMINFORMEZ) = ""
      Grid.TextMatrix(1, C_NUMINFORMEZ) = ""
      Grid.TextMatrix(0, C_VENTASACUM) = ""
      Grid.TextMatrix(1, C_VENTASACUM) = ""
      
   Else
      Grid.ColWidth(C_NUMFISCIMPR) = 1200
      Grid.ColWidth(C_NUMINFORMEZ) = 1200
      Grid.ColWidth(C_VENTASACUM) = 1200
      Grid.TextMatrix(0, C_NUMFISCIMPR) = "N° Fiscal"
      Grid.TextMatrix(1, C_NUMFISCIMPR) = "Impresora"
      Grid.TextMatrix(0, C_NUMINFORMEZ) = "N° Informe"
      Grid.TextMatrix(1, C_NUMINFORMEZ) = """Z"""
      Grid.TextMatrix(0, C_VENTASACUM) = "Ventas Acum."
      Grid.TextMatrix(1, C_VENTASACUM) = "Informe ""Z"""
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerMaqReg", Abs(Ch_ViewMaqReg.Value))
   gVarIniFile.VerMaqReg = Abs(Ch_ViewMaqReg.Value)

End Sub

Private Sub Ch_ViewExento_Click()
   
   If Ch_ViewExento = 0 Then
      Grid.ColWidth(C_EXENTO) = 0
      Grid.ColWidth(C_EX_CODCUENTA) = 0
      Grid.ColWidth(C_EX_CUENTA) = 0
      
      Grid.TextMatrix(1, C_EXENTO) = ""
      Grid.TextMatrix(0, C_EX_CODCUENTA) = ""
      Grid.TextMatrix(1, C_EX_CODCUENTA) = ""
      Grid.TextMatrix(0, C_EX_CUENTA) = ""
      Grid.TextMatrix(1, C_EX_CUENTA) = ""
   
   Else

      Grid.ColWidth(C_EXENTO) = 1200
      Grid.TextMatrix(1, C_EXENTO) = "Exento"
      
      If lOper = O_VIEWLIBLEGAL Then
         Grid.ColWidth(C_EX_CODCUENTA) = 0
         Grid.ColWidth(C_EX_CUENTA) = 0
         
         Grid.TextMatrix(0, C_EX_CODCUENTA) = ""
         Grid.TextMatrix(1, C_EX_CODCUENTA) = ""
         Grid.TextMatrix(0, C_EX_CUENTA) = ""
         Grid.TextMatrix(1, C_EX_CUENTA) = ""
         
      Else
         Grid.ColWidth(C_EX_CODCUENTA) = lWCodCuenta
         Grid.ColWidth(C_EX_CUENTA) = lWCuenta
         
         Grid.TextMatrix(0, C_EX_CODCUENTA) = "Cod.Cuenta"
         Grid.TextMatrix(1, C_EX_CODCUENTA) = "Exento"
         Grid.TextMatrix(0, C_EX_CUENTA) = "  Cuenta"
         Grid.TextMatrix(1, C_EX_CUENTA) = "  Exento"
         
      End If
   
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerExento", Abs(Ch_ViewExento.Value))
   gVarIniFile.VerExento = Abs(Ch_ViewExento.Value)

End Sub


Private Sub Form_Load()
   Dim i As Integer
   Dim j As Integer
   Dim Q1 As String

   lMsgFechaErr = False
    
   lOrientacion = ORIENT_VER    'imprimimos como libro legal
   lFNameLogImp = gImportPath & "\Log\ImpLib-" & Format(Now, "yyyymmdd") & ".log"
           
   Set lClsPaging = New ClsPaging

   Call lClsPaging.Init(Bt_ToLeft, Bt_ToRight)

   Call SetUpGrid
   
   '3217885
    If gDbType = SQL_SERVER Then
   Ch_CentralizacionFull.visible = True
   Else
   Ch_CentralizacionFull.visible = False
   End If
   '3217885
     
   If gFunciones.ExpImpLibrosAuxFile Then
      Bt_Importar.visible = True
      Bt_HlpImport.visible = True
   Else
      Bt_Importar.visible = False
      Bt_HlpImport.visible = False
   End If
     
   Fr_Opciones.visible = False
            
   Ch_LibOficial.visible = False
   Ch_ViewExento.visible = True
   Ch_ViewExento = gVarIniFile.VerExento
   Ch_ViewDTE.visible = True
   Ch_ViewDTE = gVarIniFile.VerDTE
   Ch_ViewSucursal.visible = True
   Ch_ViewSucursal = gVarIniFile.VerSucursal
   Ch_ViewNumInterno.visible = True
   Ch_ViewNumInterno = gVarIniFile.VerNumInterno
   Ch_ViewOtrosImp.visible = True
   Ch_ViewOtrosImp = gVarIniFile.VerOtrosImp
   Ch_ViewMaqReg.visible = True
   Ch_ViewMaqReg = gVarIniFile.VerMaqReg
   Ch_ViewCantBoletas.visible = True
   Ch_ViewCantBoletas = gVarIniFile.VerCantBoletas
   Ch_ViewDetOtrosImp.visible = True
   Ch_ViewDetOtrosImp = gVarIniFile.VerDetOtrosImp
   Ch_ViewPropIVA.visible = gFunciones.ProporcionalidadIVA
   Ch_ViewPropIVA = gVarIniFile.VerPropIVA
   
   Ch_RepetirGlosa.Enabled = False
   Ch_RepetirGlosa = 0
   
   Ch_Rut = 1

   If lTipoLib = LIB_COMPRAS Then
      lLibOf = LIBOF_COMPRAS
      Ch_ViewDocHasta = 0
      Ch_ViewDocHasta.Enabled = False
      Ch_ViewMaqReg = 0
      Ch_ViewMaqReg.Enabled = False
      Ch_ViewCantBoletas = 0
      Ch_ViewCantBoletas.Enabled = False
           
      
   ElseIf lTipoLib = LIB_VENTAS Then
      lLibOf = LIBOF_VENTAS
      Ch_ViewDocHasta = gVarIniFile.VerNumDocHasta
      Ch_ViewDocHasta.Enabled = True
      Ch_ViewMaqReg = gVarIniFile.VerMaqReg
      Ch_ViewMaqReg.Enabled = True
      Ch_ViewCantBoletas = gVarIniFile.VerCantBoletas
      Ch_ViewCantBoletas.Enabled = True
      Ch_ViewPropIVA = 0
      Ch_ViewPropIVA.Enabled = False
           
           
   End If
   
   Bt_Opciones.Caption = "Opciones de Vista"

   Select Case lOper
   
      Case O_VIEWLIBLEGAL
         If lTitEspecialVentas <> "" Then
            Me.Caption = lTitEspecialVentas
         Else
            Me.Caption = gTipoLib(lTipoLib)
         End If
         Ch_LibOficial.visible = True
         'Ch_ViewExento = 1
         Ch_ViewExento.Enabled = False
         Ch_ViewSucursal = 0
         Ch_ViewSucursal.Enabled = False
         'Ch_ViewNumInterno = 0
         'Ch_ViewNumInterno.Enabled = False
         Ch_ViewNumInterno = False
         Bt_Importar.visible = False
         Bt_HlpImport.visible = False
         Bt_DelAll.visible = False
         
      Case O_VIEW
         Me.Caption = "Listar " & gTipoLib(lTipoLib)
         
      Case O_EDIT
         Me.Caption = "Editar " & gTipoLib(lTipoLib)
         Ch_ViewDetOtrosImp.Enabled = False
         Ch_ViewDetOtrosImp = 0
         If lTipoLib = LIB_VENTAS Then
            Ch_RepetirGlosa.Enabled = True
            Ch_RepetirGlosa = gVarIniFile.RepetirGlosa
         End If
         Bt_Opciones.Caption = "Opciones de Edición"
         Lb_Sucursal.visible = False

      Case O_SELECT
         Me.Caption = "Seleccionar Documento del " & gTipoLib(lTipoLib)

   End Select
      
    
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Call CbAddItem(Grid.CbList(C_SUCURSAL), " ", 0)
   Call CbAddItem(Cb_Sucursal, " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales WHERE IdEmpresa = " & gEmpresa.id
   If lOper = O_EDIT Then
      Q1 = Q1 & " AND Vigente <> 0 "
   End If
   Q1 = Q1 & " ORDER BY Descripcion"
   Call FillCombo(Grid.CbList(C_SUCURSAL), DbMain, Q1, -1)
   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)
      
   Call CbAddItem(Cb_DTE, " ", 0)
   Call CbAddItem(Cb_DTE, "DTE", 1)
   Call CbAddItem(Cb_DTE, "No DTE", -1)
   Cb_DTE.ListIndex = 0
   
   
   If lOper = O_VIEW Or lOper = O_VIEWLIBLEGAL Or lOper = O_SELECT Then
      Grid.Locked = True
      Bt_Cancel.Caption = "Cerrar"
      Bt_Cancel.Left = Bt_Cancel.Left - 320
      Bt_Cancel.Width = 1095
      Bt_OK.visible = False
      Fr_BtEdit.visible = False
      Fr_BtGen.Left = Fr_BtEdit.Left
      Ch_LibOficial.Left = Fr_BtGen.Left + Fr_BtGen.Width + 300
      
      Bt_Importar.visible = False
      Bt_HlpImport.visible = False
      Bt_DelAll.visible = False

      
      If lOper = O_SELECT Then
         Bt_Centralizar.visible = False
         
         Bt_List.Left = 7980
         Fr_List.Width = 9195
         
      Else
         Bt_Sel.visible = False
         Bt_ExitNewDoc.visible = False
      End If
      
   End If
   
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Me.Caption = Me.Caption & " - " & gNomMes(lMes) & " " & lAno
      Fr_List.visible = False
      Fr_List.Enabled = False
      Bt_Centralizar.visible = False
      Bt_Sel.visible = False
      Bt_ExitNewDoc.visible = False
      Grid.Height = Grid.Height + Grid.Top - Fr_List.Top + 50
      Grid.Top = Fr_List.Top
      
   Else
      Cb_Sucursal.Left = Cb_Nombre.Left + Fr_List.Left
      Cb_Sucursal.Width = Cb_Nombre.Width
      Lb_Sucursal.Left = Cb_Sucursal.Left - Lb_Sucursal.Width - 60
      
   End If
     
   Call SetTxRO(Tx_CurrCell, True)
     
   Call FillCb
   
   Call SetOrderLst
   
   If lTipoLib = LIB_COMPRAS Then
      lIdCuentaIVA = gCtasBas.IdCtaIVACred
      lIdCuentaIVAIrrec = gCtasBas.IdCtaIVAIrrec
      lIdCuentaOtrosImp = gCtasBas.IdCtaOtrosImpCred
      lIdCuentaOtrosImpFacCompra = gCtasBas.IdCtaOtrosImpDeb
   Else
      lIdCuentaIVA = gCtasBas.IdCtaIVADeb               'LIB_VENTAS
      lIdCuentaIVAIrrec = 0
      lIdCuentaOtrosImp = gCtasBas.IdCtaOtrosImpDeb
      lIdCuentaOtrosImpFacCompra = gCtasBas.IdCtaOtrosImpCred
   End If
   
   Call LoadCuentasDef
   
   Call SetupPriv
   
   Call LoadGrid
   
   'limpiamos los títulos de la grilla de las columnas que tienen ancho cero
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) = 0 Then
         For j = 0 To Grid.FixedRows - 1
            Grid.TextMatrix(j, i) = ""
         Next j
      End If
   Next i
   
   Bt_ActivoFijo.visible = gFunciones.ActivoFijo
   Bt_ActivoFijo.Enabled = gFunciones.ActivoFijo
   
   Bt_DocCuotas.visible = gFunciones.DocCuotas
   Bt_DocCuotas.Enabled = gFunciones.DocCuotas
   
'   Ch_ViewNumInterno.Enabled = Not gAppCode.Demo
'   Ch_ViewOtrosImp.Enabled = Not gAppCode.Demo
   
   Tm_ColWi.Enabled = True
   
   
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   
   Grid.Cols = NCOLS + 1
   Grid.ColWidth(C_IDDOC) = 0
   
   If lOper = O_VIEW Then
      Grid.ColWidth(C_CHECK) = 300
   Else
      Grid.ColWidth(C_CHECK) = 0
   End If
   
   If lTipoLib = LIB_VENTAS Then
      Grid.ColWidth(C_GIRO) = 400
   Else
      Grid.ColWidth(C_GIRO) = 0
   End If
   
   
   lWCodCuenta = Me.TextWidth(gFmtCodigoCta) + 300
   lWCuenta = 1450
   
   Grid.ColWidth(C_NUMLIN) = 500
   Grid.ColWidth(C_CORRINTERNO) = 600
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Grid.ColWidth(C_FECHA) = 400
   Else
      Grid.ColWidth(C_FECHA) = 800
   End If
   
   '2814014 pipe
   'Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_TIPODOC) = 510
   'fin 2814014
   Grid.ColWidth(C_IDTIPODOC) = 0
   Grid.ColWidth(C_DOCIMPEXP) = 0
   Grid.ColWidth(C_DTE) = 400
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.ColWidth(C_NUMFISCIMPR) = 0
      Grid.ColWidth(C_NUMINFORMEZ) = 0
   Else
      Grid.ColWidth(C_NUMFISCIMPR) = 1200
      Grid.ColWidth(C_NUMINFORMEZ) = 1200
   End If
   
   Grid.ColWidth(C_NUMDOC) = 900
   
   Grid.ColWidth(C_IDPROPIVA) = 0
   Grid.ColWidth(C_PROPIVA) = 0
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.ColWidth(C_NUMDOCHASTA) = 0
      Grid.ColWidth(C_CANTBOLETAS) = 0
      If gFunciones.ProporcionalidadIVA And gVarIniFile.VerPropIVA Then
         Grid.ColWidth(C_PROPIVA) = 430
      End If
         
   Else
      Grid.ColWidth(C_NUMDOCHASTA) = 900 ' Opcional 900
      Grid.ColWidth(C_CANTBOLETAS) = 900 ' Opcional 900
   End If
      
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_NOMBRE) = 2000
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_DESCRIP) = 2000
   Grid.ColWidth(C_IDSUCURSAL) = 0
   Grid.ColWidth(C_SUCURSAL) = 1500
   Grid.ColWidth(C_EXENTO) = 1200
   Grid.ColWidth(C_EX_IDCUENTA) = 0
   Grid.ColWidth(C_EX_CODCUENTA) = lWCodCuenta
   Grid.ColWidth(C_EX_CUENTA) = lWCuenta
   Grid.ColWidth(C_AFECTO) = 1200
   Grid.ColWidth(C_AF_IDCUENTA) = 0
   Grid.ColWidth(C_AF_CODCUENTA) = lWCodCuenta
   Grid.ColWidth(C_AF_CUENTA) = lWCuenta
   Grid.ColWidth(C_IVA) = 1200
   Grid.ColWidth(C_IVA_IDCUENTA) = 0
   Grid.ColWidth(C_OTROIMP) = 1200
   Grid.ColWidth(C_OIMP_IDCUENTA) = 0
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_TOT_IDCUENTA) = 0
   Grid.ColWidth(C_TOT_CODCUENTA) = lWCodCuenta
   Grid.ColWidth(C_TOT_CUENTA) = lWCuenta
   
   Grid.ColWidth(C_IDANEG_CCOSTO) = 0
   
   For i = C_INIDETOTROIMP To C_ENDDETOTROIMP
      Grid.ColWidth(i) = 0
   Next i
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.ColWidth(C_VENTASACUM) = 0
   Else
      Grid.ColWidth(C_VENTASACUM) = 1200
   End If

   Grid.ColWidth(C_DETALLE) = 300
   
   Grid.ColWidth(C_FECHAEMIORI) = 780
   Grid.ColWidth(C_LNGFECHAEMIORI) = 0
   Grid.ColWidth(C_FECHAVENC) = 780
   Grid.ColWidth(C_LNGFECHAVENC) = 0
   If gFunciones.ActivoFijo Then
      Grid.ColWidth(C_DETACTFIJO) = 700
   Else
      Grid.ColWidth(C_DETACTFIJO) = 0    'se oculta por ahora
   End If
   Grid.ColWidth(C_ESTADO) = 1050
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_USUARIO) = 1000
   Grid.ColWidth(C_DOCASOC) = 1400
   Grid.ColWidth(C_MOVEDITED) = 0
   Grid.ColWidth(C_IDCOMPCENT) = 0
   Grid.ColWidth(C_IDCOMPPAGO) = 0
   Grid.ColWidth(C_MSGACTFIJO) = 0
   Grid.ColWidth(C_EXPORTED) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   If lOper = O_EDIT Then
      'Grid.ColWidth(C_DESCRIP) = 3000
   
   ElseIf lOper = O_VIEWLIBLEGAL Then
   
      lOrientacion = ORIENT_VER
   
      Grid.ColWidth(C_CORRINTERNO) = 0
      Grid.ColWidth(C_NOMBRE) = 2100
      Grid.ColWidth(C_FECHA) = 0
      'Grid.ColWidth(C_FECHAEMIORI) = 0
      Grid.ColWidth(C_FECHAVENC) = 0
      Grid.ColWidth(C_DESCRIP) = 0
      Grid.ColWidth(C_SUCURSAL) = 0
      Grid.ColWidth(C_ESTADO) = 0
      Grid.ColWidth(C_EX_CODCUENTA) = 0
      Grid.ColWidth(C_EX_CUENTA) = 0
      Grid.ColWidth(C_AF_CODCUENTA) = 0
      Grid.ColWidth(C_AF_CUENTA) = 0
      Grid.ColWidth(C_TOT_CODCUENTA) = 0
      Grid.ColWidth(C_TOT_CUENTA) = 0
      Grid.ColWidth(C_DETALLE) = 0
      Grid.ColWidth(C_DETACTFIJO) = 0
      Grid.ColWidth(C_USUARIO) = 0
      Grid.ColWidth(C_DOCASOC) = 0
   
   End If
      
   If Ch_ViewExento = 0 Then
      Grid.ColWidth(C_EXENTO) = 0
      Grid.ColWidth(C_EX_IDCUENTA) = 0
      Grid.ColWidth(C_EX_CODCUENTA) = 0
      Grid.ColWidth(C_EX_CUENTA) = 0
   End If
      
   
   If Ch_ViewDocHasta = 0 Then
      Grid.ColWidth(C_NUMDOCHASTA) = 0
   End If
      
   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
   End If
      
   If Ch_ViewSucursal = 0 Then
      Grid.ColWidth(C_SUCURSAL) = 0
   End If
      
   If Ch_ViewNumInterno = 0 Then
      Grid.ColWidth(C_CORRINTERNO) = 0
   End If
      
   If Ch_ViewOtrosImp = 0 Then
      Grid.ColWidth(C_OTROIMP) = 0
   End If
      
   If Ch_ViewMaqReg = 0 Then
      Grid.ColWidth(C_NUMFISCIMPR) = 0
      Grid.ColWidth(C_NUMINFORMEZ) = 0
      Grid.ColWidth(C_VENTASACUM) = 0
   End If
      
   If Ch_ViewCantBoletas = 0 Then
      Grid.ColWidth(C_CANTBOLETAS) = 0
   End If
      
      
   Grid.ColAlignment(C_NUMLIN) = flexAlignRightCenter
   Grid.ColAlignment(C_GIRO) = flexAlignCenterCenter
   Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_NUMDOCHASTA) = flexAlignRightCenter
   Grid.ColAlignment(C_PROPIVA) = flexAlignCenterCenter
   Grid.ColAlignment(C_CHECK) = flexAlignCenterCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESCRIP) = flexAlignLeftCenter
   Grid.ColAlignment(C_SUCURSAL) = flexAlignLeftCenter
   Grid.ColAlignment(C_EXENTO) = flexAlignRightCenter
   Grid.ColAlignment(C_AFECTO) = flexAlignRightCenter
   Grid.ColAlignment(C_IVA) = flexAlignRightCenter
   Grid.ColAlignment(C_OTROIMP) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_CORRINTERNO) = flexAlignRightCenter
   Grid.ColAlignment(C_DETALLE) = flexAlignCenterCenter
   Grid.ColAlignment(C_DETACTFIJO) = flexAlignCenterCenter
   
   For i = C_INIDETOTROIMP To C_ENDDETOTROIMP
      Grid.ColAlignment(i) = flexAlignRightCenter
   Next i

   
   Grid.TextMatrix(1, C_NUMLIN) = "Línea"
   Grid.TextMatrix(0, C_CORRINTERNO) = "Nro."
   Grid.TextMatrix(1, C_CORRINTERNO) = "Interno"
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Grid.TextMatrix(0, C_FECHA) = "Día"
   Else
      Grid.TextMatrix(0, C_FECHA) = "Fecha"
   End If
   Grid.TextMatrix(1, C_FECHA) = "Rec."
   Grid.TextMatrix(1, C_TIPODOC) = "TD"
   
   If Grid.ColWidth(C_DTE) <> 0 Then
      Grid.TextMatrix(1, C_DTE) = "DTE"
   End If
   
   If Grid.ColWidth(C_GIRO) <> 0 Then
      Grid.TextMatrix(0, C_GIRO) = "Del"
      Grid.TextMatrix(1, C_GIRO) = "Giro"
   End If
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.TextMatrix(0, C_NUMFISCIMPR) = ""
      Grid.TextMatrix(1, C_NUMFISCIMPR) = ""
      Grid.TextMatrix(0, C_NUMINFORMEZ) = ""
      Grid.TextMatrix(1, C_NUMINFORMEZ) = ""
      Grid.TextMatrix(0, C_CANTBOLETAS) = ""
      Grid.TextMatrix(1, C_CANTBOLETAS) = ""
   Else
      If Ch_ViewMaqReg <> 0 Then
         Grid.TextMatrix(0, C_NUMFISCIMPR) = "N° Fiscal"
         Grid.TextMatrix(1, C_NUMFISCIMPR) = "Impresora"
         Grid.TextMatrix(0, C_NUMINFORMEZ) = "N° Informe"
         Grid.TextMatrix(1, C_NUMINFORMEZ) = """Z"""
      End If
      If Ch_ViewCantBoletas <> 0 Then
         Grid.TextMatrix(0, C_CANTBOLETAS) = "Cantidad"
         Grid.TextMatrix(1, C_CANTBOLETAS) = "de Boletas"
      End If
   End If
   
   Grid.ColWidth(C_NUMDOC) = 900
   

   
   Grid.TextMatrix(1, C_NUMDOC) = "N° Doc."
   If lTipoLib = LIB_VENTAS Then
      Grid.TextMatrix(0, C_NUMDOC) = "N° Doc."
      Grid.TextMatrix(1, C_NUMDOC) = "o Máq."
   End If
   
   Grid.TextMatrix(0, C_NUMDOCHASTA) = "N° Doc"
   Grid.TextMatrix(1, C_NUMDOCHASTA) = "Hasta"
   Grid.TextMatrix(1, C_ESTADO) = "Estado"
   Grid.TextMatrix(1, C_RUT) = "RUT"
   Grid.TextMatrix(1, C_NOMBRE) = "Razón Social"
   Grid.TextMatrix(1, C_EXENTO) = "Exento"
   Grid.TextMatrix(0, C_EX_CODCUENTA) = "Cod.Cuenta"
   Grid.TextMatrix(1, C_EX_CODCUENTA) = "Exento"
   Grid.TextMatrix(0, C_EX_CUENTA) = "  Cuenta"
   Grid.TextMatrix(1, C_EX_CUENTA) = "  Exento"
   Grid.TextMatrix(1, C_AFECTO) = "Afecto"
   Grid.TextMatrix(0, C_AF_CODCUENTA) = "Cod.Cuenta"
   Grid.TextMatrix(1, C_AF_CODCUENTA) = "Afecto"
   Grid.TextMatrix(0, C_AF_CUENTA) = "  Cuenta"
   Grid.TextMatrix(1, C_AF_CUENTA) = "  Afecto"
   
   Grid.TextMatrix(0, C_IVA) = "IVA"
   If lTipoLib = LIB_COMPRAS Then
      Grid.TextMatrix(1, C_IVA) = "Crédito Fiscal"
   ElseIf lTipoLib = LIB_VENTAS Then
      Grid.TextMatrix(1, C_IVA) = "Débito Fiscal"
   End If
   
   Grid.TextMatrix(0, C_OTROIMP) = "Otros"
   Grid.TextMatrix(1, C_OTROIMP) = "Impuestos"
   Grid.TextMatrix(1, C_TOTAL) = "Total"
   Grid.TextMatrix(0, C_TOT_CODCUENTA) = "Cod.Cuenta"
   Grid.TextMatrix(1, C_TOT_CODCUENTA) = "Total"
   Grid.TextMatrix(0, C_TOT_CUENTA) = " Cuenta"
   Grid.TextMatrix(1, C_TOT_CUENTA) = " Total"
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.TextMatrix(0, C_VENTASACUM) = ""
      Grid.TextMatrix(1, C_VENTASACUM) = ""
   Else
      Grid.TextMatrix(0, C_VENTASACUM) = "Ventas Acum."
      Grid.TextMatrix(1, C_VENTASACUM) = "Informe ""Z"""
   End If
   
   
   Grid.TextMatrix(0, C_FECHAEMIORI) = "Fecha"
   Grid.TextMatrix(1, C_FECHAEMIORI) = "Emisión"
   Grid.TextMatrix(0, C_FECHAVENC) = "Fecha"
   Grid.TextMatrix(1, C_FECHAVENC) = "Vencim."
   If gFunciones.ActivoFijo Then
      Grid.TextMatrix(1, C_DETACTFIJO) = "Act. Fijo"
   End If
   Grid.TextMatrix(1, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(1, C_SUCURSAL) = "Sucursal"
   Grid.TextMatrix(1, C_USUARIO) = "Usuario"
   Grid.TextMatrix(0, C_DOCASOC) = "Documento"
   Grid.TextMatrix(1, C_DOCASOC) = "Asociado"
   
   If Grid.ColWidth(C_PROPIVA) > 0 Then
      Grid.TextMatrix(0, C_PROPIVA) = "Prop."
      Grid.TextMatrix(1, C_PROPIVA) = "IVA"
   End If
      
   If Grid.ColWidth(C_DETALLE) <> 0 Then
      Grid.Row = 1
      Grid.Col = C_DETALLE
      Grid.CellPictureAlignment = flexAlignCenterCenter
      Set Grid.CellPicture = FrmMain.Pc_Lupa
   End If
   
   'If lOper = O_EDIT Then
   '   Grid.TextMatrix(0, C_NUMCOMP) = ""
   'End If
 
   If lOper = O_VIEW Then
      Grid.Row = 1
      Grid.Col = C_CHECK
      Set Grid.CellPicture = Pc_Cent
   End If
   
   If lOper = O_EDIT Then
      Grid.FixedCols = 2
      Grid.ColWidth(C_DESCRIP) = Grid.ColWidth(C_DESCRIP) + 400
   Else
      Grid.FixedCols = 7
   End If
      
   Call FGrVRows(Grid)
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
         
   GridTot.TextMatrix(0, C_NOMBRE) = "TOTAL"

End Sub

Public Function FEdit(ByVal TipoLib As Integer, ByVal Mes As Integer, ByVal Ano As Integer, IdDoc As Long) As Integer

   lOper = O_EDIT
   lTipoLib = TipoLib
   lMes = Mes
   lAno = Ano
   
   Me.Show vbModal
   
   FEdit = lRc
   
   IdDoc = lIdDoc
   
End Function

Public Sub FView(ByVal TipoLib As Integer, Optional ByVal Mes As Integer = 0)

   lOper = O_VIEW
   lTipoLib = TipoLib
   lAno = gEmpresa.Ano
   lMes = Mes
   
   Me.Show vbModal
End Sub
Public Function FSelect(ByVal TipoLib As Integer, IdDoc As Long) As Integer

   lOper = O_SELECT
   lTipoLib = TipoLib
   lAno = gEmpresa.Ano
   
   Me.Show vbModal
   
   IdDoc = lIdDoc
   FSelect = lRc
   
End Function

Public Sub FViewLibroLeg(ByVal TipoLib As Integer, Optional ByVal Mes As Integer = 0, Optional ByVal Ano As Integer = 0, Optional ByVal Where As String = "", Optional ByVal TitEspecialVentas As String = "")

   lOper = O_VIEWLIBLEGAL
   lTipoLib = TipoLib
   lMes = Mes
   lAno = Ano
   lWhere = Where
   lTitEspecialVentas = TitEspecialVentas
   
   Me.Show vbModeless
End Sub

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Frm.FView (Valor)
      
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

Private Sub FillCb()
   Dim i As Integer
   Dim PrefLen As Integer
   Dim MesActual As Integer
   
   MesActual = GetMesActual()
   
   'If lOper <> O_EDIT Then
   
      Cb_Mes.AddItem " "
      Cb_Mes.ItemData(Cb_Mes.NewIndex) = 0
      
      Call FillMes(Cb_Mes)
                  
      If lMes > 0 Then
         Cb_Mes.ListIndex = lMes
      Else
         If MesActual > 0 Then
            Cb_Mes.ListIndex = MesActual
         Else
            Cb_Mes.ListIndex = GetUltimoMesConMovs()
         End If
      End If
           
   'Else           'O_EDIT
         
   '   Cb_Mes.AddItem gNomMes(MesActual)
   '   Cb_Mes.ItemData(Cb_Mes.NewIndex) = MesActual
   '   Cb_Mes.ListIndex = 0
            
   'End If
      
   'SF 14733340 se cambia el limite de llenado de cb de 2005 al año 2000
   'For i = gEmpresa.Ano To 2005 Step -1
   For i = gEmpresa.Ano To 2000 Step -1
      Cb_Ano.AddItem i
   Next i
   Cb_Ano.ListIndex = 0
   If lAno > 0 Then
      For i = 0 To Cb_Ano.ListCount - 1
         If Val(Cb_Ano.list(i)) = lAno Then
            Cb_Ano.ListIndex = i
            Exit For
         End If
      Next i
   End If
      
   PrefLen = Len("Libro de") + 1
   
   Call LoadTipoDoc

   Call AddItem(Cb_Estado, "(todos)", 0)
   'Cb_Estado.AddItem "(Todos)"
   'Cb_Estado.ItemData(Cb_Estado.NewIndex) = 0
   For i = 1 To UBound(gEstadoDoc)
      Call AddItem(Cb_Estado, gEstadoDoc(i), i)
      'Cb_Estado.AddItem gEstadoDoc(i)
      'Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   Cb_Estado.ListIndex = 0
   
   Call AddItem(Cb_Entidad, "", -1)
   'Cb_Entidad.AddItem ""
   'Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = -1
   For i = ENT_CLIENTE To ENT_OTRO
      Call AddItem(Cb_Entidad, gClasifEnt(i), i)
      'Cb_Entidad.AddItem gClasifEnt(i)
      'Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = i
      
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir

End Sub
Private Sub LoadTipoDoc()
   Dim i As Integer
   Dim Item As String
   Dim FindLib As Boolean
   Dim j As Integer
   
   Cb_TipoDoc.Clear
   
   If lTipoLib > 0 Then
      Cb_TipoDoc.AddItem "(Todos)"
      Cb_TipoDoc.ItemData(Cb_TipoDoc.NewIndex) = 0
      
      j = 1
      
      For i = 0 To UBound(gTipoDoc)
         
         If gTipoDoc(i).Nombre = "" Then
            Exit For
         End If
         
         If gTipoDoc(i).TipoLib = lTipoLib Then
         
            FindLib = True
            
            Cb_TipoDoc.AddItem gTipoDoc(i).Nombre
            Cb_TipoDoc.ItemData(Cb_TipoDoc.NewIndex) = gTipoDoc(i).TipoDoc
            
            Item = "[" & gTipoDoc(i).Diminutivo & "] " & gTipoDoc(i).Nombre
                     
            Load M_ItTipoDoc(j)
            M_ItTipoDoc(j).Caption = Item
            j = j + 1
         
         ElseIf FindLib Then   'se terminó el libro actual
            Exit For
         
         End If
         
      Next i
           
      If M_ItTipoDoc.Count > 1 Then
         M_ItTipoDoc(0).visible = False
      End If
      
      Cb_TipoDoc.ListIndex = 0
               
   End If
   
End Sub
Private Sub LoadCuentasMenu(ByVal Col As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
      
   If lTipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
      Q1 = Q1 & " WHERE CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " AND TipoLib = " & lTipoLib & " AND TipoValor ="
      
      Select Case Col
      
         Case C_EX_CODCUENTA
            
            If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_EXENTO
            Else
               Q1 = Q1 & LIBCOMPRAS_EXENTO
            End If
            
         Case C_AF_CODCUENTA
             
             If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_AFECTO
            Else
               Q1 = Q1 & LIBCOMPRAS_AFECTO
            End If
        
        Case C_TOT_CODCUENTA
        
            If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_TOTAL
            Else
               Q1 = Q1 & LIBCOMPRAS_TOTAL
            End If

         Case Else
            Exit Sub
            
      End Select
      
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY CuentasBasicas.Id "
      
      For i = 0 To UBound(lCuentas)
         lCuentas(i) = 0
      Next i
      
      'eliminamos los itemes del menú, menos el primero, que siempre debe estar, pero invisible
      M_ItCuenta(0).visible = True
      
      For i = M_ItCuenta.Count - 1 To 1 Step -1
         Unload M_ItCuenta(i)
      Next i
      
      Set Rs = OpenRs(DbMain, Q1)
      
      i = 1
      
      Do While Rs.EOF = False
         
         If i > MAX_CUENTAS Then
            Exit Do
         End If
      
         lCuentas(i) = vFld(Rs("IdCuenta"))
                  
         Item = Format(vFld(Rs("Codigo")), gFmtCodigoCta) & " [" & vFld(Rs("Nombre"), True) & "] " & FCase(vFld(Rs("Descripcion"), True))
         
         Load M_ItCuenta(i)
         M_ItCuenta(i).Caption = Item
      
         i = i + 1
         Rs.MoveNext
      Loop
           
      Call CloseRs(Rs)
      
      'agregamos un item (otra)
      Load M_ItCuenta(i)
      M_ItCuenta(i).Caption = MITEM_OTRA
      
      If M_ItCuenta.Count > 1 Then
         M_ItCuenta(0).visible = False
      End If
               
   End If
   
End Sub
Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 900
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 120
   GridTot.Width = Grid.Width - 230
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   Tx_CurrCell.Width = Bt_ToLeft.Left - Tx_CurrCell.Left - 200
   Bt_ToLeft.Top = Tx_CurrCell.Top
   Bt_ToRight.Top = Tx_CurrCell.Top
   Bt_Opciones.Top = Tx_CurrCell.Top
   Bt_Importar.Top = Tx_CurrCell.Top
   Bt_DelAll.Top = Tx_CurrCell.Top
   Bt_HlpImport.Top = Tx_CurrCell.Top
   Fr_Opciones.Top = Bt_Opciones.Top - Fr_Opciones.Height - 30
   
   
  
   
   Call FGrVRows(Grid)
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call UnLockAction(DbMain, lTipoLib, , , , False)
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   Dim IdDoc As Long
   Dim ValPrevLine As Boolean
   Dim F1 As Long, F2 As Long
   Dim Msg As String
   Dim IdxTipoDoc As Integer
'   Dim EntRelacionada As Boolean
   Dim Es14TER As Boolean
   '2763862
   Dim FechaLey21 As Long
   Dim FechaForm As Long
   FechaLey21 = DateSerial(2022, 4, 1)
   FechaForm = DateSerial(lAno, lMes + 1, 1) - 1
   'fin 2763862
   
   If lEditEnabled = False Then
      Exit Sub
   End If
                           
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      
      If Col <> C_FECHA Then
         MsgBox1 "Ingrese el día antes de continuar.", vbExclamation + vbOKOnly
         Exit Sub
      End If
      
      'Linea anterior tiene valor o está eliminada?
      ValPrevLine = (Row > Grid.FixedRows) And IsValidLine(Row - 1, Msg)
      'ValPrevLine = (ValPrevLine Or Grid.RowHeight(Row - 1) = 0)   'línea anterior borrada
      ValPrevLine = (ValPrevLine Or Grid.RowHeight(Row - 1) = 0 Or Val(Grid.TextMatrix(Row - 1, C_IDESTADO)) = ED_ANULADO) 'línea borrada o doc anulado
      
      If Not (Row = Grid.FixedRows Or ValPrevLine) Then
         If Not ValPrevLine Then
            MsgBox1 "Línea anterior incompleta o inválida. " & Msg, vbExclamation + vbOKOnly
         End If
         Exit Sub
      End If
                      
      'docs de año anterior que quedaron pendientes se dejan en estado centralizado
      If lAno < gEmpresa.Ano Then
         Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(ED_CENTRALIZADO)
         Grid.TextMatrix(Row, C_IDESTADO) = ED_CENTRALIZADO
      Else
         Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(ED_PENDIENTE)
         Grid.TextMatrix(Row, C_IDESTADO) = ED_PENDIENTE
      End If
      
      Grid.TextMatrix(Row, C_NUMLIN) = Row - Grid.FixedRows + 1
      If Row > Grid.FixedRows Then
         Grid.TextMatrix(Row, C_CORRINTERNO) = vFmt(Grid.TextMatrix(Row - 1, C_CORRINTERNO)) + 1
      Else
         Grid.TextMatrix(Row, C_CORRINTERNO) = 1
      End If
            
      If Row >= Grid.rows - 2 Then
         Grid.rows = Grid.rows + 1
      End If
      
   ElseIf Val(Grid.TextMatrix(Row, C_MOVEDITED)) <> 0 And Col <> C_DETALLE Then
      Call PostClick(Bt_DetDoc)
      Exit Sub
      
   ElseIf Not ValidaEstadoEdit(Row) And Col <> C_DETALLE Then
      MsgBeep vbExclamation
      Exit Sub
   
   End If
   
   Call ValidaFExported(Row)
   
   Grid.TxBox.MaxLength = 0
   
   IdxTipoDoc = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   

   Select Case Col
   
      Case C_FECHA
     
         Grid.TxBox.MaxLength = 2
         If month(Now) <> CbItemData(Cb_Mes) Then
            Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), F1, F2)
            Grid.TextMatrix(Row, Col) = Day(F2)
            LDia = Day(F2)
         Else
            Grid.TextMatrix(Row, Col) = Day(Int(lAno))
            LDia = Day(Int(lAno))
         End If
         
         If Row > Grid.FixedRows And Grid.TextMatrix(Row, C_TIPODOC) = "" Then
            Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
            Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
            Grid.TextMatrix(Row, C_GIRO) = ""
            Grid.TextMatrix(Row, C_DOCIMPEXP) = Grid.TextMatrix(Row - 1, C_DOCIMPEXP)
            
            If lTipoLib = LIB_VENTAS Then
            
               If Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VENTASINDOC And Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VALEPAGOELECTR Then
               
                  If Val(Grid.TextMatrix(Row - 1, C_NUMDOCHASTA)) <> 0 Then
                     
                     If IsNumeric(Trim(Grid.TextMatrix(Row - 1, C_NUMDOCHASTA))) Then
                        Grid.TextMatrix(Row, C_NUMDOC) = Val(Grid.TextMatrix(Row - 1, C_NUMDOCHASTA)) + 1
                     End If
                     
                  Else
                     
                     If IsNumeric(Trim(Grid.TextMatrix(Row - 1, C_NUMDOC))) Then
                        Grid.TextMatrix(Row, C_NUMDOC) = Val(Grid.TextMatrix(Row - 1, C_NUMDOC)) + 1
                     End If
                     
                  End If
                  
               End If
                  
               If Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Then
               
                  If Not lMsgNotaCred Then
                     If FechaForm > FechaLey21 Then
                        MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 6 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     Else
                        MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     End If
                     lMsgNotaCred = True
                  End If
                  
               End If
               
            End If
            
            If Val(Grid.TextMatrix(Row, C_IDSUCURSAL)) = 0 And CbItemData(Cb_Sucursal) <> 0 Then
               Grid.TextMatrix(Row, C_IDSUCURSAL) = CbItemData(Cb_Sucursal)
               Grid.TextMatrix(Row, C_SUCURSAL) = Cb_Sucursal
            End If

         End If
         
         EdType = FEG_Edit

      Case C_TIPODOC
        
         '2814014
         'Grid.TxBox.MaxLength = 3
         Grid.TxBox.MaxLength = 4
         'fin 2814014
      
         If Row > Grid.FixedRows And Grid.TextMatrix(Row, Col) = "" Then
            Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
            Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
         End If
         
         EdType = FEG_Edit
               
      Case C_NUMFISCIMPR, C_NUMINFORMEZ
     
         If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
            Grid.TxBox.MaxLength = MAX_NUMDOCMRG
            EdType = FEG_Edit
         End If
         
      Case C_NUMDOC
     
         If Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
            Grid.TxBox.MaxLength = MAX_NUMDOCLEN
            EdType = FEG_Edit
         End If
         
      Case C_NUMDOCHASTA
             
'         If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTAEX Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLEXENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then      'venta sin documento
         If IdxTipoDoc > 0 Then
            If gTipoDoc(IdxTipoDoc).TieneNumDocHasta <> 0 Then
               Grid.TxBox.MaxLength = MAX_NUMDOCLEN
               EdType = FEG_Edit
            End If
         End If
         
      Case C_CANTBOLETAS
         If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VALEPAGOELECTR Then     'vale de pago electrónico VPE
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
         End If
         
      Case C_GIRO
      
         If Grid.TextMatrix(Row, C_TIPODOC) = "FAV" Or Grid.TextMatrix(Row, C_TIPODOC) = "FVE" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDV" Then
            If Trim(Grid.TextMatrix(Row, Col)) = "" Then
               Grid.TextMatrix(Row, Col) = "No"
            Else
               Grid.TextMatrix(Row, Col) = ""
            End If
                  
            Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         End If
      
      Case C_DTE
      
         If Trim(Grid.TextMatrix(Row, Col)) = "" And Grid.TextMatrix(Row, C_TIPODOC) <> "IMP" And Grid.TextMatrix(Row, C_TIPODOC) <> "FIC" And Grid.TextMatrix(Row, C_TIPODOC) <> "FIV" And Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VALEPAGOELECTR Then
            Grid.TextMatrix(Row, Col) = "x"
         Else
            Grid.TextMatrix(Row, Col) = ""
         End If
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         
      'Case C_CHECK
      
         'If Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(ED_PENDIENTE) Then
                        
            'If Grid.CellPicture = 0 Then
            '   Call FGrSetPicture(Grid, Row, C_CHECK, Pc_Check, 0)
            'Else
            '   Set Grid.CellPicture = LoadPicture()
            'End If
            
         'End If
         
'      Case C_PROPIVA    'no se maneja acá sino en el doble-click porque si no se cambia cuando uno se mueve con el Enter
'         If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).AceptaPropIVA Then
'            Grid.TextMatrix(Row, C_IDPROPIVA) = (Val(Grid.TextMatrix(Row, C_IDPROPIVA)) + 1) Mod 4
'            Grid.TextMatrix(Row, C_PROPIVA) = Left(gStrPropIVA(Val(Grid.TextMatrix(Row, C_IDPROPIVA))), 1)
'            Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
'
'         ElseIf Val(Grid.TextMatrix(Row, C_IDPROPIVA)) <> 0 Then
'            Grid.TextMatrix(Row, C_IDPROPIVA) = PIVA_SINPROP
'            Grid.TextMatrix(Row, C_PROPIVA) = gStrPropIVA(PIVA_SINPROP)
'            Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
'         End If
        
         
      Case C_RUT
         Grid.TxBox.MaxLength = 13
         EdType = FEG_Edit
         
      'Case C_NOMBRE
         'If Val(Grid.TextMatrix(Row, C_IDENTIDAD)) = 0 Then
         '   Grid.TxBox.MaxLength = 50
         '   EdType = FEG_Edit
         'End If
                       
      Case C_SUCURSAL
         EdType = FEG_List
      
      Case C_EXENTO
      
         If IdxTipoDoc >= 0 Then
'            If gTipoDoc(IdxTipoDoc).TieneExento Then
            If gTipoDoc(IdxTipoDoc).TieneExento And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 Then
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
            End If
         End If
         
      Case C_AFECTO, C_IVA, C_OTROIMP
      
         If IdxTipoDoc >= 0 Then
      
'            If gTipoDoc(IdxTipoDoc).TieneAfecto Then
            If gTipoDoc(IdxTipoDoc).TieneAfecto <> 0 And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 Then
               
               If Col = C_OTROIMP And lIdCuentaOtrosImp = 0 Then
                  MsgBox1 "Falta definir la cuenta de otros impuestos en el menú " & vbCrLf & vbCrLf & "Configuración>>Configuración Inicial>>Definir Cuentas Básicas>>Impuestos.", vbExclamation + vbOKOnly
                  Exit Sub
               End If
               
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
               
            End If
         
         End If
         
      Case C_TOTAL
            
         If IdxTipoDoc >= 0 Then
'            If Not gTipoDoc(IdxTipoDoc).TieneExento And Not gTipoDoc(IdxTipoDoc).TieneAfecto Then   'no se ingresa ni afecto ni exento => se ingresa el total (caso de boletas de venta y devoluciones con boleta)
            If gTipoDoc(IdxTipoDoc).IngresarTotal <> 0 Then  'se ingresa el total (caso de boletas de venta y devoluciones con boleta)
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
            End If
         End If
         
      Case C_VENTASACUM
      
         If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
         End If

         
      Case C_EX_CODCUENTA
         If IdxTipoDoc >= 0 Then
            If gTipoDoc(IdxTipoDoc).TieneExento <> 0 Or vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Then
'            If gTipoDoc(IdxTipoDoc).TieneExento And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 And vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Then
               Grid.TxBox.MaxLength = 20
               EdType = FEG_Edit
            End If
         End If
         
      Case C_AF_CODCUENTA
         If IdxTipoDoc >= 0 Then
            If gTipoDoc(IdxTipoDoc).TieneAfecto <> 0 Or vFmt(Grid.TextMatrix(Row, C_AFECTO)) Then
'            If gTipoDoc(IdxTipoDoc).TieneAfecto <> 0 And vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 Then
               Grid.TxBox.MaxLength = 20
               EdType = FEG_Edit
            End If
         End If
         
      Case C_TOT_CODCUENTA
         Grid.TxBox.MaxLength = 20
         EdType = FEG_Edit
   
   
      Case C_FECHAEMIORI
         
         Grid.TxBox.MaxLength = 10
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
            Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(Grid.TextMatrix(Row, C_LNGFECHAEMIORI), SDATEFMT)
         Else
            Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))))
            Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)), SDATEFMT)
         End If
         
         EdType = FEG_Edit
         
      Case C_FECHAVENC
      
'         EntRelacionada = GetEntRelacionada(Val(Grid.TextMatrix(Row, C_IDENTIDAD))) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS   (Claudio Vollegas - 11 jul 2017)
         Es14TER = gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAVENC)) > 0 Then
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
         
         Else
            If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
               If Es14TER Then    'si es 14 TER => pago contado
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))
               Else
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
               End If
            
            Else   'usamos la fecha de recepción del documento
               
               If Es14TER Then    'si es 14 TER => pago contado
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))))
            
               Else
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA)))))
               End If
            End If
            
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)), SDATEFMT)
            
         End If
        
         Grid.TxBox.MaxLength = 10
         EdType = FEG_Edit
         
      Case C_CORRINTERNO
         Grid.TxBox.MaxLength = 9
         EdType = FEG_Edit
         
      Case C_DESCRIP
         Grid.TxBox.MaxLength = 100
         EdType = FEG_Edit
         
      'Case C_ESTADO, C_NUMCOMP
      
      'Case C_DETALLE
      '   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
         
      '   If IdDoc <> 0 Then
         
      '      Set FrmDoc = New FrmDocumento
      '      Call FrmDoc.FView(IdDoc)
      '      Set FrmDoc = Nothing
            
      '   End If
      
   End Select

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim IdEnt As Long
   Dim Rc As Integer
   Dim Nombre As String
   Static EnAcceptVal As Boolean
   Dim IdCuenta As Long
   Dim Cod As String
   Dim DescCta As String
   Dim NombCta As String
   Dim UltimoNivel As Boolean
   Dim ColId As Integer
   Dim ColCta As Integer
   Dim IdActFijo As Long
   Dim Frm As Form
   Dim OldIdCuenta As Long
   Dim IVA As Double
   Dim Afecto As Double
   Dim Exento As Double
   Dim TipoDoc As Integer
   Dim IdxDoc As Integer
   Dim NotValidRut As Boolean
   Dim TipoDocExp As Boolean
   Dim AuxRut As String
   Dim CtaIntName As String
   Dim FEmision As Long, FVenc30Dias As Long
   Dim IdxTipoDoc As Integer
'   Dim EntRelacionada As Boolean
   Dim Es14TER As Boolean
    '2763862
   Dim FechaLey21 As Long
   Dim FechaForm As Long
   FechaLey21 = DateSerial(2022, 3, 24)
   FechaForm = DateSerial(lAno, lMes, LDia)
   
   
   'fin 2763862
   If EnAcceptVal Then
      Exit Sub
   End If
   
   EnAcceptVal = True
   
   Action = vbOK
   
   IdxTipoDoc = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
'   Es14TER = gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS                       'Claudio Villegas - 11 Jul 2017
   Es14TER = gEmpresa.Franq14Ter                                                    'Claudio Villegas - 24 ago 2017
  
   Value = Trim(Value)
   Value = ReplaceStr(Value, vbCr, "")
   Value = ReplaceStr(Value, vbLf, "")

  
   Select Case Col
   
      Case C_FECHA
         LDia = Value
         Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), FirstDay, LastDay)
         
         If Val(Value) < 1 Or Val(Value) > Day(LastDay) Then
            MsgBox1 "Día inválido.", vbExclamation + vbOKOnly
            Value = Day(LastDay)
            Action = vbCancel
            EnAcceptVal = False
            Exit Sub
         End If
         
         If Grid.TextMatrix(Row, C_FECHAEMIORI) = "" Then
         
            Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Value)))
            
            If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
               Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(Grid.TextMatrix(Row, C_LNGFECHAEMIORI), SDATEFMT)
               
'               EntRelacionada = GetEntRelacionada(Val(Grid.TextMatrix(Row, C_IDENTIDAD))) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS
               
               If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) = 0 Then
               
                  If Es14TER Then
                     'pago contado
                     Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))
                  Else
                     'proponemos fecha de vencimiento a 30 días
                     Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
                  End If
               End If
               
               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
            End If
         
         End If
                  
         If Ch_RepetirGlosa <> 0 And Row > Grid.FixedRows Then
            Grid.TextMatrix(Row, C_DESCRIP) = Grid.TextMatrix(Row - 1, C_DESCRIP)
         End If
            
      Case C_TIPODOC
                                                
         Value = Trim(Value)
         
         If Value <> "" Then
         
            TipoDoc = FindTipoDoc(lTipoLib, Value)
            
            If TipoDoc > 0 Then
            
               Grid.TextMatrix(Row, C_IDTIPODOC) = TipoDoc
               Grid.TextMatrix(Row, Col) = Value
               Grid.TextMatrix(Row, C_DOCIMPEXP) = CInt(gTipoDoc(GetTipoDoc(lTipoLib, TipoDoc)).DocImpExp)
               
               Call ActCambioTipoDoc(Row)
               
               Call CalcTot
               
               If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VENTASINDOC Then
                  Grid.TextMatrix(Row, C_NUMDOC) = ""
               End If
               
               Grid.TextMatrix(Row, C_GIRO) = ""
               
               If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).AceptaPropIVA Then
                  If Grid.TextMatrix(Row, C_PROPIVA) = "" Then
                     Grid.TextMatrix(Row, C_IDPROPIVA) = PIVA_TOTAL
                     Grid.TextMatrix(Row, C_PROPIVA) = Left(gStrPropIVA(Grid.TextMatrix(Row, C_IDPROPIVA)), 1)
                  End If
               Else
                  Grid.TextMatrix(Row, C_IDPROPIVA) = 0
                  Grid.TextMatrix(Row, C_PROPIVA) = ""
               End If

               
               If Value = "NCV" Then
                  If Not lMsgNotaCred Then
                     'MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     ' 2763862 desde abril debe mostrar otro mensaje cambia de 3 meses a 6
                     If FechaForm > DateSerial(2022, 4, 1) Then
                        MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 6 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     Else
                        MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     End If
                     lMsgNotaCred = True
                  End If
               End If
               
'               If Not (Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTAEX Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLEXENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA) Then     'venta sin documento
               If gTipoDoc(GetTipoDoc(lTipoLib, TipoDoc)).TieneNumDocHasta = 0 Then
                  Grid.TextMatrix(Row, C_NUMDOCHASTA) = ""    'esto se hizo por un cliente que cambió el tipo para engañar al sistema
               End If
                  
            Else
               MsgBox1 "Tipo de documento inválido. Presione el botón derecho del mouse para Ayuda.", vbExclamation + vbOKOnly
               Grid.TextMatrix(Row, C_TIPODOC) = ""
               Grid.TextMatrix(Row, C_IDTIPODOC) = 0
               Action = vbRetry
               
            End If
            
         End If
         
'      Case C_NUMFISCIMPR, C_NUMINFORMEZ
'         Value = Trim(Value)
'         Value = ReplaceStr(Value, vbCr, "")
'         Value = ReplaceStr(Value, vbLf, "")
         
         
      Case C_NUMDOC, C_NUMDOCHASTA
         
         Grid.TextMatrix(Row, Col) = Value
         
         If Not Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VALEPAGOELECTR Then
            Grid.TextMatrix(Row, C_CANTBOLETAS) = ""
            If vFmt(Grid.TextMatrix(Row, C_NUMDOC)) > 0 Then
               If vFmt(Grid.TextMatrix(Row, C_NUMDOCHASTA)) > 0 And vFmt(Grid.TextMatrix(Row, C_NUMDOCHASTA)) >= vFmt(Grid.TextMatrix(Row, C_NUMDOC)) Then
                  Grid.TextMatrix(Row, C_CANTBOLETAS) = Format(vFmt(Grid.TextMatrix(Row, C_NUMDOCHASTA)) - vFmt(Grid.TextMatrix(Row, C_NUMDOC)), NUMFMT) + 1
               End If
            End If
         End If
         
      Case C_RUT
      
         If Value = "" Then
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbOK
                  
         ElseIf Value = "0-0" Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
         
         ElseIf (Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And Not ValidCID(Value)) Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
         
         Else
         
            IdEnt = GetIdEntidad(Value, Nombre, NotValidRut)
            
            If IdEnt = 0 Then
               If MsgBox1("Esta entidad no ha sido ingresada a la lista de entidades predefinidas." & vbNewLine & vbNewLine & "¿Desea agregarla ahora?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                  AuxRut = FmtCID(vFmtCID(Value))
                  If AuxRut = "0-0" Then
                     AuxRut = Trim(Value)
                  End If
                  If NewEntidad(Row, AuxRut) <> vbOK Then
                     Value = ""
                  End If
               Else
                  Value = ""
                  Grid.TextMatrix(Row, C_IDENTIDAD) = 0
                  Grid.TextMatrix(Row, C_NOMBRE) = ""
               End If
            Else
               Value = FmtCID(vFmtCID(Value, NotValidRut = False), NotValidRut = False)
               Grid.TextMatrix(Row, C_RUT) = Value
               Grid.TextMatrix(Row, C_NOMBRE) = Nombre
               Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
               
'               If GetEntRelacionada(IdEnt) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then    'si es Ent Relacionada y 14TER => pago contado
'               If gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then    'si es 14TER => pago contado  (Claudio Villegas - 11 Jul 2017)
'
'                  If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
'                     Grid.TextMatrix(Row, C_LNGFECHAVENC) = vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))
'                  Else
'                     Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))))
'                  End If
'
'                  Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
'               End If
               
            End If
            
            Call ValidaNumDoc(Row)
         
         
         End If
         
      'Case C_NOMBRE
         'If Trim(Value) = "" Then
         '   MsgBox1 "Nombre o razón social inválido.", vbExclamation + vbOKOnly
         '   Action = vbCancel
         'End If
         
      Case C_SUCURSAL
         Grid.TextMatrix(Row, C_IDSUCURSAL) = CbItemData(Grid.CbList(C_SUCURSAL))
         
      Case C_EXENTO, C_AFECTO, C_IVA, C_OTROIMP, C_VENTASACUM
      
         If Grid.TextMatrix(Row, C_TIPODOC) = "IMP" And Grid.TextMatrix(Row, C_RUT) = "" Then
            AddRutDefault (Row)
         End If
      
         If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).EsRebaja Then
            If Col <> C_OTROIMP Then
               Value = Format(Abs(vFmt(Value)) * -1, NEGNUMFMT)
            Else
               Value = Format(vFmt(Value) * -1, NEGNUMFMT)   'para permitir ingresar valores negativos en los otros impuestos. Al ser rebaja y el usuario ingresa un valor negativo, queda positivo
            End If
         
         Else
            
            If vFmt(Value) < 0 And Col <> C_OTROIMP Then
               MsgBox1 "Valor inválido.", vbExclamation + vbOKOnly
               Action = vbRetry
            End If
            
            Value = Format(vFmt(Value), NEGNUMFMT)
            
         End If
         
         Grid.TextMatrix(Row, Col) = Value
            
         If Col = C_EXENTO Then
            If vFmt(Value) <> 0 Then
               If Val(Grid.TextMatrix(Row, C_EX_IDCUENTA)) = 0 Then
                  Grid.TextMatrix(Row, C_EX_IDCUENTA) = lCtaExento.id
                  Grid.TextMatrix(Row, C_EX_CODCUENTA) = FmtCodCuenta(lCtaExento.Codigo)
                  Grid.TextMatrix(Row, C_EX_CUENTA) = lCtaExento.Descripcion
               End If
            Else
               Grid.TextMatrix(Row, C_EX_IDCUENTA) = 0
               Grid.TextMatrix(Row, C_EX_CODCUENTA) = ""
               Grid.TextMatrix(Row, C_EX_CUENTA) = ""
            End If
            
            Call GridActivoFijo(lCtaExento.id, Row, C_EX_CODCUENTA)

            Call CalcTotRow(Row, True)
         
         ElseIf Col = C_AFECTO Then
            If vFmt(Value) <> 0 Then
               If Val(Grid.TextMatrix(Row, C_AF_IDCUENTA)) = 0 Then
                  Grid.TextMatrix(Row, C_AF_IDCUENTA) = lCtaAfecto.id
                  Grid.TextMatrix(Row, C_AF_CODCUENTA) = FmtCodCuenta(lCtaAfecto.Codigo)
                  Grid.TextMatrix(Row, C_AF_CUENTA) = lCtaAfecto.Descripcion
               End If
            Else
               Grid.TextMatrix(Row, C_AF_IDCUENTA) = 0
               Grid.TextMatrix(Row, C_AF_CODCUENTA) = ""
               Grid.TextMatrix(Row, C_AF_CUENTA) = ""
            End If
            
            Call GridActivoFijo(lCtaAfecto.id, Row, C_AF_CODCUENTA)
            
            Call CalcTotRow(Row, True)
         
         Else
            Call CalcTotRow(Row, False)
            
            If Col = C_IVA Then    'esto se hace por si el usuario cambia el valor y luego se arrepiente en función CalcTotRow
               Value = Grid.TextMatrix(Row, C_IVA)
            End If
            
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 Then
            If Val(Grid.TextMatrix(Row, C_TOT_IDCUENTA)) = 0 Then
               Grid.TextMatrix(Row, C_TOT_IDCUENTA) = lCtaTotal.id
               Grid.TextMatrix(Row, C_TOT_CODCUENTA) = FmtCodCuenta(lCtaTotal.Codigo)
               Grid.TextMatrix(Row, C_TOT_CUENTA) = lCtaTotal.Descripcion
            End If
         Else
            Grid.TextMatrix(Row, C_TOT_IDCUENTA) = 0
            Grid.TextMatrix(Row, C_TOT_CODCUENTA) = ""
            Grid.TextMatrix(Row, C_TOT_CUENTA) = ""
         End If
        
         Call CalcTot
         
         If vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Then
            If Val(Grid.TextMatrix(Row, C_IVA_IDCUENTA)) = 0 Then
               Grid.TextMatrix(Row, C_IVA_IDCUENTA) = lIdCuentaIVA
            End If
         Else
            Grid.TextMatrix(Row, C_IVA_IDCUENTA) = 0
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Then
            'si es factura de compras, nota de crédito de fac. compras o nota de débito de fac. compras, se pone la cuenta al revés
            If Grid.TextMatrix(Row, C_TIPODOC) = "FCC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCF" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDF" Or Grid.TextMatrix(Row, C_TIPODOC) = "FCV" Then
               Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = lIdCuentaOtrosImpFacCompra
            Else
               Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = lIdCuentaOtrosImp
            End If
         Else
            Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = 0
         End If
                  
      Case C_EX_CODCUENTA, C_AF_CODCUENTA, C_TOT_CODCUENTA
      
         If Value = "" Then
            IdCuenta = 0
            Cod = ""
            DescCta = ""
         Else
            Cod = Trim(ReplaceStr(Value, "-", ""))
            If Len(Cod) < Len(VFmtCodigoCta(gFmtCodigoCta)) Then   'asumimos que está usando nombre corto
               NombCta = UCase(Trim(Value))
               Cod = ""
            Else
               NombCta = ""
            End If
            
            IdCuenta = GetIdCuenta(NombCta, Cod, DescCta, UltimoNivel)
         End If
         
         Select Case Col
            Case C_EX_CODCUENTA
               ColId = C_EX_IDCUENTA
               ColCta = C_EX_CUENTA
            Case C_AF_CODCUENTA
               ColId = C_AF_IDCUENTA
               ColCta = C_AF_CUENTA
            Case C_TOT_CODCUENTA
               ColId = C_TOT_IDCUENTA
               ColCta = C_TOT_CUENTA
         End Select
         
         If IdCuenta = 0 Then
         
            If Value <> "" Then
               'IdCuenta = 0
               'Cod = ""
               'DescCta = ""
               MsgBeep vbExclamation
               Action = vbRetry
            End If
            
         ElseIf UltimoNivel = False Then
            'IdCuenta = 0
            'Cod = ""
            'DescCta = ""
            MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
            Action = vbRetry
                        
'         ElseIf Not EsCuentaBasica(IdCuenta, Col) Then
'            'IdCuenta = 0
'            'Cod = ""
'            'DescCta = ""
'
'            MsgBox1 "Esta cuenta no es válida para este tipo de valor, de acuerdo a la configuración básica de la empresa.", vbExclamation + vbOKOnly
'            Action = vbRetry
            
         Else
                     
            Grid.TextMatrix(Row, ColId) = IdCuenta       'asignamos nuevo valor
            Value = Format(Cod, gFmtCodigoCta)
            Grid.TextMatrix(Row, ColCta) = DescCta
            
            Call GridActivoFijo(IdCuenta, Row, Col)
            
            'vemos si la cuenta de total es de Caja o Banco para mandar warning
            If C_TOT_CODCUENTA Then
               
               If GetAtribCuenta(IdCuenta, ATRIB_CAJA) <> 0 Or GetAtribCuenta(IdCuenta, ATRIB_CONCILIACION) <> 0 Then
                  If lTipoLib = LIB_COMPRAS Then
                     CtaIntName = "Proveedor"
                  Else
                     CtaIntName = "Cliente"
                  End If
                  MsgBox1 "Atención!!" & vbNewLine & vbNewLine & "Si selecciona una cuenta de pago contado (Banco o Caja) el informe analítico reflejará valores erróneos." & vbNewLine & vbNewLine & "Se sugiere utilizar una cuenta intermedia de " & CtaIntName & ", para registrar la centralización del documento, y luego, realizar el pago correspondiente, en dos operaciones separadas. De esta forma, quedará registrada la operación en la cuenta de " & CtaIntName & " correspondiente, para su posterior análisis de gestión.", vbExclamation + vbOKOnly
               End If
                                          
            End If
            
         End If
         
      Case C_TOTAL
         
         If vFmt(Value) < 0 Then
            MsgBox1 "Valor inválido.", vbExclamation + vbOKOnly
            Action = vbRetry
               
         ElseIf EsIngresoTotal(Row) Then
         
            Call CalcIngresoTotal(Row, Col, Value)
            '3071158
            Call CalcTot
            '3071158
    
         End If
            
      Case C_FECHAEMIORI
         
         Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), FirstDay, LastDay)
         
         FEmision = GetDate(Value, "dmy")
         
         If lTipoLib = LIB_VENTAS Then
            If FEmision >= FirstDay And FEmision <= LastDay Then
               Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = FEmision
               
            ElseIf Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Then    'para notas de crédito, permitimos ingreso fuera de plazo (solicitado el día 29/06/2011 por Victro Morales)
               '2763862
               If FEmision >= FechaLey21 Then
                    MsgBox1 "Recuerde verificar la rebaja Débito Fiscal según instrucciones de Ley sobre impuesto a las ventas y Servicios y Ley 21.398", vbOKOnly + vbInformation
               Else
                    MsgBox1 "Recuerde verificar la rebaja Débito Fiscal según instrucciones del Articulo 21 Ley sobre Impuesto a las ventas y servicios", vbOKOnly + vbInformation
               End If
               ' fin 2763862
               Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = FEmision
            
            Else
               MsgBox1 "Fecha de emisión inválida.", vbExclamation + vbOKOnly
               Action = vbCancel
               EnAcceptVal = False
               Exit Sub
            End If
         
         Else
            If FEmision > LastDay Then
               MsgBox1 "Fecha de emisión inválida.", vbExclamation + vbOKOnly
               Action = vbCancel
               EnAcceptVal = False
               Exit Sub
            End If
               
            Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = FEmision
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
            Value = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)), SDATEFMT)
'            EntRelacionada = GetEntRelacionada(Val(Grid.TextMatrix(Row, C_IDENTIDAD))) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS
            
            'proponemos fecha de vencimiento a 30 días
            If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) = 0 Then
               If Es14TER Then    'pago contado
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))
               Else        'Crédito 30 días
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
               End If
               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
               
            Else
               If Es14TER Then
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))
               
               Else
                  FVenc30Dias = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
                  
                  If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) <> FVenc30Dias Then
                     If MsgBox1("¿Desea ajustar la fecha de vencimiento a 30 días?", vbQuestion + vbYesNo) = vbYes Then
                        Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
                     End If
                  End If
                  
               End If
               
               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
            End If
         End If
         
         'se elimina esta validación ya que se puede aprovechar los dos períodos siguientes (Victor Morales 4 ago 2020)
'         If Grid.TextMatrix(Row, C_TIPODOC) = "IMP" Then
'            'si el doc de importación no es del mismo mes, hay que mostrar advertencia
'            If CbItemData(Cb_Mes) <> Month(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))) Or gEmpresa.Ano <> Year(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))) Then
'               MsgBox1 "Recuerde que según las normas de la Ley sobre Impuesto a las Ventas y Servicios, el IVA se puede aprovechar dentro del mismo periodo de emisión del documento.", vbInformation + vbOKOnly
'            End If
'         End If

      
      Case C_FECHAVENC
         
         Grid.TextMatrix(Row, C_LNGFECHAVENC) = GetDate(Value, "dmy")
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAVENC)) > 0 Then
            Value = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
         End If
         
      'Case C_DESCRIP
         
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
   End If
   
   EnAcceptVal = False

End Sub

Private Sub Grid_BeforeFocus(ByVal Row As Integer, ByVal Col As Integer, ByVal EdType As FlexEdGrid2.FEG2_EdType)
   
   If Col = C_TIPODOC Then
      Grid.CbEdList(Col).Width = 1800
   End If
   
End Sub

Private Sub Grid_CbEditKeyPress(ByVal Col As Integer, KeyAscii As Integer)

   If Col = C_TIPODOC Then
      Call KeyUpper(KeyAscii)
   End If

End Sub

Private Sub Grid_Click()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Fr_Opciones.visible Then
      Fr_Opciones.visible = False
   End If
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If

   Call OrdenaPorCol(Col)
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   
   Col = Grid.MouseCol
   Row = Grid.MouseRow
   
   'primero entra a Grid_DblClick y luego a Grid_BeforeEdit
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      Exit Sub
   End If

   If Col = C_CHECK Then
   
      If Bt_Centralizar.visible And Bt_Centralizar.Enabled Then
         If Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PENDIENTE Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_APROBADO Or (Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PAGADO And vFmt(Grid.TextMatrix(Row, C_IDCOMPCENT)) = 0) Then
            
            'permitimos centralizar los documentos que están en estado Pagado pero que no han sido centralizados
            
            If lMsgPagadoNoCent = False And Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PAGADO And vFmt(Grid.TextMatrix(Row, C_IDCOMPCENT)) = 0 Then
               If MsgBox1("ATENCIÓN:" & vbNewLine & vbNewLine & "Este documento está Pagado pero no Centralizado." & vbNewLine & vbNewLine & "¿Desea marcarlo para centralizar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               Else
                  lMsgPagadoNoCent = True
               End If
            End If
            
            Grid.Row = Row
            Grid.Col = Col
            
            If Grid.CellPicture = 0 Then
               Call FGrSetPicture(Grid, Row, C_CHECK, Pc_Check, 0)
            Else
               Set Grid.CellPicture = LoadPicture()
            End If
            
         End If
      End If
         
   ElseIf Col = C_DETACTFIJO Then
      If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO And Grid.TextMatrix(Row, C_DETACTFIJO) <> "" Then
         Call PostClick(Bt_ActivoFijo)
      End If
   
   ElseIf Col = C_DETALLE Or Col = C_DOCASOC Then
      Call PostClick(Bt_DetDoc)
      
   ElseIf Col = C_PROPIVA Then
         
      If lEditEnabled = False Then
         Exit Sub
      End If

      If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).AceptaPropIVA Then
         Grid.TextMatrix(Row, C_IDPROPIVA) = (Val(Grid.TextMatrix(Row, C_IDPROPIVA)) + 1) Mod 4
         Grid.TextMatrix(Row, C_PROPIVA) = Left(gStrPropIVA(Val(Grid.TextMatrix(Row, C_IDPROPIVA))), 1)
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)

      ElseIf Val(Grid.TextMatrix(Row, C_IDPROPIVA)) <> 0 Then
         Grid.TextMatrix(Row, C_IDPROPIVA) = PIVA_SINPROP
         Grid.TextMatrix(Row, C_PROPIVA) = gStrPropIVA(PIVA_SINPROP)
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)

      End If
      
      Call Grid_SelChange
      
   ElseIf Bt_Sel.visible And Bt_Sel.Enabled Then
      Call PostClick(Bt_Sel)
      
   ElseIf lOper = O_VIEW Or lOper = O_VIEWLIBLEGAL Or lEditEnabled = False Then
      Call PostClick(Bt_DetDoc)
      
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   Select Case Grid.Col
   
      Case C_FECHA
         Call KeyNum(KeyAscii)
      
      Case C_TIPODOC
         Call KeyUpper(KeyAscii)
         
      Case C_NUMDOC, C_NUMDOCHASTA
         
         If (lTipoLib = LIB_COMPRAS And Grid.TextMatrix(Grid.Row, C_TIPODOC) = "IMP") Or (lTipoLib = LIB_VENTAS And Grid.TextMatrix(Grid.Row, C_TIPODOC) = TDOC_VALEPAGOELECTR) Then
            Call KeyName(KeyAscii)
         Else
            Call KeyNum(KeyAscii)
         End If
         
       Case C_CORRINTERNO
         Call KeyNum(KeyAscii)
        
      Case C_NUMFISCIMPR, C_NUMINFORMEZ
         Call KeyNum(KeyAscii)
            
      Case C_CANTBOLETAS
         Call KeyNum(KeyAscii)
            
      Case C_RUT
         Call KeyName(KeyAscii)
         Call KeyUpper(KeyAscii)
         
      Case C_NOMBRE
         Call KeyName(KeyAscii)
         
      Case C_EXENTO, C_AFECTO, C_IVA, C_OTROIMP, C_TOTAL, C_VENTASACUM
         Call KeyNum(KeyAscii)
         
      Case C_FECHAEMIORI, C_FECHAVENC
         Call KeyDate(KeyAscii)
      
      Case C_DESCRIP
         Call KeyName(KeyAscii)
         
   End Select
   
End Sub

Private Sub CalcTotRow(ByVal Row As Integer, Optional ByVal RecalcIVA As Boolean)
   Dim Tot As Double
   Dim IVA As Double
   
   IVA = Round(vFmt(Grid.TextMatrix(Row, C_AFECTO)) * gIVA)
   
   If Not RecalcIVA Then
      If vFmt(Grid.TextMatrix(Row, C_IVA)) <> IVA And (vFmt(Grid.TextMatrix(Row, C_IVA)) + 2 <= IVA Or vFmt(Grid.TextMatrix(Row, C_IVA)) - 2 >= IVA) Then
         If MsgBox1("En la línea " & Grid.TextMatrix(Row, C_NUMLIN) & ", " & Grid.TextMatrix(Row, C_TIPODOC) & " " & Grid.TextMatrix(Row, C_NUMDOC) & ", el valor del IVA difiere en más de dos unidades del valor calculado por el sistema." & vbNewLine & vbNewLine & "¿Está seguro que desea dejar este valor?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
         End If
      End If
   Else
      Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
   End If
   
   Tot = vFmt(Grid.TextMatrix(Row, C_EXENTO)) + vFmt(Grid.TextMatrix(Row, C_AFECTO)) + vFmt(Grid.TextMatrix(Row, C_IVA)) + vFmt(Grid.TextMatrix(Row, C_OTROIMP))
   Grid.TextMatrix(Row, C_TOTAL) = Format(Tot, NEGNUMFMT)

   
End Sub

Private Function valida() As Boolean
   Dim Row As Integer
   Dim ValLine As Boolean
   Dim Msg As String
   Dim DbName As String
#If DATACON = 1 Then
   Dim AuxDb As Database
#End If
   
   valida = False
   lHayPropIVA = False
   
   If gCtasBas.IdCtaIVACred <= 0 Or gCtasBas.IdCtaIVADeb <= 0 Then
      MsgBox1 "No es posible ingresar documentos sin antes definir la configuración las cuentas de IVA y Otros Impuestos.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   'vemos si las líneas están completas
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If LineaEnBlanco(Row) Then
         Exit For
      End If
            
      If Grid.RowHeight(Row) > 0 And Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then    'no ha sido borrada
            
            ValLine = IsValidLine(Row, Msg)
'         ValLine = (ValLine Or Grid.RowHeight(Row) = 0 Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO)    'línea borrada o doc anulado
         
         If ValLine = False Then
            If Msg <> "" Then
               MsgBox1 "Línea " & Row - 1 & " inválida. " & Msg, vbExclamation + vbOKOnly
            End If
            Exit Function
         End If
         
         'vemos si algún documento tiene proporcionalidad de IVA, para preguntar si queire recálculo de proporcionalidad una vez grabado el libro
         If lTipoLib = LIB_COMPRAS And lHayPropIVA = False Then
         
            If Val(Grid.TextMatrix(Row, C_IDPROPIVA)) <> 0 Then
               lHayPropIVA = True
            End If
            
         End If
      End If
      
   Next Row
   
   
   'abrimos base de datos año anterior para ver algún doc ya fue ingresado en año anterior

#If DATACON = 1 Then       'Access

   If gEmpresa.TieneAnoAnt Then

      If gEmprSeparadas Then

         DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
         If ExistFile(DbName) Then

            Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
            'hacemos CorrigeBase de año anterior por si las moscas
            Set AuxDb = DbMain
            Set DbMain = lDbAnoAnt
            'Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
            Call CorrigeBase
            Set DbMain = AuxDb
            'Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
         End If
      End If
   End If

#End If

   'vemos si no hay documentos repetidos
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If LineaEnBlanco(Row) Then
         Exit For
      End If
      
      'lo hacemos para todos, no solo los modificados, por si se nos pasó algo (FCA 29/03/2012)
      'If Grid.TextMatrix(Row, C_UPDATE) <> "" And Grid.RowHeight(Row) > 0 Then
      If Grid.RowHeight(Row) > 0 Then
         If Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VALEPAGOELECTR Then   'no validamos año anterior ya que no tiene sentido
            If Not ValidaNumDocAnoAnt(Row) Then
               Call FGrSelRow(Grid, Row)
               Exit Function
            End If
         End If
      End If
      
   Next Row

#If DATACON = 1 Then

  If Not lDbAnoAnt Is Nothing Then
    'Call AddLog("Paso por aca 12 : " & DbMain.Name & "  la otra base : " & lDbAnoAnt.Name)
      Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
      Call CloseDb(lDbAnoAnt)
      Set lDbAnoAnt = Nothing
   End If
     
#End If
   valida = True

End Function
Private Sub LoadGrid(Optional ByVal Row As Integer = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim EsRebaja As Boolean
   Dim ValOtros As Double
   Dim NotValidRut As Boolean
   Dim EditEnable As Boolean
   Dim TotOtrosImp As Double
   Dim IVAActFijo As Double
   Dim IVAIrrecuperable As Double
   
   TotOtrosImp = 0
   IVAActFijo = 0
   IVAIrrecuperable = 0
      
   Grid.FlxGrid.Redraw = False
   
   If CbItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
      If CbItemData(Cb_TipoDoc) > 0 Then
         Where = Where & " AND Documento.TipoDoc = " & CbItemData(Cb_TipoDoc)
      End If
      
      If CbItemData(Cb_Estado) > 0 Then
         Where = Where & " AND Documento.Estado = " & CbItemData(Cb_Estado)
      End If
      
      If Val(Tx_NumDoc) <> 0 Then
         Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
      End If
      
      If Val(Tx_NumDocAsoc) <> 0 Then
         Where = Where & " AND Documento.NumDocAsoc = '" & Trim(Tx_NumDocAsoc) & "'"
      End If
      
      If Trim(Tx_Rut) <> "" Then
         IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
         If IdEnt > 0 Then
            Where = Where & " AND Documento.IdEntidad = " & IdEnt
         Else
            Tx_Rut = ""
            Cb_Entidad.ListIndex = 0
            Cb_Nombre.ListIndex = 0
         End If
      
      End If
      
      If Trim(Tx_Descrip) <> "" Then
         Where = Where & " AND " & GenLike(DbMain, Tx_Descrip, "Documento.Descrip", 3)
      End If
      
      If vFmt(Tx_Valor) <> 0 Then
         Where = Where & " AND Documento.Afecto = " & vFmt(Tx_Valor)
      End If
      
   End If
   
   If Row > 0 Then
      Where = Where & " AND Documento.IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC))
   End If
   
   If CbItemData(Cb_Sucursal) > 0 Then
      Where = Where & " AND Documento.IdSucursal = " & CbItemData(Cb_Sucursal)
   End If
   
   If CbItemData(Cb_DTE) > 0 Then
      Where = Where & " AND Documento.DTE <> 0 "
   ElseIf CbItemData(Cb_DTE) < 0 Then
      Where = Where & " AND Documento.DTE = 0 "
   End If
      
   If Ch_EsSupermercado > 0 Then
      Where = Where & " AND Entidades.EsSupermercado <> 0 "
   End If
      

   Q1 = "SELECT IdDoc, Documento.TipoDoc, DTE, Documento.Giro, NumDoc, CorrInterno, NumDocHasta, Documento.IdEntidad, Documento.RutEntidad, Documento.MovEdited, Documento.PropIVA,"
   Q1 = Q1 & " Documento.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre, FEmision, FEmisionOri, FVenc, Exento, Afecto, IVA, Documento.IdCompCent, Documento.IdCompPago, "
   Q1 = Q1 & " OtroImp, OtrosVal, Total, Descrip, Documento.Estado, Documento.IdANegCCosto, IdCuentaExento, Usuarios.Usuario, Cuentas1.Codigo as CodCtaEx, "
   Q1 = Q1 & " Cuentas1.Descripcion as DescCtaEx, Cuentas1.Atrib" & ATRIB_ACTIVOFIJO & " as ActFijoCtaEx, IdCuentaAfecto, "
   Q1 = Q1 & " Cuentas2.Codigo as CodCtaAf, Cuentas2.Descripcion as DescCtaAf, Cuentas2.Atrib" & ATRIB_ACTIVOFIJO & " as ActFijoCtaAf, "
   Q1 = Q1 & " IdCuentaIVA, IdCuentaOtroImp, IdCuentaTotal, Cuentas3.Codigo as CodCtaTot,Cuentas3.Descripcion as DescCtaTot, "
   Q1 = Q1 & " Documento.IdSucursal, Sucursales.Descripcion as DescSucursal, EsSupermercado, Entidades.EntRelacionada, "
   Q1 = Q1 & " iif(" & SqlMonthLng("FEmisionOri") & "= " & CbItemData(Cb_Mes) & ",0,1) as MesActual, "
   Q1 = Q1 & " NumFiscImpr, NumInformeZ, CantBoletas, VentasAcumInfZ, IdDocAsoc, FExported "
   Q1 = Q1 & " FROM (((((((Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib=TipoDocs.TipoLib AND Documento.TipoDOC=TipoDocs.TipoDoc)"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON Documento.IdUsuario = Usuarios.IdUsuario )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas1") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaAfecto = Cuentas2.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas2") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas3 ON Documento.IdCuentaTotal = Cuentas3.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas3") & " )"
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Sucursales", True, True) & " )"
   Q1 = Q1 & " WHERE Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & Where
   Q1 = Q1 & lWhere
   
   If lOrdenSel = C_FECHAEMIORI Then
      Q1 = Q1 & " ORDER BY IIf(" & SqlMonthLng("FEmisionOri") & "= " & CbItemData(Cb_Mes) & ", 0, 1), " & lOrdenGr(lOrdenSel)
   Else
      Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   End If
   
   If Row = 0 Then
      'Q1 = Q1 & SqlPaging(gDbType, lClsPaging.CurReg - 1, gPageNumReg)
   End If
      
   lCurWhere = " Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   lCurWhere = lCurWhere & Where & lWhere
   
   Call DetOtrosImp(lTipoLib, lCurWhere)
   
   Set Rs = OpenRs(DbMain, Q1)

   If Row <= 0 Then
      Grid.rows = Grid.FixedRows
      i = Grid.FixedRows
   Else
      i = Row
   End If
   
   Do While Rs.EOF = False
      
      If Row <= 0 Then
         Grid.rows = Grid.rows + 1
      End If
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If

'      If vFld(Rs("IdDoc")) = 8450 Then
'      MsgBox ""
'      End If
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1 + (lClsPaging.CurReg - 1)
      If Row = 0 Then
         lClsPaging.NumReg = vFmt(Grid.TextMatrix(i, C_NUMLIN)) - (lClsPaging.CurReg - 1)
      End If
      
      If lOrdenSel = C_NUMLIN And vFld(Rs("CorrInterno")) = 0 Then
         Grid.TextMatrix(i, C_CORRINTERNO) = Grid.TextMatrix(i, C_NUMLIN)
         Call FGrModRow(Grid, i, FGR_U, C_IDDOC, C_UPDATE)
      Else
         Grid.TextMatrix(i, C_CORRINTERNO) = vFld(Rs("CorrInterno"))
      End If
      
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(lTipoLib, vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_IDTIPODOC) = vFld(Rs("TipoDoc"))
      Grid.TextMatrix(i, C_DOCIMPEXP) = CInt(gTipoDoc(GetTipoDoc(lTipoLib, vFld(Rs("TipoDoc")))).DocImpExp)
      If vFld(Rs("Giro")) <> 0 Then
         Grid.TextMatrix(i, C_GIRO) = ""
      Else
         If Grid.TextMatrix(i, C_TIPODOC) = "FAV" Or Grid.TextMatrix(i, C_TIPODOC) = "FVE" Or Grid.TextMatrix(i, C_TIPODOC) = "NCV" Or Grid.TextMatrix(i, C_TIPODOC) = "NDV" Then
            Grid.TextMatrix(i, C_GIRO) = "No"
         Else
            Call FGrModRow(Grid, i, FGR_U, C_IDDOC, C_UPDATE)   'esto no debiera ocurrir nunca, pero por si quedó con el valor 0, que cambie a -1 y lo grabe
         End If
      End If
      If vFld(Rs("DTE")) <> 0 Then
         Grid.TextMatrix(i, C_DTE) = "x"
      Else
         Grid.TextMatrix(i, C_DTE) = ""
      End If
      
      If Grid.TextMatrix(i, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
         Grid.TextMatrix(i, C_NUMFISCIMPR) = vFld(Rs("NumFiscImpr"))
         Grid.TextMatrix(i, C_NUMINFORMEZ) = vFld(Rs("NumInformeZ"))
      End If
      
      If Grid.TextMatrix(i, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
         Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
         Grid.TextMatrix(i, C_NUMDOCHASTA) = IIf(Val(vFld(Rs("NumDocHasta"))) <> 0, vFld(Rs("NumDocHasta")), "")
      End If
      
      If lTipoLib = LIB_COMPRAS And gTipoDoc(GetTipoDoc(lTipoLib, vFld(Rs("TipoDoc")))).AceptaPropIVA Then

         Grid.TextMatrix(i, C_IDPROPIVA) = vFld(Rs("PropIVA"))
         Grid.TextMatrix(i, C_PROPIVA) = Left(gStrPropIVA(Grid.TextMatrix(i, C_IDPROPIVA)), 1)
         
      End If
      
      
      Grid.TextMatrix(i, C_CANTBOLETAS) = IIf(vFld(Rs("CantBoletas")) > 0, Format(vFld(Rs("CantBoletas")), NUMFMT), "")
      
      Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      
      If vFld(Rs("Estado")) <> ED_ANULADO Then
      
         If vFld(Rs("IdEntidad")) <= 0 Then
            If vFld(Rs("RutEntidad")) <> "" And vFld(Rs("RutEntidad")) <> "0" Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RutEntidad")))
               Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("NombreEntidad"), True)
            End If
         Else
            If vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0" Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
               Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
            End If
         End If
      
      Else
         Grid.TextMatrix(i, C_RUT) = ""
         Grid.TextMatrix(i, C_NOMBRE) = "NULO"
         
      End If
      
      If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
         Grid.TextMatrix(i, C_FECHA) = Day(vFld(Rs("FEmision")))
      Else
         Grid.TextMatrix(i, C_FECHA) = Format(vFld(Rs("FEmision")), SDATEFMT)
      End If
      
      If vFld(Rs("FEmisionOri")) > 0 Then
         Grid.TextMatrix(i, C_FECHAEMIORI) = Format(vFld(Rs("FEmisionOri")), SDATEFMT)
         Grid.TextMatrix(i, C_LNGFECHAEMIORI) = vFld(Rs("FEmisionOri"))
      End If
      
      'parche para docs ya ingresados, que tienen fecha emisión cero (ahora es obligatoria) y no están en estado pendiente (los de estado pendiente el usuario puede ingresarle la fecha)
      'le asignamos FEmision (que corresponde al día de recepción)
      If Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI)) = 0 And Val(Grid.TextMatrix(i, C_IDESTADO)) <> ED_PENDIENTE Then
         Grid.TextMatrix(i, C_FECHAEMIORI) = Format(vFld(Rs("FEmision")), SDATEFMT)
         Grid.TextMatrix(i, C_LNGFECHAEMIORI) = vFld(Rs("FEmision"))
         Call FGrModRow(Grid, i, FGR_U, C_IDDOC, C_UPDATE)
      End If
      
      If vFld(Rs("FVenc")) > 0 Then
         Grid.TextMatrix(i, C_FECHAVENC) = Format(vFld(Rs("FVenc")), SDATEFMT)
         Grid.TextMatrix(i, C_LNGFECHAVENC) = vFld(Rs("FVenc"))
      End If
      
      'parche para docs ya ingresados, que tienen fecha vencimiento cero (ahora es obligatoria) y no están en estado pendiente (los de estado pendiente el usuario puede ingresarle la fecha)
      'le asignamos FVenc (30 días después de la fecha de emisión ori)
'      If Val(Grid.TextMatrix(i, C_LNGFECHAVENC)) = 0 And vFld(Rs("Estado")) <> ED_PENDIENTE Then
      If Val(Grid.TextMatrix(i, C_LNGFECHAVENC)) = 0 Then
'         If vFld(Rs("EntRelacionada")) <> 0 Then 'pago contado
         If gEmpresa.Franq14Ter Then 'pago contado
            Grid.TextMatrix(i, C_LNGFECHAVENC) = CLng(vFld(Rs("FEmision")))
         Else                                      'crédito 30 dias propuesto
            Grid.TextMatrix(i, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, vFld(Rs("FEmision"))))
         End If
         Grid.TextMatrix(i, C_FECHAVENC) = Format(Grid.TextMatrix(i, C_LNGFECHAVENC), SDATEFMT)
         Call FGrModRow(Grid, i, FGR_U, C_IDDOC, C_UPDATE)
      End If
      
      EsRebaja = gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(i, C_IDTIPODOC)))).EsRebaja              'FCA 30 nov 2017
      
      ValOtros = vFld(Rs("Total")) - (vFld(Rs("Exento")) + vFld(Rs("Afecto")) + vFld(Rs("IVA")))
            
'      If vFld(Rs("IVA")) = "1140000" Or vFld(Rs("IVA")) = "152273" Then
'      Dim a As String
'      a = ValOtros
'
'      End If
'
      TotOtrosImp = 0
      If Ch_ViewDetOtrosImp <> 0 Then
         TotOtrosImp = FillDetOtroImp(i, lTipoLib, vFld(Rs("IdDoc")), EsRebaja, IVAActFijo, IVAIrrecuperable)              'esta función, a diferencia de la que se usa en Resumen Supermercado, retorna el total de otros impuestos que no están clasificados como impuestos adicionales, ya que estos se desglozan en la misma grilla
         If EsRebaja And TotOtrosImp <> 0 Then
            ValOtros = ValOtros * -1
         End If
'         ValOtros = ValOtros - Abs(TotOtrosImp)
         ValOtros = ValOtros - TotOtrosImp      'FCA 30 nov 2017
         lDetOtrosImpFilled = True
      Else
         lDetOtrosImpFilled = False
      End If
      
            
      If EsRebaja Then                             'nota de crédito = valores negativos
         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")) * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")) * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - Abs(IVAActFijo) - Abs(IVAIrrecuperable)) * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_OTROIMP) = Format(ValOtros * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")) * -1, NEGNUMFMT)
      
      Else
         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")), NEGNUMFMT)
         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")), NEGNUMFMT)
         Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - IVAActFijo - IVAIrrecuperable), NEGNUMFMT)
         Grid.TextMatrix(i, C_OTROIMP) = Format(ValOtros, NEGNUMFMT)
         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NEGNUMFMT)
         
      End If
      

               
      If Grid.TextMatrix(i, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
         Grid.TextMatrix(i, C_VENTASACUM) = Format(vFld(Rs("VentasAcumInfZ")), NUMFMT)
      End If
         
      Grid.TextMatrix(i, C_EX_IDCUENTA) = vFld(Rs("IdCuentaExento"))
'      If vFld(Rs("MovEdited")) = 0 Then
         Grid.TextMatrix(i, C_EX_CODCUENTA) = Format(vFld(Rs("CodCtaEx")), gFmtCodigoCta)
         Grid.TextMatrix(i, C_EX_CUENTA) = vFld(Rs("DescCtaEx"), True)
'      Else
'         Grid.TextMatrix(i, C_EX_CODCUENTA) = ""
'         Grid.TextMatrix(i, C_EX_CUENTA) = ""
'      End If
      
      Grid.TextMatrix(i, C_AF_IDCUENTA) = vFld(Rs("IdCuentaAfecto"))
'      If vFld(Rs("MovEdited")) = 0 Then
         Grid.TextMatrix(i, C_AF_CODCUENTA) = Format(vFld(Rs("CodCtaAf")), gFmtCodigoCta)
         Grid.TextMatrix(i, C_AF_CUENTA) = vFld(Rs("DescCtaAf"), True)
'      Else
'         Grid.TextMatrix(i, C_AF_CODCUENTA) = ""
'         Grid.TextMatrix(i, C_AF_CUENTA) = ""
'      End If
      
      Grid.TextMatrix(i, C_IVA_IDCUENTA) = vFld(Rs("IdCuentaIVA"))
      Grid.TextMatrix(i, C_OIMP_IDCUENTA) = vFld(Rs("IdCuentaOtroImp"))
      
      If vFmt(Grid.TextMatrix(i, C_IVA)) <> 0 And vFmt(Grid.TextMatrix(i, C_IVA_IDCUENTA)) = 0 Then
         Grid.TextMatrix(i, C_IVA_IDCUENTA) = lIdCuentaIVA
      End If
      
      If vFmt(Grid.TextMatrix(i, C_OTROIMP)) <> 0 And vFmt(Grid.TextMatrix(i, C_OIMP_IDCUENTA)) = 0 Then
         'si es factura de compras, nota de crédito de fac. compras o nota de débito de fac. compras, se pone la cuenta al revés
         If Grid.TextMatrix(i, C_TIPODOC) = "FCC" Or Grid.TextMatrix(i, C_TIPODOC) = "NCF" Or Grid.TextMatrix(i, C_TIPODOC) = "NDF" Or Grid.TextMatrix(i, C_TIPODOC) = "FCV" Then
            Grid.TextMatrix(i, C_OIMP_IDCUENTA) = lIdCuentaOtrosImpFacCompra
         Else
            Grid.TextMatrix(i, C_OIMP_IDCUENTA) = lIdCuentaOtrosImp
         End If
      End If
      
      Grid.TextMatrix(i, C_TOT_IDCUENTA) = vFld(Rs("IdCuentaTotal"))
'      If vFld(Rs("MovEdited")) = 0 Then
         Grid.TextMatrix(i, C_TOT_CODCUENTA) = Format(vFld(Rs("CodCtaTot")), gFmtCodigoCta)
         Grid.TextMatrix(i, C_TOT_CUENTA) = vFld(Rs("DescCtaTot"), True)
'      Else
'         Grid.TextMatrix(i, C_TOT_CODCUENTA) = ""
'         Grid.TextMatrix(i, C_TOT_CUENTA) = ""
'      End If
      
      Grid.TextMatrix(i, C_DETALLE) = TX_DETALLE
      
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(i, C_IDANEG_CCOSTO) = vFld(Rs("IdANegCCosto"), True)
      Grid.TextMatrix(i, C_IDSUCURSAL) = vFld(Rs("IdSucursal"), True)
      Grid.TextMatrix(i, C_SUCURSAL) = vFld(Rs("DescSucursal"), True)
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(i, C_ESTADO) = gEstadoDoc(vFld(Rs("Estado")))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.TextMatrix(i, C_MOVEDITED) = IIf(vFld(Rs("MovEdited")) <> 0, -1, 0)
      Grid.TextMatrix(i, C_EXPORTED) = IIf(vFld(Rs("FExported")) <> 0, -1, 0)
      
      If vFld(Rs("IdDocAsoc")) <> 0 Then
         Grid.TextMatrix(i, C_DOCASOC) = GetInfoDoc(vFld(Rs("IdDocAsoc")))
      End If
      
      Grid.TextMatrix(i, C_IDCOMPCENT) = vFld(Rs("IdCompCent"))
      Grid.TextMatrix(i, C_IDCOMPPAGO) = vFld(Rs("IdCompPago"))
      
      If vFld(Rs("MovEdited")) <> 0 Then   'tiene que estar antes del C_CHECK por el color de la celda
         Call FGrSetRowStyle(Grid, i, "FC", COLOR_AZULOSCURO)
      End If
      
      If Grid.ColWidth(C_CHECK) > 0 Then
         If vFld(Rs("Estado")) = ED_CENTRALIZADO Then
            Grid.TextMatrix(i, C_CHECK) = "C"
         ElseIf vFld(Rs("Estado")) = ED_PAGADO Then
            Grid.TextMatrix(i, C_CHECK) = "P"
              
            'ponemos el texto "P" en gris porque perimtimos centralizar docuemntos pagados que no han sido centralizados
            If vFmt(Grid.TextMatrix(i, C_IDCOMPCENT)) = 0 Then
               Call FGrSetRowStyle(Grid, i, "FC", COLOR_GRIS, C_CHECK, C_CHECK)
            End If

         ElseIf vFld(Rs("Estado")) = ED_ANULADO Then
            Grid.TextMatrix(i, C_CHECK) = "A"
         End If
      End If
            
      'If lOper = O_VIEW Then
      '   Call FGrSetPicture(Grid, i, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)
      'End If
       
      If vFld(Rs("ActFijoCtaEx")) <> 0 Or vFld(Rs("ActFijoCtaAf")) <> 0 Then
         Grid.TextMatrix(i, C_DETACTFIJO) = TX_ACTFIJO
      End If
               
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   If Row <= 0 Then
      Call FGrVRows(Grid)
      Grid.rows = Grid.rows + 1
      Grid.TopRow = Grid.FixedRows
      
      'Marco la columna Ordenada
      Grid.Row = Grid.FixedRows - 1
      Grid.Col = lOrdenSel
      Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Else
      Grid.Row = Row
      Grid.Col = C_TIPODOC
   
   End If
   
      
   Tx_CurrCell = ""
      
   Call CalcTot
   
   Grid.FlxGrid.Redraw = True
   
   Call UnLockAction(DbMain, lTipoLib, , , , False)

   EditEnable = LockAction(DbMain, lTipoLib, CbItemData(Cb_Mes))

   If EditEnable = False Then    'alguien más lo está editando, no podemos editarlo (esto se hace sólo una vez ya que en lOper = O_EDIT no se puede cambiar el mes)
      MsgBox1 "El " & gTipoLib(lTipoLib) & " del mes de " & gNomMes(CbItemData(Cb_Mes)) & " se está editando en el equipo '" & IsLockedAction(DbMain, lTipoLib, CbItemData(Cb_Mes)) & "'. Sólo se abrirá de lectura.", vbInformation
      lEditEnabled = False
      lAdmDocsEnabled = False
      lIngDocsEnabled = False
   Else
      lIngDocsEnabled = ChkPriv(PRV_ING_DOCS)
      lAdmDocsEnabled = ChkPriv(PRV_ADM_DOCS)
   End If
   
   If lOper = O_EDIT Then
      Call EnableForm(Me, lEditEnabled)
      Call SetTxRO(Tx_CurrCell, True)
      Bt_DetDoc.Enabled = True   'Bt_DetDoc lo abrirá sólo de lectura
      Ch_ViewDetOtrosImp.Enabled = False
            
      If lTipoLib = LIB_COMPRAS And gFunciones.ProporcionalidadIVA Then
         Ch_ViewDocHasta.Enabled = False
         Ch_RepetirGlosa.Enabled = False
         Ch_ViewMaqReg.Enabled = False
         Ch_ViewCantBoletas.Enabled = False
      Else
         Ch_ViewPropIVA.Enabled = False
      
      End If

   End If
   
   Bt_Centralizar.Enabled = lAdmDocsEnabled
   Bt_ExitNewDoc.Enabled = lIngDocsEnabled

   Bt_List.Enabled = False
   Bt_Sel.Enabled = True
   
   'botones de paginamiento
   Call lClsPaging.ActivateButtons(Grid.TextMatrix(Grid.FixedRows, C_NUMLIN) = "")
   
   
End Sub
Private Sub SaveGrid()
   Dim i As Integer
   Dim Lin As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim j As Integer
   Dim EsRebaja As Boolean
   Dim NumDocVSD As String
   Dim Descrip As String
   Dim FldArray(4) As AdvTbAddNew_t

   lIdDoc = 0
   'Call OpenDbEmp
   Lin = Grid.FixedRows
   For i = Grid.FixedRows To Grid.rows - 1
   
   '626924 opcion solo para cliente que no existe detalle de documentos
'   Grid.TextMatrix(i, C_UPDATE) = FGR_U
'   Grid.TextMatrix(i, C_MOVEDITED) = 0
   '626924
      If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If
      
     If gAppCode.Demo Then
         If i - Grid.FixedRows >= MAX_DOCDEMO And Not W.InDesign Then
            MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
            Exit For
         End If
      End If
            
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert     OJO REVISAR CON PABLO
'         Set Rs = DbMain.OpenRecordset("Documento")
'         Rs.AddNew
'
'         IdDoc = vFld(Rs("IdDoc"))
'         Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'         Rs.Fields("FechaCreacion") = CLng(Int(Now))
'         Rs.Fields("FEmision") = CLng(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), Val(Grid.TextMatrix(i, C_FECHA))))    ' 14 jul 2017
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            Rs.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            Rs.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'            Err.Clear
'         Next j
'
'         Rs.Close
'
'         Set Rs = Nothing
         
         

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
         
         FldArray(4).FldName = "FEmision"
         FldArray(4).FldValue = CLng(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), Val(Grid.TextMatrix(i, C_FECHA))))
         FldArray(4).FldIsNum = True
         
      
         IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
         
         'Tracking 3217833
         
         Grid.TextMatrix(i, C_IDDOC) = IdDoc
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
         
         If lIdDoc = 0 Then   'selecciona el primero que se insertó
            lIdDoc = IdDoc
         End If
         
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM Documento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Call DeleteSQL(DbMain, "Documento", " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC)))

'         Q1 = "DELETE FROM MovDocumento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Call DeleteSQL(DbMain, "MovDocumento", " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC)))
         
         '3133008
       
'        Dim PathDbAnoAnt As String
'        Dim ConnStr As String
'
'        #If DATACON = 1 Then
'        Dim DbAnoAnt As Database
'        #Else
'        Dim DbAnoAnt As ADODB.Connection
'        Set DbAnoAnt = DbMain
'        #End If
'
'   If gDbType = SQL_ACCESS Then
'        PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")
'
'        If ExistFile(PathDbAnoAnt) Then
'          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
'          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
'
'        Else
'          'Exit Sub
'        End If
'    End If
'
'             Q1 = ""
'             Q1 = "Update Documento Set FExported = null WHERE NumDoc = '" & Val(Grid.TextMatrix(i, C_NUMDOC)) & "'"
'             Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
'             Q1 = Q1 & " And TipoLib = " & lTipoLib
'             Q1 = Q1 & " And TipoDoc = " & FindTipoDoc(lTipoLib, Grid.TextMatrix(i, C_TIPODOC))
'             Q1 = Q1 & " And FEmisionOri = " & Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI))
'             Q1 = Q1 & " And Total = " & Abs(vFmt(Grid.TextMatrix(i, C_TOTAL)))
'
'             Call ExecSQL(DbAnoAnt, Q1)
'
'             If gDbType = SQL_ACCESS Then
'             Call CloseDb(DbAnoAnt)
'             End If

        '3133008

'         Q1 = "DELETE FROM LibroCaja WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Call DeleteSQL(DbMain, "LibroCaja", " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC)))

      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then 'Update
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & lTipoLib
         Q1 = Q1 & ", TipoDoc = " & Val(Grid.TextMatrix(i, C_IDTIPODOC))
         Q1 = Q1 & ", CorrInterno = " & vFmt(Grid.TextMatrix(i, C_CORRINTERNO))
         
          'pipe2807009
            If gDbType = SQL_ACCESS Then
            Q1 = Q1 & ", DTE = " & IIf(Trim(Grid.TextMatrix(i, C_DTE)) <> "", -1, 0)
            Else
             Q1 = Q1 & ", DTE = " & IIf(Trim(Grid.TextMatrix(i, C_DTE)) <> "", 1, 0)
            End If
        'fin 2807009
      
         Q1 = Q1 & ", Giro = " & IIf(Trim(Grid.TextMatrix(i, C_GIRO)) <> "", 0, -1)
         
         Q1 = Q1 & ", NumFiscImpr = '" & RemoveNoPrtChars(Left(Grid.TextMatrix(i, C_NUMFISCIMPR), 20)) & "'"
         Q1 = Q1 & ", NumInformeZ = '" & RemoveNoPrtChars(Left(Grid.TextMatrix(i, C_NUMINFORMEZ), 20)) & "'"
        
         If Grid.TextMatrix(i, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
            Q1 = Q1 & ", NumDoc = '" & RemoveNoPrtChars(Left(Grid.TextMatrix(i, C_NUMDOC), 20)) & "'"
            Q1 = Q1 & ", NumDocHasta = '" & RemoveNoPrtChars(Left(Grid.TextMatrix(i, C_NUMDOCHASTA), 20)) & "'"
         Else
            NumDocVSD = GetNumDocVSD(lTipoLib, Val(Grid.TextMatrix(i, C_IDTIPODOC)))
            Q1 = Q1 & ", NumDoc = '" & NumDocVSD & "'"
            Q1 = Q1 & ", NumDocHasta = '0'"
         End If
         
         Q1 = Q1 & ", CantBoletas = " & vFmt(Grid.TextMatrix(i, C_CANTBOLETAS))
            
         Q1 = Q1 & ", IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
         Q1 = Q1 & ", EntRelacionada = " & Abs(GetEntRelacionada(Val(Grid.TextMatrix(i, C_IDENTIDAD))))
         Q1 = Q1 & ", IdSucursal = " & vFmt(Grid.TextMatrix(i, C_IDSUCURSAL))

         If lTipoLib = LIB_COMPRAS Then
            Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR
            'por si acaso, ponemos la clasificación de la entidad
            Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_PROVEEDOR & " = 1 WHERE IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD)))
         Else
            Q1 = Q1 & ", TipoEntidad = " & ENT_CLIENTE
            'por si acaso, ponemos la clasificación de la entidad
            Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_CLIENTE & " = 1 WHERE IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD)))
         End If
         
         If Val(Grid.TextMatrix(i, C_IDENTIDAD)) = 0 Then
            Q1 = Q1 & ", RutEntidad ='" & vFmtCID(Grid.TextMatrix(i, C_RUT)) & "'"
            Q1 = Q1 & ", NombreEntidad = '" & ParaSQL(Left(Grid.TextMatrix(i, C_NOMBRE), 50)) & "'"
         End If
         
         If lTipoLib = LIB_COMPRAS Then
            Q1 = Q1 & ", PropIVA = " & Val(Grid.TextMatrix(i, C_IDPROPIVA))
         End If

         Q1 = Q1 & ", FEmision = " & CLng(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), Val(Grid.TextMatrix(i, C_FECHA))))
         Q1 = Q1 & ", FEmisionOri = " & Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI))
         Q1 = Q1 & ", FVenc = " & Val(Grid.TextMatrix(i, C_LNGFECHAVENC))
         Q1 = Q1 & ", Exento = " & Abs(vFmt(Grid.TextMatrix(i, C_EXENTO)))
         
         If vFmt(Grid.TextMatrix(i, C_EXENTO)) <> 0 Then
            Q1 = Q1 & ", IdCuentaExento = " & vFmt(Grid.TextMatrix(i, C_EX_IDCUENTA))
         Else
            Q1 = Q1 & ", IdCuentaExento = 0"
         End If
         
         Q1 = Q1 & ", Afecto = " & Abs(vFmt(Grid.TextMatrix(i, C_AFECTO)))
         If vFmt(Grid.TextMatrix(i, C_AFECTO)) <> 0 Then
            Q1 = Q1 & ", IdCuentaAfecto = " & vFmt(Grid.TextMatrix(i, C_AF_IDCUENTA))
         Else
            Q1 = Q1 & ", IdCuentaAfecto = 0 "
         End If

         EsRebaja = gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(i, C_IDTIPODOC)))).EsRebaja
'         Descrip = ParaSQL(Left(RemoveSpcChars(Grid.TextMatrix(i, C_DESCRIP)), 100))
         Descrip = RemoveNoPrtChars(Grid.TextMatrix(i, C_DESCRIP))
         Q1 = Q1 & ", IVA = " & Abs(vFmt(Grid.TextMatrix(i, C_IVA)))
         Q1 = Q1 & ", IdCuentaIVA = " & vFmt(Grid.TextMatrix(i, C_IVA_IDCUENTA))
         If EsRebaja Then
            Q1 = Q1 & ", OtroImp = " & vFmt(Grid.TextMatrix(i, C_OTROIMP)) * -1
         Else
            Q1 = Q1 & ", OtroImp = " & vFmt(Grid.TextMatrix(i, C_OTROIMP))
         End If
         Q1 = Q1 & ", IdCuentaOtroImp = " & vFmt(Grid.TextMatrix(i, C_OIMP_IDCUENTA))
         Q1 = Q1 & ", Total = " & Abs(vFmt(Grid.TextMatrix(i, C_TOTAL)))
         Q1 = Q1 & ", VentasAcumInfZ = " & vFmt(Grid.TextMatrix(i, C_VENTASACUM))
         Q1 = Q1 & ", IdCuentaTotal = " & vFmt(Grid.TextMatrix(i, C_TOT_IDCUENTA))
         Q1 = Q1 & ", Descrip = '" & ParaSQL(Left(Descrip, 100)) & "'"
         Q1 = Q1 & ", IdANegCCosto = '" & ParaSQL(Grid.TextMatrix(i, C_IDANEG_CCOSTO)) & "'"
         Q1 = Q1 & ", Estado = " & Val(Grid.TextMatrix(i, C_IDESTADO))
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

         Call ExecSQL(DbMain, Q1)

         If Val(Grid.TextMatrix(i, C_MOVEDITED)) = 0 Then
         
'            Q1 = "DELETE FROM MovDocumento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'            Call ExecSQL(DbMain, Q1)

            Call DeleteSQL(DbMain, "MovDocumento", " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC)))

            Call GenMovDocumento(i)

            
         ElseIf Val(Grid.TextMatrix(i, C_IDESTADO)) = ED_ANULADO Then
         
'            Q1 = "DELETE FROM MovDocumento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'            Call ExecSQL(DbMain, Q1)

            Call DeleteSQL(DbMain, "MovDocumento", " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC)))

         End If
         
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
         Lin = Lin + 1
      End If

      Grid.TextMatrix(i, C_UPDATE) = ""     'lo limpiampos dado que esta función es invocada en Bt_DetDoc
      
   Next i

   If gFunciones.ProporcionalidadIVA Then
      If lTipoLib = LIB_VENTAS Then
         Call PropIVA_UpdateTblTotMensual(CbItemData(Cb_Mes), True)
      End If
   End If
   
End Sub

Private Function NewEntidad(ByVal Row As Integer, ByVal Rut As String) As Integer
   Dim Frm As FrmEntidad
   Dim Entidad As Entidad_t
   Dim i As Integer
   Dim Rc As Integer
 
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   Entidad.Clasif = 0    'ItemData(Cb_Entidad)
   Entidad.Rut = Rut
   
   Rc = Frm.FNew(Entidad, Rut)
   
   If Rc <> vbCancel Then
   
      'If Entidad.Clasif = ItemData(Cb_Entidad) Then
         
         Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
         Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
         Grid.TextMatrix(Row, C_RUT) = Entidad.Rut
         
'         If GetEntRelacionada(Entidad.id) And gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then    'si es Ent Relacionada and 14TER => pago contado
         If gEmpresa.Franq14Ter And lTipoLib = LIB_VENTAS Then                                       'si es 14TER => se propone pago contado  (Claudio Villegas - 11 Jul 2017)
            Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))))
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
         End If
         
      'Else
      '   MsgBox1 "La clasificación de la nueva entidad no coincide con la que está seleccionada. Vuelva a seleccionar el tipo de entidad para que la muestre en la lista.", vbOKOnly + vbInformation
      'End If
      
   Else
      Grid.TextMatrix(Row, C_NOMBRE) = ""
      Grid.TextMatrix(Row, C_IDENTIDAD) = 0
      
   End If
   
   Set Frm = Nothing
   MousePointer = vbDefault
   
   NewEntidad = Rc
End Function

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call bt_Copy_Click
   ElseIf KeyPaste(KeyCode, Shift) Then
      Call Bt_Paste_Click
   End If

End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   Grid.Row = Grid.MouseRow
   Grid.Col = Grid.MouseCol
   
   If lOper <> O_EDIT And lOper <> O_NEW Then
      Exit Sub
   End If

   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Or Not ValidaEstadoEdit(Grid.Row) Then
      Exit Sub
   End If
   
'   If Not ValidaFExported(Grid.Row) Then
'      Exit Sub
'   End If
   Call ValidaFExported(Grid.Row)
   
   If Button = vbRightButton Then
      If (Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC) Then
         Call PopupMenu(M_TipoDoc)
      ElseIf Grid.Col = C_EX_CODCUENTA Or Grid.Col = C_AF_CODCUENTA Or Grid.Col = C_TOT_CODCUENTA Then
         Call Bt_Cuentas_Click
      End If
   End If
   
End Sub

Private Sub Grid_SelChange()
   Dim EdType As FlexEdGrid2.FEG2_EdType
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Col = C_FECHA And Grid.TextMatrix(Grid.Row, Grid.Col) = "" And (Grid.Row = Grid.FixedRows Or Grid.TextMatrix(Grid.Row - 1, Grid.Col) <> "") Then
      Call Grid_BeforeEdit(Grid.Row, Grid.Col, EdType)
   ElseIf Grid.Col = C_PROPIVA Then
      Tx_CurrCell = gDescPropIVA(Val(Grid.TextMatrix(Grid.Row, C_IDPROPIVA)))
   Else
      Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)
   End If

End Sub

Private Sub M_ItCuenta_Click(Index As Integer)
   Dim Row As Integer
   Dim Col As Integer
   Dim ColId As Integer
   Dim ColCta As Integer
   Dim OldIdCuenta As Long
   Dim CodCta As String
   Dim DescCta As String
   Dim NombCta As String
   Dim IdCuenta As Long
   Dim Frm As FrmPlanCuentas

   Row = Grid.Row
   Col = Grid.Col
   
   Select Case Col
      Case C_EX_CODCUENTA
         ColId = C_EX_IDCUENTA
         ColCta = C_EX_CUENTA
      Case C_AF_CODCUENTA
         ColId = C_AF_IDCUENTA
         ColCta = C_AF_CUENTA
      Case C_TOT_CODCUENTA
         ColId = C_TOT_IDCUENTA
         ColCta = C_TOT_CUENTA
      Case Else
         Exit Sub
   End Select
  
   If M_ItCuenta(Index).Caption <> MITEM_OTRA Then
  
      Grid.TextMatrix(Row, ColId) = lCuentas(Index) 'se asigna porque lo necesita la función de activo fijo
      
      Call GridActivoFijo(lCuentas(Index), Row, Col)
      
      Grid.TextMatrix(Row, Col) = Left(M_ItCuenta(Index).Caption, InStr(M_ItCuenta(Index).Caption, " [") - 1)
      Grid.TextMatrix(Row, ColCta) = Mid(M_ItCuenta(Index).Caption, InStr(M_ItCuenta(Index).Caption, "] ") + 2)
   
   Else
      Set Frm = New FrmPlanCuentas
      If FrmPlanCuentas.FSelEdit(IdCuenta, CodCta, DescCta, NombCta, False) = vbOK Then
         
         Grid.TextMatrix(Row, ColId) = IdCuenta 'se asigna porque lo necesita la función de activo fijo
         
         Call GridActivoFijo(IdCuenta, Row, Col)
         
         Grid.TextMatrix(Row, Col) = FmtCodCuenta(CodCta)
         Grid.TextMatrix(Row, ColCta) = DescCta
         
         Set Frm = Nothing
      Else
         Set Frm = Nothing
         Exit Sub
      End If
   
   End If
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         
End Sub

Private Sub M_ItTipoDoc_Click(Index As Integer)
   Dim Value As String
   Dim TipoDoc As Integer
   Dim Row As Integer
   '2763862
   Dim FechaLey21 As Long
   Dim FechaForm As Long
   FechaLey21 = DateSerial(2022, 3, 24)
   Dim Dia As Integer
   'fin 2763862
   
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Then
      Exit Sub
   End If

   FechaForm = DateSerial(lAno, lMes, LDia)
   Row = Grid.Row
   
   TipoDoc = Cb_TipoDoc.ItemData(Index)
   
   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      Value = GetDiminutivoDoc(lTipoLib, TipoDoc)
      Grid.TextMatrix(Row, C_TIPODOC) = Value
      Grid.TextMatrix(Row, C_IDTIPODOC) = TipoDoc
      Grid.TextMatrix(Row, C_DOCIMPEXP) = CInt(gTipoDoc(GetTipoDoc(lTipoLib, TipoDoc)).DocImpExp)
      
      Call ActCambioTipoDoc(Row)
               
      Call CalcTot
     
      If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VENTASINDOC Then
         Grid.TextMatrix(Row, C_NUMDOC) = ""
      
      ElseIf Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Then
         If Not lMsgNotaCred Then
            'MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
            If FechaForm >= FechaLey21 Then
               MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 6 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
            Else
               MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
            End If
            lMsgNotaCred = True
         End If
      End If
      
      Grid.TextMatrix(Row, C_GIRO) = ""

      If gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).AceptaPropIVA Then
         If Grid.TextMatrix(Row, C_PROPIVA) = "" Then
            Grid.TextMatrix(Row, C_IDPROPIVA) = PIVA_TOTAL
            Grid.TextMatrix(Row, C_PROPIVA) = Left(gStrPropIVA(Grid.TextMatrix(Row, C_IDPROPIVA)), 1)
         End If
      Else
         Grid.TextMatrix(Row, C_IDPROPIVA) = 0
         Grid.TextMatrix(Row, C_PROPIVA) = ""
      End If

      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
   End If

End Sub

Private Function IsValidLine(ByVal Row As Integer, Msg As String) As Boolean
   Dim ValLine As Boolean
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim FRecep As Long
   Dim TxtImp As String
   
   IsValidLine = False
   
   ValLine = True
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_FECHA)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la fecha de recepción o ingreso al libro."
      Exit Function
   End If
   
   ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> "")
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el Tipo de Documento."
      Exit Function
   End If
   
   If Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> TDOC_VENTASINDOC And Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> TDOC_VALEPAGOELECTR Then    'venta sin documento o Vale Pago Electronico
      If Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> "IMP" Then
         ValLine = (ValLine And (Val(Grid.TextMatrix(Row, C_NUMDOC)) <> 0 Or Val(Grid.TextMatrix(Row, C_MOVEDITED)) <> 0))   'valor distinto de cero     OJO SOLO PARA PRUEBAS DE IMPORTAR FACTURACION
         If ValLine = False And Msg = "" Then
            Msg = "Falta ingresar el Número de Documento."
            Exit Function
         End If
      Else
         ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_NUMDOC)) <> "")   'string distinto de cero (IMP puede tener letras y el Val puede dar cero)
         If ValLine = False And Msg = "" Then
            Msg = "Falta ingresar el Número de Documento."
            Exit Function
         End If
      End If
   End If
   
   If Trim(Grid.TextMatrix(Row, C_TIPODOC)) = TDOC_VALEPAGOELECTR Then    'venta Vale Pago Electronico
      ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_CANTBOLETAS)) <> 0)
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la Cantidad de Boletas."
         Exit Function
      End If
   End If
   
   If IsNumeric(Trim(Grid.TextMatrix(Row, C_NUMDOCHASTA))) Then
      ValLine = (ValLine And (Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) = 0 Or (Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) <> 0 And vFmt(Grid.TextMatrix(Row, C_NUMDOCHASTA)) >= vFmt(Grid.TextMatrix(Row, C_NUMDOC)))))
      If ValLine = False And Msg = "" Then
         Msg = "El rango de Números de Documentos es inválido."
         Exit Function
      End If
   End If
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la fecha de emisión."
      Exit Function
   End If
      
   Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) <= LastDay)
   If ValLine = False And Msg = "" Then
      Msg = "La fecha de emisión es posterior al último día de este mes."
      Exit Function
   End If
         
   FRecep = DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA)))
         
'   If FRecep < Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) Then
'      MsgBox1 "La fecha de recepción del documento o de ingreso al libro, es anterior a la fecha de emisión del mismo.", vbExclamation
'      Exit Function
'   End If
   
'   If Abs(DateDiff("d", FRecep, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))) > 90 Then
'      If MsgBox1("En la línea " & Grid.TextMatrix(Row, C_NUMLIN) & " hay más de 90 días entre la fecha de ingreso o recepción del documento y la fecha de emisión del mismo." & vbNewLine & vbNewLine & "¿Está seguro que desea almacenar esta información?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   
   'validamos sólo si no está borrada la línea y si la línea fue modificada
   
   If lTipoLib = LIB_COMPRAS Then         'se restringe sólo a las compras, solicitado por Victor Morales, 19 oct 2020
      If Not lMsgFechaErr And Grid.RowHeight(Row) > 0 And Grid.TextMatrix(Row, C_UPDATE) <> "" Then
      
         If Abs(DateDiff("m", FRecep, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))) > 2 Then
   '         If MsgBox1("En la línea " & Grid.TextMatrix(Row, C_NUMLIN) & " la fecha de ingreso o recepción del documento es posterior a los dos meses siguientes de la fecha de emisión del mismo." & vbNewLine & vbNewLine & "¿Está seguro que desea almacenar esta información?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            
            TxtImp = IIf(Grid.TextMatrix(Row, C_TIPODOC) = "IMP", "inciso final", "")
            If MsgBox1("En la línea " & Grid.TextMatrix(Row, C_NUMLIN) & ", de acuerdo a lo establecido en el art. 24 " & TxtImp & " del D.L. 825, solo es posible recuperar el IVA Crédito Fiscal dentro de los dos períodos tributarios siguientes." & vbNewLine & vbNewLine & "¿Desea reclasificar el IVA CF como IVA Irrecuperable?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               If MsgBox1("Si usted esta seguro de utilizar el IVA Crédito Fiscal fuera del plazo legal presione SI/YES." & vbNewLine & vbNewLine & "Si desea considerar el IVA como Irrecuperable presione NO.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Grid.TextMatrix(Row, C_IVA_IDCUENTA) = lIdCuentaIVAIrrec
               End If
            Else
               Grid.TextMatrix(Row, C_IVA_IDCUENTA) = lIdCuentaIVAIrrec
            End If
   '         lMsgFechaErr = True     'solicitado por Nicolás Cartrín que se muestre simpre, no solo la primera vez (21 nov 2017)
         End If
      
      End If
   End If
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la fecha de vencimiento."
      Exit Function
   End If
   
   If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) > 0 And Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) Then
      Msg = "La fecha de emisión es posterior a la fecha de vencimiento."
      Exit Function
   End If
    
   ValLine = (ValLine And (Trim(Grid.TextMatrix(Row, C_RUT)) <> "" Or (Trim(Grid.TextMatrix(Row, C_RUT)) = "" And (Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO Or Not DocExigeRut(Row)))))
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el RUT."
      Exit Function
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) <> "" Then
      If Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And Not ValidCID(Grid.TextMatrix(Row, C_RUT)) Then
         Msg = "Debe ingresar un RUT válido para este tipo de documento."
         Exit Function
      End If
   End If
   
   If Not ValidaNumDoc(Row) Then
      Msg = "Documento inválido."
      Exit Function
   End If

   If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
      If Trim(Grid.TextMatrix(Row, C_NUMFISCIMPR)) = "" Then
         Msg = "Debe ingresar el número Fiscal de la Impresora."
         Exit Function
      End If
      
      If Trim(Grid.TextMatrix(Row, C_NUMINFORMEZ)) = "" Then
         Msg = "Debe ingresar el número de Informe ""Z""."
         Exit Function
      End If
      
      If Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) = 0 Then
         Msg = "Debe ingresar el número correlativo de vale-boleta final, según informe ""Z"", columna ""Num. Doc. Hasta""."
         Exit Function
      End If
      
      If vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 And vFmt(Grid.TextMatrix(Row, C_VENTASACUM)) = 0 Then
         Msg = "Debe ingresar las ventas acumuladas según informe ""Z"", columna ""Ventas Acum. Informe Z""."
         Exit Function
      End If
   End If
   
   
'   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
'      ValLine = (ValLine And (vfmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Or (vfmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 And vfmt(Grid.TextMatrix(Row, C_IVA)) <> 0)))
'      If ValLine = False And Msg = "" Then
'         Msg = "El total del documento es cero."
'         Exit Function
'      End If
'   End If

   If Grid.TextMatrix(Row, C_TIPODOC) = "IMP" And Grid.TextMatrix(Row, C_RUT) = "" Then
      AddRutDefault (Row)
   End If

   If Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAC" Or Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAV" Then    'factura de venta
      'ValLine = (ValLine And vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0)
       '3387590 if gAfectoCero
       If gAfectoCero = 0 Then
         ValLine = (ValLine And vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0)
            If ValLine = False And Msg = "" Then
               Msg = "Falta ingresar el valor Afecto."
               Exit Function
            End If
        End If
       '3387590
   End If

   If gEmpresa.Ano >= 2017 And (gEmpresa.Franq14Ter Or gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp) Then  'Solicitado por Víctor 1 mar 2021
      ValLine = ValLine And (Trim(Grid.TextMatrix(Row, C_DESCRIP)) <> "" Or Val(Grid.TextMatrix(Row, C_MOVEDITED)) <> 0)
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la descripción del documento." & vbCrLf & vbCrLf & "Ésta es obligatoria para el Libro de Caja, en empresas acogidas a 14TER/14D."
         Exit Function
      End If
   End If
   
   If Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0 Then  'estas validaciones no tienen sentido si el documento ha sido editado en detalle,porque las cuentas se asignan en el detalle
   
      'ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0, vFmt(Grid.TextMatrix(Row, C_AF_IDCUENTA)) <> 0, True)))
      
       '3387590 if gAfectoCero
       If gAfectoCero = 1 And (Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAC" Or Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAV") Then
       
       Else
         ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0, vFmt(Grid.TextMatrix(Row, C_AF_IDCUENTA)) <> 0, True)))
      
            If ValLine = False And Msg = "" Then
             Msg = "Falta ingresar la cuenta para Afecto."
             Exit Function
            End If
      
      End If
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0, vFmt(Grid.TextMatrix(Row, C_EX_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta para Exento."
         Exit Function
      End If
      
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0, vFmt(Grid.TextMatrix(Row, C_IVA_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta para IVA."
         Exit Function
      End If
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0, vFmt(Grid.TextMatrix(Row, C_OIMP_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta para Otros Impuestos."
         Exit Function
      End If
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0, vFmt(Grid.TextMatrix(Row, C_TOT_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta para Total."
         Exit Function
      End If
   
      'si es factura de compra, n. cred de factura de compra o n. debito de factura de compra, debe tener IVA ret Parcial o Total
      If Not lInBt_DetDoc And Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
         If Grid.TextMatrix(Row, C_TIPODOC) = "FCC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCF" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDF" Or Grid.TextMatrix(Row, C_TIPODOC) = "FCV" Then
         
            ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_OTROIMP)) = 0 Or (vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 And Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0), False, True)))
            If ValLine = False And Msg = "" Then
               Msg = "Falta ingresar el detalle de IVA Retenido Parcial o Total. Utilice el botón 'Detalle documento seleccionado' para ingresar este dato."
               Exit Function
            End If
         End If
      End If
   
   End If
   
   'puede haber documentos con valor 0
   'If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
   '   ValLine = (ValLine And vfmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0)
   'End If

   IsValidLine = ValLine

End Function

Private Function LineaEnBlanco(ByVal Row As Integer) As Boolean
   Dim i As Integer
   
   LineaEnBlanco = True
   
   For i = C_NUMDOC To Grid.Cols - 1
   
      If Grid.TextMatrix(Row, i) <> "" And i <> C_ESTADO And i <> C_IDESTADO And i <> C_CORRINTERNO And i <> C_LNGFECHAEMIORI And i <> C_FECHAEMIORI And i <> C_UPDATE Then
         LineaEnBlanco = False
      End If
      
   Next i
     
   If LineaEnBlanco Then
      Grid.TextMatrix(Row, C_FECHA) = ""
      Grid.TextMatrix(Row, C_TIPODOC) = ""
   End If
      
End Function

Private Sub CalcTot()
   Dim Tot(C_TOTAL) As Double
   Dim i As Integer, j As Integer
   
  
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 Then

         Tot(C_AFECTO) = Tot(C_AFECTO) + vFmt(Grid.TextMatrix(i, C_AFECTO))
         Tot(C_EXENTO) = Tot(C_EXENTO) + vFmt(Grid.TextMatrix(i, C_EXENTO))
         Tot(C_IVA) = Tot(C_IVA) + vFmt(Grid.TextMatrix(i, C_IVA))
         Tot(C_OTROIMP) = Tot(C_OTROIMP) + vFmt(Grid.TextMatrix(i, C_OTROIMP))
         Tot(C_TOTAL) = Tot(C_TOTAL) + vFmt(Grid.TextMatrix(i, C_TOTAL))
         
         If Ch_ViewDetOtrosImp <> 0 Then
         
            For j = C_INIDETOTROIMP To C_ENDDETOTROIMP
               Tot(j) = Tot(j) + vFmt(Grid.TextMatrix(i, j))
            Next j

         End If
         
'         If Tot(C_AFECTO) = 0 Then
'            MsgBeep (vbCritical)
'         End If
         
      End If
      
   Next i
   
   GridTot.TextMatrix(0, C_AFECTO) = Format(Tot(C_AFECTO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_EXENTO) = Format(Tot(C_EXENTO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_IVA) = Format(Tot(C_IVA), NEGNUMFMT)
   GridTot.TextMatrix(0, C_OTROIMP) = Format(Tot(C_OTROIMP), NEGNUMFMT)
   GridTot.TextMatrix(0, C_TOTAL) = Format(Tot(C_TOTAL), NEGNUMFMT)

   If Ch_ViewDetOtrosImp <> 0 Then
   
      For j = C_INIDETOTROIMP To C_ENDDETOTROIMP
         GridTot.TextMatrix(0, j) = Format(Tot(j), NEGNUMFMT)
      Next j

   End If
   
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub

Private Function ValidaNumDoc(ByVal Row As Integer) As Boolean
   Dim NumDoc As String
   Dim TipoDoc As Integer
   Dim IdDoc As Long
   Dim DTE As Boolean
   Dim IdEnt As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim DocNotVal As Boolean
   Dim EqDoc As Boolean
   Dim Wh As String, WhEquDoc As String
  
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   NumDoc = Trim(Grid.TextMatrix(Row, C_NUMDOC))
   TipoDoc = Val(Grid.TextMatrix(Row, C_IDTIPODOC))
   DTE = (Grid.TextMatrix(Row, C_DTE) <> "")
   IdEnt = Val(Grid.TextMatrix(Row, C_IDENTIDAD))
   
   ValidaNumDoc = False
   
   'veamos si faltan algunos datos
   If NumDoc = "" Or TipoDoc = 0 Then
      ValidaNumDoc = True
      Exit Function
   End If
   
   If Grid.TextMatrix(Row, C_TIPODOC) = "LFV" Then    'este tipo de documento puede tener correlativo repetido '¡!??  (Victor Morales 12/01/15)
      ValidaNumDoc = True
      Exit Function
   End If

   If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VALEPAGOELECTR Then    'el Num. Doc. para este tipo de documento corresponde al número de terminal o máquina que se usó para hacer el pago en el comercio, por lo tanto puede estar repetido
      ValidaNumDoc = True
      Exit Function
   End If

   
'   If IdEnt = 0 Then   'si no se ingresa la entidad, no se valida duplicidad de documentos
'      ValidaNumDoc = True
'      Exit Function
'   End If

   
   
   'primero vemos si está en la grilla
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      If i <> Row And Grid.RowHeight(i) > 0 Then
         
         EqDoc = (TipoDoc = Val(Grid.TextMatrix(i, C_IDTIPODOC)) And NumDoc = Trim(Grid.TextMatrix(i, C_NUMDOC))) And DTE = (Grid.TextMatrix(i, C_DTE) <> "")
         If lTipoLib = LIB_COMPRAS Then
            DocNotVal = EqDoc And IdEnt = Val(Grid.TextMatrix(i, C_IDENTIDAD))
         ElseIf lTipoLib = LIB_VENTAS Then
            DocNotVal = EqDoc
         End If
         
         If DocNotVal Then
         
            If IdEnt = 0 Then  'no se ha ingresada la entidad, sólo se da un mensaje de advertencia
               If MsgBox1("¡Atención!" & vbNewLine & "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & ", sin entidad asociada, ya ha sido ingresado en este libro. Es posible que esté repetido." & vbNewLine & vbNewLine & "¿Desea verificar los datos antes de continuar?", vbQuestion + vbYesNo) = vbYes Then
                  Exit Function
               End If
               
            Else
               MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en este libro.", vbExclamation + vbOKOnly
               Exit Function
            End If
            
         End If
         
      End If
   Next i
      
   'ahora vemos si está en la DB, en otros meses si es ACCESS o en este u otros meses si es SQL Server(pensando en el paginamiento)
      
   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = '" & NumDoc & "' AND Abs(DTE) = " & Abs(CInt(DTE))
   Wh = " WHERE "
   If lTipoLib = LIB_COMPRAS Then
      Wh = Wh & "TipoLib = " & lTipoLib & WhEquDoc & " AND IdEntidad = " & IdEnt
   ElseIf lTipoLib = LIB_VENTAS Then
      Wh = Wh & "TipoLib = " & lTipoLib & WhEquDoc
   End If
   If gDbType = SQL_ACCESS Then    'en otros meses si es ACCESS
      Wh = Wh & " AND " & SqlMonthLng("FEmision") & " <> " & CbItemData(Cb_Mes)
   End If
   
   Q1 = "SELECT IdDoc, FEmision FROM Documento " & Wh
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
   
      If IdEnt = 0 Then  'no se ha ingresado la entidad, sólo se da un mensaje de advertencia
         If MsgBox1("¡Atención!" & vbNewLine & "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & ", sin entidad asociada, ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ". Es posible que esté repetido." & vbNewLine & vbNewLine & "¿Desea verificar antes de continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Call CloseRs(Rs)
            Exit Function
         End If
         
         
      ElseIf vFld(Rs("IdDoc")) <> IdDoc Then
         MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ".", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      
      End If
      
   End If
      
   Call CloseRs(Rs)
   
   'ahora vemos si está en el año anterior
      
   'se hace sólo en función Valida: al final se recorre todo el libro y se ve si ha sido ingresado el año anterior.
'   If lAno = gEmpresa.Ano Then
'      If Not ValidaNumDocAnoAnt(Row) Then
'         Exit Function
'      End If
'   End If
   
   ValidaNumDoc = True

End Function

Private Function ValidaNumDocAnoAnt(ByVal Row As Integer) As Boolean
   Dim NumDoc As String
   Dim TipoDoc As Integer
   Dim DTE As Boolean
   Dim IdEnt As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim DocNotVal As Boolean
   Dim EqDoc As Boolean
   Dim Wh As String, WhEquDoc As String
   Dim RutEnt As String
   
   'si estamos editando el libro del año anterior, no validamos porque perfectamente puede estar en la db del año anterior, cuando el sistema lo trae automáticamente
   If lAno < gEmpresa.Ano Then
      ValidaNumDocAnoAnt = True
      Exit Function
   End If
      
   If gEmprSeparadas Then
#If DATACON = 1 Then       'Access
      If lDbAnoAnt Is Nothing Then
         ValidaNumDocAnoAnt = True
         Exit Function
      End If
#End If
   ElseIf Not gEmpresa.TieneAnoAnt Then
      ValidaNumDocAnoAnt = True
      Exit Function
   End If
  
   NumDoc = Trim(Grid.TextMatrix(Row, C_NUMDOC))
   TipoDoc = Val(Grid.TextMatrix(Row, C_IDTIPODOC))
   DTE = (Grid.TextMatrix(Row, C_DTE) <> "")
   IdEnt = Val(Grid.TextMatrix(Row, C_IDENTIDAD))
   RutEnt = vFmtCID(Grid.TextMatrix(Row, C_RUT))
   
   ValidaNumDocAnoAnt = False
   
   'veamos si faltan algunos datos
   If NumDoc = "" Or TipoDoc = 0 Then
      ValidaNumDocAnoAnt = True
      Exit Function
   End If
   
'   DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
'   Call OpenDb(DbAnoAnt, DbName)
   
   'vemos si está en la DB del año anterior
      
'   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = " & NumDoc & " AND DTE = " & CInt(DTE)
'   Wh = " WHERE (TipoLib = " & LIB_COMPRAS & WhEquDoc & " AND IdEntidad = " & IdEnt & ")"
'   Wh = Wh & " OR (TipoLib = " & LIB_VENTAS & WhEquDoc & ")"
'   Wh = Wh & " AND Month(FEmision) <> " & ItemData(Cb_Mes)
   
   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = '" & NumDoc & "' AND DTE = " & CInt(DTE)
   Wh = " WHERE "
   If lTipoLib = LIB_COMPRAS Then
      Wh = Wh & "TipoLib = " & lTipoLib & WhEquDoc & " AND Rut = '" & RutEnt & "'"
   ElseIf lTipoLib = LIB_VENTAS Then
      Wh = Wh & "TipoLib = " & lTipoLib & WhEquDoc & " AND Year(FEmision) <> '" & gEmpresa.Ano & "'"
   End If
      
   Q1 = "SELECT IdDoc, FEmision FROM Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Documento.IdEmpresa = Entidades.IdEmpresa "
   Q1 = Q1 & Wh
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano - 1
   
   If gEmpresa.TieneAnoAnt And gEmprSeparadas Then
#If DATACON = 1 Then       'Access
      Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
      Set Rs = OpenRs(lDbAnoAnt, Q1)
#End If
   Else
      Set Rs = OpenRs(DbMain, Q1)
   End If
   
   If Rs.EOF = False Then
   
      If IdEnt = 0 Then  'no se ha ingresado la entidad, sólo se da un mensaje de advertencia
         If MsgBox1("¡Atención!" & vbNewLine & "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & ", sin entidad asociada, ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ". Es posible que esté repetido." & vbNewLine & vbNewLine & "¿Desea verificar antes de continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Call CloseRs(Rs)
            Exit Function
         End If
         
      Else
         MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ".", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      
      End If
      
   End If
      
   Call CloseRs(Rs)
      
   ValidaNumDocAnoAnt = True

End Function

Private Sub GenMovDocumento(ByVal Row As Integer)
   Dim Q1 As String
   Dim QBase As String
   Dim i As Integer
   Dim Glosa As String
   Dim TipoDocNC As Boolean
   Dim Idx As Integer
   Dim IdANeg As Long, IdCCosto As Long
   Dim ANegCCosto As String
   Dim IdCtaAfecto As Long, IdCtaExento As Long, IdCtaIVA As Long, IdCtaOImp As Long, IdCtaTot As Long
   Dim ANegAfecto As Boolean, ANegExento As Boolean, ANegIVA As Boolean, ANegOImp As Boolean, ANegTot As Boolean
   Dim CCostoAfecto As Boolean, CCostoExento As Boolean, CCostoIVA As Boolean, CCostoOImp As Boolean, CCostoTot As Boolean
   
   
   IdCtaAfecto = Val(Grid.TextMatrix(Row, C_AF_IDCUENTA))
   IdCtaExento = Val(Grid.TextMatrix(Row, C_EX_IDCUENTA))
   IdCtaIVA = Val(Grid.TextMatrix(Row, C_IVA_IDCUENTA))
   IdCtaOImp = Val(Grid.TextMatrix(Row, C_OIMP_IDCUENTA))
   IdCtaTot = Val(Grid.TextMatrix(Row, C_TOT_IDCUENTA))
   
   ANegCCosto = Grid.TextMatrix(Row, C_IDANEG_CCOSTO)
   If ANegCCosto <> "" Then
      Idx = InStr(ANegCCosto, "-")
      If Idx > 0 Then
         IdANeg = Left(ANegCCosto, Idx - 1)
         IdCCosto = Mid(ANegCCosto, Idx + 1)
      End If
   End If
   
   If IdANeg <> 0 Then
      If GetAtribCuenta(IdCtaAfecto, ATRIB_AREANEG) <> 0 Then
         ANegAfecto = True
      End If
      If GetAtribCuenta(IdCtaExento, ATRIB_AREANEG) <> 0 Then
         ANegExento = True
      End If
      If GetAtribCuenta(IdCtaIVA, ATRIB_AREANEG) <> 0 Then
         ANegIVA = True
      End If
      If GetAtribCuenta(IdCtaOImp, ATRIB_AREANEG) <> 0 Then
         ANegOImp = True
      End If
      If GetAtribCuenta(IdCtaTot, ATRIB_AREANEG) <> 0 Then
         ANegTot = True
      End If
   End If
      
   If IdCCosto <> 0 Then
      If GetAtribCuenta(IdCtaAfecto, ATRIB_CCOSTO) <> 0 Then
         CCostoAfecto = True
      End If
   
      If GetAtribCuenta(IdCtaExento, ATRIB_CCOSTO) <> 0 Then
         CCostoExento = True
      End If
   
      If GetAtribCuenta(IdCtaIVA, ATRIB_CCOSTO) <> 0 Then
         CCostoIVA = True
      End If
            
      If GetAtribCuenta(IdCtaOImp, ATRIB_CCOSTO) <> 0 Then
         CCostoOImp = True
      End If
   
      If GetAtribCuenta(IdCtaTot, ATRIB_CCOSTO) <> 0 Then
         CCostoTot = True
      End If
   End If
         
   QBase = "INSERT INTO MovDocumento"
   QBase = QBase & "(IdDoc, IdEmpresa, Ano, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg) "
   QBase = QBase & " VALUES(" & Grid.TextMatrix(Row, C_IDDOC) & "," & gEmpresa.id & "," & gEmpresa.Ano & ","
   
   Glosa = ParaSQL(Left(Trim(Grid.TextMatrix(Row, C_DESCRIP)), 50))
   
   Idx = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   If Idx >= 0 Then
      TipoDocNC = gTipoDoc(Idx).EsRebaja
   End If
   
   
   
   i = 1
      
   If lTipoLib = LIB_COMPRAS Then
      
      'Exento
      
      If vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaExento & ","                        'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))) & ","  'Haber
         Else
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBCOMPRAS_EXENTO & ","                     'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                              'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoExento Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegExento Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
                  
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'Afecto (idem Exento)
      If vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaAfecto & ","                        'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) & ","  'Haber
         Else
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBCOMPRAS_AFECTO & ","                     'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                              'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoAfecto Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegAfecto Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
         
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
         
      'IVA Crédito Fiscal
      If vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaIVA & ","                              'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_IVA))) & ","     'Haber
         Else
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_IVA))) & ","     'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         
         'si es activo fijo, ponemos IVA Activo Fijo
         If EsCuentaActFijo(IdCtaAfecto) Then
            Q1 = Q1 & LIBCOMPRAS_IVAACTFIJO & ","              'IdTipoValLib
         ElseIf IdCtaIVA = gCtasBas.IdCtaIVAIrrec Then
            Q1 = Q1 & LIBCOMPRAS_IVAIRREC2 & ","               'IdTipoValLib
         Else
            Q1 = Q1 & LIBCOMPRAS_IVACREDFISC & ","             'IdTipoValLib
         End If
         
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoIVA Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegIVA Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
         
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'Otros Impuestos Crédito Fiscal
      If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaOImp & ","                             'IdCuenta
         
         If TipoDocNC Then
            If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
               Q1 = Q1 & "0" & ","                                         'Debe
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Haber
            Else
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Debe
               Q1 = Q1 & "0" & ","                                         'Haber
            End If
         Else
            If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) > 0 Then
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Debe
               Q1 = Q1 & "0" & ","                                         'Haber
            Else
               Q1 = Q1 & "0" & ","                                         'Debe
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Haber
            End If
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBCOMPRAS_OTROSIMP & ","                   'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoOImp Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegOImp Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
      
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'Total
      If vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaTot & ","                              'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) & ","   'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) & ","   'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBCOMPRAS_TOTAL & ","                      'IdTipoValLib
'         Q1 = Q1 & "1,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "1,"                                        'EsTotalDoc
         
         If CCostoTot Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegTot Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
         
         Call ExecSQL(DbMain, Q1)
      End If

   ElseIf lTipoLib = LIB_VENTAS Then
      
      'Exento
      If vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaExento & ","                           'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))) & ","  'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBVENTAS_EXENTO & ","                      'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoExento Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegExento Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
      
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'Afecto (idem Exento)
      If vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaAfecto & ","                           'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) & ","  'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBVENTAS_AFECTO & ","                      'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoAfecto Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegAfecto Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
        
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'IVA Débito Fiscal
      If vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaIVA & ","                              'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_IVA))) & ","      'Debe
            Q1 = Q1 & "0" & ","                                         'Haber
         Else
            Q1 = Q1 & "0" & ","                                         'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_IVA))) & ","      'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBVENTAS_IVADEBFISC & ","                  'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoIVA Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegIVA Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
         
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
      
      'Otros Impuestos Débito Fiscal
      If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaOImp & ","                             'IdCuenta
         
         If TipoDocNC Then
            If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Debe
               Q1 = Q1 & "0" & ","                                         'Haber
            Else
               Q1 = Q1 & "0" & ","                                         'Debe
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Haber
            End If
         Else
            If vFmt(Grid.TextMatrix(Row, C_OTROIMP)) > 0 Then
               Q1 = Q1 & "0" & ","                                         'Debe
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Haber
            Else
               Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) & ","  'Debe
               Q1 = Q1 & "0" & ","                                         'Haber
            End If
        End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBVENTAS_OTROSIMP & ","                    'IdTipoValLib
'         Q1 = Q1 & "0,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,"                                        'EsTotalDoc
         
         If CCostoOImp Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegOImp Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
      
         Call ExecSQL(DbMain, Q1)
         
         i = i + 1
      End If
         
      'Total
      If vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 Then
         Q1 = QBase & i & ","                                  'Orden
         Q1 = Q1 & IdCtaTot & ","                              'IdCuenta
         
         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) & ","   'Haber
         Else
            Q1 = Q1 & Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) & ","   'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If
         
         Q1 = Q1 & "'" & Glosa & "',"                          'Glosa
         Q1 = Q1 & LIBVENTAS_TOTAL & ","                       'IdTipoValLib
'         Q1 = Q1 & "1,0,0" & ")"                               'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "1,"                                        'EsTotalDoc
         
         If CCostoTot Then
            Q1 = Q1 & IdCCosto & ","                            'Centro de Costo
         Else
            Q1 = Q1 & "0,"
         End If
         
         If ANegTot Then
            Q1 = Q1 & IdANeg & ")"                             'Area de Negocio
         Else
            Q1 = Q1 & "0)"
         End If
         
         Call ExecSQL(DbMain, Q1)
      End If

   End If
   
End Sub

Private Sub Tm_ColWi_Timer()
   Call FGrTotales(Grid, GridTot)
   
End Sub

Private Sub Tx_Descrip_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

Private Sub Tx_NumDoc_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

Private Sub Tx_NumDoc_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_NumDocAsoc_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

Private Sub Tx_NumDocAsoc_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_Rut_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

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


Private Function EsCuentaBasica(ByVal IdCuenta, ByVal Col As Integer) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
   
   EsCuentaBasica = False
      
   If lTipoLib > 0 Then
   
      Q1 = "SELECT IdCuenta "
      Q1 = Q1 & " FROM CuentasBasicas "
      Q1 = Q1 & " WHERE TipoLib = " & lTipoLib & " AND IdCuenta = " & IdCuenta & " AND TipoValor ="
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
      Select Case Col
      
         Case C_EX_CODCUENTA
            
            If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_EXENTO
            Else
               Q1 = Q1 & LIBCOMPRAS_EXENTO
            End If
            
         Case C_AF_CODCUENTA
             
             If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_AFECTO
            Else
               Q1 = Q1 & LIBCOMPRAS_AFECTO
            End If
        
        Case C_TOT_CODCUENTA
        
            If lTipoLib = LIB_VENTAS Then
               Q1 = Q1 & LIBVENTAS_TOTAL
            Else
               Q1 = Q1 & LIBCOMPRAS_TOTAL
            End If

         Case Else
            Exit Function
            
      End Select
 
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
           
      Set Rs = OpenRs(DbMain, Q1)
            
      If Rs.EOF = False Then
         EsCuentaBasica = True
      End If
                             
      Call CloseRs(Rs)
                          
   End If

End Function

Private Sub LoadCuentasDef()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
      
   If lTipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion, TipoValor "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
      Q1 = Q1 & " WHERE TipoLib = " & lTipoLib
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, CuentasBasicas.Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         Select Case vFld(Rs("TipoValor"))
         
            Case LIBVENTAS_AFECTO, LIBCOMPRAS_AFECTO
            
               If lCtaAfecto.id = 0 Then
                  lCtaAfecto.id = vFld(Rs("IdCuenta"))
                  lCtaAfecto.Codigo = vFld(Rs("Codigo"))
                  lCtaAfecto.Descripcion = FCase(vFld(Rs("Descripcion"), True))
               End If
               
            Case LIBVENTAS_EXENTO, LIBCOMPRAS_EXENTO
            
               If lCtaExento.id = 0 Then
                  lCtaExento.id = vFld(Rs("IdCuenta"))
                  lCtaExento.Codigo = vFld(Rs("Codigo"))
                  lCtaExento.Descripcion = FCase(vFld(Rs("Descripcion"), True))
               End If
            
            Case LIBVENTAS_TOTAL, LIBCOMPRAS_TOTAL
            
               If lCtaTotal.id = 0 Then
                  lCtaTotal.id = vFld(Rs("IdCuenta"))
                  lCtaTotal.Codigo = vFld(Rs("Codigo"))
                  lCtaTotal.Descripcion = FCase(vFld(Rs("Descripcion"), True))
               End If
               
         End Select
                           
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
End Sub

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Cb_Entidad.ListIndex = 0  'en blanco
      Exit Sub
   End If
   
'   If Not MsgValidCID(Tx_Rut) Then    'está en Tx_Rut_Validate
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
      For i = 0 To MAX_ENTCLASIF
         If Cb_Entidad.ItemData(i) >= 0 Then
            If vFld(Rs("Clasif" & Cb_Entidad.ItemData(i))) <> 0 Then
               Cb_Entidad.ListIndex = i
               Exit For
            End If
         End If
      Next i
   
      'ahora seleccionamos la entidad
      For i = 0 To Cb_Nombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
      Bt_List.Enabled = True
      Bt_Sel.Enabled = False
      Bt_Centralizar.Enabled = False

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
Private Sub cb_Nombre_Click()
   
   If lcbNombre.ListIndex >= 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
      Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
   End If
   
   Bt_List.Enabled = True
   Bt_Sel.Enabled = False
   Bt_Centralizar.Enabled = False

End Sub
Private Sub Cb_Entidad_Click()
      
   Cb_Nombre.Clear
   If CbItemData(Cb_Entidad) >= 0 Then
      Call SelCbEntidad(CbItemData(Cb_Entidad))
   Else
      Tx_Rut = ""
   End If
   
   Bt_List.Enabled = True
   Bt_Sel.Enabled = False
   Bt_Centralizar.Enabled = False

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
   End If
End Sub
Private Sub SetOrderLst()
   Dim i As Integer
   Dim StrOrder As String
   
   StrOrder = ", Documento.TipoDoc, Documento.NumDoc"
   
   lOrdenGr(C_NUMLIN) = "Documento.IdDoc"
   lOrdenGr(C_CORRINTERNO) = "Documento.CorrInterno"
   lOrdenGr(C_FECHA) = "Documento.FEmision" & StrOrder
   lOrdenGr(C_TIPODOC) = "Documento.TipoDoc, Documento.DTE, Documento.FEmision, Documento.NumDoc"
   lOrdenGr(C_DTE) = "Documento.DTE, Documento.FEmision, Documento.NumDoc"
   lOrdenGr(C_GIRO) = "Documento.Giro, Documento.FEmision, Documento.NumDoc"
   lOrdenGr(C_PROPIVA) = "Documento.PropIVA, Documento.FEmision, Documento.NumDoc"
   lOrdenGr(C_NUMDOC) = "Documento.NumDoc, Documento.TipoDoc"
   lOrdenGr(C_RUT) = "Entidades.Rut" & StrOrder
   lOrdenGr(C_NOMBRE) = "Entidades.Nombre" & StrOrder
   lOrdenGr(C_DESCRIP) = "Documento.Descrip" & StrOrder
   lOrdenGr(C_SUCURSAL) = "Sucursales.Descripcion" & StrOrder
   lOrdenGr(C_EXENTO) = "Documento.Exento" & StrOrder
   lOrdenGr(C_EX_CODCUENTA) = "Cuentas1.Codigo" & StrOrder
   lOrdenGr(C_EX_CUENTA) = "Cuentas1.Descripcion" & StrOrder
   lOrdenGr(C_AFECTO) = "Documento.Afecto" & StrOrder
   lOrdenGr(C_AF_CODCUENTA) = "Cuentas2.Codigo" & StrOrder
   lOrdenGr(C_AF_CUENTA) = "Cuentas2.Descripcion" & StrOrder
   lOrdenGr(C_IVA) = "Documento.IVA" & StrOrder
   lOrdenGr(C_OTROIMP) = "Documento.OtroImp" & StrOrder
   lOrdenGr(C_TOTAL) = "Documento.Total" & StrOrder
   lOrdenGr(C_TOT_CODCUENTA) = "Cuentas3.Codigo" & StrOrder
   lOrdenGr(C_TOT_CUENTA) = "Cuentas3.Descripcion" & StrOrder
   lOrdenGr(C_FECHAEMIORI) = "Documento.FEmisionOri " & StrOrder
   lOrdenGr(C_FECHAVENC) = "Documento.FVenc" & StrOrder
   lOrdenGr(C_ESTADO) = "Documento.Estado" & StrOrder
   lOrdenGr(C_USUARIO) = "Usuarios.Usuario" & StrOrder

   If lOper = O_VIEWLIBLEGAL Then
      lOrdenSel = C_FECHAEMIORI
   Else
      lOrdenSel = C_NUMLIN
   End If
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)

   If lOrdenGr(Col) = "" Then
      Exit Sub
   End If
   
   If lOper = O_EDIT Or lOper = O_NEW Then
      If GrabarParaContinuar("cambiar el ordenamiento de las columnas") = False Then
         Exit Sub
      End If
   End If
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 1
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadGrid
      
   Me.MousePointer = vbDefault
      
End Sub

Private Sub ActivoFijo(ByVal Row As Integer, ByVal Col As Integer)
   Dim IdDoc As Long
   Dim n As Integer
   Dim ValNeto As Double
   Dim ValIVA As Double
   Dim Frm As Form
   Dim Fecha As Long
   Dim IdCuenta As Long
   Dim IdCtaEx As Long
   Dim IdCtaAf As Long
   
   '2861733
   Dim Q1 As String
   Dim Rs As Recordset
   '2861733

   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
   
   'si no hay IdDoc en esta fila, hay que grabar primero
   
   If Val(Grid.TextMatrix(Row, C_IDDOC)) = 0 Then
      If MsgBox1("Para ingresar Activos Fijos, es necesario grabar los cambios hechos a este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
         Exit Sub
      End If
      
      If valida() = False Then
         Exit Sub
      End If

      Call SaveGrid
      
      If Val(Grid.TextMatrix(Grid.Row, C_IDDOC)) = 0 Then  'algo falló
         Exit Sub
      End If
   End If
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   Fecha = DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA)))

   'vemos si hay algún mov. de activo fijo asociado a este doc
   n = CountActFijo(IdDoc)
   
   If n = 0 And lTipoLib = LIB_COMPRAS Then         'no hay ninguno aún y es libro de compras, llamamos al form de activo fijo que permite crear un mov directamente, sin pasar por la lista
      
      Set Frm = New FrmActivoFijo
      IdCtaEx = 0
      IdCtaAf = 0
      If Col = C_EX_CODCUENTA Or Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) = 0 Then
         ValNeto = Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO)))
         ValIVA = 0
         
         IdCuenta = Val(Grid.TextMatrix(Row, C_EX_IDCUENTA))

         If EsCuentaActFijo(IdCuenta) Then
            IdCtaEx = IdCuenta
         End If
      Else              'lo más probable
         ValNeto = Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO)))
         ValIVA = Abs(vFmt(Grid.TextMatrix(Row, C_IVA)))
         
         IdCuenta = Val(Grid.TextMatrix(Row, C_AF_IDCUENTA))
         If EsCuentaActFijo(IdCuenta) Then
            IdCtaAf = IdCuenta
           
         End If
      End If
      
      '2861733
      Dim vArea As Long
      Dim vCentro As Long

      Q1 = ""
      Q1 = Q1 & "SELECT IdAreaNeg,IdCCosto from MovDocumento "
      Q1 = Q1 & " where iddoc = " & IdDoc
      Q1 = Q1 & " and IdTipoValLib = 1 "

      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = False Then

      vArea = vFld(Rs("IdAreaNeg"))
      vCentro = vFld(Rs("IdCCosto"))
      End If

      If Frm.FNewFromDoc(IdDoc, Fecha, ValNeto, ValIVA, Grid.TextMatrix(Row, C_DESCRIP), lTipoLib, IdCtaAf, IdCtaEx) = vbOK Then
         Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO
      End If

      If Frm.FNewFromDocActFijo(IdDoc, Fecha, ValNeto, ValIVA, Grid.TextMatrix(Row, C_DESCRIP), lTipoLib, IdCtaAf, IdCtaEx, vArea, vCentro) = vbOK Then
         Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO
      End If
      ''2861733
            
      Set Frm = Nothing
   
   Else              'hay uno o más o es libro de ventas => llamamos a la lista de activos fijos para este doc
      
      'primero actualizamos las cuentas en los Mov de Activo Fijo ya existentes, por si el usuario cambió la cuenta de exento o afecto
      
'      If Col = C_AF_CODCUENTA Then
'
'         IdCuenta = Val(Grid.TextMatrix(Row, C_AF_IDCUENTA))
'         If EsCuentaActFijo(IdCuenta) Then
'            Call ExecSQL(DbMain, "UPDATE MovActivoFijo SET IdCuenta = " & IdCuenta & " WHERE IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC)) & " AND IVA <> 0 ")
'         End If
'
'      ElseIf Col = C_EX_CODCUENTA Then
'         IdCuenta = Val(Grid.TextMatrix(Row, C_EX_IDCUENTA))
'
'         If EsCuentaActFijo(IdCuenta) Then
'            Call ExecSQL(DbMain, "UPDATE MovActivoFijo SET IdCuenta = " & IdCuenta & " WHERE IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC)) & " AND IVA = 0 ")
'         End If
'
'      End If
      
      'ahora listamos los movs de activo fijo
      Set Frm = New FrmLstActFijo
      If lTipoLib = LIB_COMPRAS Then
         Call Frm.FEditFromDoc(Val(Grid.TextMatrix(Grid.Row, C_IDDOC)), lTipoLib, Fecha)
      Else
         If MsgBox1("¿Desea registrar la venta en alguno de los activos fijos ya ingresados en el sistema?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Call Frm.FEditFromDoc(0, lTipoLib, Fecha)
         End If
      End If
      Set Frm = Nothing
 
   End If

End Sub

Private Function CountActFijo(ByVal IdDoc As Long) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   CountActFijo = 0
   
   If IdDoc = 0 Then
      Exit Function
   End If
   
   If lTipoLib <> LIB_COMPRAS Then
      Exit Function
   End If
     
   Q1 = "SELECT Count(*) FROM MovActivoFijo WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      CountActFijo = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)

End Function

Private Sub GridActivoFijo(ByVal IdCuenta As Long, ByVal Row As Integer, ByVal Col As Integer)
   Dim HayCtaActFijo As Boolean
   Dim Fecha As Long
   Dim Frm As FrmLstActFijo
      
   If Col <> C_AF_CODCUENTA And Col <> C_EX_CODCUENTA Then
      Exit Sub
   End If
   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
   
   'vemos si es cuenta de activo fijo
   If EsCuentaActFijo(IdCuenta) And Grid.TextMatrix(Row, C_MSGACTFIJO) = "" Then
   
      'si no es Nota de Credito
      If Not gTipoDoc(GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).EsRebaja Then
      
         If CountActFijo(Val(Grid.TextMatrix(Row, C_IDDOC))) = 0 Then
                 
            Grid.TextMatrix(Row, C_MSGACTFIJO) = True
            
            If MsgBox1("¿Desea ingresar la información de este activo fijo para llevar el control financiero y tributario correspondiente?", vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
               Exit Sub
            End If
            
            Call ActivoFijo(Row, Col)
            
         End If
         
      Else        'es nota de crédito con cuenta de activo fijo
         Grid.TextMatrix(Row, C_MSGACTFIJO) = True
         MsgBox1 "Se recomienda revisar la lista de activos fijos por si se requiere modificar algún valor, producto de la emisión de este documento.", vbInformation
     End If
   
   'no es de activo fijo, vemos si tenía un activo fijo definido, para sugerir que el usuario revise la lista de activos fijos asociados
   ElseIf CountActFijo(Val(Grid.TextMatrix(Row, C_IDDOC))) > 0 Then
         
      'vemos si la otra cuenta es de activo fijo
      If Col = C_AF_CODCUENTA Then
         If EsCuentaActFijo(Val(Grid.TextMatrix(Row, C_AF_IDCUENTA))) Then
            HayCtaActFijo = True
         End If
      ElseIf EsCuentaActFijo(Val(Grid.TextMatrix(Row, C_EX_IDCUENTA))) Then
         HayCtaActFijo = True
      End If
      
      If Not HayCtaActFijo Then
'         If MsgBox1("Esta cuenta NO es de Activo Fijo. Se eliminarán los Activos Fijos asociados a este documento." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
'            Call ExecSQL(DbMain, "DELETE * FROM MovActivoFijo WHERE IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC)))
'         End If
         If MsgBox1("Esta cuenta NO es de Activo Fijo. Se recomienda revisar la lista de Activos Fijos asociados a este documento, por si es necesario hacer alguna modificación." & vbNewLine & vbNewLine & "¿Desea hacerlo ahora?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Set Frm = New FrmLstActFijo
            Fecha = DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA)))
            Call Frm.FEditFromDoc(Val(Grid.TextMatrix(Grid.Row, C_IDDOC)), lTipoLib, Fecha)
            Set Frm = Nothing
         End If
      
      End If
           
      Grid.TextMatrix(Row, C_DETACTFIJO) = ""
      
   End If
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid, Grid.FlxGrid.Row, Grid.FlxGrid.RowSel, Grid.FlxGrid.Col, Grid.FlxGrid.ColSel)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_Sum_Click_Old()
   Dim Col1 As Integer, Col2 As Integer
   Dim Row1 As Integer, Row2 As Integer
   Dim Tot As Double
   Dim i As Integer
   
   Col1 = Grid.FlxGrid.Col
   Col2 = Grid.FlxGrid.ColSel
   
   Row1 = Grid.FlxGrid.Row
   Row2 = Grid.FlxGrid.RowSel
   
   If Row1 <> Row2 And Col1 <> Col2 Then
      MsgBox1 "Seleccione celdas en una sola fila o en una sola columna.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Tot = 0
   If Row1 <> Row2 Then
      For i = Row1 To Row2
         Tot = Tot + vFmt(Grid.TextMatrix(i, Col1))
      Next i
   Else
      For i = Col1 To Col2
         Tot = Tot + vFmt(Grid.TextMatrix(Row1, i))
      Next i
   
   End If
   
   MsgBox "Total calculado = " & Format(Tot, NEGNUMFMT), vbInformation + vbOKOnly
   
End Sub

Private Function SetupPriv()
     
   If lOper = O_EDIT Or lOper = O_NEW Then
      lEditEnabled = True
   End If
   lIngDocsEnabled = True
   lAdmDocsEnabled = True
   
   If lOper = O_EDIT Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         Call EnableForm(Me, False)
         lEditEnabled = False
         Bt_DetDoc.Enabled = True   'Bt_DetDoc lo abrirá sólo de lectura
         'Bt_ActivoFijo.Enabled = True
      End If
   
   Else
   
      If Not ChkPriv(PRV_ING_DOCS) Then
         Bt_ExitNewDoc.Enabled = False
         lIngDocsEnabled = False
      End If
         
      If Not ChkPriv(PRV_ADM_DOCS) Then
         Bt_Centralizar.Enabled = False
         lAdmDocsEnabled = False
      End If
      
   End If
   
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If



End Function

Private Function EsIngresoTotal(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
   
   EsIngresoTotal = False
   
   Idx = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
'   If Not gTipoDoc(Idx).TieneAfecto And Not gTipoDoc(Idx).TieneExento Then 'no se ingresa afecto y exento, sólo total
   If gTipoDoc(Idx).IngresarTotal <> 0 Then 'no se ingresa afecto y exento, sólo total
      EsIngresoTotal = True
   End If

End Function

Private Function EsDocExento(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
   
   EsDocExento = False
   
   Idx = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
'   If gTipoDoc(Idx).TieneExento And Not gTipoDoc(Idx).TieneAfecto Then 'sólo exento
   If gTipoDoc(Idx).TieneExento <> 0 And gTipoDoc(Idx).TieneAfecto = 0 Then  'sólo exento
      EsDocExento = True
   End If

End Function
Private Function DocExigeRut(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
      
   Idx = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   DocExigeRut = gTipoDoc(Idx).ExigeRUT

End Function
Private Sub Tx_Valor_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub
Private Function ValidaEstadoEdit(ByVal Row As Integer) As Boolean

   ValidaEstadoEdit = lEditEnabled And (Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0 And (Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PENDIENTE Or (lAno < gEmpresa.Ano And Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_CENTRALIZADO)))

End Function
Private Function ValidaFExported(ByVal Row As Integer) As Boolean
   ValidaFExported = True

   
   'mostramos un mensaje de advertencia. Si lo cambiamos por una consulta nos afecta el manejo de eventos
   If Val(Grid.TextMatrix(Row, C_EXPORTED)) <> 0 Then
      Call MsgBox1("ADVERTENCIA: Este documento ya ha sido traspasado al año siguiente." & vbCrLf & vbCrLf & "Si lo modifica, tendrá que:" & vbCrLf & "- marcarlo para 'Volver a Exportar'" & vbCrLf & "- eliminarlo en el año siguiente" & vbCrLf & "- volver a generar la Apertura", vbOKOnly + vbExclamation)
      Grid.TextMatrix(Row, C_EXPORTED) = 0    'para que no vuelva a mostrar este mensaje para ESTE documento
   End If
   

End Function
Private Sub ActCambioTipoDoc(ByVal Row As Integer)
   Dim IdxDoc As Integer

   If vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 Then
      
      IdxDoc = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
      
      If gTipoDoc(IdxDoc).EsRebaja Then
         Grid.TextMatrix(Row, C_EXENTO) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))) * -1, NEGNUMFMT)
         Grid.TextMatrix(Row, C_AFECTO) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))) * -1, NEGNUMFMT)
         Grid.TextMatrix(Row, C_IVA) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_IVA))) * -1, NEGNUMFMT)
         Grid.TextMatrix(Row, C_OTROIMP) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))) * -1, NEGNUMFMT)
         Grid.TextMatrix(Row, C_TOTAL) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) * -1, NEGNUMFMT)
      Else
         Grid.TextMatrix(Row, C_EXENTO) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_EXENTO))), NEGNUMFMT)
         Grid.TextMatrix(Row, C_AFECTO) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_AFECTO))), NEGNUMFMT)
         Grid.TextMatrix(Row, C_IVA) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_IVA))), NEGNUMFMT)
         Grid.TextMatrix(Row, C_OTROIMP) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_OTROIMP))), NEGNUMFMT)
         Grid.TextMatrix(Row, C_TOTAL) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))), NEGNUMFMT)
      End If
      
      If InStr(LCase(gTipoDoc(IdxDoc).Nombre), "exent") > 0 Then
         'es doc exento
         
         Grid.TextMatrix(Row, C_AF_IDCUENTA) = ""
         Grid.TextMatrix(Row, C_AF_CODCUENTA) = ""
         Grid.TextMatrix(Row, C_AF_CUENTA) = ""
         Grid.TextMatrix(Row, C_AFECTO) = 0
      
         Call CalcTotRow(Row)
      End If
            
   End If

End Sub
Private Function PrtResumenIVA(PrtObj As Object, ByVal Pag As Integer, ByVal LeftX As Integer, ByVal RightX As Integer, Optional ByVal CallNewPage As Boolean = False) As Integer
   Dim PrtPage As Object
   Dim CurY As Integer
   Dim ResHeight As Integer
   Dim ResWidth As Integer
   Dim TotIVACred As Double
   Dim TotIVADeb As Double
   Dim OldFntSize As Integer
   Dim StrVal As String
   Dim CurX As Integer
   Dim TotIVA As Double
   Dim ResOImp() As ResOImp_t
   Dim Where As String
   Dim CurrLib As String
   Dim TotOImp As Double
   Dim i As Integer
   Dim Lib As String
   Dim TopY As Integer
   Dim TabX As Integer
   Dim RemMesAntUTM As Double
   Dim ValUTM As Double
   Dim Fecha As Double
   Dim TotRemMesAnt As Double
   Dim TotRemUTM As Double
   Dim IVAIrrec As Double, IVARetParcial As Double, IVARetTotal As Double
   Dim TotIEPDGen As Double, TotIEPDTransp As Double
   Dim dx As Integer
   Dim AuxTxt As String
   

   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   PrtPage.Print
   PrtPage.Print
   
   Where = " " & SqlYearLng("FEmision") & " = " & Val(Cb_Ano)
   
'   If lTipoLib > 0 Then
'      Where = Where & " AND Documento.TipoLib = " & lTipoLib
'   End If
   
   If CbItemData(Cb_Mes) > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & CbItemData(Cb_Mes)
   End If
   
   Call GenResOImp(Where, ResOImp)
   
   If UBound(ResOImp) = 0 And ResOImp(0).CodValLib = 0 Then 'no hay otros impuestos
      ResHeight = 1500 + Grid.RowHeight(0) + 3
   Else
      ResHeight = 1500 + (UBound(ResOImp) + 1) * Grid.RowHeight(0) + 3
      
      'buscamos el IVA irrecuperable, IVA Ret Parcual e IVA Ret Total
      For i = 0 To UBound(ResOImp)
         If ResOImp(i).TipoLib = LIB_COMPRAS And (ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC1 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC2 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC3 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC4 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC9) Then
            IVAIrrec = ResOImp(i).Valor
         End If
         
         If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_PARCIAL Then
            IVARetParcial = ResOImp(i).Valor
         End If
         
         If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_TOTAL Then
            IVARetTotal = ResOImp(i).Valor
         End If

      Next i

   End If
   
   ResWidth = 4500
   
   CurY = PrtPage.CurrentY
   
   If PrtPage.CurrentY >= PrtPage.Height - 2000 - ResHeight Or CallNewPage Then
      Call gPrtReportes.PrtFooter(PrtPage, "Continua >>>", RightX)
      Set PrtPage = NewPage(PrtObj)
      Pag = Pag + 1
      
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      CurY = PrtPage.CurrentY

   End If
   
   'LeftX = RightX - ResWidth   'se imprime alineado a la izquierda con la grilla (se mantiene LexfX que viene como parámetro)
   
   TopY = CurY
   PrtPage.Line (LeftX, CurY)-(LeftX + ResWidth, CurY)
'   PrtPage.Line (LeftX, CurY)-(LeftX, CurY + ResHeight)
'   PrtPage.Line (LeftX + ResWidth, CurY)-(LeftX + ResWidth, CurY + ResHeight)
'   PrtPage.Line (LeftX, CurY + ResHeight)-(LeftX + ResWidth, CurY + ResHeight)
   
   '3340329
   Dim Mes As Integer
   
   Mes = CbItemData(Cb_Mes)
   For i = 1 To Mes
   
   ''3340329
   'If GetRemIVAUTM(CbItemData(Cb_Mes), Val(Cb_Ano), RemMesAntUTM) < 0 Then
   
   '3389677 FPR se creo para que no traspase si es un saldo a pagar al mes siguiente, ya que no es remanente
   If i = Mes Then
        If Not remMesAnt Then
            'Tx_RemMesAnt.Text = 0
            TotRemMesAnt = 0
            vTotalRemMesAnt = 0
            'TotIVA = 0
        End If
   End If
   'FIN 3389677 FPR
   
   
   If GetRemIVAUTM_New(i, Val(Cb_Ano), RemMesAntUTM, vFmt(Format(Abs(vTotalRemMesAnt), NEGNUMFMT))) < 0 Then
   '3340329
      RemMesAntUTM = 0
   End If
   
   '3340329
   'Fecha = DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes) + 1, 1)      'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011)
   Fecha = DateSerial(Val(Cb_Ano), i + 1, 1)      'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011)
   '3340329
   
   If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
      TotRemMesAnt = RemMesAntUTM * ValUTM
   Else
      TotRemMesAnt = 0
   End If

   ''   '3340329
   'Call GetResIVA(CbItemData(Cb_Mes), Val(Cb_Ano), TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
   Call GetResIVA(i, Val(Cb_Ano), TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
   '   '3340329
   
   TotIVACred = TotIVACred - IVAIrrec          'se resta IVA Irrecuperable que se obtiene en la función LoadValOImp
      
   TotIVADeb = TotIVADeb - IVARetParcial - IVARetTotal        'se resta IVA Retenido Parcial o Total que se obtiene en la función LoadValOImp
    
    
   Dim AjusteIvaMen As Double
   'Dim Mes As Integer
   'Mes = ItemData(Cb_Mes)
   AjusteIvaMen = GetAjusteIVAMensual(i)
    
    
   TotIVA = TotIVADeb - (TotIVACred + TotIEPDGen + TotIEPDTransp + TotRemMesAnt + AjusteIvaMen)     'FCA 28/09/2021
   '3340329
   vTotalRemMesAnt = TotIVA
   '3340329
   
   If TotIVA < 0 Then
      remMesAnt = True
   Else
   remMesAnt = False
   End If
   '3389677 FPR
        
   Next i
   '3340329
    
   OldFntSize = Printer.FontSize
   Printer.FontSize = 10
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) / 2
   AuxTxt = "RESUMEN"
   PrtPage.CurrentX = LeftX + (ResWidth - PrtPage.TextWidth(AuxTxt)) / 2   '1750
   PrtPage.Print AuxTxt;
   
   PrtPage.FontSize = 8
   
   TabX = 3000    '2500
   dx = 1800
      
   'IVA
   
   CurY = PrtPage.CurrentY
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) / 2
   PrtPage.CurrentX = LeftX + 1860
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 2
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Remanente periodo anterior:";
   PrtPage.CurrentX = LeftX + 1600
   'PrtPage.Print "$ ";
   Call gPrtReportes.PrtAlign_(PrtPage, Format(TotRemMesAnt, NEGNUMFMT), LeftX + TabX, 1200, vbRightJustify)
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 3
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Total IVA Débito:";
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
   Call gPrtReportes.PrtAlign_(PrtPage, Format(TotIVADeb, NEGNUMFMT), LeftX + TabX, 1200, vbRightJustify)
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 4
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Total IVA Crédito:";
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
   Call gPrtReportes.PrtAlign_(PrtPage, Right("               " & Format(TotIVACred, NEGNUMFMT), 18), LeftX + TabX, 1200, vbRightJustify)
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 5
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Total IEPD General CF:";
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
   Call gPrtReportes.PrtAlign_(PrtPage, Right("               " & Format(TotIEPDGen, NEGNUMFMT), 18), LeftX + TabX, 1200, vbRightJustify)
   
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 6
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Total IEPD Transporte CF:";
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
'   PrtPage.FontUnderline = True
   Call gPrtReportes.PrtAlign_(PrtPage, Right("               " & Format(TotIEPDTransp, NEGNUMFMT), 18), LeftX + TabX, 1200, vbRightJustify)
   PrtPage.FontUnderline = False
   
   
'   Dim AjusteIvaMen As Double
'   Dim Mes As Integer
'   Mes = ItemData(Cb_Mes)
'   AjusteIvaMen = GetAjusteIVAMensual(Mes)
   
   'ajuste remanente
   'gcb22092021
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 7
   PrtPage.CurrentX = LeftX + 300
   PrtPage.Print "Ajuste Remanente:";
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
   PrtPage.FontUnderline = True
   Call gPrtReportes.PrtAlign_(PrtPage, Right("               " & Format(AjusteIvaMen, NEGNUMFMT), 18), LeftX + TabX, 1200, vbRightJustify)
   PrtPage.FontUnderline = False
   
   '   '3340329
'   TotIVA = TotIVADeb - (TotIVACred + TotIEPDGen + TotIEPDTransp + TotRemMesAnt + AjusteIvaMen)     'FCA 28/09/2021

'   vTotalRemMesAnt = TotIVA
'   '3340329
'

   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 8
   PrtPage.CurrentX = LeftX + 300
   If TotIVA < 0 Then    'remanente
      PrtPage.Print "Remanente periodo siguiente:";
      
   Else                          'a pagar
      PrtPage.Print "Total a Pagar:";
   End If
   PrtPage.CurrentX = LeftX + dx
   'PrtPage.Print "$ ";
   Call gPrtReportes.PrtAlign_(PrtPage, Format(Abs(TotIVA), NEGNUMFMT), LeftX + TabX, 1200, vbRightJustify)
   
   If TotIVA < 0 Then
         
      If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
         
         '3340329
         'TotRemUTM = TotIVA / ValUTM
         TotRemUTM = vFmt(Format(Abs(TotIVA), NEGNUMFMT)) / ValUTM
         '3340329
         PrtPage.Print
         PrtPage.CurrentX = LeftX + 300
         PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 9
         PrtPage.Print "Remanente IVA Crédito Fiscal UTM:";
         PrtPage.CurrentX = LeftX + dx
         Call gPrtReportes.PrtAlign_(PrtPage, Format(Abs(TotRemUTM), DBLFMT2), LeftX + TabX, 1200, vbRightJustify)
      End If
      
   End If
   
   'Otros Impuestos
      
   If UBound(ResOImp) = 0 And ResOImp(0).CodValLib = 0 Then 'no hay otros impuestos
      CurY = PrtPage.CurrentY + Grid.RowHeight(0) * 1.5
      
      PrtPage.Line (LeftX, TopY)-(LeftX, CurY)
      PrtPage.Line (LeftX + ResWidth, TopY)-(LeftX + ResWidth, CurY)
      PrtPage.Line (LeftX, CurY)-(LeftX + ResWidth, CurY)
   
      Printer.FontSize = OldFntSize
      Exit Function
   End If
   CurY = PrtPage.CurrentY
   PrtPage.CurrentY = CurY + Grid.RowHeight(0) * 2
   AuxTxt = "Otros Impuestos " & gTipoLib(lTipoLib)
   PrtPage.CurrentX = LeftX + (ResWidth - PrtPage.TextWidth(AuxTxt)) / 2
   PrtPage.FontUnderline = True
   PrtPage.Print AuxTxt;
   PrtPage.FontUnderline = False
   CurY = PrtPage.CurrentY

   CurrLib = ""
   TotOImp = 0
   
   For i = 0 To UBound(ResOImp)
      
      If ResOImp(i).TipoLib = lTipoLib And ResOImp(i).CodValLib <> 0 Then
         CurY = CurY + Grid.RowHeight(0)
         PrtPage.CurrentY = CurY
         PrtPage.CurrentX = LeftX + 300
         
         Lib = ReplaceStr(gTipoLib(ResOImp(i).TipoLib), "Libro de ", "")
         
         If Lib <> CurrLib Then
            CurrLib = Lib
            'Grid.TextMatrix(j, C_LIBRO) = Lib
         End If
         
         PrtPage.Print Left(ResOImp(i).DescValLib, 40)
         PrtPage.CurrentY = CurY
         PrtPage.CurrentX = LeftX + 300
'         If i = UBound(ResOImp) Then   'es el último
'            PrtPage.FontUnderline = True
'         End If
         Call gPrtReportes.PrtAlign_(PrtPage, Format(ResOImp(i).Valor, NEGNUMFMT), LeftX + TabX, 1200, vbRightJustify)
         PrtPage.FontUnderline = False
         
         TotOImp = TotOImp + ResOImp(i).Valor
      
      End If
   Next i
   
   CurY = CurY + Grid.RowHeight(0)
   
'   No se eimprime el total de otros impuestos (7 Jul 2008)

'   PrtPage.CurrentY = CurY
'   PrtPage.CurrentX = LeftX + 300
'   PrtPage.Print "Total Otros Impuestos " & CurrLib
'   PrtPage.CurrentY = CurY
'   PrtPage.CurrentX = LeftX + 300
'   Call gPrtReportes.PrtAlign_(PrtPage, Format(TotOImp, NEGNUMFMT), LeftX + TabX, 1200, vbRightJustify)
'   CurY = CurY + Grid.RowHeight(0)
   
   
   CurY = PrtPage.CurrentY + Grid.RowHeight(0) * 1.5
   
   PrtPage.Line (LeftX, TopY)-(LeftX, CurY)
   PrtPage.Line (LeftX + ResWidth, TopY)-(LeftX + ResWidth, CurY)
   PrtPage.Line (LeftX, CurY)-(LeftX + ResWidth, CurY)

   Printer.FontSize = OldFntSize

End Function

Private Function GrabarParaContinuar(ByVal MsgOper As String) As Boolean
   Dim i As Integer
   Dim Upd As Boolean

   GrabarParaContinuar = False

   'vemos si el usuario hizo algún cambio
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> "" Then
         Upd = True
         Exit For
      End If
   Next i
      
   If Upd Then
      If MsgBox1("Para " & MsgOper & ", es necesario grabar los cambios hechos en el libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
         Exit Function
      End If
      
      If valida() Then
         Call SaveGrid
      Else
         Exit Function
      End If
   End If
   
   GrabarParaContinuar = True

End Function
Private Function ImportFromFile() As Boolean
   Dim fname As String
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ImpEnable As Boolean
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim i As Integer, l As Integer
   Dim j As Integer, p As Long, k As Integer
   Dim NUpd As Long
   Dim NIns As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim Aux As String
   Dim DtRec As Long, DtEmi As Long, DtVenc As Long
   Dim CampoInvalido As String
   Dim IdTipoDoc As Integer
   Dim DTE As Integer, DelGiro As Integer, TxtDTE As String
   Dim NumDoc As String, NumDocHasta As String
   Dim RutEnt As String, CodEnt As String, NombEnt As String
   Dim ClasifEnt As Integer
   Dim Afecto As Double, Exento As Double, IVA As Double, OtroImp As Double, Total As Double
   Dim NumInterno As Long
   Dim Row As Integer, r As Integer
   Dim TipoDoc As String, DocImpExp As Boolean
   Dim CodSuc As String, IdSucursal As Long, Sucursal As String
   Dim Descrip As String
   Dim Dt1 As Long, Dt2 As Long
   Dim IdCtaAfecto As Long, CodCtaAfecto As String, DescCtaAfecto As String
   Dim IdCtaExento As Long, CodCtaExento As String, DescCtaExento As String
   Dim IdCtaTotal As Long, CodCtaTotal As String, DescCtaTotal As String
   Dim IdCtaIVA As Long, IdCtaOtroImp As Long, CodCtaOtroImp As String
   Dim Estado As Integer
   Dim NRecErroneos As Integer, StrNRecErroneos As String
   Dim IdPropIVA As Integer, PropIVA As String, AceptaPropIVA As Boolean
   Dim AuxCodCtaAfecto As String, AuxIdCtaAfecto As Long, AuxDescCtaAfecto As String
   Dim AuxCodCtaExento As String, AuxIdCtaExento As Long, AuxDescCtaExento As String
   Dim AuxCodCtaTotal As String, AuxIdCtaTotal As Long, AuxDescCtaTotal As String
   Dim AuxCodCtaOtroImp As String, AuxIdCtaOtroImp As Long, AuxDescCtaOtroImp As String
   Dim NomCta As String, UltNivel As Boolean
   Dim NumFiscImp As String, NumInfZ As String
   Dim CantBoletas As Long, VentasAcumInfZ As Double
   Dim IdxTipoDoc As Long
   Dim CodANeg As String, CodCCosto As String, IdANeg As Long, IdCCosto As Long
   Dim AtribANeg As Boolean, AtribCCosto As Boolean, pANeg As Long, pCCosto As Long
   Dim FldArray(3) As AdvTbAddNew_t
   Dim Mayor3000reg As Boolean
   
   'se comenta segun lo solicitado en req. 2764744
    'pipe 2738156 tema 3
'         If gDbType = SQL_SERVER And lTipoLib = LIB_COMPRAS Or gDbType = SQL_SERVER And lTipoLib = LIB_VENTAS Then
'         MsgBox1 "En version SQL para Compras y Ventas se debe capturar la información mediante registro CSV (Menu Procesos - Importar Registros SII)", vbExclamation + vbOKOnly
'        Bt_Importar.Enabled = False
'         Bt_HlpImport.Enabled = False
'         Exit Function
'        End If
    'fin
     'fin se comenta segun lo solicitado en req. 2764744
   
   Mayor3000reg = False
   ImportFromFile = False
   
   Estado = ED_PENDIENTE
      
   Call FirstLastMonthDay(DateSerial(lAno, lMes, 1), Dt1, Dt2)
      
   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Function
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Function
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Function
   End If
   ERR.Clear
   
   fname = FrmMain.Cm_ComDlg.Filename
   
   MousePointer = vbHourglass
   DoEvents
      
   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & fname & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
   
   'abrimos el archivo
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportFromFile = -ERR
      Exit Function
   End If
   
   If lTipoLib = LIB_COMPRAS Then
      ClasifEnt = ENT_CLIENTE
   ElseIf lTipoLib = LIB_VENTAS Then
      ClasifEnt = ENT_PROVEEDOR
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista
         Exit For
      End If
   Next i
   
   Row = i
   r = 0
   
   Grid.FlxGrid.Redraw = False
   
   Do Until EOF(Fd)
   
      Estado = ED_PENDIENTE
      DocImpExp = False
      AceptaPropIVA = False
      
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      

      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "RUT", vbTextCompare) Then
         GoTo NextRec
      End If
      
      CampoInvalido = ""
      
      'Fecha recepción/emisión
      Aux = Trim(NextField2(Buf, p))
      DtRec = ValFmtDate(Aux, False)
      If DtRec = 0 Or DtRec < Dt1 Or DtRec > Dt2 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha recepción inválida o fuera del mes en edición.")
      End If
      
      'Tipo Doc
      TipoDoc = Trim(NextField2(Buf, p))
      IdTipoDoc = FindTipoDoc(lTipoLib, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Tipo de documento inválido o no corresponde al libro en edición.")
      Else
         DocImpExp = CInt(gTipoDoc(GetTipoDoc(lTipoLib, IdTipoDoc)).DocImpExp)
         AceptaPropIVA = CInt(gTipoDoc(GetTipoDoc(lTipoLib, IdTipoDoc)).AceptaPropIVA)

      End If
         
      IdxTipoDoc = GetTipoDoc(lTipoLib, IdTipoDoc)
         
      'Del Giro
      If lTipoLib = LIB_VENTAS Then
         DelGiro = IIf(Trim(NextField2(Buf, p)) = "", 1, 0)
      End If
      
      
      'DTE
      TxtDTE = Trim(NextField2(Buf, p))
      DTE = IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)
      
      If lTipoLib = LIB_VENTAS Then
         'N° Fiscal Impresora
         NumFiscImp = Trim(NextField2(Buf, p))
         
         'N° Informe Z
         NumInfZ = Trim(NextField2(Buf, p))
         
         If TipoDoc <> "MRG" And (NumFiscImp <> "" Or NumInfZ <> "") Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "N° Fiscal Impresora y/o N° Informe Z ser cero o blanco")
            NumFiscImp = ""
            NumInfZ = ""
         End If
      End If
      
      'NumDoc
      NumDoc = Trim(NextField2(Buf, p))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "N° de documento inválido.")
      End If
      
      If lTipoLib = LIB_VENTAS Then
         'NumDocHasta
         NumDocHasta = Trim(NextField2(Buf, p))
         If Val(NumDocHasta) <> 0 Then
         
            'NumDocHasta no se permite para VPE y otros documentos
'            If gTipoDoc(IdxTipoDoc).Diminutivo = "VPE" Then
            If gTipoDoc(IdxTipoDoc).TieneNumDocHasta = VAL_NOPERMITIDO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "N° de Documento Hasta debe ser cero o blanco.")
         
            ElseIf Val(NumDocHasta) < Val(NumDoc) Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "N° de documento hasta inválido.")
            
'            ElseIf Val(NumDocHasta) = Val(NumDoc) Then
'               NumDocHasta = "0"
            End If
            
         ElseIf gTipoDoc(IdxTipoDoc).TieneNumDocHasta = VAL_OBLIGATORIO Then   'NumDocHasta = 0
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Num. Doc. Hasta debe ser mayor que cero.")
         
         End If
         
         'CantBoletas
         CantBoletas = Val(Trim(NextField2(Buf, p)))
         
         If CantBoletas > 0 Then
            If gTipoDoc(IdxTipoDoc).TieneCantBoletas = VAL_NOPERMITIDO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Cantidad de Boletas debe ser cero o blanco.")
            End If
         ElseIf gTipoDoc(IdxTipoDoc).TieneCantBoletas = VAL_OBLIGATORIO Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Cantidad de Boletas debe ser mayor que cero.")
         End If
      End If
      
      If lTipoLib = LIB_COMPRAS Then
      
         'Prop IVA
         Aux = Trim(NextField2(Buf, p))
         IdPropIVA = -1
         PropIVA = ""
         For k = 0 To UBound(gStrPropIVA)
            If Aux = Left(gStrPropIVA(k), 1) Then
               IdPropIVA = k
               PropIVA = Aux
            End If
         Next k
         
         If IdPropIVA < 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Opción Proporcionalidad de IVA inválida.")
         
         ElseIf Not AceptaPropIVA And IdPropIVA > 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Este tipo de documento no acepta Proporcionalidad de IVA.")
         End If
         
         
         'Fecha emisión
         Aux = Trim(NextField2(Buf, p))
         DtEmi = ValFmtDate(Aux, False)
         If DtEmi = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha emisión inválida.")
         End If
      
      Else
         DtEmi = DtRec
         
      End If
      
      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      ElseIf Aux = "NULO" Then
         RutEnt = "NULO"
         Estado = ED_ANULADO
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "RUT inválido")
         End If
      End If
      
      CodEnt = RutEnt
      
      NombEnt = Trim(NextField2(Buf, p))
      If NombEnt = "" And RutEnt <> "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar nombre o razón social entidad.")
      End If
      
      Descrip = Trim(NextField2(Buf, p))
      If Descrip = "NULO" Then
         Estado = ED_ANULADO
      End If
      CodSuc = Trim(NextField2(Buf, p))
      IdSucursal = 0
      Sucursal = ""
      
      If CodSuc <> "" Then
         Q1 = "SELECT IdSucursal, Descripcion FROM Sucursales WHERE Codigo ='" & CodSuc & "'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de sucursal inválido")
         Else
            IdSucursal = vFld(Rs("IdSucursal"))
            Sucursal = vFld(Rs("Descripcion"))
         End If
         
         Call CloseRs(Rs)
      End If
      
      'Valores
      Afecto = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta Afecto
      AuxCodCtaAfecto = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      'Afecto sólo si corresponde
      If (Afecto <> 0 Or AuxCodCtaAfecto <> "") And gTipoDoc(IdxTipoDoc).TieneAfecto = VAL_NOPERMITIDO Then
         CampoInvalido = CampoInvalido & "," & "Exento"
         Call AddLogImp(lFNameLogImp, fname, l, "Valor Afecto o Código de Cuenta Afecto no permitido.")
      End If
      
      'Exento
      Exento = vFmt(Trim(NextField2(Buf, p)))
      
      'código cuenta Exento
      AuxCodCtaExento = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      'Exento sólo si corresponde
      If (Exento <> 0 Or AuxCodCtaExento <> "") And gTipoDoc(IdxTipoDoc).TieneExento = VAL_NOPERMITIDO Then
         CampoInvalido = CampoInvalido & "," & "Exento"
         Call AddLogImp(lFNameLogImp, fname, l, "Valor Exento o Código de Cuenta Exento no permitido.")
      End If
      
      IVA = vFmt(Trim(NextField2(Buf, p)))
      
      OtroImp = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta OtrosImp
      AuxCodCtaOtroImp = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      If TipoDoc = "VPE" And OtroImp <> 0 Then
         CampoInvalido = CampoInvalido & "," & "Otro Impuesto"
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Otro Impuesto inválido.")
      End If

      Total = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta Total
      AuxCodCtaTotal = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      If gTipoDoc(IdxTipoDoc).EsRebaja Then
         Afecto = Abs(Afecto) * -1
         Exento = Abs(Exento) * -1
         IVA = Abs(IVA) * -1
         Total = Abs(Total) * -1
         OtroImp = OtroImp * -1   'para permitir ingresar valores negativos en los otros impuestos. Al ser rebaja y el usuario ingresa un valor negativo, queda positivo
      
      Else
         If Afecto < 0 Or Exento < 0 Or IVA < 0 Or Total < 0 Then
            CampoInvalido = CampoInvalido & "," & "Afecto, Exento, IVA, Total"
            Call AddLogImp(lFNameLogImp, fname, l, "Valor de Afecto, Exento, IVA y/o Total inválido.")
         End If
                     
      End If
      
      If lTipoLib = LIB_VENTAS Then
         'Ventas Acum Informe Z
         VentasAcumInfZ = vFmt(Trim(NextField2(Buf, p)))
                  
         If TipoDoc <> "MRG" And VentasAcumInfZ <> 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Ventas Acum. Informe Z inválido")
            VentasAcumInfZ = 0
         End If
         
      End If

      
      'Fecha Vencim
      Aux = Trim(NextField2(Buf, p))
      DtVenc = 0
      If Aux <> "" Then
         DtVenc = ValFmtDate(Aux, False)
         If DtVenc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha vencimiento inválida.")
         End If
      Else
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar Fecha Vencimiento.")
      End If
            
      NumInterno = vFmt(Trim(NextField2(Buf, p)))
      CodANeg = Trim(NextField2(Buf, p))
      CodCCosto = Trim(NextField2(Buf, p))
      
      IdANeg = GetAreaNegocio(CodANeg)
      IdCCosto = GetCentroCosto(CodCCosto)
      
      If CodANeg <> "" And IdANeg = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         pCCosto = p
         Call AddLogImp(lFNameLogImp, fname, l, "Área de Negocio inválida.")
      End If
           
      If CodCCosto <> "" And IdCCosto = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         pANeg = p
         Call AddLogImp(lFNameLogImp, fname, l, "Centro de Gestión inválido.")
      End If
      
      'códigos cuentas
      
      NomCta = ""
      
      If AuxCodCtaAfecto <> "" Then
         If Afecto = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Afecto debe ser cero o blanco")
         
         Else
            AuxIdCtaAfecto = GetIdCuenta(NomCta, AuxCodCtaAfecto, AuxDescCtaAfecto, UltNivel)
            If AuxIdCtaAfecto <= 0 Or Not UltNivel Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Afecto inválido")
            End If
         End If
      Else
         AuxIdCtaAfecto = 0
         AuxDescCtaAfecto = ""
      End If
      
      NomCta = ""
      
      If AuxCodCtaExento <> "" Then
         If Exento = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Exento debe ser cero o blanco")
         Else
            AuxIdCtaExento = GetIdCuenta(NomCta, AuxCodCtaExento, AuxDescCtaExento, UltNivel)
            If AuxIdCtaExento <= 0 Or Not UltNivel Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Exento inválido")
            End If
         End If
      Else
         AuxIdCtaExento = 0
         AuxDescCtaExento = ""
      End If
            
      NomCta = ""
      
      If AuxCodCtaTotal <> "" Then
         AuxIdCtaTotal = GetIdCuenta(NomCta, AuxCodCtaTotal, AuxDescCtaTotal, UltNivel)
         If AuxIdCtaTotal <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Total inválido")
         End If
      Else
         AuxIdCtaTotal = 0
         AuxDescCtaTotal = ""
      End If
            
      NomCta = ""
    
      If AuxCodCtaOtroImp <> "" Then
         AuxIdCtaOtroImp = GetIdCuenta(NomCta, AuxCodCtaOtroImp, AuxDescCtaOtroImp, UltNivel)
         If AuxIdCtaOtroImp <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta Otros Impuestos inválido")
         End If
      Else
         AuxIdCtaOtroImp = 0
         AuxDescCtaOtroImp = ""
      End If
            
      NomCta = ""
    
      'Cuentas Contables Default
                  
      If Exento <> 0 Then
         If AuxIdCtaExento > 0 Then
            IdCtaExento = AuxIdCtaExento
            CodCtaExento = FmtCodCuenta(AuxCodCtaExento)
            DescCtaExento = AuxDescCtaExento
         Else
            IdCtaExento = lCtaExento.id
            CodCtaExento = FmtCodCuenta(lCtaExento.Codigo)
            DescCtaExento = lCtaExento.Descripcion
         End If
      Else
         IdCtaExento = 0
         CodCtaExento = ""
         DescCtaExento = ""
      End If
         
      
      If Afecto <> 0 Then
         If AuxIdCtaAfecto > 0 Then
            IdCtaAfecto = AuxIdCtaAfecto
            CodCtaAfecto = FmtCodCuenta(AuxCodCtaAfecto)
            DescCtaAfecto = AuxDescCtaAfecto
         Else
            IdCtaAfecto = lCtaAfecto.id
            CodCtaAfecto = FmtCodCuenta(lCtaAfecto.Codigo)
            DescCtaAfecto = lCtaAfecto.Descripcion
         End If
      Else
         IdCtaAfecto = 0
         CodCtaAfecto = 0
         DescCtaAfecto = 0
      End If
                     
                  
      If AuxIdCtaTotal > 0 Then
         IdCtaTotal = AuxIdCtaTotal
         CodCtaTotal = FmtCodCuenta(AuxCodCtaTotal)
         DescCtaTotal = AuxDescCtaTotal
      Else
         IdCtaTotal = lCtaTotal.id
         CodCtaTotal = FmtCodCuenta(lCtaTotal.Codigo)
         DescCtaTotal = lCtaTotal.Descripcion
      End If
            
      If IVA <> 0 Then
         IdCtaIVA = lIdCuentaIVA
      Else
         IdCtaIVA = 0
      End If
      
      If OtroImp <> 0 Then
         If AuxIdCtaOtroImp > 0 Then
            IdCtaOtroImp = AuxIdCtaOtroImp
'            CodCtaOtroImp = FmtCodCuenta(AuxCodCtaOtroImp)
'            DescCtaOtroImp = AuxDescCtaOtroImp
         Else
            'si es factura de compras, nota de crédito de fac. compras o nota de débito de fac. compras, se pone la cuenta al revés
            If TipoDoc = "FCC" Or TipoDoc = "NCF" Or TipoDoc = "NDF" Or TipoDoc = "FCV" Then
               IdCtaOtroImp = lIdCuentaOtrosImpFacCompra
            Else
               IdCtaOtroImp = lIdCuentaOtrosImp
            End If
         End If
      Else
         IdCtaOtroImp = 0
      End If
      
      'validamos si ingresó area de negocio y centro de costo si corresponde
      AtribANeg = GetAtribCuenta(IdCtaAfecto, ATRIB_AREANEG) Or GetAtribCuenta(IdCtaExento, ATRIB_AREANEG) Or GetAtribCuenta(IdCtaTotal, ATRIB_AREANEG)
      
      AtribCCosto = GetAtribCuenta(IdCtaAfecto, ATRIB_CCOSTO) Or GetAtribCuenta(IdCtaExento, ATRIB_CCOSTO) Or GetAtribCuenta(IdCtaTotal, ATRIB_CCOSTO)
      
      If AtribANeg And IdANeg = 0 Then
         CampoInvalido = CampoInvalido & "," & pANeg
         Call AddLogImp(lFNameLogImp, fname, l, "Falta indicar Área de Negocio")
      End If
      
      If AtribCCosto And IdCCosto = 0 Then
         CampoInvalido = CampoInvalido & "," & pCCosto
         Call AddLogImp(lFNameLogImp, fname, l, "Falta indicar Centro de Costo")
      End If
      
      'si no hay errores y la entidad no existe, la insertamos
      
      If CampoInvalido = "" Then
      
         If RutEnt <> "" And RutEnt <> "NULO" Then
      
            Q1 = "SELECT IdEntidad, Nombre FROM Entidades WHERE Rut = '" & RutEnt & "'"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdEnt = vFld(Rs("IdEntidad"))
               NombEnt = vFld(Rs("Nombre"))
            End If
            Call CloseRs(Rs)
            
            If IdEnt = 0 Then  'no existe
         
               'insertamos la nueva entidad
               
'               Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
'               Rs.AddNew
'
'               IdEnt = Rs("IdEntidad")
'               Rs("RUT") = RutEnt
'               Rs("Codigo") = CodEnt
'               Rs("Nombre") = NombEnt
'               Rs("Clasif" & ClasifEnt) = 1
'
'               Rs.Update
'               Rs.Close


               FldArray(0).FldName = "RUT"
               FldArray(0).FldValue = RutEnt
               FldArray(0).FldIsNum = False
               
               FldArray(1).FldName = "Codigo"
               FldArray(1).FldValue = CodEnt
               FldArray(1).FldIsNum = False
                     
               FldArray(2).FldName = "Nombre"
               FldArray(2).FldValue = NombEnt
               FldArray(2).FldIsNum = False
                           
               FldArray(3).FldName = "IdEmpresa"
               FldArray(3).FldValue = gEmpresa.id
               FldArray(3).FldIsNum = True
                     
               IdEnt = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)
               
'               Q1 = "UPDATE Entidades SET "
'               Q1 = Q1 & "  Codigo = '" & CodEnt & "'"
'               Q1 = Q1 & ", Nombre = '" & NombEnt & "'"
'               Q1 = Q1 & ", Clasif" & ClasifEnt & " = 1"
'               Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
'               Q1 = Q1 & " WHERE IdEntidad = " & IdEnt
'
'               Call ExecSQL(DbMain, Q1)
               
            End If
            
         End If
         
         'para que no ingrese documento que ya existen en la grilla
         
         Q1 = "SELECT IdDoc FROM Documento "
         Q1 = Q1 & " WHERE TipoLib = " & lTipoLib & " AND TipoDoc = " & IdTipoDoc
         Q1 = Q1 & " AND NumDoc = '" & NumDoc & "'"
         Q1 = Q1 & " AND IdEntidad = " & IdEnt
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = True Then 'documento no existe
      
         'si no hay errores, ingresamos el registro a la grilla
         Grid.TextMatrix(Row, C_NUMLIN) = vFmt(Grid.TextMatrix(Row - 1, C_NUMLIN)) + 1
         Grid.TextMatrix(Row, C_FECHA) = Day(DtRec)
         Grid.TextMatrix(Row, C_IDTIPODOC) = IdTipoDoc
         Grid.TextMatrix(Row, C_TIPODOC) = TipoDoc
         Grid.TextMatrix(Row, C_DOCIMPEXP) = DocImpExp
         
         If lTipoLib = LIB_VENTAS Then
            Grid.TextMatrix(Row, C_GIRO) = IIf(DelGiro = 0, "No", "")
         Else
            Grid.TextMatrix(Row, C_GIRO) = ""
         End If
         
         Grid.TextMatrix(Row, C_DTE) = IIf(DTE <> 0, "x", "")
                  
         If lTipoLib = LIB_VENTAS Then
            Grid.TextMatrix(Row, C_NUMFISCIMPR) = NumFiscImp
            Grid.TextMatrix(Row, C_NUMINFORMEZ) = NumInfZ
            Grid.TextMatrix(Row, C_VENTASACUM) = Format(VentasAcumInfZ, NUMFMT)
         End If
         
         Grid.TextMatrix(Row, C_NUMDOC) = NumDoc
         Grid.TextMatrix(Row, C_NUMDOCHASTA) = NumDocHasta
         
         
         If lTipoLib = LIB_VENTAS Then
            If Val(NumDocHasta) > 0 And Val(NumDocHasta) >= Val(NumDoc) Then
               Grid.TextMatrix(Row, C_CANTBOLETAS) = Format(Val(NumDocHasta) - Val(NumDoc), NUMFMT) + 1
            ElseIf CantBoletas > 0 Then
               Grid.TextMatrix(Row, C_CANTBOLETAS) = Format(CantBoletas, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_CANTBOLETAS) = ""
            End If
         End If
         
         If lTipoLib = LIB_COMPRAS Then
            Grid.TextMatrix(Row, C_IDPROPIVA) = IdPropIVA
            Grid.TextMatrix(Row, C_PROPIVA) = PropIVA
         End If
         
         Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(DtEmi, SDATEFMT)
         Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = DtEmi
         Grid.TextMatrix(Row, C_RUT) = FmtRut(RutEnt)
         Grid.TextMatrix(Row, C_NOMBRE) = NombEnt
         Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
         Grid.TextMatrix(Row, C_DESCRIP) = Descrip
         Grid.TextMatrix(Row, C_IDSUCURSAL) = IdSucursal
         Grid.TextMatrix(Row, C_SUCURSAL) = Sucursal
         
         Grid.TextMatrix(Row, C_AFECTO) = Format(Afecto, NUMFMT)
         Grid.TextMatrix(Row, C_AF_IDCUENTA) = IdCtaAfecto
         Grid.TextMatrix(Row, C_AF_CODCUENTA) = CodCtaAfecto
         Grid.TextMatrix(Row, C_AF_CUENTA) = DescCtaAfecto
         
         Grid.TextMatrix(Row, C_EXENTO) = Format(Exento, NUMFMT)
         Grid.TextMatrix(Row, C_EX_IDCUENTA) = IdCtaExento
         Grid.TextMatrix(Row, C_EX_CODCUENTA) = CodCtaExento
         Grid.TextMatrix(Row, C_EX_CUENTA) = DescCtaExento
         
         Grid.TextMatrix(Row, C_IVA) = Format(IVA, NUMFMT)
         Grid.TextMatrix(Row, C_IVA_IDCUENTA) = IdCtaIVA
         
         Grid.TextMatrix(Row, C_OTROIMP) = Format(OtroImp, NUMFMT)
         Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = IdCtaOtroImp
         
         Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NUMFMT)
         Grid.TextMatrix(Row, C_TOT_IDCUENTA) = IdCtaTotal
         Grid.TextMatrix(Row, C_TOT_CODCUENTA) = CodCtaTotal
         Grid.TextMatrix(Row, C_TOT_CUENTA) = DescCtaTotal
         
         If IdANeg > 0 Or IdCCosto > 0 Then
            Grid.TextMatrix(Row, C_IDANEG_CCOSTO) = IdANeg & "-" & IdCCosto
         End If
         
         Grid.TextMatrix(Row, C_DETALLE) = TX_DETALLE
         Grid.TextMatrix(Row, C_FECHAVENC) = IIf(DtVenc <> 0, Format(DtVenc, SDATEFMT), "")
         Grid.TextMatrix(Row, C_LNGFECHAVENC) = DtVenc
         Grid.TextMatrix(Row, C_CORRINTERNO) = NumInterno
         Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO
         Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(Estado)
         Grid.TextMatrix(Row, C_IDESTADO) = Estado
         Grid.TextMatrix(Row, C_USUARIO) = gUsuario.Nombre
         Grid.TextMatrix(Row, C_UPDATE) = FGR_I
                  
                 
         If EsIngresoTotal(Row) Then
            Dim Value As String
            Value = Total
            Call CalcIngresoTotal(Row, C_TOTAL, Value)
            '3071158
            'Call CalcTot
            '3071158

         Else
            Call CalcTotRow(Row, False)   'no recalcula IVA, deja el que viene
            
         End If
         
         
         Row = Row + 1
         
         r = r + 1
         
        '3071158
         Call CloseRs(Rs)
         
          End If
        '3071158
         
         If gDbType = SQL_ACCESS Then
            If r = 3000 Then
                Exit Do
            End If
         Else
            If r = 3000 Then
                Mayor3000reg = True
                Exit Do
            End If
         End If
      
         Grid.rows = Grid.rows + 1
         
'         If FGrChkMaxSize(Grid) = True Then
'            Exit loop
'         End If
    
   

      Else
         NRecErroneos = NRecErroneos + 1
         
         
      End If
      
NextRec:
   Loop

   Close #Fd
   
   '3071158
   Call CalcTot
   '3071158
   
   Grid.FlxGrid.Redraw = True
   
   Me.MousePointer = vbDefault
   
   If NRecErroneos = 0 Then
      If r = 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg = False Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
      End If
   
   Else
      If NRecErroneos > 1 Then
         StrNRecErroneos = "- Se encontraron " & NRecErroneos & " registros con errores en el archivo."
      Else
         StrNRecErroneos = "- Se encontró " & NRecErroneos & " registro con errores en el archivo."
      End If
   
      If r = 1 Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg = False Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If


   ImportFromFile = True
   
End Function

'agrega las columnas de detalle de otros impuestos (sólo las que usa el usuario) y los títulos correspondientes
Private Sub DetOtrosImp(ByVal TipoLib As Integer, Optional ByVal Where As String = "")
   Dim Q1 As String
   Dim Rs As Recordset
   '3031651
   Dim i, T As Long ' T As integer
   '3031651
   Dim PrimerDetalle As Long
   Dim TituloIva As String
   
   For i = C_INIDETOTROIMP To C_ENDDETOTROIMP
      Grid.TextMatrix(0, i) = ""
      Grid.TextMatrix(1, i) = ""
      Grid.ColWidth(i) = 0
   Next i
   
   If Ch_ViewDetOtrosImp = 0 Then
   
      If lTipoLib = LIB_COMPRAS Then
         Grid.TextMatrix(1, C_IVA) = ""
      ElseIf lTipoLib = LIB_VENTAS Then
         Grid.TextMatrix(1, C_IVA) = ""
      End If

      Exit Sub
   End If
   
   If TipoLib = LIB_COMPRAS Then
      PrimerDetalle = LIBCOMPRAS_IVAIRREC
   Else
      PrimerDetalle = LIBVENTAS_REBAJA65
   End If
   
   
   Q1 = "SELECT DISTINCT TipoValor.idTValor, TipoValor.Codigo, TipoValor.Valor, TipoValor.Tit1, TipoValor.Tit2"
   Q1 = Q1 & " FROM (((Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc) "
   Q1 = Q1 & " INNER JOIN TipoValor ON (MovDocumento.IdTipoValLib = TipoValor.Codigo) AND (Documento.TipoLib = TipoValor.TipoLib))"
   Q1 = Q1 & " INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   
   If Where = "" Then
      Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib
   Else
      Q1 = Q1 & " WHERE " & Where
   End If
   
   Q1 = Q1 & " AND TipoValor.Codigo > " & LIBCOMPRAS_OTROSIMP
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY TipoValor.Codigo "

   Set Rs = OpenRs(DbMain, Q1)
   TituloIva = ""
   T = 0
   Do While Not Rs.EOF
   If gDbType = SQL_SERVER Then
      Grid.TextMatrix(0, C_INIDETOTROIMP + vFld(Rs("Codigo")) - PrimerDetalle) = vFld(Rs("Tit1"))
      Grid.TextMatrix(1, C_INIDETOTROIMP + vFld(Rs("Codigo")) - PrimerDetalle) = vFld(Rs("Tit2"))
      Grid.ColWidth(C_INIDETOTROIMP + vFld(Rs("Codigo")) - PrimerDetalle) = 1200
    Else
      'Grid.TextMatrix(0, 63) = vFld(Rs("Tit1"))
      'Grid.TextMatrix(1, 63) = vFld(Rs("Tit2"))
      T = T + 1
      Grid.ColWidth(63) = 1200
      '2755620 se crea variable para concatenar el titulo en caso que sea mas de 1
      TituloIva = TituloIva & vFld(Rs("Tit1")) & " " & vFld(Rs("Tit2")) & "/"
    End If
      
      Rs.MoveNext
   Loop
   
   ''2755620 se asigna el valor a la del string creado a la grilla
   If TituloIva <> "" Then
   If T > 1 Then
    Grid.ColWidth(63) = 1200 + (500 * T)
   End If
    Grid.TextMatrix(1, 63) = Mid(TituloIva, 1, Len(TituloIva) - 1)
   End If
   Call CloseRs(Rs)
   
   If lTipoLib = LIB_COMPRAS Then
      Grid.TextMatrix(1, C_IVA) = "Crédito Fiscal"
   ElseIf lTipoLib = LIB_VENTAS Then
      Grid.TextMatrix(1, C_IVA) = "Débito Fiscal"
   End If

   Call FGrTotales(Grid, GridTot)

End Sub
'llena los datos de otros impuest0s para el documento (registro) indicado
'los valores se calculan restando Debe - Haber manteniendo el signo, es decir sin hacer Abs
'por lo tanto el valor retornado por esta función es el valor obtenido al sumar todos loe detalled de Debe - Haber con los signos correspondientes
Private Function FillDetOtroImp(ByVal Row As Integer, ByVal TipoLib As Integer, ByVal IdDoc As Long, ByVal EsRebaja As Boolean, IVAActFijo As Double, IVAIrrecuperable As Double) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PrimerDetalle As Long
   Dim TotOtrosImp As Double
   Dim Col As Integer
   Dim Valor As Double
   
   FillDetOtroImp = 0
   IVAActFijo = 0
   IVAIrrecuperable = 0
   
   
   If Ch_ViewDetOtrosImp = 0 Then
      Exit Function
   End If
   
   If IdDoc <= 0 Then
      Exit Function
   End If
   
   If TipoLib = LIB_COMPRAS Then
      PrimerDetalle = LIBCOMPRAS_IVAIRREC
   Else
      PrimerDetalle = LIBVENTAS_REBAJA65
   End If
   
   Q1 = "SELECT IdTipoValLib, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber FROM MovDocumento "
   Q1 = Q1 & " WHERE IdDoc = " & IdDoc & " AND IdTipoValLib >" & LIBCOMPRAS_OTROSIMP
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdTipoValLib"

'Q1 = "SELECT IdTipoValLib, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber FROM MovDocumento as md , documento as d "
'   Q1 = Q1 & " WHERE md.iddoc = d.iddoc AND md.IdDoc = " & IdDoc & " AND md.IdTipoValLib >" & LIBCOMPRAS_OTROSIMP
'   Q1 = Q1 & " AND md.IdEmpresa = " & gEmpresa.id & " AND md.Ano = " & gEmpresa.Ano & " AND d.otroimp > 0 "
'   Q1 = Q1 & " GROUP BY md.IdTipoValLib"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   TotOtrosImp = 0
   IVAActFijo = 0
   IVAIrrecuperable = 0
   
   Do While Not Rs.EOF
   
   If gDbType = SQL_SERVER Then
      Col = C_INIDETOTROIMP + vFld(Rs("IdTipoValLib")) - PrimerDetalle
    Else
     Col = 63
   End If
'      Valor = Abs(vFld(Rs("SumDebe")) - vFld(Rs("SumHaber")))             'FCA 28 nov 2017
'      Valor = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
      
'      If TipoLib = LIB_COMPRAS Then
'
'         If EsRebaja Then
'            Grid.TextMatrix(Row, Col) = Format(Valor * -1, NEGNUMFMT)
'         Else
'            Grid.TextMatrix(Row, Col) = Format(Valor, NEGNUMFMT)
'         End If
'
'      Else
'         If EsRebaja Then
'            Grid.TextMatrix(Row, Col) = Format(Valor, NEGNUMFMT)
'         Else
'            Grid.TextMatrix(Row, Col) = Format(Valor * -1, NEGNUMFMT)
'         End If
'      End If
                      
'      Grid.TextMatrix(Row, Col) = Format(Valor, NEGNUMFMT)
                     
      If TipoLib = LIB_COMPRAS Then
         Valor = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))             'FCA 4 sep 2018
         
         If EsRebaja Then
            Grid.TextMatrix(Row, Col) = Format(Valor * -1, NEGNUMFMT)
         End If
         
         Grid.TextMatrix(Row, Col) = Format(Valor, NEGNUMFMT)
         
         If vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO Then
            IVAActFijo = IVAActFijo + vFmt(Grid.TextMatrix(Row, Col))
         
         ElseIf vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC1 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC2 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC3 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC4 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC9 Then
            IVAIrrecuperable = IVAIrrecuperable + vFmt(Grid.TextMatrix(Row, Col))
         
         Else
            TotOtrosImp = TotOtrosImp + vFmt(Grid.TextMatrix(Row, Col))
         
         End If
     
      Else    'LIB_VENTAS
         Valor = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))             'FCA 4 sep 2018
         
         If EsRebaja Then
            Grid.TextMatrix(Row, Col) = Format(Valor * -1, NEGNUMFMT)
         End If
        
         Grid.TextMatrix(Row, Col) = Format(Valor, NEGNUMFMT)
         
         TotOtrosImp = TotOtrosImp + vFmt(Grid.TextMatrix(Row, Col))
         
      End If
      
      
      Rs.MoveNext
   
   Loop
   
   Call CloseRs(Rs)
   
   FillDetOtroImp = TotOtrosImp
   
End Function

Private Function CalcPropIVALibro() As Boolean
   Dim i As Integer
   Dim Mes As Integer
   
   CalcPropIVALibro = False
   Mes = CbItemData(Cb_Mes)

'   If lTipoLib = LIB_COMPRAS And lHayPropIVA And gValPropIVA(Mes).CalcProp Then     'habíamos agregado que sólo se preguntara si correspde calcular proporcionalidad de IVA en ese mes, pero para el caso de N o T igual se debe preguntar, aún cuando no corresponda proporcinalidad de IVA en ese mes. Esto hace muy compleja la condición por lo cual se pregunta siempre. (22 oct 2018 Joshua Nicolás Catrin)
   If lTipoLib = LIB_COMPRAS And lHayPropIVA Then
      If MsgBox1("¿Desea aplicar la Proporcionalidad del IVA para este mes?" & vbCrLf & vbCrLf & "Recuerde que este proceso sólo se realiza para los documentos" & vbCrLf & "en estado PENDIENTE y que tienen seleccionada alguna opción" & vbCrLf & "(T, P o N) en la columna Prop. IVA." & vbCrLf & vbCrLf & "ATENCIÓN: Este cálculo afecta a la columna Otros Impuestos," & vbCrLf & "al Resumen de los Libros y las respectivas impresiones.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
      
      If gCtasBas.IdCtaIVAIrrec = 0 Then
         MsgBox "ATENCIÓN:  Falta definir la cuenta de IVA Irrecuperable." & vbCrLf & vbCrLf & "Utilice la opción: " & vbCrLf & vbCrLf & "Configuración >> Configuración Inicial >> Definir Cuentas Básicas >> Otros", vbExclamation + vbOKOnly
         Exit Function
      End If
      
      Me.MousePointer = vbHourglass
      
'      For i = Grid.FixedRows To Grid.Rows - 1
'
'         If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista de mov.
'            Exit For
'         End If
'
'         If Grid.RowHeight(i) > 0 And Val(Grid.TextMatrix(i, C_IDESTADO)) = ED_PENDIENTE And Val(Grid.TextMatrix(i, C_IDPROPIVA)) <> 0 Then
'            Call PropIVA_UpdateMovDoc(Val(Grid.TextMatrix(i, C_IDDOC)))
'         End If
'
'      Next i

      Call PropIVA_UpdateMovDoc(Mes)

      Me.MousePointer = vbDefault
      
      'MsgBox1 "El proceso de aplicación de Proporcionalidad de IVA Crédito Fiscal ha finalizado.", vbInformation
   End If
   
   CalcPropIVALibro = True

End Function
Private Sub CalcIngresoTotal(ByVal Row As Integer, ByVal Col As Integer, Value As String)
   Dim IdxDoc As Long
   Dim IVA As Double, Exento As Double, Afecto As Double
         
   IdxDoc = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   
   If gTipoDoc(IdxDoc).EsRebaja Then
      Value = Format(vFmt(Value) * -1, NEGNUMFMT)
   Else
      Value = Format(vFmt(Value), NEGNUMFMT)
   End If
   
   Grid.TextMatrix(Row, Col) = Value
   
   If Val(Grid.TextMatrix(Row, C_TOT_IDCUENTA)) = 0 Then
      Grid.TextMatrix(Row, C_TOT_IDCUENTA) = lCtaTotal.id
      Grid.TextMatrix(Row, C_TOT_CODCUENTA) = FmtCodCuenta(lCtaTotal.Codigo)
      Grid.TextMatrix(Row, C_TOT_CUENTA) = lCtaTotal.Descripcion
      
   End If
   
'   If InStr(LCase(gTipoDoc(IdxDoc).Nombre), "exent") > 0 Or InStr(LCase(gTipoDoc(IdxDoc).Nombre), "sin") > 0 Then
   If gTipoDoc(IdxDoc).TieneExento <> 0 Then
      
      'es exenta o venta SIN documento
      If Val(Grid.TextMatrix(Row, C_EX_IDCUENTA)) = 0 Then
         Grid.TextMatrix(Row, C_EX_IDCUENTA) = lCtaExento.id
         Grid.TextMatrix(Row, C_EX_CODCUENTA) = FmtCodCuenta(lCtaExento.Codigo)
         Grid.TextMatrix(Row, C_EX_CUENTA) = lCtaExento.Descripcion
         
         Grid.TextMatrix(Row, C_AF_IDCUENTA) = ""
         Grid.TextMatrix(Row, C_AF_CODCUENTA) = ""
         Grid.TextMatrix(Row, C_AF_CUENTA) = ""
      End If
      
      IVA = 0
      Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
      
      Exento = vFmt(Grid.TextMatrix(Row, C_TOTAL))
      Grid.TextMatrix(Row, C_EXENTO) = Format(Exento, NEGNUMFMT)
      Grid.TextMatrix(Row, C_AFECTO) = Format(0, NEGNUMFMT)
                     
   Else   'es afecta
   
      If Val(Grid.TextMatrix(Row, C_AF_IDCUENTA)) = 0 Then
         Grid.TextMatrix(Row, C_AF_IDCUENTA) = lCtaAfecto.id
         Grid.TextMatrix(Row, C_AF_CODCUENTA) = FmtCodCuenta(lCtaAfecto.Codigo)
         Grid.TextMatrix(Row, C_AF_CUENTA) = lCtaAfecto.Descripcion
         
         Grid.TextMatrix(Row, C_EX_IDCUENTA) = ""
         Grid.TextMatrix(Row, C_EX_CODCUENTA) = ""
         Grid.TextMatrix(Row, C_EX_CUENTA) = ""
      End If
   
      IVA = Round(vFmt(Grid.TextMatrix(Row, C_TOTAL)) * (1 - (1 / (1 + gIVA))))
      Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
      
      Afecto = vFmt(Grid.TextMatrix(Row, C_TOTAL)) - IVA
      Grid.TextMatrix(Row, C_AFECTO) = Format(Afecto, NEGNUMFMT)
      Grid.TextMatrix(Row, C_EXENTO) = Format(0, NEGNUMFMT)
      
   End If
   
   If vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Then
      Grid.TextMatrix(Row, C_IVA_IDCUENTA) = lIdCuentaIVA
   Else
      Grid.TextMatrix(Row, C_IVA_IDCUENTA) = 0
   End If
   
   'Call CalcTot
   
End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Sub AddRutDefault(Row)
   Dim IdEnt As Long
   Dim Nombre As String
   Dim NotValidRut As Boolean
   
   If Grid.TextMatrix(Row, C_TIPODOC) = "IMP" And Grid.TextMatrix(Row, C_RUT) = "" Then
      IdEnt = GetIdEntidad(ENTIMP_RUT, Nombre, NotValidRut)
      If IdEnt > 0 Then
         Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
         Grid.TextMatrix(Row, C_RUT) = ENTIMP_RUT
         Grid.TextMatrix(Row, C_NOMBRE) = Nombre
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
      Else
         MsgBox1 "No se encuentra definida la entidad especial para Formulario de Importaciones:" & vbCrLf & "RUT: " & ENTIMP_RUT & vbCrLf & "Razón social: " & ENTIMP_RSOCIAL & vbCrLf & vbCrLf & "El sistema la define automáticamente. Si no está por alguna razón, vuelva a ingresarla directamente en la lista de entidades.", vbExclamation
      End If
   End If
  
End Sub

'3217885
Private Sub CentralizarFull()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim EsRebaja As Boolean
   Dim ValOtros As Double
   Dim NotValidRut As Boolean
   Dim EditEnable As Boolean
   Dim TotOtrosImp As Double
   Dim IVAActFijo As Double
   Dim IVAIrrecuperable As Double
    Dim StrIdDoc As String
   Dim idcomp As Long
   
   TotOtrosImp = 0
   IVAActFijo = 0
   IVAIrrecuperable = 0
      
   
   If CbItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
      If CbItemData(Cb_TipoDoc) > 0 Then
         Where = Where & " AND Documento.TipoDoc = " & CbItemData(Cb_TipoDoc)
      End If
      
      If CbItemData(Cb_Estado) > 0 Then
         Where = Where & " AND Documento.Estado = " & CbItemData(Cb_Estado)
      End If
      
      If Val(Tx_NumDoc) <> 0 Then
         Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
      End If
      
      If Val(Tx_NumDocAsoc) <> 0 Then
         Where = Where & " AND Documento.NumDocAsoc = '" & Trim(Tx_NumDocAsoc) & "'"
      End If
      
      If Trim(Tx_Rut) <> "" Then
         IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
         If IdEnt > 0 Then
            Where = Where & " AND Documento.IdEntidad = " & IdEnt
         Else
            Tx_Rut = ""
            Cb_Entidad.ListIndex = 0
            Cb_Nombre.ListIndex = 0
         End If
      
      End If
      
      If Trim(Tx_Descrip) <> "" Then
         Where = Where & " AND " & GenLike(DbMain, Tx_Descrip, "Documento.Descrip", 3)
      End If
      
      If vFmt(Tx_Valor) <> 0 Then
         Where = Where & " AND Documento.Afecto = " & vFmt(Tx_Valor)
      End If
      
   End If
   
   
   If CbItemData(Cb_Sucursal) > 0 Then
      Where = Where & " AND Documento.IdSucursal = " & CbItemData(Cb_Sucursal)
   End If
   
   If CbItemData(Cb_DTE) > 0 Then
      Where = Where & " AND Documento.DTE <> 0 "
   ElseIf CbItemData(Cb_DTE) < 0 Then
      Where = Where & " AND Documento.DTE = 0 "
   End If
      
   If Ch_EsSupermercado > 0 Then
      Where = Where & " AND Entidades.EsSupermercado <> 0 "
   End If
      

   Q1 = "SELECT IdDoc, Documento.TipoDoc, DTE, Documento.Giro, NumDoc, CorrInterno, NumDocHasta, Documento.IdEntidad, Documento.RutEntidad, Documento.MovEdited, Documento.PropIVA,"
   Q1 = Q1 & " Documento.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre, FEmision, FEmisionOri, FVenc, Exento, Afecto, IVA, Documento.IdCompCent, Documento.IdCompPago, "
   Q1 = Q1 & " OtroImp, OtrosVal, Total, Descrip, Documento.Estado, Documento.IdANegCCosto, IdCuentaExento, Usuarios.Usuario, Cuentas1.Codigo as CodCtaEx, "
   Q1 = Q1 & " Cuentas1.Descripcion as DescCtaEx, Cuentas1.Atrib" & ATRIB_ACTIVOFIJO & " as ActFijoCtaEx, IdCuentaAfecto, "
   Q1 = Q1 & " Cuentas2.Codigo as CodCtaAf, Cuentas2.Descripcion as DescCtaAf, Cuentas2.Atrib" & ATRIB_ACTIVOFIJO & " as ActFijoCtaAf, "
   Q1 = Q1 & " IdCuentaIVA, IdCuentaOtroImp, IdCuentaTotal, Cuentas3.Codigo as CodCtaTot,Cuentas3.Descripcion as DescCtaTot, "
   Q1 = Q1 & " Documento.IdSucursal, Sucursales.Descripcion as DescSucursal, EsSupermercado, Entidades.EntRelacionada, "
   Q1 = Q1 & " iif(" & SqlMonthLng("FEmisionOri") & "= " & CbItemData(Cb_Mes) & ",0,1) as MesActual, "
   Q1 = Q1 & " NumFiscImpr, NumInformeZ, CantBoletas, VentasAcumInfZ, IdDocAsoc, FExported "
   Q1 = Q1 & " FROM (((((((Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib=TipoDocs.TipoLib AND Documento.TipoDOC=TipoDocs.TipoDoc)"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON Documento.IdUsuario = Usuarios.IdUsuario )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas1") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaAfecto = Cuentas2.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas2") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas3 ON Documento.IdCuentaTotal = Cuentas3.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas3") & " )"
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Sucursales", True, True) & " )"
   Q1 = Q1 & " WHERE Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND Documento.ESTADO in (" & ED_PENDIENTE & "," & ED_PAGADO & ")"
   Q1 = Q1 & Where
   Q1 = Q1 & lWhere
   
   If lOrdenSel = C_FECHAEMIORI Then
      Q1 = Q1 & " ORDER BY IIf(" & SqlMonthLng("FEmisionOri") & "= " & CbItemData(Cb_Mes) & ", 0, 1), " & lOrdenGr(lOrdenSel)
   Else
      Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   End If
   
'   If Row = 0 Then
'      Q1 = Q1 & SqlPaging(gDbType, lClsPaging.CurReg - 1, gPageNumReg)
'   End If
      
   lCurWhere = " Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   lCurWhere = lCurWhere & Where & lWhere
   
   Call DetOtrosImp(lTipoLib, lCurWhere)
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF = False Then
     MsgBox1 "No hay documentos para centralizar.", vbExclamation + vbOKOnly
            Exit Sub
   End If
   
   Do While Rs.EOF = False
    
    StrIdDoc = StrIdDoc & "," & vFld(Rs("IdDoc"))
    
    
    Rs.MoveNext
   Loop
  Call CloseRs(Rs)
   
    StrIdDoc = Mid(StrIdDoc, 2)
     idcomp = GenComprobante(StrIdDoc, lTipoLib, CbItemData(Cb_Mes), Val(Cb_Ano), 0, 1) 'se asigna el valor 1 al final de la funcion para indentificar que es un comprobante full
   
   If idcomp > 0 Then
   
      If FrmComprobante.FEditCentraliz(idcomp, CbItemData(Cb_Mes), Val(Cb_Ano), 1) = vbOK Then
       Call LoadGrid
      End If
      
  End If
   
  End Sub
'3217885

'3292777
Private Function EliminarDocCentralizadoPagado(ByVal vIdDoc As String) As Boolean
 Dim Q1 As String
 Dim Rs As Recordset
   
 EliminarDocCentralizadoPagado = False
 
      Me.MousePointer = vbHourglass
      
        Q1 = "Select IdDoc,numdoc from  Documento "
        Q1 = Q1 & " Where Documento.IdDoc not in ( select IdDoc from MovComprobante where ano = " & gEmpresa.Ano & "  and IdEmpresa = " & gEmpresa.id & ") "
        Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
        Q1 = Q1 & " AND Ano = " & gEmpresa.Ano
        Q1 = Q1 & " AND Tipolib =  " & lTipoLib
        Q1 = Q1 & " AND IdDoc = " & vIdDoc
        Q1 = Q1 & " AND Year(FEmision) = " & gEmpresa.Ano
        Q1 = Q1 & " And Estado in (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") And IdCompCent <> 0 And IdCompPago <> 0 "
        
        Set Rs = OpenRs(DbMain, Q1)
        If Not Rs.EOF Then
         
         EliminarDocCentralizadoPagado = True
        End If
        Call CloseRs(Rs)

      Me.MousePointer = vbDefault
      
End Function
'3292777
