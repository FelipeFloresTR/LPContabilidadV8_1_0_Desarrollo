VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmLibCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Caja"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14085
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_ImportTotalDocs 
      Caption         =   "Traer Docs Compras Anual"
      Height          =   315
      Left            =   6960
      TabIndex        =   58
      ToolTipText     =   "Traer otros "
      Top             =   8280
      Width           =   2870
   End
   Begin VB.CommandButton Bt_ImportOIngEgAnual 
      Caption         =   "Traer Otros Ingresos Anual"
      Height          =   315
      Left            =   9960
      TabIndex        =   57
      ToolTipText     =   "Traer documentos desde libro de "
      Top             =   8280
      Width           =   2870
   End
   Begin VB.CommandButton Bt_ImportOIngEg 
      Caption         =   "Traer otros Ingresos"
      Height          =   315
      Left            =   9120
      TabIndex        =   46
      ToolTipText     =   "Traer otros "
      Top             =   7800
      Width           =   1995
   End
   Begin VB.CommandButton Bt_Opciones 
      Caption         =   "Opciones de Ediciï¿½n"
      Height          =   315
      Left            =   11160
      TabIndex        =   47
      Top             =   7800
      Width           =   1635
   End
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones de Vista/Ediciï¿½n"
      Height          =   2535
      Left            =   10320
      TabIndex        =   53
      Top             =   5220
      Width           =   2475
      Begin VB.CheckBox Ch_ViewNotaCred 
         Caption         =   "ï¿½Tiene NC Asociada?"
         Height          =   195
         Left            =   300
         TabIndex        =   60
         Top             =   2100
         Width           =   2085
      End
      Begin VB.CheckBox Ch_ViewDTE 
         Caption         =   "Ver  DTE"
         Height          =   195
         Left            =   300
         TabIndex        =   49
         Top             =   660
         Width           =   1020
      End
      Begin VB.CheckBox Ch_ViewOper 
         Caption         =   "Ver  Operaciï¿½n"
         Height          =   195
         Left            =   300
         TabIndex        =   48
         Top             =   300
         Width           =   1755
      End
      Begin VB.CheckBox Ch_ViewNombre 
         Caption         =   "Ver  Nombre"
         Height          =   195
         Left            =   300
         TabIndex        =   50
         Top             =   1020
         Width           =   1275
      End
      Begin VB.CheckBox Ch_ViewIVAIrrec 
         Caption         =   "Ver IVA No Recuperable"
         Height          =   195
         Left            =   300
         TabIndex        =   51
         Top             =   1380
         Width           =   2085
      End
      Begin VB.CheckBox Ch_ViewOtrosImp 
         Caption         =   "Ver Otros Impuestos"
         Height          =   195
         Left            =   300
         TabIndex        =   52
         Top             =   1740
         Width           =   1845
      End
      Begin VB.CommandButton Bt_CerrarOpt 
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.CommandButton Bt_ImportDocs 
      Caption         =   "Traer Docs Compras"
      Height          =   315
      Left            =   6960
      TabIndex        =   45
      ToolTipText     =   "Traer documentos desde libro de "
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   7800
      Width           =   6855
   End
   Begin VB.Frame Fr_List 
      ForeColor       =   &H00FF0000&
      Height          =   945
      Left            =   0
      TabIndex        =   36
      Top             =   600
      Width           =   13755
      Begin VB.CheckBox Ch_SaldoInicial 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Left            =   7920
         TabIndex        =   55
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.ComboBox Cb_TipoOper 
         Height          =   315
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1335
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   2175
      End
      Begin VB.CommandButton Bt_List 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   12480
         Picture         =   "FrmLibroCaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   4140
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   9420
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   2595
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   7920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   1455
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Tx_Glosa 
         Height          =   315
         Left            =   9900
         MaxLength       =   100
         TabIndex        =   11
         Top             =   540
         Width           =   2115
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Oper.:"
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   44
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc.:"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   40
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Neto:"
         Height          =   195
         Index           =   9
         Left            =   6120
         TabIndex        =   39
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   1
         Left            =   9420
         TabIndex        =   38
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   6120
         TabIndex        =   37
         Top             =   240
         Width           =   390
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5835
      Left            =   30
      TabIndex        =   0
      Top             =   1620
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10292
      Cols            =   2
      Rows            =   4
      FixedCols       =   1
      FixedRows       =   2
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   13755
      Begin VB.CommandButton Bt_ViewRes 
         Caption         =   "Vista Resumida"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   9780
         TabIndex        =   29
         ToolTipText     =   "Vista resumida para impresión"
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton Bt_DelAll 
         Caption         =   "Eliminar TODOS los Registros"
         Height          =   315
         Left            =   7260
         TabIndex        =   28
         Top             =   180
         Width           =   2355
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   12540
         TabIndex        =   31
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   11400
         TabIndex        =   30
         Top             =   180
         Width           =   1035
      End
      Begin VB.Frame Fr_BtGen 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   3720
         TabIndex        =   34
         Top             =   180
         Width           =   3495
         Begin VB.CommandButton Bt_SaldoTotal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   19.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            Picture         =   "FrmLibroCaja.frx":043E
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Saldo y Totales Libro de Cajas"
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
            Left            =   1200
            Picture         =   "FrmLibroCaja.frx":07F1
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Vista previa de la impresión"
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
            Left            =   1560
            Picture         =   "FrmLibroCaja.frx":0C98
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Imprimir"
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
            Left            =   1920
            Picture         =   "FrmLibroCaja.frx":1152
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Copiar Excel"
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
            Left            =   720
            Picture         =   "FrmLibroCaja.frx":1597
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Sumar datos seleccionados"
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
            Left            =   2640
            Picture         =   "FrmLibroCaja.frx":163B
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Calculadora"
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
            Left            =   2280
            Picture         =   "FrmLibroCaja.frx":199C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Convertir moneda"
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
            Left            =   3120
            Picture         =   "FrmLibroCaja.frx":1D3A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Calendario"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Frame Fr_BtEdit 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   60
         TabIndex        =   33
         Top             =   180
         Width           =   3615
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
            Left            =   3180
            Picture         =   "FrmLibroCaja.frx":2163
            Style           =   1  'Graphical
            TabIndex        =   20
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
            Left            =   2760
            Picture         =   "FrmLibroCaja.frx":2640
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Detalle documento o comprobante seleccionado"
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
            Left            =   1320
            Picture         =   "FrmLibroCaja.frx":2AA5
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Pegar dato copiado"
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
            Left            =   900
            Picture         =   "FrmLibroCaja.frx":2E8E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Copiar dato"
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
            Left            =   2220
            Picture         =   "FrmLibroCaja.frx":326E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Eliminar documento seleccionado"
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
            Left            =   1740
            Picture         =   "FrmLibroCaja.frx":366A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Duplicar documento seleccionado"
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
            Picture         =   "FrmLibroCaja.frx":3ABC
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Seleccionar tipo de documento"
            Top             =   0
            Width           =   375
         End
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
            Left            =   420
            Picture         =   "FrmLibroCaja.frx":3E61
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Seleccionar Entidad"
            Top             =   0
            Width           =   375
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7440
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   11
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin FlexEdGrid2.FEd2Grid GridAnual 
      Height          =   315
      Left            =   240
      TabIndex        =   59
      Top             =   9000
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   556
      Cols            =   2
      Rows            =   4
      FixedCols       =   1
      FixedRows       =   2
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Menu M_TipoDoc 
      Caption         =   "Tipo Documento"
      Visible         =   0   'False
      Begin VB.Menu M_ItTipoDoc 
         Caption         =   "TipoDoc0"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmLibCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDLIBROCAJA = 0
Const C_IDDOC = 1
Const C_NUMLIN = 2
Const C_TIPOOPER = 3
Const C_IDTIPOOPER = 4
Const C_NUMDOC = 5
Const C_NUMDOCHASTA = 6
Const C_IDTIPODOC = 7
Const C_TIPODOC = 8
Const C_IDTIPOLIB = 9
Const C_DTE = 10
Const C_TIPODOCEXT = 11
Const C_DOCIMPEXP = 12
Const C_IDENTIDAD = 13
Const C_RUT = 14
Const C_NOMBRE = 15
Const C_FECHAOPER = 16
Const C_LNGFECHAOPER = 17
Const C_LNGFECHAINGRESOLIBRO = 18
Const C_AFECTO = 19
Const C_IVA = 20
Const C_IVAIRREC = 21
Const C_EXENTO = 22
Const C_OTROIMP = 23
Const C_TOTAL = 24
Const C_ESREBAJA = 25
Const C_PAGADO = 26
Const C_DESCRIP = 27
Const C_CONENTREL = 28
Const C_OPERDEVENGADA = 29
Const C_PAGOAPLAZO = 30
Const C_FECHAEXIGPAGO = 31
Const C_LNGFECHAEXIGPAGO = 32
Const C_INGRESO = 33
Const C_EGRESO = 34
Const C_SALDO = 35
Const C_MONTOAFECTABASEIMP = 36
Const C_IDESTADO = 37
Const C_ESTADO = 38
Const C_IDENTREAL = 39
Const C_IDCOMP = 40
Const C_IDUSUARIO = 41
Const C_USUARIO = 42
Const C_UPDATE = 43

'2690461
Const C_NOTACRED = 44
'fin 2690461

'2690461
Const NCOLS = C_NOTACRED
'Const NCOLS = C_UPDATE
'fin 2690461

Const RUT_VARIOS = "VARIOS"


Const O_VIEWLIBLEGAL = -1


Dim lTipoLib As Integer
Dim lTipoOper As Integer   'Ingreso o Egreso
Dim lAno As Integer
Dim lMes As Integer
Dim lIdLibroCaja As Long

Dim lOper As Integer
Dim lRc As Integer

Dim lOrdenGr(NCOLS) As String
Dim lOrdenSel As Integer            'orden seleccionado o actual
Dim lEditEnabled As Boolean
Dim lMsgNotaCred As Boolean         'para indicar que ya no deseo ver mensaje de Nota de Crédito
Dim lMsgFechaErr As Boolean

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

Dim lcbNombre As ClsCombo

Dim lHayDocsLibComprasVentas As Boolean   'para indicar si ya se importaron los libros de compras y ventas
Dim lMsgHayDocsLibCompasVentas As Boolean  'para indicar si ya se mostró el mensaje que ya se  importaron los libros de compras y ventas

Dim lModifica As Boolean

Dim lInLoad As Boolean

Dim lViewRes As Boolean

'2802201
Dim IdDoc As String
Dim NumDoc As String
'fin 2802201




Public Function FEdit(ByVal TipoOper As Integer, ByVal Mes As Integer, ByVal Ano As Integer) As Integer

   lOper = O_EDIT
   lTipoOper = TipoOper
   lMes = Mes
   lAno = Ano
      
   Me.Show vbModal
   
   FEdit = lRc
      
End Function

Public Sub FView(Optional ByVal Mes As Integer = 0)

   lOper = O_VIEW
   lAno = gEmpresa.Ano
   lMes = Mes
   lTipoOper = 0
   
   Me.Show vbModal
End Sub

Public Sub FViewLibroLeg(Optional ByVal Mes As Integer = 0, Optional ByVal Ano As Integer = 0)

   lOper = O_VIEWLIBLEGAL
   lMes = Mes
   lAno = Ano
   
   Me.Show vbModeless
End Sub

Private Sub SetUpGrid(Grid As FEd2Grid)

   Grid.Cols = NCOLS + 1
   Grid.FixedCols = C_NUMLIN + 1
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDLIBROCAJA) = 0
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_NUMLIN) = 500
   Grid.ColWidth(C_IDTIPOOPER) = 0
   Grid.ColWidth(C_TIPOOPER) = 500
   Grid.ColWidth(C_NUMDOC) = 1200
   Grid.ColWidth(C_NUMDOCHASTA) = 0
   Grid.ColWidth(C_IDTIPODOC) = 0
   '2814014 pipe
   'Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_TIPODOC) = 500
   'fin 2814014
   Grid.ColWidth(C_IDTIPOLIB) = 0
   Grid.ColWidth(C_DTE) = 400
   Grid.ColWidth(C_TIPODOCEXT) = 2000
   Grid.ColWidth(C_DOCIMPEXP) = 0
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_IDENTREAL) = 0         'esta columna ya no se usa, nueva disposición SII (2 jun 2021) 'ANTES de jun 2021: sólo para las ventas, dado que en este caso, C_IDENTIDAD es cero porque se usa en RUT y NOMBRE los datos de la empresa emisora, es decir la empresa con la que estamos trabajando ahora y que tenemos abierta
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_NOMBRE) = 2000
   Grid.ColWidth(C_FECHAOPER) = 1000
   Grid.ColWidth(C_LNGFECHAOPER) = 0
   Grid.ColWidth(C_LNGFECHAINGRESOLIBRO) = 0
   Grid.ColWidth(C_AFECTO) = 1200
   Grid.ColWidth(C_IVA) = 1200
   Grid.ColWidth(C_IVAIRREC) = 1200
   Grid.ColWidth(C_EXENTO) = 1200
   Grid.ColWidth(C_OTROIMP) = 1200
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_ESREBAJA) = 0
   Grid.ColWidth(C_PAGADO) = 1200
   Grid.ColWidth(C_DESCRIP) = 1800
   Grid.ColWidth(C_CONENTREL) = 1000
   If gEmpresa.Ano < 2021 Then
      Grid.ColWidth(C_OPERDEVENGADA) = 1000
   Else
      Grid.ColWidth(C_OPERDEVENGADA) = 0    '1000     'Se elimina esta columna. Solicitado por Victor Morales 31/5/2021
   End If
   Grid.ColWidth(C_PAGOAPLAZO) = 900
   Grid.ColWidth(C_FECHAEXIGPAGO) = 1100
   Grid.ColWidth(C_LNGFECHAEXIGPAGO) = 0
   Grid.ColWidth(C_INGRESO) = 1200
   Grid.ColWidth(C_EGRESO) = 1200
   Grid.ColWidth(C_SALDO) = 1200
   Grid.ColWidth(C_MONTOAFECTABASEIMP) = 1200
   
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_ESTADO) = 0
   Grid.ColWidth(C_IDUSUARIO) = 0
   Grid.ColWidth(C_USUARIO) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   '2690461
    'If lTipoOper = TOPERCAJA_INGRESO Then
        If Ch_ViewNotaCred = 0 Then
         Grid.ColWidth(C_NOTACRED) = 0
        End If
    'End If
   'fin 2690461
   
      
   If Ch_ViewOper = 0 Then
      Grid.ColWidth(C_TIPOOPER) = 0
   End If
   
   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
   End If
   
   If Ch_ViewNombre = 0 Then
      Grid.ColWidth(C_NOMBRE) = 0
   End If

   If Ch_ViewIVAIrrec = 0 Then
      Grid.ColWidth(C_IVAIRREC) = 0
   End If

   If Ch_ViewOtrosImp = 0 Then
      Grid.ColWidth(C_OTROIMP) = 0
   End If

   
   Grid.ColAlignment(C_NUMLIN) = flexAlignRightCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPOOPER) = flexAlignCenterCenter
   Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_AFECTO) = flexAlignRightCenter
   Grid.ColAlignment(C_IVA) = flexAlignRightCenter
   Grid.ColAlignment(C_IVAIRREC) = flexAlignRightCenter
   Grid.ColAlignment(C_EXENTO) = flexAlignRightCenter
   Grid.ColAlignment(C_OTROIMP) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_PAGADO) = flexAlignRightCenter
   Grid.ColAlignment(C_CONENTREL) = flexAlignCenterCenter
   Grid.ColAlignment(C_OPERDEVENGADA) = flexAlignCenterCenter
   Grid.ColAlignment(C_PAGOAPLAZO) = flexAlignCenterCenter
   Grid.ColAlignment(C_FECHAOPER) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHAEXIGPAGO) = flexAlignRightCenter
   Grid.ColAlignment(C_INGRESO) = flexAlignRightCenter
   Grid.ColAlignment(C_EGRESO) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOAFECTABASEIMP) = flexAlignRightCenter

   '2690461
   Grid.ColAlignment(C_NOTACRED) = flexAlignCenterCenter
   '2690461
   
   Grid.TextMatrix(0, C_NUMLIN) = "N°"
   Grid.TextMatrix(1, C_NUMLIN) = "Corr."
   If Grid.ColWidth(C_TIPOOPER) > 0 Then
      Grid.TextMatrix(1, C_TIPOOPER) = "Oper."
   End If
   Grid.TextMatrix(1, C_NUMDOC) = "N° Doc."
   Grid.TextMatrix(1, C_TIPODOC) = "TD"
   If Grid.ColWidth(C_DTE) > 0 Then
      Grid.TextMatrix(1, C_DTE) = "DTE"
   End If
   Grid.TextMatrix(1, C_TIPODOCEXT) = "Tipo Documento"
   Grid.TextMatrix(0, C_RUT) = "RUT"
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Grid.TextMatrix(1, C_RUT) = "Receptor Doc."
   ElseIf lTipoOper = TOPERCAJA_EGRESO Then
      Grid.TextMatrix(1, C_RUT) = "Emisor Doc."
   Else
      Grid.TextMatrix(1, C_RUT) = "Emis/Recep"
   End If
   
   If Grid.ColWidth(C_NOMBRE) > 0 Then
      Grid.TextMatrix(1, C_NOMBRE) = "Nombre"
   End If
   Grid.TextMatrix(0, C_FECHAOPER) = "Fecha"
   Grid.TextMatrix(1, C_FECHAOPER) = "Operación"
   Grid.TextMatrix(0, C_AFECTO) = "Afectas"
   Grid.TextMatrix(1, C_AFECTO) = "Monto Neto"
   Grid.TextMatrix(0, C_IVA) = "Afectas"
   Grid.TextMatrix(1, C_IVA) = "Monto IVA"
   If Grid.ColWidth(C_IVAIRREC) > 0 Then
      Grid.TextMatrix(0, C_IVAIRREC) = "Afectas"
      Grid.TextMatrix(1, C_IVAIRREC) = "IVA No Recup."
   End If
   Grid.TextMatrix(1, C_EXENTO) = "Monto Exento"
   If Grid.ColWidth(C_OTROIMP) > 0 Then
      Grid.TextMatrix(1, C_OTROIMP) = "Otros Imp."
   End If
   Grid.TextMatrix(1, C_TOTAL) = "Monto Total"
   Grid.TextMatrix(0, C_PAGADO) = "Monto Percib."
   Grid.TextMatrix(1, C_PAGADO) = "o Pagado"
   Grid.TextMatrix(0, C_DESCRIP) = "Glosa"
   Grid.TextMatrix(1, C_DESCRIP) = "Operación"
   Grid.TextMatrix(0, C_CONENTREL) = "Oper. Ent."
   Grid.TextMatrix(1, C_CONENTREL) = "Relacionada"
   
   If gEmpresa.Ano < 2021 Then
      Grid.TextMatrix(0, C_OPERDEVENGADA) = "Operación"
      Grid.TextMatrix(1, C_OPERDEVENGADA) = "Devengada"
   End If
   
   Grid.TextMatrix(0, C_PAGOAPLAZO) = "Pago a"
   Grid.TextMatrix(1, C_PAGOAPLAZO) = "Plazo"
   Grid.TextMatrix(0, C_FECHAEXIGPAGO) = "Fecha Exig."
   Grid.TextMatrix(1, C_FECHAEXIGPAGO) = "Pago"
   Grid.TextMatrix(0, C_INGRESO) = "Monto"
   Grid.TextMatrix(1, C_INGRESO) = "Ingreso"
   Grid.TextMatrix(0, C_EGRESO) = "Monto"
   Grid.TextMatrix(1, C_EGRESO) = "Egreso"
   Grid.TextMatrix(1, C_SALDO) = "Saldo"
   Grid.TextMatrix(0, C_MONTOAFECTABASEIMP) = "Monto Afecta"
   Grid.TextMatrix(1, C_MONTOAFECTABASEIMP) = "a Base Imp."
   
   
   Grid.TextMatrix(1, C_ESTADO) = "Estado"
   
   If lOper = O_VIEWLIBLEGAL Then
      Grid.ColWidth(C_TIPODOC) = 0
      Grid.ColWidth(C_DTE) = 0
      Grid.ColWidth(C_NOMBRE) = 0
      Grid.TextMatrix(0, C_TIPODOC) = ""
      Grid.TextMatrix(1, C_TIPODOC) = ""
      Grid.TextMatrix(0, C_DTE) = ""
      Grid.TextMatrix(1, C_DTE) = ""
      Grid.TextMatrix(0, C_NOMBRE) = ""
      Grid.TextMatrix(1, C_NOMBRE) = ""
   End If
   
   If lViewRes Then
   
      Grid.ColWidth(C_CONENTREL) = 0
      Grid.ColWidth(C_OPERDEVENGADA) = 0
      Grid.ColWidth(C_PAGOAPLAZO) = 0
      Grid.ColWidth(C_FECHAEXIGPAGO) = 0
      
      Grid.TextMatrix(0, C_CONENTREL) = ""
      Grid.TextMatrix(1, C_CONENTREL) = ""
      Grid.TextMatrix(0, C_OPERDEVENGADA) = ""
      Grid.TextMatrix(1, C_OPERDEVENGADA) = ""
      Grid.TextMatrix(0, C_PAGOAPLAZO) = ""
      Grid.TextMatrix(1, C_PAGOAPLAZO) = ""
      Grid.TextMatrix(0, C_FECHAEXIGPAGO) = ""
      Grid.TextMatrix(1, C_FECHAEXIGPAGO) = ""
   
      If lTipoOper = TOPERCAJA_INGRESO Then
         Grid.ColWidth(C_EGRESO) = 0
         Grid.TextMatrix(0, C_EGRESO) = ""
         Grid.TextMatrix(1, C_EGRESO) = ""
         
          '2690461
            Grid.TextMatrix(0, C_NOTACRED) = "Nota de"
            Grid.TextMatrix(1, C_NOTACRED) = "Crédito"
          '2690461
         
      ElseIf lTipoOper = TOPERCAJA_EGRESO Then
         Grid.ColWidth(C_INGRESO) = 0
         Grid.TextMatrix(0, C_INGRESO) = ""
         Grid.TextMatrix(1, C_INGRESO) = ""
      End If
      
   End If
   

   Call FGrVRows(Grid, 1)
  
   Call FGrTotales(Me.Grid, GridTot)
End Sub

Private Sub FillCb()
   Dim i As Integer
   Dim PrefLen As Integer
   Dim MesActual As Integer
   
   MesActual = GetMesActual()
      
   Cb_Mes.AddItem " "
   Cb_Mes.ItemData(Cb_Mes.NewIndex) = 0
   
   Call FillMes(Cb_Mes)
               
   If lMes > 0 Then
      Cb_Mes.ListIndex = lMes
   ElseIf lMes = 0 Then
      If MesActual > 0 Then
         Cb_Mes.ListIndex = MesActual
      Else
         Cb_Mes.ListIndex = GetUltimoMesConMovs()
      End If
   Else  'lmes <0
      Cb_Mes.ListIndex = 0
   End If
           
      
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex
'   If lAno > 0 Then
'      For i = 0 To Cb_Ano.ListCount - 1
'         If Val(Cb_Ano.List(i)) = lAno Then
'            Cb_Ano.ListIndex = i
'            Exit For
'         End If
'      Next i
'   End If
      
   Call CbAddItem(Cb_TipoOper, "(Todas)", 0)
   Call CbAddItem(Cb_TipoOper, gTipoOperCaja(TOPERCAJA_INGRESO), TOPERCAJA_INGRESO)
   Call CbAddItem(Cb_TipoOper, gTipoOperCaja(TOPERCAJA_EGRESO), TOPERCAJA_EGRESO)
   Call CbSelItem(Cb_TipoOper, lTipoOper)
   
   PrefLen = Len("Libro de") + 1
   
   Call AddItem(Cb_Entidad, "", -1)
   For i = ENT_CLIENTE To ENT_OTRO
      Call AddItem(Cb_Entidad, gClasifEnt(i), i)
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir

End Sub
Private Sub LoadTipoDoc()
   Dim i As Integer
   Dim Item As String
   Dim FindLib As Boolean
   Dim j As Integer
   Dim TipoDoc As Integer
   
   If lInLoad Then
      Exit Sub
   End If
   
   Cb_TipoDoc.Clear
   
   If lTipoOper > 0 Then
      Cb_TipoDoc.AddItem "(Todos)"
      Cb_TipoDoc.ItemData(Cb_TipoDoc.NewIndex) = 0
      
      j = 1
      
      'primero agregamos los tipos de docs de Compras, Ventas y Retenciones si corresponde
      For i = 0 To UBound(gTipoDoc)
         
         If gTipoDoc(i).Nombre = "" Then
            Exit For
         End If
         
         If (lTipoOper = TOPERCAJA_EGRESO And (gTipoDoc(i).TipoLib = LIB_COMPRAS Or gTipoDoc(i).TipoLib = LIB_RETEN Or gTipoDoc(i).TipoLib = LIB_CAJAEGR)) Or (lTipoOper = TOPERCAJA_INGRESO And (gTipoDoc(i).TipoLib = LIB_VENTAS Or gTipoDoc(i).TipoLib = LIB_CAJAING)) Then
         
            FindLib = True
            
            Item = "[" & gTipoDoc(i).Diminutivo & "] " & gTipoDoc(i).Nombre
            
            If gTipoDoc(i).TipoLib = LIB_RETEN Then
               TipoDoc = gTipoDoc(i).TipoDoc + BASELIBCAJA_RETEN
            ElseIf gTipoDoc(i).TipoLib = LIB_CAJAEGR Or gTipoDoc(i).TipoLib = LIB_CAJAING Then
               TipoDoc = gTipoDoc(i).TipoDoc + BASELIBCAJA_INGEGR
            Else
               TipoDoc = gTipoDoc(i).TipoDoc
            End If
               
            Call CbAddItem(Cb_TipoDoc, Item, TipoDoc)
                        
            If lOper = O_EDIT Then       ' llenamos el menú desplegable
            
               If Not lHayDocsLibComprasVentas Or gTipoDoc(i).TipoLib = LIB_CAJAEGR Or gTipoDoc(i).TipoLib = LIB_CAJAING Then    'Importó documentos desde Libros de Compras y Venta => no puede agregar documentos de estos libros en forma manual, pero si otros ingresos y egresos
                  Load M_ItTipoDoc(j)
                  M_ItTipoDoc(j).Caption = Item
               End If
            
            End If
            
            j = j + 1
            
            
'         ElseIf FindLib Then   'se terminó el libro actual
'            Exit For
         
         End If
         
      Next i
                      
      If M_ItTipoDoc.Count > 1 Then
         M_ItTipoDoc(0).visible = False
      End If
      
      Cb_TipoDoc.ListIndex = 0
               
   End If
   
End Sub


Private Sub Bt_DelAll_Click()
   Dim Row As Integer
   
   If MsgBox1("¿Está seguro de eliminar todos los registros de este libro?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(Row, C_TIPOOPER) = "" Then   'terminó la lista
         Exit For
      End If
         
      Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE)
      
      
      
      lModifica = True
      
      Grid.rows = Grid.rows + 1
   Next Row
   
   lHayDocsLibComprasVentas = False
         
End Sub

Private Sub Bt_ImportDocs_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer
   Dim Msg As String, MsgLib As String
   Dim TmpTbl As String, TmpTbl2 As String, TmpTbl3 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
      
      
   If CbItemData(Cb_Mes) = 0 Or Val(Cb_Ano) = 0 Then
      Exit Sub
   End If
   
   If lOper <> O_EDIT Then
      Exit Sub
   End If
      
'   If lTipoOper = 0 Or lTipoLib = 0 Then
   If lTipoOper = 0 Then
      Exit Sub
   End If
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      MsgLib = "del Libro de Ventas"
   Else
      MsgLib = "de los Libros de Compras y Retenciones"
   End If
   
   If lModifica Then
      If MsgBox1("Antes de continuar se grabarán los cambios realizados en este libro." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If valida() Then
         Call SaveGrid(Grid)
      Else
         Exit Sub
      End If
   End If
   
   Msg = "Esta operación traerá los nuevos documentos " & MsgLib & " y actualizará los traidos con anterioridad."
   Msg = Msg & vbCrLf & vbCrLf & "Además actualizará los montos Percibidos/Pagados de cada documentoen este mes."
   Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
   Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"
   
   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
        
        
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
  
   'primero actualizamos los que ya están en el libro de caja

   Tbl = " LibroCaja "
   sFrom = " (LibroCaja INNER JOIN Documento ON Documento.IdDoc = LibroCaja.IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "LibroCaja") & " )"
   sFrom = sFrom & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   sFrom = sFrom & "  AND Documento.IdEmpresa = Entidades.IdEmpresa "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   sSet = " LibroCaja.TipoDoc = Documento.TipoDoc "
   sSet = sSet & " , LibroCaja.TipoLib = Documento.TipoLib "
   sSet = sSet & " , LibroCaja.NumDoc = iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ', Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc) "
   sSet = sSet & " , LibroCaja.NumDocHasta = Documento.NumDocHasta "
   sSet = sSet & " , LibroCaja.DTE = Documento.DTE "
   sSet = sSet & " , LibroCaja.IdEntidad = Documento.IdEntidad "
   sSet = sSet & " , LibroCaja.RutEntidad = Documento.RutEntidad "
   sSet = sSet & " , LibroCaja.NombreEntidad = Documento.NombreEntidad "
   sSet = sSet & " , LibroCaja.FechaOperacion = Documento.FEmisionOri "
'   Q1 = Q1 & " , LibroCaja.FechaIngresoLibro = Documento.FEmision "       'no corresponde hacerlo ya que esto se hace sólo la primera vez
   sSet = sSet & " , LibroCaja.Afecto = Documento.Afecto "
'   sSet = sSet & " , LibroCaja.IVA =  Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec) "
'   sSet = sSet & " , LibroCaja.IVAIrrec = Documento.ValIVAIrrec "
   
   '2772366
   sSet = sSet & " , LibroCaja.IVA =  Documento.IVA - iif(Documento.tipolib = 2 and Documento.tipodoc= 2,0, iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec)) "
   sSet = sSet & " , LibroCaja.IVAIrrec = iif(Documento.tipolib =2 and Documento.tipodoc= 2,0, Documento.ValIVAIrrec  ) "
   
   sSet = sSet & " , LibroCaja.Exento = Documento.Exento "
   sSet = sSet & " , LibroCaja.OtroImp = Documento.OtroImp "
   sSet = sSet & " , LibroCaja.Total = Documento.Total "
   'sSet = sSet & " , LibroCaja.Descrip = LibroCaja.Descrip"
   sSet = sSet & " , LibroCaja.Descrip = iif(Documento.Descrip <> '' AND Documento.Descrip <> ' ', Documento.Descrip, LibroCaja.Descrip)"
   sSet = sSet & " , LibroCaja.Estado = Documento.Estado "
   sSet = sSet & " , LibroCaja.IdEntReal = Documento.IdEntidad "           'en el caso de las ventas, esta es la entidad a la que se le emitió en documento (destinatario), ya que el IdEntidad está en cero para indicar de poner la Empresa Emisora
   sSet = sSet & " , LibroCaja.ConEntRel = Entidades.EntRelacionada "
   sSet = sSet & " , LibroCaja.PagoAPlazo = iif(Documento.NumCuotas > 0,-1,0) "
      
   If lTipoOper = TOPERCAJA_INGRESO Then
      sWhere = " WHERE Documento.TipoLib = " & LIB_VENTAS
   Else           'egreso
      sWhere = " WHERE Documento.TipoLib IN (" & LIB_COMPRAS & "," & LIB_RETEN & ")"
   End If
   
   sWhere = sWhere & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'luego importamos los nuevos
   Q1 = "INSERT INTO LibroCaja  (IdDoc, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc, DTE, NumDocHasta "
   Q1 = Q1 & ", IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec, Exento, OtroImp, Total "
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   '2840454
   'Q1 = Q1 & "  SELECT Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
    Q1 = Q1 & "  SELECT distinct Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
   'FIn 2840454
   
   Q1 = Q1 & ", iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ' , Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc) "
   Q1 = Q1 & ", Documento.DTE, Documento.NumDocHasta "
   Q1 = Q1 & ", Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad "
'   Q1 = Q1 & ", Documento.FEmisionOri, Documento.FEmision, Documento.Afecto, Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec) "
   'Q1 = Q1 & ", Documento.ValIVAIrrec, Documento.Exento, Documento.OtroImp, Documento.Total"
   '2772366
   Q1 = Q1 & ", Documento.FEmisionOri, Documento.FEmision, Documento.Afecto,  Documento.IVA - iif(Documento.tipolib = 2 and Documento.tipodoc= 2,0, iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec)) "
   Q1 = Q1 & ", iif(Documento.tipolib =2 and Documento.tipodoc= 2,0, Documento.ValIVAIrrec  ) , Documento.Exento, Documento.OtroImp, Documento.Total"
   Q1 = Q1 & ", Documento.Descrip, Entidades.EntRelacionada, iif(Documento.NumCuotas > 0,-1,0), Documento.Estado, " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", Documento.IdEntidad "
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad  "
   Q1 = Q1 & "  AND Documento.IdEmpresa = Entidades.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN LibroCaja ON Documento.IdDoc = LibroCaja.IdDoc "
   Q1 = Q1 & "  AND Documento.IdEmpresa = LibroCaja.IdEmpresa AND Documento.Ano = LibroCaja.Ano"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "LibroCaja")
   Q1 = Q1 & " WHERE LibroCaja.IdDoc IS NULL "
   Q1 = Q1 & "  AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & "  AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.TipoLib = " & LIB_VENTAS
   Else           'egreso
      Q1 = Q1 & " AND Documento.TipoLib IN (" & LIB_COMPRAS & "," & LIB_RETEN & ")"
   End If
   
   Q1 = Q1 & " ORDER BY Documento.IdDoc"
   
   Call ExecSQL(DbMain, Q1)
   
   'OJO: no se eliminan del libro de caja los documentos que fueron eliminados en el libro de compras, ventas o retenciones, porque el sistema lo hace automáticamente al momento de grabar libro de compras, ventas o retenciones
   
   
   'Ingresos pagados o percibidos del mes
   
   TmpTbl = DbGenTmpName2(gDbType, "tmplibcaja_")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   'primero obtenemos los saldos pagados o percibidos en este mes y los tiramos a una tabla temporal
   Q1 = "SELECT MovComprobante.IdDoc, Sum(Debe - Haber) as Saldo, MovComprobante.IdEmpresa, MovComprobante.Ano,  MovComprobante.Glosa "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & " INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
   Q1 = Q1 & " WHERE IdDoc > 0 "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_INGRESO
   Else           'egreso
      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_EGRESO
   End If
   
 
   Q1 = Q1 & " AND Comprobante.OtrosIngEg14TER = 0 "    ' los que no están marcados como otros ingresos/egresos
   Q1 = Q1 & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdDoc, MovComprobante.IdEmpresa, MovComprobante.Ano,  MovComprobante.Glosa"
   
   Call ExecSQL(DbMain, Q1)
   
   'segundo: los documentos de este mes que ya están en el libro de caja, en otra tabla temporal:
   ' - actualizar los valores percibidos y pagados sólo en estos documentos
   ' - determinar cuáles no están y debemos insertar, para registrar su percepción
   TmpTbl2 = DbGenTmpName2(gDbType, "tmplibcaja_2")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)

   Q1 = "SELECT IdLibroCaja, IdDoc, IdEmpresa, Ano,Descrip "
   Q1 = Q1 & " INTO " & TmpTbl2
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " WHERE IdDoc > 0 AND TipoOper = " & lTipoOper
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   'ahora asignamos los saldos pagados o percibidos a los documentos que ya están en el libro de este mes
   
   Tbl = " LibroCaja "
   sFrom = " (LibroCaja "
   sFrom = sFrom & " INNER JOIN " & TmpTbl2 & " ON LibroCaja.IdLibroCaja = " & TmpTbl2 & ".IdLibroCaja "
   sFrom = sFrom & JoinEmpAno(gDbType, TmpTbl2, "LibroCaja") & " )"
   sFrom = sFrom & " LEFT JOIN " & TmpTbl & " ON LibroCaja.IdDoc = " & TmpTbl & ".IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, TmpTbl, "LibroCaja")
   sSet = " LibroCaja.Pagado = iif( Saldo IS NULL, 0, Abs(" & TmpTbl & ".Saldo))"
   sWhere = " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'tercero: los pagos de documentos que no están en el libro de caja de este mes
   'OJO: si el cliente hace el pago antes de la centralización del documento es posible que los documentos se dupliquen en el libro de caja
   'Esto ocurrió con un cliente con unas NCC que pagó antes de tener el documento y centralizarlo (FCA 24 may 2019)

   TmpTbl3 = DbGenTmpName2(gDbType, "tmplibcaja_3")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl3)

   Q1 = "SELECT " & TmpTbl & ".IdDoc, " & TmpTbl & ".IdEmpresa, " & TmpTbl & ".Ano, " & TmpTbl & ".Glosa  "
   Q1 = Q1 & " INTO " & TmpTbl3
   Q1 = Q1 & " FROM " & TmpTbl & " LEFT JOIN " & TmpTbl2 & " ON " & TmpTbl & ".IdDoc = " & TmpTbl2 & ".IdDoc"
'   Q1 = Q1 & "  AND " & TmpTbl & ".IdEmpresa = " & TmpTbl2 & ".IdEmpresa AND " & TmpTbl & ".Ano = " & TmpTbl2 & ".Ano "
   Q1 = Q1 & JoinEmpAno(gDbType, TmpTbl, TmpTbl2)
   Q1 = Q1 & " WHERE " & TmpTbl2 & ".IdDoc IS NULL "
   Q1 = Q1 & " AND " & TmpTbl & ".IdEmpresa = " & gEmpresa.id & " AND " & TmpTbl & ".Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
      
     
      
      '2857886
    Call deleteCompraIngreso
    'fin 2857886
   
   'y finalmente agregamos los nuevos pagos o percepciones de este mes utilizando las tablas temporales (en la tabla temporal 1 ya están los pagos de este mes y en la tabla temporal 2 ya están los documentos del libro de caja de este mes)
   Q1 = "INSERT INTO LibroCaja  (IdDoc, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc "
   Q1 = Q1 & ", DTE, NumDocHasta, IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec "
   Q1 = Q1 & ", Exento, OtroImp, Total, Pagado "
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   
    '2840454
   'Q1 = Q1 & "  SELECT  Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
   Q1 = Q1 & "  SELECT distinct Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
   'FIn 2840454
   
   Q1 = Q1 & ", iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ' , Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc)"
   Q1 = Q1 & ", Documento.DTE, Documento.NumDocHasta, Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad "
   Q1 = Q1 & ", Documento.FEmisionOri, " & FirstDay & ", Documento.Afecto, Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec), Documento.ValIVAIrrec"
   Q1 = Q1 & ", Documento.Exento, Documento.OtroImp, Documento.Total, Abs(" & TmpTbl & ".Saldo)"
   Q1 = Q1 & ", iif(Documento.Descrip <> '' AND Documento.Descrip <> '  ' , Documento.Descrip, " & TmpTbl & ".glosa), Documento.EntRelacionada, iif(Documento.NumCuotas > 0,-1,0), Documento.Estado, " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", Documento.IdEntidad "
   Q1 = Q1 & " FROM ((Documento INNER JOIN " & TmpTbl & " ON Documento.IdDoc = " & TmpTbl & ".IdDoc  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl) & " )"
   Q1 = Q1 & " INNER JOIN " & TmpTbl3 & " ON Documento.IdDoc = " & TmpTbl3 & ".IdDoc   "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl3) & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True)
   Q1 = Q1 & " WHERE Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   
    'fin 2802201


   'Actualizamos los campos Ingreso y Egreso de cada documento
  
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
'      Q1 = Q1 & "  Ingreso = iif(EsRebaja <> 0, 0, LibroCaja.Pagado)"     'para ser consistentes con la función CalcSaldo
'      Q1 = Q1 & ", Egreso = iif(EsRebaja <> 0, LibroCaja.Pagado, 0 )"
      sSet = sSet & "  Ingreso = LibroCaja.Pagado"
      sSet = sSet & ", Egreso =  0 "
  Else
'      Q1 = Q1 & "  Ingreso = iif(EsRebaja <> 0, LibroCaja.Pagado, 0)"
'      Q1 = Q1 & ", Egreso = iif(EsRebaja <> 0, 0, LibroCaja.Pagado)"
      sSet = sSet & "  Ingreso =  0"
      sSet = sSet & ", Egreso = LibroCaja.Pagado"
   End If
   
   sWhere = " WHERE IdDoc > 0 AND TipoOper = " & lTipoOper
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"    'Fecha Ingreso corresponde a FEmision (es decir la fecha de recepción o ingreso a libro de compras o ventas)
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
  
   'En el caso de Notas de Crédito, tanto de compras como ventas, el TipoOper se debe invertir. Es decir, si es Ingreso pasa a Egreso y viceversa
   
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " TipoOper = " & TOPERCAJA_EGRESO
   sWhere = " WHERE TipoOper = " & TOPERCAJA_INGRESO & " AND LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo IN ( 'NCV', 'NCE', 'DVB' )"
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
  

   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " TipoOper = " & TOPERCAJA_INGRESO
   sWhere = " WHERE TipoOper = " & TOPERCAJA_EGRESO & " AND LibroCaja.TipoLib = " & LIB_COMPRAS & " AND TipoDocs.Diminutivo IN( 'NCC', 'NCF' )"
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'En el caso de las Retenciones, tanto el bruto como el neto van a la columna exento del libro de caja y no se ingresan los impuestos. El total corresponde al valor neto
   Q1 = "UPDATE LibroCaja  "
   Q1 = Q1 & " SET LibroCaja.Exento = LibroCaja.Exento + LibroCaja.Afecto, OtroImp = 0, Total = LibroCaja.Exento + LibroCaja.Afecto"
   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'y ahora limpiamos el Afecto
   Q1 = "UPDATE LibroCaja  "
   Q1 = Q1 & " SET LibroCaja.Afecto = 0 "
   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'OJO: Desde el 2 jun 2021 esto ya no se hace por disposición del SII. Solicitado por Victor Morales
   
   'En el caso de los ingresos el RUT es el de la empresa (no el que está en el libro de ventas), salvo en el caso de (FCV) y (LFV) y las notas de crédito/débito asociadas a dichos documentos.
   'por lo tento hacemos un UPDATE para dejar el Rut de la empresa en los ingresos, donde corresponde
   
'   Tbl = " LibroCaja "
'   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
'   sSet = " LibroCaja.IdEntidad = 0, LibroCaja.RutEntidad = '" & gEmpresa.Rut & "', LibroCaja.NombreEntidad = '" & gEmpresa.RazonSocial & "'"
'   sWhere = " WHERE LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo NOT IN ('FCV', 'LFV', 'NCV', 'NDV' )"
'   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
'   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'y en las notas de crédito asociadas
'   Tbl = " LibroCaja "
'   sFrom = " (((LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
'   sFrom = sFrom & JoinEmpAno(gDbType, "LibroCaja", "Documento") & " )"
'   sFrom = sFrom & " INNER JOIN Documento as Doc1 ON Documento.IdDocAsoc = Doc1.IdDoc "
'   sFrom = sFrom & JoinEmpAno(gDbType, "Doc1", "Documento") & " )"
'   sFrom = sFrom & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc )"
'   sFrom = sFrom & " INNER JOIN TipoDocs as TDocs1 ON Doc1.TipoLib = TDocs1.TipoLib AND Doc1.TipoDoc = TDocs1.TipoDoc "
'   sSet = " LibroCaja.IdEntidad = 0, LibroCaja.RutEntidad = '" & gEmpresa.Rut & "', LibroCaja.NombreEntidad = '" & gEmpresa.RazonSocial & "'"
'   sWhere = " WHERE LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo IN ('NCV', 'NDV') AND TDocs1.Diminutivo NOT IN ('FCV', 'LFV') "
'   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
'   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
'
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
    If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
  
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl3)
   
   Call LoadGrid
   
   Call GetMontoQueAfectaBaseImp(Grid)
   
   Me.MousePointer = vbDefault
      
End Sub

Private Sub ImportDocumentos(Mes As Integer)

   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer
   Dim Msg As String, MsgLib As String
   Dim TmpTbl As String, TmpTbl2 As String, TmpTbl3 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
      
      
   If CbItemData(Cb_Mes) = 0 Or Val(Cb_Ano) = 0 Then
      Exit Sub
   End If
   
   If lOper <> O_EDIT Then
      Exit Sub
   End If
      
'   If lTipoOper = 0 Or lTipoLib = 0 Then
   If lTipoOper = 0 Then
      Exit Sub
   End If
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      MsgLib = "del Libro de Ventas"
   Else
      MsgLib = "de los Libros de Compras y Retenciones"
   End If
   
'   If lModifica Then
'      If MsgBox1("Antes de continuar se grabarán los cambios realizados en este libro." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'         Exit Sub
'      End If
'
'      If Valida() Then
'         Call SaveGrid(Grid)
'      Else
'         Exit Sub
'      End If
'   End If
   
   'Msg = "Esta operación traerá los nuevos documentos " & MsgLib & " y actualizará los traidos con anterioridad."
   'Msg = Msg & vbCrLf & vbCrLf & "Además actualizará los montos Percibidos/Pagados de cada documentoen este mes."
   'Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
   'Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"
   
'   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'      Exit Sub
'   End If
   
   Me.MousePointer = vbHourglass
   
   Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), Mes, 1), FirstDay, LastDay)
        
        
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
  
   'primero actualizamos los que ya están en el libro de caja

   Tbl = " LibroCaja "
   sFrom = " (LibroCaja INNER JOIN Documento ON Documento.IdDoc = LibroCaja.IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "LibroCaja") & " )"
   sFrom = sFrom & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   sFrom = sFrom & "  AND Documento.IdEmpresa = Entidades.IdEmpresa "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   sSet = " LibroCaja.TipoDoc = Documento.TipoDoc "
   sSet = sSet & " , LibroCaja.TipoLib = Documento.TipoLib "
   sSet = sSet & " , LibroCaja.NumDoc = iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ', Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc) "
   sSet = sSet & " , LibroCaja.NumDocHasta = Documento.NumDocHasta "
   sSet = sSet & " , LibroCaja.DTE = Documento.DTE "
   sSet = sSet & " , LibroCaja.IdEntidad = Documento.IdEntidad "
   sSet = sSet & " , LibroCaja.RutEntidad = Documento.RutEntidad "
   sSet = sSet & " , LibroCaja.NombreEntidad = Documento.NombreEntidad "
   sSet = sSet & " , LibroCaja.FechaOperacion = Documento.FEmisionOri "
'   Q1 = Q1 & " , LibroCaja.FechaIngresoLibro = Documento.FEmision "       'no corresponde hacerlo ya que esto se hace sólo la primera vez
   sSet = sSet & " , LibroCaja.Afecto = Documento.Afecto "
'   sSet = sSet & " , LibroCaja.IVA = Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec) "
'   sSet = sSet & " , LibroCaja.IVAIrrec = Documento.ValIVAIrrec "
   
    '2772366
   sSet = sSet & " , LibroCaja.IVA =  Documento.IVA - iif(Documento.tipolib = 2 and Documento.tipodoc= 2,0, iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec)) "
   sSet = sSet & " , LibroCaja.IVAIrrec = iif(Documento.tipolib =2 and Documento.tipodoc= 2,0, Documento.ValIVAIrrec  ) "
   
   
   
   sSet = sSet & " , LibroCaja.Exento = Documento.Exento "
   sSet = sSet & " , LibroCaja.OtroImp = Documento.OtroImp "
   sSet = sSet & " , LibroCaja.Total = Documento.Total "
   'sSet = sSet & " , LibroCaja.Descrip = LibroCaja.Descrip"
   sSet = sSet & " , LibroCaja.Descrip = iif(Documento.Descrip <> '' AND Documento.Descrip <> ' ', Documento.Descrip, LibroCaja.Descrip)"
   sSet = sSet & " , LibroCaja.Estado = Documento.Estado "
   sSet = sSet & " , LibroCaja.IdEntReal = Documento.IdEntidad "           'en el caso de las ventas, esta es la entidad a la que se le emitió en documento (destinatario), ya que el IdEntidad está en cero para indicar de poner la Empresa Emisora
   sSet = sSet & " , LibroCaja.ConEntRel = Entidades.EntRelacionada "
   sSet = sSet & " , LibroCaja.PagoAPlazo = iif(Documento.NumCuotas > 0,-1,0) "
      
   If lTipoOper = TOPERCAJA_INGRESO Then
      sWhere = " WHERE Documento.TipoLib = " & LIB_VENTAS
   Else           'egreso
      sWhere = " WHERE Documento.TipoLib IN (" & LIB_COMPRAS & "," & LIB_RETEN & ")"
   End If
   
   sWhere = sWhere & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'luego importamos los nuevos
   Q1 = "INSERT INTO LibroCaja  (IdDoc, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc, DTE, NumDocHasta "
   Q1 = Q1 & ", IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec, Exento, OtroImp, Total "
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   Q1 = Q1 & "  SELECT  Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
   Q1 = Q1 & ", iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ' , Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc) "
   Q1 = Q1 & ", Documento.DTE, Documento.NumDocHasta "
   Q1 = Q1 & ", Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad "
   'Q1 = Q1 & ", Documento.FEmisionOri, Documento.FEmision, Documento.Afecto, Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec) "
   'Q1 = Q1 & ", Documento.ValIVAIrrec, Documento.Exento, Documento.OtroImp, Documento.Total"
   
   '2772366
   Q1 = Q1 & ", Documento.FEmisionOri, Documento.FEmision, Documento.Afecto,  Documento.IVA - iif(Documento.tipolib = 2 and Documento.tipodoc= 2,0, iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec)) "
   Q1 = Q1 & ", iif(Documento.tipolib =2 and Documento.tipodoc= 2,0, Documento.ValIVAIrrec  ) , Documento.Exento, Documento.OtroImp, Documento.Total"
   
   Q1 = Q1 & ", Documento.Descrip, Entidades.EntRelacionada, iif(Documento.NumCuotas > 0,-1,0), Documento.Estado, " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", Documento.IdEntidad "
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad  "
   Q1 = Q1 & "  AND Documento.IdEmpresa = Entidades.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN LibroCaja ON Documento.IdDoc = LibroCaja.IdDoc "
   Q1 = Q1 & "  AND Documento.IdEmpresa = LibroCaja.IdEmpresa AND Documento.Ano = LibroCaja.Ano"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "LibroCaja")
   Q1 = Q1 & " WHERE LibroCaja.IdDoc IS NULL "
   Q1 = Q1 & "  AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & "  AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.TipoLib = " & LIB_VENTAS
   Else           'egreso
      Q1 = Q1 & " AND Documento.TipoLib IN (" & LIB_COMPRAS & "," & LIB_RETEN & ")"
   End If
   
   Q1 = Q1 & " ORDER BY Documento.IdDoc"
   
   Call ExecSQL(DbMain, Q1)
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl3)
   
   'OJO: no se eliminan del libro de caja los documentos que fueron eliminados en el libro de compras, ventas o retenciones, porque el sistema lo hace automáticamente al momento de grabar libro de compras, ventas o retenciones
   
   
   'Ingresos pagados o percibidos del mes
   
   TmpTbl = DbGenTmpName2(gDbType, "tmplibcaja_")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   'primero obtenemos los saldos pagados o percibidos en este mes y los tiramos a una tabla temporal
   Q1 = "SELECT MovComprobante.IdDoc, Sum(Debe - Haber) as Saldo, MovComprobante.IdEmpresa, MovComprobante.Ano,  MovComprobante.Glosa "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & " INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
   Q1 = Q1 & " WHERE IdDoc > 0 "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_INGRESO
   Else           'egreso
      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_EGRESO
   End If
   
  
   Q1 = Q1 & " AND Comprobante.OtrosIngEg14TER = 0 "    ' los que no están marcados como otros ingresos/egresos
   Q1 = Q1 & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdDoc, MovComprobante.IdEmpresa, MovComprobante.Ano,  MovComprobante.Glosa"
   
   Call ExecSQL(DbMain, Q1)
   
   'segundo: los documentos de este mes que ya están en el libro de caja, en otra tabla temporal:
   ' - actualizar los valores percibidos y pagados sólo en estos documentos
   ' - determinar cuáles no están y debemos insertar, para registrar su percepción
   TmpTbl2 = DbGenTmpName2(gDbType, "tmplibcaja_2")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)

   Q1 = "SELECT IdLibroCaja, IdDoc, IdEmpresa, Ano,Descrip "
   Q1 = Q1 & " INTO " & TmpTbl2
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " WHERE IdDoc > 0 AND TipoOper = " & lTipoOper
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   'ahora asignamos los saldos pagados o percibidos a los documentos que ya están en el libro de este mes
   
   Tbl = " LibroCaja "
   sFrom = " (LibroCaja "
   sFrom = sFrom & " INNER JOIN " & TmpTbl2 & " ON LibroCaja.IdLibroCaja = " & TmpTbl2 & ".IdLibroCaja "
   sFrom = sFrom & JoinEmpAno(gDbType, TmpTbl2, "LibroCaja") & " )"
   sFrom = sFrom & " LEFT JOIN " & TmpTbl & " ON LibroCaja.IdDoc = " & TmpTbl & ".IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, TmpTbl, "LibroCaja")
   sSet = " LibroCaja.Pagado = iif( Saldo IS NULL, 0, Abs(" & TmpTbl & ".Saldo))"
   sWhere = " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'tercero: los pagos de documentos que no están en el libro de caja de este mes
   'OJO: si el cliente hace el pago antes de la centralización del documento es posible que los documentos se dupliquen en el libro de caja
   'Esto ocurrió con un cliente con unas NCC que pagó antes de tener el documento y centralizarlo (FCA 24 may 2019)

   TmpTbl3 = DbGenTmpName2(gDbType, "tmplibcaja_3")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl3)

   Q1 = "SELECT " & TmpTbl & ".IdDoc, " & TmpTbl & ".IdEmpresa, " & TmpTbl & ".Ano, " & TmpTbl & ".Glosa  "
   Q1 = Q1 & " INTO " & TmpTbl3
   Q1 = Q1 & " FROM " & TmpTbl & " LEFT JOIN " & TmpTbl2 & " ON " & TmpTbl & ".IdDoc = " & TmpTbl2 & ".IdDoc"
'   Q1 = Q1 & "  AND " & TmpTbl & ".IdEmpresa = " & TmpTbl2 & ".IdEmpresa AND " & TmpTbl & ".Ano = " & TmpTbl2 & ".Ano "
   Q1 = Q1 & JoinEmpAno(gDbType, TmpTbl, TmpTbl2)
   Q1 = Q1 & " WHERE " & TmpTbl2 & ".IdDoc IS NULL "
   Q1 = Q1 & " AND " & TmpTbl & ".IdEmpresa = " & gEmpresa.id & " AND " & TmpTbl & ".Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
  '2857886
   Call deleteCompraIngreso
    'fin 2857886
   
   'y finalmente agregamos los nuevos pagos o percepciones de este mes utilizando las tablas temporales (en la tabla temporal 1 ya están los pagos de este mes y en la tabla temporal 2 ya están los documentos del libro de caja de este mes)
   Q1 = "INSERT INTO LibroCaja  (IdDoc, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc "
   Q1 = Q1 & ", DTE, NumDocHasta, IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec "
   Q1 = Q1 & ", Exento, OtroImp, Total, Pagado"
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   Q1 = Q1 & "  SELECT  Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, " & lTipoOper & ", Documento.TipoDoc, Documento.TipoLib "
   Q1 = Q1 & ", iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ' , Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc)"
   Q1 = Q1 & ", Documento.DTE, Documento.NumDocHasta, Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad "
   Q1 = Q1 & ", Documento.FEmisionOri, " & FirstDay & ", Documento.Afecto, Documento.IVA - iif(Documento.ValIVAIrrec IS NULL, 0, Documento.ValIVAIrrec), Documento.ValIVAIrrec"
   Q1 = Q1 & ", Documento.Exento, Documento.OtroImp, Documento.Total, Abs(" & TmpTbl & ".Saldo)"
   Q1 = Q1 & ", iif(Documento.Descrip <> '' AND Documento.Descrip <> '  ' , Documento.Descrip, " & TmpTbl & ".glosa), Documento.EntRelacionada, iif(Documento.NumCuotas > 0,-1,0), Documento.Estado, " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", Documento.IdEntidad "
   Q1 = Q1 & " FROM ((Documento INNER JOIN " & TmpTbl & " ON Documento.IdDoc = " & TmpTbl & ".IdDoc  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl) & " )"
   Q1 = Q1 & " INNER JOIN " & TmpTbl3 & " ON Documento.IdDoc = " & TmpTbl3 & ".IdDoc   "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl3) & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True)
   Q1 = Q1 & " WHERE Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)


   'Actualizamos los campos Ingreso y Egreso de cada documento
  
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
'      Q1 = Q1 & "  Ingreso = iif(EsRebaja <> 0, 0, LibroCaja.Pagado)"     'para ser consistentes con la función CalcSaldo
'      Q1 = Q1 & ", Egreso = iif(EsRebaja <> 0, LibroCaja.Pagado, 0 )"
      sSet = sSet & "  Ingreso = LibroCaja.Pagado"
      sSet = sSet & ", Egreso =  0 "
  Else
'      Q1 = Q1 & "  Ingreso = iif(EsRebaja <> 0, LibroCaja.Pagado, 0)"
'      Q1 = Q1 & ", Egreso = iif(EsRebaja <> 0, 0, LibroCaja.Pagado)"
      sSet = sSet & "  Ingreso =  0"
      sSet = sSet & ", Egreso = LibroCaja.Pagado"
   End If
   
   sWhere = " WHERE IdDoc > 0 AND TipoOper = " & lTipoOper
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"    'Fecha Ingreso corresponde a FEmision (es decir la fecha de recepción o ingreso a libro de compras o ventas)
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
  
   'En el caso de Notas de Crédito, tanto de compras como ventas, el TipoOper se debe invertir. Es decir, si es Ingreso pasa a Egreso y viceversa
   
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " TipoOper = " & TOPERCAJA_EGRESO
   sWhere = " WHERE TipoOper = " & TOPERCAJA_INGRESO & " AND LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo IN ( 'NCV', 'NCE', 'DVB' )"
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
  

   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   sSet = " TipoOper = " & TOPERCAJA_INGRESO
   sWhere = " WHERE TipoOper = " & TOPERCAJA_EGRESO & " AND LibroCaja.TipoLib = " & LIB_COMPRAS & " AND TipoDocs.Diminutivo IN( 'NCC', 'NCF' )"
   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'En el caso de las Retenciones, tanto el bruto como el neto van a la columna exento del libro de caja y no se ingresan los impuestos. El total corresponde al valor neto
   Q1 = "UPDATE LibroCaja  "
   Q1 = Q1 & " SET LibroCaja.Exento = LibroCaja.Exento + LibroCaja.Afecto, OtroImp = 0, Total = LibroCaja.Exento + LibroCaja.Afecto"
   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'y ahora limpiamos el Afecto
   Q1 = "UPDATE LibroCaja  "
   Q1 = Q1 & " SET LibroCaja.Afecto = 0 "
   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'OJO: Desde el 2 jun 2021 esto ya no se hace por disposición del SII. Solicitado por Victor Morales
   
   'En el caso de los ingresos el RUT es el de la empresa (no el que está en el libro de ventas), salvo en el caso de (FCV) y (LFV) y las notas de crédito/débito asociadas a dichos documentos.
   'por lo tento hacemos un UPDATE para dejar el Rut de la empresa en los ingresos, donde corresponde
   
'   Tbl = " LibroCaja "
'   sFrom = " LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
'   sSet = " LibroCaja.IdEntidad = 0, LibroCaja.RutEntidad = '" & gEmpresa.Rut & "', LibroCaja.NombreEntidad = '" & gEmpresa.RazonSocial & "'"
'   sWhere = " WHERE LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo NOT IN ('FCV', 'LFV', 'NCV', 'NDV' )"
'   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
'   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'y en las notas de crédito asociadas
'   Tbl = " LibroCaja "
'   sFrom = " (((LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
'   sFrom = sFrom & JoinEmpAno(gDbType, "LibroCaja", "Documento") & " )"
'   sFrom = sFrom & " INNER JOIN Documento as Doc1 ON Documento.IdDocAsoc = Doc1.IdDoc "
'   sFrom = sFrom & JoinEmpAno(gDbType, "Doc1", "Documento") & " )"
'   sFrom = sFrom & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc )"
'   sFrom = sFrom & " INNER JOIN TipoDocs as TDocs1 ON Doc1.TipoLib = TDocs1.TipoLib AND Doc1.TipoDoc = TDocs1.TipoDoc "
'   sSet = " LibroCaja.IdEntidad = 0, LibroCaja.RutEntidad = '" & gEmpresa.Rut & "', LibroCaja.NombreEntidad = '" & gEmpresa.RazonSocial & "'"
'   sWhere = " WHERE LibroCaja.TipoLib = " & LIB_VENTAS & " AND TipoDocs.Diminutivo IN ('NCV', 'NDV') AND TDocs1.Diminutivo NOT IN ('FCV', 'LFV') "
'   sWhere = sWhere & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
'   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
'
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
    If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
  
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl3)
   
   Call LoadGridAnual(Mes)
   
   Call GetMontoQueAfectaBaseImp(GridAnual)
   
   Me.MousePointer = vbDefault


End Sub

Private Sub Bt_DetDoc_Click()
   Dim Frm As Form
   Dim IdDoc As Long, idcomp As Long
   Dim Rc As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc > 0 Then
      Set Frm = New FrmDoc
      Call Frm.FView(IdDoc)
      Set Frm = Nothing
   Else
      idcomp = Val(Grid.TextMatrix(Grid.Row, C_IDCOMP))
      If idcomp > 0 Then
         Set Frm = New FrmComprobante
         Call Frm.FView(idcomp, False)
         Set Frm = Nothing
      End If
   End If
   
      
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
   
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   If Grid.TextMatrix(Row, C_IDDOC) = 0 Then    'registro en blanco
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_RETEN Then
      MsgBox1 "Este tipo de documento no permite pago en cuotas", vbExclamation
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
   
   
   Set Frm = New FrmDocCuotas
   Call Frm.FView(IdDoc)
   Set Frm = Nothing
         
End Sub
Private Sub Bt_ImportOIngEg_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer, TipoLib As Integer, TipoCompStr As String, TipoComp As Integer
   Dim Msg As String, MsgLib As String
   Dim TmpTbl As String
   Dim sDelFrom As String, sDelWhere As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
      
   If CbItemData(Cb_Mes) = 0 Or Val(Cb_Ano) = 0 Then
      Exit Sub
   End If
   
   If lOper <> O_EDIT Then
      Exit Sub
   End If
      
'   If lTipoOper = 0 Or lTipoLib = 0 Then
   If lTipoOper = 0 Then
      Exit Sub
   End If
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      MsgLib = "Ingresos"
   Else
      MsgLib = "Egresos"
   End If
   
   If lModifica Then
      If MsgBox1("Antes de continuar se grabarán los cambios realizados en este libro." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If valida() Then
         Call SaveGrid(Grid)
      Else
         Exit Sub
      End If
   End If
   
   Msg = "Esta operación traerá los nuevos " & MsgLib & " obtenidos de los comprobantes de este mes y actualizará los traídos con anterioridad."
   Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
   Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"
   
   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      TipoDoc = LIBCAJA_OTROSING
      TipoLib = LIB_CAJAING
      TipoComp = TC_INGRESO
      TipoCompStr = Left(gTipoComp(TC_INGRESO), 1)
   Else
      TipoDoc = LIBCAJA_OTROSEGR
      TipoLib = LIB_CAJAEGR
      TipoComp = TC_EGRESO
      TipoCompStr = Left(gTipoComp(TC_EGRESO), 1)
   End If
       
   Where = " Comprobante.Tipo = " & TipoComp
   Where = Where & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
   Where = Where & " AND Comprobante.Estado =" & EC_APROBADO
   Where = Where & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Where = Where & " AND Comprobante.OtrosIngEg14TER <> 0 "
       
   'primero Eliminamos los que ya no están marcados con Comprobante.OtrosIngEg14TER <> 0
'   Q1 = "DELETE LibroCaja.* FROM LibroCaja INNER JOIN Comprobante ON Comprobante.IdComp = LibroCaja.IdComp "
'   Q1 = Q1 & " WHERE Comprobante.Tipo = " & TipoComp
'   Q1 = Q1 & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
'   Q1 = Q1 & " AND (Comprobante.OtrosIngEg14TER = 0 OR Comprobante.Estado <> " & EC_APROBADO & ")"
'   Call ExecSQL(DbMain, Q1)
   
   sDelFrom = " LibroCaja INNER JOIN Comprobante ON Comprobante.IdComp = LibroCaja.IdComp "
   sDelFrom = sDelFrom & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   
   sDelWhere = sDelWhere & " WHERE Comprobante.Tipo = " & TipoComp
   sDelWhere = sDelWhere & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
   
   sDelWhere = sDelWhere & " AND (Comprobante.OtrosIngEg14TER = 0 OR Comprobante.Estado <> " & EC_APROBADO & ")"
   sDelWhere = sDelWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call DeleteJSQL(DbMain, "LibroCaja", sDelFrom, sDelWhere)
   
   'luego actualizamos los que ya están en el libro de caja
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN Comprobante ON LibroCaja.IdComp = Comprobante.IdComp "
   sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   sSet = " LibroCaja.FechaOperacion = Comprobante.Fecha "
'   sSet = sSet & " , LibroCaja.NumDoc = '" & TipoCompStr & "' & Comprobante.Correlativo & '-' & Year(Comprobante.Fecha) & Right('0' & Month(Comprobante.Fecha),2)"
   sSet = sSet & " , LibroCaja.NumDoc = " & SqlConcat(gDbType, "'" & TipoCompStr & "'", "Comprobante.Correlativo", "'-'", SqlYearLng("Comprobante.Fecha"), "Right(" & SqlConcat(gDbType, "'0'", SqlMonthLng("Comprobante.Fecha")) & ",2)")
   sSet = sSet & " , LibroCaja.FechaIngresoLibro = Comprobante.Fecha "
   sSet = sSet & " , LibroCaja.Exento = Comprobante.TotalDebe "
   sSet = sSet & " , LibroCaja.Total = Comprobante.TotalDebe "
   '3338188
   If gDbType = SQL_ACCESS Then
    sSet = sSet & " , LibroCaja.Descrip = Comprobante.Glosa "
   Else
    sSet = sSet & " , LibroCaja.Descrip =  SUBSTRING(Comprobante.Glosa,0,50) "
   End If
   '3338188
   sSet = sSet & " , LibroCaja.ConEntRel = 0 "
   sSet = sSet & " , LibroCaja.PagoAPlazo = 0 "
   sSet = sSet & " , LibroCaja.FechaExigPago = Comprobante.Fecha "
   sSet = sSet & " , LibroCaja.IdEntReal = 0 "
   sWhere = " WHERE " & Where
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'finalmente importamos los nuevos
   Q1 = "INSERT INTO LibroCaja  (IdComp, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc, DTE, NumDocHasta "
   Q1 = Q1 & ", IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec, Exento, OtroImp, Total "
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, FechaExigPago, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   
   Q1 = Q1 & "  SELECT  Comprobante.IdComp, Comprobante.IdEmpresa, Comprobante.Ano, " & lTipoOper & "," & TipoDoc & "," & TipoLib
'   Q1 = Q1 & ", '" & TipoCompStr & "' & Comprobante.Correlativo & '-' & Year(Comprobante.Fecha) & Right('0' & Month(Comprobante.Fecha),2), 0, 0 "
   Q1 = Q1 & ", " & SqlConcat(gDbType, "'" & TipoCompStr & "'", "Comprobante.Correlativo", "'-'", SqlYearLng("Comprobante.Fecha"), "Right(" & SqlConcat(gDbType, "'0'", SqlMonthLng("Comprobante.Fecha")) & ",2)") & ", 0, 0"
   Q1 = Q1 & ", 0, '" & RUT_VARIOS & "', '' "
   Q1 = Q1 & ", Comprobante.Fecha, Comprobante.Fecha, 0, 0, 0, Comprobante.TotalDebe, 0, Comprobante.TotalDebe"
   
   '3338188
   If gDbType = SQL_ACCESS Then
    Q1 = Q1 & ", Comprobante.Glosa, 0, 0, Comprobante.Fecha, " & ED_PAGADO & ", " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", 0 "
   Else
    Q1 = Q1 & ", SUBSTRING(Comprobante.Glosa,0,50), 0, 0, Comprobante.Fecha, " & ED_PAGADO & ", " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", 0 "
   End If
   '3338188
   
   
   Q1 = Q1 & " FROM Comprobante LEFT JOIN LibroCaja ON Comprobante.IdComp = LibroCaja.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   Q1 = Q1 & " WHERE LibroCaja.IdComp IS NULL "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Q1 = Q1 & " ORDER BY Comprobante.IdComp"
   
   Call ExecSQL(DbMain, Q1)
   
   'ahora actualizamos los ingresos percibidos o egresos pagados
   
   TmpTbl = DbGenTmpName2(gDbType, "timplibcaja_")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = "SELECT MovComprobante.IdComp, Sum(Debe) as SumPago, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Else
      Q1 = "SELECT MovComprobante.IdComp, Sum(Haber) as SumPago, MovComprobante.IdEmpresa, MovComprobante.Ano "
   End If

   
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN  Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & " WHERE InStr( '" & gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas & "', ',' & MovComprobante.IdCuenta & ',' )"
   Q1 = Q1 & " WHERE " & SqlInStr("'" & gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas & "'", SqlConcat(gDbType, "','", "MovComprobante.IdCuenta", "','")) & " > 0 "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY MovComprobante.IdComp, MovComprobante.IdEmpresa, MovComprobante.Ano ORDER BY MovComprobante.IdComp"

   Call ExecSQL(DbMain, Q1)

'   Q1 = "UPDATE (LibroCaja INNER JOIN " & TmpTbl & " ON LibroCaja.IdComp = " & TmpTbl & ".IdComp "
'   Q1 = Q1 & "  AND LibroCaja.IdEmpresa = " & TmpTbl & ".IdEmpresa AND LibroCaja.Ano = " & TmpTbl & ".Ano )"
'   Q1 = Q1 & " SET LibroCaja.Pagado = SumPago"
'   Q1 = Q1 & " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " LibroCaja "
   sFrom = " (LibroCaja INNER JOIN " & TmpTbl & " ON LibroCaja.IdComp = " & TmpTbl & ".IdComp "
   sFrom = sFrom & JoinEmpAno(gDbType, "LibroCaja", TmpTbl) & " )"
   sSet = " LibroCaja.Pagado = SumPago"
   sWhere = " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

   'Finalmente, actualizamos los campos Ingreso y Egreso de cada documento
   Q1 = "UPDATE LibroCaja SET "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & "  Ingreso = LibroCaja.Pagado"
      Q1 = Q1 & ", Egreso = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_CAJAING
   Else
      Q1 = Q1 & "  Ingreso = 0"
      Q1 = Q1 & ", Egreso = LibroCaja.Pagado"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_CAJAEGR
   End If
   
   Q1 = Q1 & " AND TipoOper = " & lTipoOper
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)

   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   
   Call LoadGrid
   
   Call GetMontoQueAfectaBaseImp(Grid)
   
   Me.MousePointer = vbDefault
      
End Sub
Private Sub ImportOtrosDocu(Mes As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer, TipoLib As Integer, TipoCompStr As String, TipoComp As Integer
   Dim Msg As String, MsgLib As String
   Dim TmpTbl As String
   Dim sDelFrom As String, sDelWhere As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
      
   If CbItemData(Cb_Mes) = 0 Or Val(Cb_Ano) = 0 Then
      Exit Sub
   End If
   
   If lOper <> O_EDIT Then
      Exit Sub
   End If
      
'   If lTipoOper = 0 Or lTipoLib = 0 Then
   If lTipoOper = 0 Then
      Exit Sub
   End If
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      MsgLib = "Ingresos"
   Else
      MsgLib = "Egresos"
   End If
   
'   If lModifica Then
'      If MsgBox1("Antes de continuar se grabarán los cambios realizados en este libro." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'         Exit Sub
'      End If
'
'      If Valida() Then
'         Call SaveGrid(Grid)
'      Else
'         Exit Sub
'      End If
'   End If
'
'   Msg = "Esta operación traerá los nuevos " & MsgLib & " obtenidos de los comprobantes de este mes y actualizará los traídos con anterioridad."
'   Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
'   Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"
'
'   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'      Exit Sub
'   End If
   
   Me.MousePointer = vbHourglass
   
   Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), Mes, 1), FirstDay, LastDay)
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      TipoDoc = LIBCAJA_OTROSING
      TipoLib = LIB_CAJAING
      TipoComp = TC_INGRESO
      TipoCompStr = Left(gTipoComp(TC_INGRESO), 1)
   Else
      TipoDoc = LIBCAJA_OTROSEGR
      TipoLib = LIB_CAJAEGR
      TipoComp = TC_EGRESO
      TipoCompStr = Left(gTipoComp(TC_EGRESO), 1)
   End If
       
   Where = " Comprobante.Tipo = " & TipoComp
   Where = Where & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
   Where = Where & " AND Comprobante.Estado =" & EC_APROBADO
   Where = Where & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Where = Where & " AND Comprobante.OtrosIngEg14TER <> 0 "
       
   'primero Eliminamos los que ya no están marcados con Comprobante.OtrosIngEg14TER <> 0
'   Q1 = "DELETE LibroCaja.* FROM LibroCaja INNER JOIN Comprobante ON Comprobante.IdComp = LibroCaja.IdComp "
'   Q1 = Q1 & " WHERE Comprobante.Tipo = " & TipoComp
'   Q1 = Q1 & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
'   Q1 = Q1 & " AND (Comprobante.OtrosIngEg14TER = 0 OR Comprobante.Estado <> " & EC_APROBADO & ")"
'   Call ExecSQL(DbMain, Q1)
   
   sDelFrom = " LibroCaja INNER JOIN Comprobante ON Comprobante.IdComp = LibroCaja.IdComp "
   sDelFrom = sDelFrom & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   
   sDelWhere = sDelWhere & " WHERE Comprobante.Tipo = " & TipoComp
   sDelWhere = sDelWhere & " AND (Comprobante.Fecha BETWEEN " & FirstDay & " AND " & LastDay & ") "
   
   sDelWhere = sDelWhere & " AND (Comprobante.OtrosIngEg14TER = 0 OR Comprobante.Estado <> " & EC_APROBADO & ")"
   sDelWhere = sDelWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call DeleteJSQL(DbMain, "LibroCaja", sDelFrom, sDelWhere)
   
   'luego actualizamos los que ya están en el libro de caja
   Tbl = " LibroCaja "
   sFrom = " LibroCaja INNER JOIN Comprobante ON LibroCaja.IdComp = Comprobante.IdComp "
   sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   sSet = " LibroCaja.FechaOperacion = Comprobante.Fecha "
'   sSet = sSet & " , LibroCaja.NumDoc = '" & TipoCompStr & "' & Comprobante.Correlativo & '-' & Year(Comprobante.Fecha) & Right('0' & Month(Comprobante.Fecha),2)"
   sSet = sSet & " , LibroCaja.NumDoc = " & SqlConcat(gDbType, "'" & TipoCompStr & "'", "Comprobante.Correlativo", "'-'", SqlYearLng("Comprobante.Fecha"), "Right(" & SqlConcat(gDbType, "'0'", SqlMonthLng("Comprobante.Fecha")) & ",2)")
   sSet = sSet & " , LibroCaja.FechaIngresoLibro = Comprobante.Fecha "
   sSet = sSet & " , LibroCaja.Exento = Comprobante.TotalDebe "
   sSet = sSet & " , LibroCaja.Total = Comprobante.TotalDebe "
   '3338188
   If gDbType = SQL_ACCESS Then
    sSet = sSet & " , LibroCaja.Descrip = Comprobante.Glosa "
   Else
    sSet = sSet & " , LibroCaja.Descrip = SUBSTRING(Comprobante.Glosa,0,50) "
   End If
   '3338188
   sSet = sSet & " , LibroCaja.ConEntRel = 0 "
   sSet = sSet & " , LibroCaja.PagoAPlazo = 0 "
   sSet = sSet & " , LibroCaja.FechaExigPago = Comprobante.Fecha "
   sSet = sSet & " , LibroCaja.IdEntReal = 0 "
   sWhere = " WHERE " & Where
   sWhere = sWhere & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'finalmente importamos los nuevos
   Q1 = "INSERT INTO LibroCaja  (IdComp, IdEmpresa, Ano, TipoOper, TipoDoc, TipoLib, NumDoc, DTE, NumDocHasta "
   Q1 = Q1 & ", IdEntidad, RutEntidad, NombreEntidad  "
   Q1 = Q1 & ", FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec, Exento, OtroImp, Total "
   Q1 = Q1 & ", Descrip, ConEntRel, PagoAPlazo, FechaExigPago, Estado, IdUsuario, FechaCreacion, IdEntReal ) "
   
   Q1 = Q1 & "  SELECT  Comprobante.IdComp, Comprobante.IdEmpresa, Comprobante.Ano, " & lTipoOper & "," & TipoDoc & "," & TipoLib
'   Q1 = Q1 & ", '" & TipoCompStr & "' & Comprobante.Correlativo & '-' & Year(Comprobante.Fecha) & Right('0' & Month(Comprobante.Fecha),2), 0, 0 "
   Q1 = Q1 & ", " & SqlConcat(gDbType, "'" & TipoCompStr & "'", "Comprobante.Correlativo", "'-'", SqlYearLng("Comprobante.Fecha"), "Right(" & SqlConcat(gDbType, "'0'", SqlMonthLng("Comprobante.Fecha")) & ",2)") & ", 0, 0"
   Q1 = Q1 & ", 0, '" & RUT_VARIOS & "', '' "
   Q1 = Q1 & ", Comprobante.Fecha, Comprobante.Fecha, 0, 0, 0, Comprobante.TotalDebe, 0, Comprobante.TotalDebe"
   
   '3338188
   If gDbType = SQL_ACCESS Then
    Q1 = Q1 & ", Comprobante.Glosa, 0, 0, Comprobante.Fecha, " & ED_PAGADO & ", " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", 0 "
   Else
    Q1 = Q1 & ", SUBSTRING(Comprobante.Glosa,0,50), 0, 0, Comprobante.Fecha, " & ED_PAGADO & ", " & gUsuario.IdUsuario & "," & CLng(Int(Now)) & ", 0 "
   End If
   '3338188
   
   Q1 = Q1 & " FROM Comprobante LEFT JOIN LibroCaja ON Comprobante.IdComp = LibroCaja.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "LibroCaja")
   Q1 = Q1 & " WHERE LibroCaja.IdComp IS NULL "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Q1 = Q1 & " ORDER BY Comprobante.IdComp"
   
   Call ExecSQL(DbMain, Q1)
   
   'ahora actualizamos los ingresos percibidos o egresos pagados
   
   TmpTbl = DbGenTmpName2(gDbType, "timplibcaja_")
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = "SELECT MovComprobante.IdComp, Sum(Debe) as SumPago, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Else
      Q1 = "SELECT MovComprobante.IdComp, Sum(Haber) as SumPago, MovComprobante.IdEmpresa, MovComprobante.Ano "
   End If

   
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN  Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & " WHERE InStr( '" & gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas & "', ',' & MovComprobante.IdCuenta & ',' )"
   Q1 = Q1 & " WHERE " & SqlInStr("'" & gCtasAjusteExtraCont(TAEC_DISPONIBLES, TAEC_ITEMDISPONIBLE).LstCuentas & "'", SqlConcat(gDbType, "','", "MovComprobante.IdCuenta", "','")) & " > 0 "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY MovComprobante.IdComp, MovComprobante.IdEmpresa, MovComprobante.Ano ORDER BY MovComprobante.IdComp"

   Call ExecSQL(DbMain, Q1)

'   Q1 = "UPDATE (LibroCaja INNER JOIN " & TmpTbl & " ON LibroCaja.IdComp = " & TmpTbl & ".IdComp "
'   Q1 = Q1 & "  AND LibroCaja.IdEmpresa = " & TmpTbl & ".IdEmpresa AND LibroCaja.Ano = " & TmpTbl & ".Ano )"
'   Q1 = Q1 & " SET LibroCaja.Pagado = SumPago"
'   Q1 = Q1 & " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " LibroCaja "
   sFrom = " (LibroCaja INNER JOIN " & TmpTbl & " ON LibroCaja.IdComp = " & TmpTbl & ".IdComp "
   sFrom = sFrom & JoinEmpAno(gDbType, "LibroCaja", TmpTbl) & " )"
   sSet = " LibroCaja.Pagado = SumPago"
   sWhere = " WHERE LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

   'Finalmente, actualizamos los campos Ingreso y Egreso de cada documento
   Q1 = "UPDATE LibroCaja SET "
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      Q1 = Q1 & "  Ingreso = LibroCaja.Pagado"
      Q1 = Q1 & ", Egreso = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_CAJAING
   Else
      Q1 = Q1 & "  Ingreso = 0"
      Q1 = Q1 & ", Egreso = LibroCaja.Pagado"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_CAJAEGR
   End If
   
   Q1 = Q1 & " AND TipoOper = " & lTipoOper
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)

   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   
   Call LoadGridAnual(Mes)
   
   Call GetMontoQueAfectaBaseImp(GridAnual)
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_ImportOIngEgAnual_Click()
Dim Msg As String, MsgLib As String
Dim Index As Integer

   Msg = "Esta operación traerá los nuevos " & MsgLib & " obtenidos de los comprobantes de este año y actualizará los traídos con anterioridad."
   Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
   Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"

   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

For Index = 1 To 12
    GridAnual.Clear
    Call SetUpGrid(GridAnual)
    Call ImportOtrosDocu(Index)
    Call SaveGrid(GridAnual)
Next
Call LoadGrid
Call GetMontoQueAfectaBaseImp(Grid)
Me.MousePointer = vbDefault
End Sub

Private Sub Bt_ImportTotalDocs_Click()
Dim Msg As String, MsgLib As String
Dim Index As Integer

   Msg = "Esta operación traerá los nuevos documentos " & MsgLib & " y actualizará los traidos con anterioridad."
   Msg = Msg & vbCrLf & vbCrLf & "Además actualizará los montos Percibidos/Pagados de cada documento en este año."
   Msg = Msg & vbCrLf & vbCrLf & "Una vez realizada la importación, ésta no podrá ser cancelada posteriormente."
   Msg = Msg & vbCrLf & vbCrLf & "¿Desea continuar?"
   
   If MsgBox1(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If


For Index = 1 To 12
    GridAnual.Clear
    Call SetUpGrid(GridAnual)
    '2955019
    If lTipoOper = TOPERCAJA_EGRESO Then
     Call deleteVentaEgreso(Index)
    End If
   '2955019
    Call ImportDocumentos(Index)
    Call SaveGrid(GridAnual)
    Call LoadGrid
    Call GetMontoQueAfectaBaseImp(Grid)
Next
'Call loadGrid
'Call GetMontoQueAfectaBaseImp(Grid)
Me.MousePointer = vbDefault
End Sub

Private Sub Bt_List_Click()

   If Trim(Tx_Rut) <> "" And Val(lcbNombre.Matrix(M_IDENTIDAD)) = 0 Then
      MsgBox1 "El RUT ingresado no es válido o no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Tx_Rut.SetFocus
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Call LoadGrid
   Me.MousePointer = vbDefault
End Sub

Private Sub Bt_SaldoTotal_Click()
   '3289932
  If lOper = O_EDIT Then
        Dim Msg As String
        Msg = "Antes de ingresar a Saldos y Totales libro caja, Favor de ingresar a los libros de Ingreso y Egreso, "
        Msg = Msg & "esto para mostrar los montos correctos en item Monto que afectan a base imponible."
        Msg = Msg & vbLf & vbLf & "¿ Desea continuar ?"
         If MsgBox1(Msg, vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
              Exit Sub
         End If
   
   End If
   '3289932

   Dim Frm As FrmSalyTotLibCajas
   Set Frm = New FrmSalyTotLibCajas
   If CbItemData(Cb_Mes) = 0 Then
    Frm.GetMonth = False
   Else
    Frm.GetMonth = True
   End If
   Frm.Fecha = DateSerial(Me.Cb_Ano, CbItemData(Cb_Mes), 1)
   Call Frm.FView
End Sub

Private Function FormSaldos() As FrmSalyTotLibCajas
   Dim Frm As FrmSalyTotLibCajas
   Set Frm = New FrmSalyTotLibCajas
   If CbItemData(Cb_Mes) = 0 Then
    Frm.GetMonth = False
   Else
    Frm.GetMonth = True
   End If
   Frm.Fecha = DateSerial(Me.Cb_Ano, CbItemData(Cb_Mes), 1)
   Set FormSaldos = FrmSalyTotLibCajas
End Function

Private Sub Bt_SelEnt_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   Dim Row As Integer
   Dim TipoEnt As Integer
   Dim Col As Integer
   Dim Rc As Integer
      
   Col = Grid.Col
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lTipoOper = TOPERCAJA_EGRESO Then
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
   
   If Grid.TextMatrix(Row, C_FECHAOPER) = "" Or (Grid.Col <> C_RUT And Grid.Col <> C_NOMBRE) Or Not ValidaEstadoEdit(Row) Or Val(Grid.TextMatrix(Row, C_IDDOC)) <> 0 Then
      Exit Sub
   End If
         
   If Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And Entidad.NotValidRut <> 0 Then
      MsgBox1 "Rut inválido para este tipo de documento.", vbExclamation
      Exit Sub
   End If
   
   Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
   Grid.TextMatrix(Row, C_RUT) = FmtCID(Entidad.Rut, Entidad.NotValidRut = False)
   Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
   
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
   lModifica = True
      
End Sub


Private Sub Bt_ViewRes_Click()

   lViewRes = Not lViewRes

   Call SetUpGrid(Grid)
   
   If lViewRes Then
      Bt_ViewRes.Caption = "Vista Completa"
   Else
      Bt_ViewRes.Caption = "Vista Resumida"
   End If
   
End Sub

Private Sub Cb_Ano_Click()
   Bt_List.Enabled = True

End Sub


Private Sub Cb_Ano_LostFocus()
  vMontoBaseImpoIngreso = 0
   vMontoBaseImpoEgreso = 0
End Sub

Private Sub Cb_Mes_Click()
   Bt_List.Enabled = True

End Sub


Private Sub Cb_Mes_LostFocus()
   vMontoBaseImpoIngreso = 0
   vMontoBaseImpoEgreso = 0
End Sub

Private Sub Cb_TipoOper_Click()

   lTipoOper = CbItemData(Cb_TipoOper)
   Call LoadTipoDoc
   Bt_List.Enabled = True

End Sub

Private Sub Cb_TipoDoc_Click()
   Bt_List.Enabled = True

End Sub

Private Sub Ch_Rut_Click()
   Bt_List.Enabled = True

End Sub

Private Sub Ch_SaldoInicial_Click()
   Bt_List.Enabled = True

End Sub

'2690461
Private Sub Ch_ViewNotaCred_Click()
 If Ch_ViewNotaCred = 0 Then
      Grid.ColWidth(C_NOTACRED) = 0
      Grid.TextMatrix(0, C_NOTACRED) = ""
      Grid.TextMatrix(1, C_NOTACRED) = ""

   Else
      Grid.ColWidth(C_NOTACRED) = 900
      Grid.TextMatrix(0, C_NOTACRED) = "Nota de"
      Grid.TextMatrix(1, C_NOTACRED) = "Crédito"

   End If

   Call SetIniString(gIniFile, "Opciones", "VerLCajaNotCred", Abs(Ch_ViewNotaCred.Value))
   gVarIniFile.VerLCajaNotCred = Abs(Ch_ViewNotaCred.Value)
End Sub
'2690461

Private Sub Form_Unload(Cancel As Integer)
   Call UnLockAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes), , , False)

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub M_ItTipoDoc_Click(Index As Integer)
   Dim Value As String
   Dim TipoDoc As Integer, TipoLib As Integer
   Dim Row As Integer
   Dim AuxTipoLib As Integer
   Dim IdxTipoDoc As Integer
   Dim i As Integer
   Dim TipoDoc2 As Integer, TipoLib2 As Integer
   
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_TIPOOPER) = "" Then
      Exit Sub
   End If
   
   Row = Grid.Row
   
   TipoDoc = Cb_TipoDoc.ItemData(Index)
   
   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      
      If lTipoOper = TOPERCAJA_INGRESO Then        '(LIB_VENTAS y LIB_CAJAING)
         If TipoDoc > BASELIBCAJA_INGEGR Then
            AuxTipoLib = LIB_CAJAING
            TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
         Else
            AuxTipoLib = LIB_VENTAS
         End If
         
      ElseIf TipoDoc > BASELIBCAJA_INGEGR Then     '(LIB_COMPRAS, LIB_RETEN y LIB_CAJAEGR)
         AuxTipoLib = LIB_CAJAEGR
         TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
      ElseIf TipoDoc > BASELIBCAJA_RETEN Then
         AuxTipoLib = LIB_RETEN
         TipoDoc = TipoDoc - BASELIBCAJA_RETEN
      Else
         AuxTipoLib = LIB_COMPRAS
      End If
      
      
'      If lTipoLib = LIB_VENTAS Or lTipoLib = LIB_CAJAING Or lTipoLib = LIB_CAJAEGR Then
'         AuxTipoLib = lTipoLib
'      ElseIf TipoDoc > BASELIBCAJA_RETEN Then    'es compras o retenciones?
'         AuxTipoLib = LIB_RETEN
'         TipoDoc = TipoDoc - BASELIBCAJA_RETEN
'      ElseIf lTipoLib = LIB_COMPRAS Then
'         AuxTipoLib = LIB_COMPRAS
'      End If


      If (TipoDoc = LIBCAJA_OTROSINGINI And AuxTipoLib = LIB_CAJAING) Or (TipoDoc = LIBCAJA_OTROSEGRINI And AuxTipoLib = LIB_CAJAEGR) Then
         If CbItemData(Cb_Mes) <> 1 Then    'estos tipos de ingresos/egresos sólo se pueden ingresar en enero
            MsgBox1 "Sólo se permite ingresar este tipo de Ingresos/Egresos en el mes de enero.", vbExclamation
            TipoDoc = 0
            AuxTipoLib = 0
         Else
            For i = Grid.FixedRows To Grid.rows - 1
            
               If Grid.RowHeight(i) > 0 Then
                  TipoDoc2 = Val(Grid.TextMatrix(i, C_IDTIPODOC))
                  TipoLib2 = Val(Grid.TextMatrix(i, C_IDTIPOLIB))
                  
                  If (TipoDoc2 = LIBCAJA_OTROSINGINI And TipoLib2 = LIB_CAJAING) Or (TipoDoc2 = LIBCAJA_OTROSEGRINI And TipoLib2 = LIB_CAJAEGR) Then
                     If i <> Row Then
                        MsgBox1 "No se perimte ingresar más de un Ingreso/Egreso Inicial en el año.", vbExclamation
                        TipoDoc = 0
                        AuxTipoLib = 0
                     End If
                  End If
               End If
               
            Next i
                     
         End If
      End If


      
      Value = GetDiminutivoDoc(AuxTipoLib, TipoDoc)
      Grid.TextMatrix(Row, C_TIPODOC) = Value
      Grid.TextMatrix(Row, C_IDTIPODOC) = TipoDoc
      Grid.TextMatrix(Row, C_IDTIPOLIB) = AuxTipoLib
      IdxTipoDoc = GetTipoDoc(AuxTipoLib, TipoDoc)
      
      If IdxTipoDoc > 0 Then
         Grid.TextMatrix(Row, C_TIPODOCEXT) = gTipoDoc(IdxTipoDoc).Nombre & IIf(Grid.TextMatrix(Row, C_DTE) <> "", " E", "")
         Grid.TextMatrix(Row, C_DOCIMPEXP) = CInt(gTipoDoc(IdxTipoDoc).DocImpExp)
         Grid.TextMatrix(Row, C_ESREBAJA) = Int(gTipoDoc(IdxTipoDoc).EsRebaja)
      Else
         Grid.TextMatrix(Row, C_TIPODOCEXT) = ""
         Grid.TextMatrix(Row, C_DOCIMPEXP) = ""
         Grid.TextMatrix(Row, C_ESREBAJA) = ""
      End If
      
      Call ActCambioTipoDoc(Row)
               
      Call CalcTot
     
      If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VENTASINDOC Then
         Grid.TextMatrix(Row, C_NUMDOC) = ""
      
      ElseIf Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Then
         If Not lMsgNotaCred Then
            MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
            lMsgNotaCred = True
         End If
      End If
      
      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
      lModifica = True
   End If

End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   
   If valida() Then
      Call SaveGrid(Grid)
      lRc = vbOK
      Unload Me
   End If
      
End Sub

Private Sub Bt_TipoDoc_Click()

   If Grid.TextMatrix(Grid.Row, C_TIPOOPER) = "" Or Not ValidaEstadoEdit(Grid.Row) Or Val(Grid.TextMatrix(Grid.Row, C_IDDOC)) <> 0 Then
      Exit Sub
   End If
   
'   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      If M_ItTipoDoc.Count > 1 Then
         Call PopupMenu(M_TipoDoc, , Grid.FlxGrid.ColPos(Grid.Col) + Grid.Left + 200, Grid.FlxGrid.RowPos(Grid.Row) + Grid.Top + 100)
      End If
'   End If
   
End Sub

Private Sub Form_Load()

   lInLoad = True
   
   lViewRes = False

   If lTipoOper = TOPERCAJA_INGRESO Then
      lTipoLib = LIB_VENTAS
   Else
      lTipoLib = LIB_COMPRAS + LIB_RETEN
   End If
   
   If gAppCode.Demo Then
    Bt_ImportTotalDocs.Enabled = False
   End If
   
   lMsgHayDocsLibCompasVentas = Val(GetIniString(gIniFile, "Msg", "LibCajaDocs", "0"))

   Call SetUpGrid(Grid)
   Call SetUpGrid(GridAnual)
   
   Fr_Opciones.visible = False
   Ch_ViewOper = gVarIniFile.VerLCajaOper
   Ch_ViewDTE = gVarIniFile.VerLCajaDTE
   Ch_ViewNombre = gVarIniFile.VerLCajaNombre
   Ch_ViewIVAIrrec = gVarIniFile.VerLCajaIVAIrrec
   Ch_ViewOtrosImp = gVarIniFile.VerLCajaOtrosImp
   
   '2690461
   Ch_ViewNotaCred = gVarIniFile.VerLCajaNotCred
   '2690461

   Bt_ImportOIngEg.visible = gFunciones.OtrosIngEgresos
   
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Me.Caption = Me.Caption & " " & gTipoOperCaja(lTipoOper) & "s - " & gNomMes(lMes) & " " & lAno     'en este caso lTipoOper no vale nunca cero
      Fr_List.visible = False
      Fr_List.Enabled = False
      Grid.Height = Grid.Height + Grid.Top - Fr_List.Top + 50
      Grid.Top = Fr_List.Top + 60
      
      If lOper = O_EDIT Then
         If lTipoOper = TOPERCAJA_INGRESO Then
            Bt_ImportDocs.Caption = "Traer Docs. Ventas"
            Bt_ImportDocs.ToolTipText = "Traer/Actualizar Docs. desde Libro de Ventas"
            Bt_ImportOIngEg.Caption = "Traer Otros Ingresos"
            Bt_ImportOIngEg.ToolTipText = "Traer/Actualizar Otros Ingresos"
            Bt_ImportTotalDocs.Caption = "Traer Docs. Ventas Anual"
            Bt_ImportTotalDocs.ToolTipText = "Traer/Actualizar Docs. desde Libro de Ventas Anual"
            Bt_ImportOIngEgAnual.Caption = "Traer Otros Ingresos Anual"
            Bt_ImportOIngEgAnual.ToolTipText = "Traer/Actualizar Otros Ingresos Anual"
         Else
            Bt_ImportDocs.Caption = "Traer Docs.Compras/Ret."
            Bt_ImportDocs.ToolTipText = "Traer Docs. Libro de Compras y Retenciones"
            Bt_ImportOIngEg.Caption = "Traer Otros Egresos"
            Bt_ImportOIngEg.ToolTipText = "Traer/Actualizar Otros Egresos"
            Bt_ImportTotalDocs.Caption = "Traer Docs.Compras/Ret. Anual"
            Bt_ImportTotalDocs.ToolTipText = "Traer Docs. Libro de Compras y Retenciones Anual"
            Bt_ImportOIngEgAnual.Caption = "Traer Otros Egresos Anual"
            Bt_ImportOIngEgAnual.ToolTipText = "Traer/Actualizar Otros Egresos Anual"
         End If
         Bt_Opciones.Caption = "Opciones de Edición"

      End If
      
   ElseIf lOper = O_VIEW Then
      Me.Caption = Me.Caption & " Consolidado"
      Bt_Opciones.Caption = "Opciones de Vista"
   
   End If
   
   If lOper = O_VIEW Or lOper = O_VIEWLIBLEGAL Then
      Grid.Locked = True
      Bt_Cancel.Caption = "Cerrar"
      Bt_OK.visible = False
      Fr_BtEdit.visible = False
      Fr_BtGen.Left = Fr_BtEdit.Left
      
      Bt_ImportDocs.visible = False
      Bt_ImportOIngEg.visible = False
      Bt_ImportTotalDocs.visible = False
      Me.Bt_ImportOIngEgAnual.visible = False
      Ch_Rut = 1
   End If
   
   If gFunciones.DocCuotas Then
      Bt_DocCuotas.visible = True
      Bt_DocCuotas.Enabled = True
   End If
   

   Set lcbNombre = New ClsCombo
   
   Call lcbNombre.SetControl(Cb_Nombre)
  
   Call SetTxRO(Tx_CurrCell, True)
     
   Call FillCb
   
   Call SetOrderLst
      
   Call SetupPriv
   
   Call LoadGrid
   
   lInLoad = False
   
   Call LoadTipoDoc     'debe estar después de FillCb y después de LoadGrid y después de lInLoad = false

   
   
End Sub
Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un documento a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_TIPOOPER) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

'   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then   'permitimos eliminar, total lo puede reimportar si lo desea
'      MsgBeep vbExclamation
'      Exit Sub
'   End If

   If MsgBox1("¿Está seguro que desea borrar este documento?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE)
   lModifica = True
      
   Grid.rows = Grid.rows + 1
   lHayDocsLibComprasVentas = False
    
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then   'vemos si quedan docs con IdDoc > 0
      
      For i = Grid.FixedRows To Grid.rows - 1
      
         If Grid.TextMatrix(i, C_TIPOOPER) = "" Then    'ya terminó la lista de docs.
            Exit For
         End If
         
         If Grid.RowHeight(i) > 0 And Val(Grid.TextMatrix(i, C_IDDOC)) > 0 Then
            lHayDocsLibComprasVentas = True
            Exit For
         End If
      Next i
   End If
   
   Call CalcTot
End Sub
Private Sub LoadGrid(Optional ByVal Row As Integer = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rs1 As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EditEnable As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer
   Dim TipoLib As Integer
   Dim SaldoInicial As Boolean
   Dim IdxTipoDoc As Integer
   
   '2690461
    Dim Q2 As String
   Dim Rs2 As Recordset
   Dim MontoNotaCred As Double
   '2690461
   
   Grid.FlxGrid.Redraw = False
   
   If CbItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
   
      If CbItemData(Cb_TipoOper) > 0 Then
         
         lTipoOper = CbItemData(Cb_TipoOper)
         TipoDoc = CbItemData(Cb_TipoDoc)
         
         If TipoDoc > 0 Then
         
            If lTipoOper = TOPERCAJA_INGRESO Then        '(LIB_VENTAS y LIB_CAJAING)
               If TipoDoc > BASELIBCAJA_INGEGR Then
                  TipoLib = LIB_CAJAING
                  TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
               Else
                  TipoLib = LIB_VENTAS
               End If
               
               'Egresos
            ElseIf TipoDoc > BASELIBCAJA_INGEGR Then     '(LIB_COMPRAS, LIB_RETEN y LIB_CAJAEGR)
               TipoLib = LIB_CAJAEGR
               TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
            ElseIf TipoDoc > BASELIBCAJA_RETEN Then
               TipoLib = LIB_RETEN
               TipoDoc = TipoDoc - BASELIBCAJA_RETEN
            Else
               TipoLib = LIB_COMPRAS
            End If
            
            Where = Where & " AND LibroCaja.TipoLib = " & TipoLib & " AND LibroCaja.TipoDoc = " & TipoDoc
         End If
         
      Else
         lTipoOper = 0
         TipoDoc = 0
         TipoLib = 0
      End If
            
      If Trim(Tx_Rut) <> "" Then
         IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
         If IdEnt > 0 Then
            Where = Where & " AND LibroCaja.IdEntidad = " & IdEnt
         Else
            Tx_Rut = ""
            Cb_Entidad.ListIndex = 0
            Cb_Nombre.ListIndex = 0
         End If
      
      End If
      
      If Trim(Tx_Glosa) <> "" Then
         Where = Where & " AND " & GenLike(DbMain, Tx_Glosa, "LibroCaja.Descrip", 3)
      End If
      
      If vFmt(Tx_Valor) <> 0 Then
         Where = Where & " AND LibroCaja.Afecto = " & vFmt(Tx_Valor)
      End If
      
      If Trim(Tx_NumDoc) <> "" Then
         Where = Where & " AND LibroCaja.NumDoc = '" & Trim(Tx_NumDoc) & "'"
      End If
      
   End If
   
   If Row > 0 Then
      Where = Where & " AND IdLibroCaja=" & Val(Grid.TextMatrix(Row, C_IDLIBROCAJA))
   End If
   
   If lTipoOper = 0 And Where = "" And Row = 0 And Ch_SaldoInicial <> 0 Then
      SaldoInicial = True
   End If
   
   Q1 = "SELECT IdLibroCaja, IdDoc, TipoOper, LibroCaja.TipoDoc, LibroCaja.TipoLib, NumDoc, DTE, NumDocHasta, LibroCaja.IdEntidad, LibroCaja.RutEntidad "
   Q1 = Q1 & " , LibroCaja.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre, Entidades.EntRelacionada as EntsEntRelacionada, FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec "
   Q1 = Q1 & " , Exento, OtroImp, Total, Pagado, Descrip, ConEntRel, OperDevengada, PagoAPlazo, FechaExigPago, LibroCaja.IdUsuario, Usuarios.Usuario, LibroCaja.Estado, LibroCaja.IdEntReal, IdComp "
   'Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, MontoAfectaBaseImp "
   
   '2699582
   'Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, iif( Afecto + Exento < Pagado, Afecto + Exento, Pagado ) as MontoAfectaBaseImp "
    ''Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, iif(Entidades.EntRelacionada = -1, afecto + exento,  iif( Afecto + Exento < Pagado, Afecto + Exento, Pagado ) ) as MontoAfectaBaseImp "
   'fin 2699582
   
   '2841464
   Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, "
   
   'Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) ,  iif(iif(OtroImp < 0,Afecto + Exento +OtroImp  < Pagado,Afecto + Exento < Pagado) ,iif(OtroImp<0,Afecto + "
   
   'Q1 = Q1 & " Exento+OtroImp,Afecto + Exento ) , Pagado ) ) as MontoAfectaBaseImp "
   
   'fin 2841464
   
   
   '2882638
   Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento), iif(OtroImp< 0 and afecto + exento +OtroImp < Pagado,afecto + exento,"
   Q1 = Q1 & "iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) ) ) as MontoAfectaBaseImp    "
   'fin 2882638
   
   
   '2690461
   'Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) , Pagado ) as MontoAfectaBaseImp   "
      
   'fin 2690461
   
   Q1 = Q1 & " FROM ((LibroCaja LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON LibroCaja.IdUsuario = Usuarios.IdUsuario )"
   'Q1 = Q1 & " INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   If lTipoOper > 0 Then
      Q1 = Q1 & " AND LibroCaja.TipoOper = " & lTipoOper
   End If
   Q1 = Q1 & Where
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Q1 = Q1 & " ORDER BY iif((LibroCaja.TipoDoc = " & LIBCAJA_OTROSINGINI & " AND LibroCaja.TipoLib = " & LIB_CAJAING & ") OR (LibroCaja.TipoDoc = " & LIBCAJA_OTROSEGRINI & " AND LibroCaja.TipoLib = " & LIB_CAJAEGR & "), 0, 1), " & lOrdenGr(lOrdenSel)
      
   Set Rs = OpenRs(DbMain, Q1)

   If Row <= 0 Then
      Grid.rows = Grid.FixedRows
      i = Grid.FixedRows
      
      If lOper = O_VIEW And SaldoInicial Then
         Grid.rows = Grid.rows + 1
         
         Call InsertSaldoInicial(i, FirstDay)
         i = i + 1
           
      End If

   Else
      i = Row
   End If
   
   Do While Rs.EOF = False
      
      If Row <= 0 Then
         Grid.rows = Grid.rows + 1
                  
         If gAppCode.Demo Then
            If Grid.rows > MAX_DOCDEMO Then
               MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
               Exit Do
            End If
         End If
        
      End If
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If

      
      Grid.TextMatrix(i, C_IDLIBROCAJA) = vFld(Rs("IdLibroCaja"))
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
      
      If vFld(Rs("IdDoc")) <> 0 Then   'tiene que estar antes del C_CHECK por el color de la celda
         Call FGrSetRowStyle(Grid, i, "FC", COLOR_AZULOSCURO)
         lHayDocsLibComprasVentas = True
      End If
      
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1
      Grid.TextMatrix(i, C_TIPOOPER) = UCase(Left(gTipoOperCaja(vFld(Rs("TipoOper"))), 1))
      Grid.TextMatrix(i, C_IDTIPOOPER) = vFld(Rs("TipoOper"))
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(i, C_NUMDOCHASTA) = vFld(Rs("NumDochasta"))
      Grid.TextMatrix(i, C_IDTIPOLIB) = vFld(Rs("TipoLib"))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_IDTIPODOC) = vFld(Rs("TipoDoc"))
      If vFld(Rs("DTE")) <> 0 Then
         Grid.TextMatrix(i, C_DTE) = "x"
      Else
         Grid.TextMatrix(i, C_DTE) = ""
      End If
      IdxTipoDoc = GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      If IdxTipoDoc >= 0 Then
         Grid.TextMatrix(i, C_TIPODOCEXT) = gTipoDoc(IdxTipoDoc).Nombre & IIf(Grid.TextMatrix(i, C_DTE) <> "", " E", "")
         Grid.TextMatrix(i, C_DOCIMPEXP) = CInt(gTipoDoc(IdxTipoDoc).DocImpExp)
      End If
      
'      If Grid.TextMatrix(i, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
'         Grid.TextMatrix(i, C_NUMFISCIMPR) = vFld(Rs("NumFiscImpr"))
'         Grid.TextMatrix(i, C_NUMINFORMEZ) = vFld(Rs("NumInformeZ"))
'      End If
      
'      If Grid.TextMatrix(i, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
'         Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
'         Grid.TextMatrix(i, C_NUMDOCHASTA) = IIf(Val(vFld(Rs("NumDocHasta"))) <> 0, vFld(Rs("NumDocHasta")), "")
'      End If
            
      Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      
      If vFld(Rs("Estado")) <> ED_ANULADO Then
      
         If vFld(Rs("IdEntidad")) = 0 Then
            If vFld(Rs("RutEntidad")) <> "" And vFld(Rs("RutEntidad")) <> "0" Then
'               Grid.TextMatrix(i, C_RUT) = IIf(Val(Grid.TextMatrix(i, C_DOCIMPEXP)) = 0, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
               Grid.TextMatrix(i, C_RUT) = IIf(Val(Grid.TextMatrix(i, C_DOCIMPEXP)) = 0 And UCase(vFld(Rs("RutEntidad"))) <> RUT_VARIOS, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
'               If vFld(Rs("RutEntidad")) = gEmpresa.Rut Then
'                  Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False)
'               Else
'                  Grid.TextMatrix(i, C_RUT) = IIf(Val(Grid.TextMatrix(i, C_DOCIMPEXP)) = 0 And UCase(vFld(Rs("RutEntidad"))) <> RUT_VARIOS, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
'               End If
               Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("NombreEntidad"), True)
            End If
         Else
            If vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0" Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
               Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
'               If vFld(Rs("IdDoc")) > 0 Then       'es importado
'                  Grid.TextMatrix(i, C_CONENTREL) = IIf(vFld(Rs("ConEntRel")) <> 0, "x", "")
'               End If
            End If
         End If
      
      Else
         Grid.TextMatrix(i, C_RUT) = ""
         Grid.TextMatrix(i, C_NOMBRE) = "NULO"
         
      End If
      
'      Grid.TextMatrix(i, C_IDENTREAL) = vFld(Rs("IdEntReal"))    ya no es necesaria esta columna ya que el SII recapacitó y en las ventas se usa el rut de receptor, en vez del emisor (que es siempre el RUT de la empresa actual)
      Grid.TextMatrix(i, C_FECHAOPER) = Format(vFld(Rs("FechaOperacion")), EDATEFMT)
      Grid.TextMatrix(i, C_LNGFECHAOPER) = vFld(Rs("FechaOperacion"))
      Grid.TextMatrix(i, C_LNGFECHAINGRESOLIBRO) = vFld(Rs("FechaIngresoLibro"))
      
      If vFld(Rs("FechaExigPago")) > 0 Then
         Grid.TextMatrix(i, C_FECHAEXIGPAGO) = Format(vFld(Rs("FechaExigPago")), EDATEFMT)
         Grid.TextMatrix(i, C_LNGFECHAEXIGPAGO) = vFld(Rs("FechaExigPago"))
      End If
      
      If IdxTipoDoc >= 0 Then
         EsRebaja = gTipoDoc(IdxTipoDoc).EsRebaja
      End If
      Grid.TextMatrix(i, C_ESREBAJA) = IIf(EsRebaja, 1, 0)
            
'      If EsRebaja Then                             'nota de crédito = valores negativos
'         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVAIRREC) = Format(vFld(Rs("IVAIrrec")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_OTROIMP) = Format(vFld(Rs("OtroImp")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")) * -1, NEGNUMFMT)
'
'      Else
         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")), NEGNUMFMT)
         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")), NEGNUMFMT)
         Grid.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")), NEGNUMFMT)
         Grid.TextMatrix(i, C_IVAIRREC) = Format(vFld(Rs("IVAIrrec")), NEGNUMFMT)
         Grid.TextMatrix(i, C_OTROIMP) = Format(vFld(Rs("OtroImp")), NEGNUMFMT)
         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NEGNUMFMT)
         
'      End If
      
      
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      
      Grid.TextMatrix(i, C_PAGADO) = Format(vFld(Rs("Pagado")), NEGNUMFMT)
      
      '2690461
      Q2 = ""
     '2970116
      'Q2 = "select sum(afecto) as monto from documento "
      Q2 = "select sum(afecto + exento +OtroImp) as monto from documento "
      '2970116
      Q2 = Q2 & " Where tipoDoc = 3 and TipoLib = " & vFld(Rs("Tipolib"))
      Q2 = Q2 & " and IdDocAsoc = " & vFld(Rs("IdDoc"))

      Set Rs2 = OpenRs(DbMain, Q2)

      MontoNotaCred = 0

      If Rs2.EOF = False Then

        If vFld(Rs2("monto")) <> 0 Then
           MontoNotaCred = vFld(Rs2("monto"))
           Grid.TextMatrix(i, C_NOTACRED) = "x"
        Else
           Grid.TextMatrix(i, C_NOTACRED) = ""
        End If
      End If
      Call CloseRs(Rs2)
      'fin 2690461
      
      
      '2752418
      If vFld(Rs("OtroImp")) > 0 Then
      
        Q1 = "SELECT count(*) as resultado "
        Q1 = Q1 & "FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc) INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.idCuenta "
        Q1 = Q1 & " where numdoc = '" & vFld(Rs("NumDoc")) & "'"
        Q1 = Q1 & " and MovDocumento.debe = " & vFld(Rs("OtroImp"))
        Q1 = Q1 & " and Cuentas.codigo like '3*'  " ' num 3 corresponde a cuentas tipo 3 ejem:3010101
           
        Set Rs1 = OpenRs(DbMain, Q1)
        
        If Rs1.EOF = False Then
        '2896692
         If vFld(Rs("Pagado")) <> "0" Then
        '2896692
          If vFld(Rs1("resultado")) > 0 Then
             
             '2690461
             If (vFld(Rs("Afecto")) + vFld(Rs("exento")) - MontoNotaCred) < vFld(Rs("Pagado")) Then

             Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento")) - MontoNotaCred), NEGNUMFMT)

             Else

             Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento"))), NEGNUMFMT)

             End If
             
             '2690461
          
             'Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento"))), NEGNUMFMT)
             
             
          Else
             Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)
          End If
        End If
        
        Else
            Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(0, NEGNUMFMT)
        End If
        
       Call CloseRs(Rs1)
      Else
      '2841464
           'If vFld(Rs("OtroImp")) < 0 Then
           ' Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")) + vFld(Rs("OtroImp")), NEGNUMFMT)
    
           'Else
           
           ''Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)
           'End If
      'fin 2841464
      
       
       '2896692
       If vFld(Rs("Pagado")) <> "0" Then
       'fin 2896692
       
       '2690461
             If (vFld(Rs("Afecto")) + vFld(Rs("exento")) - MontoNotaCred) < vFld(Rs("Pagado")) Then

             Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format((vFld(Rs("Afecto")) + vFld(Rs("exento")) - MontoNotaCred), NEGNUMFMT)

             Else

             Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)

             End If
'
             '2690461
       Else
        Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(0, NEGNUMFMT)
       End If
       '2896692
      End If
      
      'fin 2752418
        
      
      If gEmpresa.Ano < 2020 Then
   '      If vFld(Rs("IdDoc")) = 0 Then       'NO es importado
            If vFld(Rs("ConEntRel")) <> 0 Then
               Grid.TextMatrix(i, C_CONENTREL) = "x"
            Else
               Grid.TextMatrix(i, C_CONENTREL) = ""
            End If
   '      End If
      Else
         If vFld(Rs("IdEntidad")) > 0 Then
            If ValidaEnt14D(vFld(Rs("IdEntidad"))) Then
               If vFld(Rs("ConEntRel")) <> 0 Then
                  Grid.TextMatrix(i, C_CONENTREL) = "x"
               Else
                  Grid.TextMatrix(i, C_CONENTREL) = ""
               End If
            End If
            
'         ElseIf vFld(Rs("IdEntReal")) > 0 Then
'            If ValidaEnt14D(vFld(Rs("IdEntReal"))) Then
'               If vFld(Rs("ConEntRel")) <> 0 Then
'                  Grid.TextMatrix(i, C_CONENTREL) = "x"
'               Else
'                  Grid.TextMatrix(i, C_CONENTREL) = ""
'               End If
'            End If
         End If
      End If

      If vFld(Rs("OperDevengada")) <> 0 Then
         Grid.TextMatrix(i, C_OPERDEVENGADA) = "x"
      Else
         Grid.TextMatrix(i, C_OPERDEVENGADA) = ""
      End If
     
      If vFld(Rs("PagoAPlazo")) <> 0 Then
         Grid.TextMatrix(i, C_PAGOAPLAZO) = "x"
      Else
         Grid.TextMatrix(i, C_PAGOAPLAZO) = ""
      End If
     
      Grid.TextMatrix(i, C_FECHAEXIGPAGO) = IIf(vFld(Rs("FechaExigPago")) > 0, Format(vFld(Rs("FechaExigPago")), EDATEFMT), "")
      
      Call CalcSaldo(i, Grid)
      
      Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
               
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   If Row <= 0 Then
      Call FGrVRows(Grid)
      Grid.rows = Grid.rows + 1
      Grid.TopRow = Grid.FixedRows
      
      'Marco la columna Ordenada
      Grid.Row = 0
      Grid.Col = lOrdenSel
      Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Else
      Grid.Row = Row
      Grid.Col = C_TIPODOC
   
   End If
   
   Tx_CurrCell = ""
      
   Call CalcTot
   
   Grid.FlxGrid.Redraw = True
   
   If lOper = O_EDIT Then
      Call UnLockAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes), , , False)
   
      EditEnable = LockAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes))
   
      If EditEnable = False Then    'alguien más lo está editando, no podemos editarlo (esto se hace sólo una vez ya que en lOper = O_EDIT no se puede cambiar el mes)
         MsgBox1 "El Libro de Caja del mes de " & gNomMes(CbItemData(Cb_Mes)) & " se está editando en el equipo '" & IsLockedAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes)) & "'. Sólo se abrirá de lectura.", vbInformation
         lEditEnabled = False
      End If
   
      Call EnableForm(Me, lEditEnabled)
      Call SetTxRO(Tx_CurrCell, True)
   End If
   
   Bt_List.Enabled = False
   
End Sub
Private Sub LoadGridAnual(Mes As Integer, Optional ByVal Row As Integer = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rs1 As Recordset
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EditEnable As Boolean
   Dim EsRebaja As Boolean
   Dim TipoDoc As Integer
   Dim TipoLib As Integer
   Dim SaldoInicial As Boolean
   Dim IdxTipoDoc As Integer
   
   '2690461
    Dim Q2 As String
   Dim Rs2 As Recordset
   Dim MontoNotaCred As Double
   '2690461
   
   
   GridAnual.FlxGrid.Redraw = False
   
   If CbItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), Mes, 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
   
      If CbItemData(Cb_TipoOper) > 0 Then
         
         lTipoOper = CbItemData(Cb_TipoOper)
         TipoDoc = CbItemData(Cb_TipoDoc)
         
         If TipoDoc > 0 Then
         
            If lTipoOper = TOPERCAJA_INGRESO Then        '(LIB_VENTAS y LIB_CAJAING)
               If TipoDoc > BASELIBCAJA_INGEGR Then
                  TipoLib = LIB_CAJAING
                  TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
               Else
                  TipoLib = LIB_VENTAS
               End If
               
               'Egresos
            ElseIf TipoDoc > BASELIBCAJA_INGEGR Then     '(LIB_COMPRAS, LIB_RETEN y LIB_CAJAEGR)
               TipoLib = LIB_CAJAEGR
               TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
            ElseIf TipoDoc > BASELIBCAJA_RETEN Then
               TipoLib = LIB_RETEN
               TipoDoc = TipoDoc - BASELIBCAJA_RETEN
            Else
               TipoLib = LIB_COMPRAS
            End If
            
            Where = Where & " AND LibroCaja.TipoLib = " & TipoLib & " AND LibroCaja.TipoDoc = " & TipoDoc
         End If
         
      Else
         lTipoOper = 0
         TipoDoc = 0
         TipoLib = 0
      End If
            
      If Trim(Tx_Rut) <> "" Then
         IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
         If IdEnt > 0 Then
            Where = Where & " AND LibroCaja.IdEntidad = " & IdEnt
         Else
            Tx_Rut = ""
            Cb_Entidad.ListIndex = 0
            Cb_Nombre.ListIndex = 0
         End If
      
      End If
      
      If Trim(Tx_Glosa) <> "" Then
         Where = Where & " AND " & GenLike(DbMain, Tx_Glosa, "LibroCaja.Descrip", 3)
      End If
      
      If vFmt(Tx_Valor) <> 0 Then
         Where = Where & " AND LibroCaja.Afecto = " & vFmt(Tx_Valor)
      End If
      
      If Trim(Tx_NumDoc) <> "" Then
         Where = Where & " AND LibroCaja.NumDoc = '" & Trim(Tx_NumDoc) & "'"
      End If
      
   End If
   
   If Row > 0 Then
      Where = Where & " AND IdLibroCaja=" & Val(GridAnual.TextMatrix(Row, C_IDLIBROCAJA))
   End If
   
   If lTipoOper = 0 And Where = "" And Row = 0 And Ch_SaldoInicial <> 0 Then
      SaldoInicial = True
   End If
   
   'Q1 = "SELECT IdLibroCaja, IdDoc, TipoOper, LibroCaja.TipoDoc, LibroCaja.TipoLib, NumDoc, DTE, NumDocHasta, LibroCaja.IdEntidad, LibroCaja.RutEntidad "
   'Q1 = Q1 & " , LibroCaja.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre, Entidades.EntRelacionada as EntsEntRelacionada, FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec "
   'Q1 = Q1 & " , Exento, OtroImp, Total, Pagado, Descrip, ConEntRel, OperDevengada, PagoAPlazo, FechaExigPago, LibroCaja.IdUsuario, Usuarios.Usuario, LibroCaja.Estado, LibroCaja.IdEntReal, IdComp "
   'Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, MontoAfectaBaseImp "
   
   
    Q1 = "SELECT IdLibroCaja, IdDoc, TipoOper, LibroCaja.TipoDoc, LibroCaja.TipoLib, NumDoc, DTE, NumDocHasta, LibroCaja.IdEntidad, LibroCaja.RutEntidad "
   Q1 = Q1 & " , LibroCaja.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre, Entidades.EntRelacionada as EntsEntRelacionada, FechaOperacion, FechaIngresoLibro, Afecto, IVA, IVAIrrec "
   Q1 = Q1 & " , Exento, OtroImp, Total, Pagado, Descrip, ConEntRel, OperDevengada, PagoAPlazo, FechaExigPago, LibroCaja.IdUsuario, Usuarios.Usuario, LibroCaja.Estado, LibroCaja.IdEntReal, IdComp "
   'Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, MontoAfectaBaseImp "
   
   '2699582
   'Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, iif( Afecto + Exento < Pagado, Afecto + Exento, Pagado ) as MontoAfectaBaseImp "
    ''Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, iif(Entidades.EntRelacionada = -1, afecto + exento,  iif( Afecto + Exento < Pagado, Afecto + Exento, Pagado ) ) as MontoAfectaBaseImp "
   'fin 2699582
   
   '2841464
   Q1 = Q1 & " , iif(" & SqlMonthLng("FechaOperacion") & "=" & CbItemData(Cb_Mes) & ",0,1) as MesActual, "
'   Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) ,  iif(iif(OtroImp < 0,Afecto + Exento +OtroImp  < Pagado,Afecto + Exento < Pagado) ,iif(OtroImp<0,Afecto + "
'   Q1 = Q1 & " Exento+OtroImp,Afecto + Exento ) , Pagado ) ) as MontoAfectaBaseImp "
   'fin 2841464
   
   
    '2882638
   Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento), iif(OtroImp< 0 and afecto + exento +OtroImp < Pagado,afecto + exento,"
   Q1 = Q1 & "iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) ) ) as MontoAfectaBaseImp    "
   'fin 2882638
   
   
    '2690461
   'Q1 = Q1 & " iif(Entidades.EntRelacionada = -1,iif(OtroImp< 0,afecto + exento +OtroImp ,afecto + exento) , Pagado ) as MontoAfectaBaseImp   "
      
   'fin 2690461
   
   Q1 = Q1 & " FROM ((LibroCaja LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON LibroCaja.IdUsuario = Usuarios.IdUsuario )"
'   Q1 = Q1 & " INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   If lTipoOper > 0 Then
      Q1 = Q1 & " AND LibroCaja.TipoOper = " & lTipoOper
   End If
   Q1 = Q1 & Where
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Q1 = Q1 & " ORDER BY iif((LibroCaja.TipoDoc = " & LIBCAJA_OTROSINGINI & " AND LibroCaja.TipoLib = " & LIB_CAJAING & ") OR (LibroCaja.TipoDoc = " & LIBCAJA_OTROSEGRINI & " AND LibroCaja.TipoLib = " & LIB_CAJAEGR & "), 0, 1), " & lOrdenGr(lOrdenSel)
      
   Set Rs = OpenRs(DbMain, Q1)

   If Row <= 0 Then
      GridAnual.rows = GridAnual.FixedRows
      i = GridAnual.FixedRows
      
      If lOper = O_VIEW And SaldoInicial Then
         GridAnual.rows = GridAnual.rows + 1
         
         Call InsertSaldoInicial(i, FirstDay)
         i = i + 1
           
      End If

   Else
      i = Row
   End If
   
   Do While Rs.EOF = False
      
      If Row <= 0 Then
         GridAnual.rows = GridAnual.rows + 1
                  
         If gAppCode.Demo Then
            If GridAnual.rows > MAX_DOCDEMO Then
               MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
               Exit Do
            End If
         End If
        
      End If
      
      If FGrChkMaxSize(GridAnual) = True Then
         Exit Do
      End If

      
      GridAnual.TextMatrix(i, C_IDLIBROCAJA) = vFld(Rs("IdLibroCaja"))
      GridAnual.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
      
      
      If vFld(Rs("IdDoc")) <> 0 Then   'tiene que estar antes del C_CHECK por el color de la celda
         Call FGrSetRowStyle(GridAnual, i, "FC", COLOR_AZULOSCURO)
         lHayDocsLibComprasVentas = True
      End If
      
      GridAnual.TextMatrix(i, C_NUMLIN) = i - GridAnual.FixedRows + 1
      GridAnual.TextMatrix(i, C_TIPOOPER) = UCase(Left(gTipoOperCaja(vFld(Rs("TipoOper"))), 1))
      GridAnual.TextMatrix(i, C_IDTIPOOPER) = vFld(Rs("TipoOper"))
      GridAnual.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      GridAnual.TextMatrix(i, C_NUMDOCHASTA) = vFld(Rs("NumDochasta"))
      GridAnual.TextMatrix(i, C_IDTIPOLIB) = vFld(Rs("TipoLib"))
      GridAnual.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      GridAnual.TextMatrix(i, C_IDTIPODOC) = vFld(Rs("TipoDoc"))
      If vFld(Rs("DTE")) <> 0 Then
         GridAnual.TextMatrix(i, C_DTE) = "x"
      Else
         GridAnual.TextMatrix(i, C_DTE) = ""
      End If
      IdxTipoDoc = GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      If IdxTipoDoc >= 0 Then
         GridAnual.TextMatrix(i, C_TIPODOCEXT) = gTipoDoc(IdxTipoDoc).Nombre & IIf(GridAnual.TextMatrix(i, C_DTE) <> "", " E", "")
         GridAnual.TextMatrix(i, C_DOCIMPEXP) = CInt(gTipoDoc(IdxTipoDoc).DocImpExp)
      End If
      
'      If GridAnual.TextMatrix(i, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
'         GridAnual.TextMatrix(i, C_NUMFISCIMPR) = vFld(Rs("NumFiscImpr"))
'         GridAnual.TextMatrix(i, C_NUMINFORMEZ) = vFld(Rs("NumInformeZ"))
'      End If
      
'      If GridAnual.TextMatrix(i, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
'         GridAnual.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
'         GridAnual.TextMatrix(i, C_NUMDOCHASTA) = IIf(Val(vFld(Rs("NumDocHasta"))) <> 0, vFld(Rs("NumDocHasta")), "")
'      End If
            
      GridAnual.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      
      If vFld(Rs("Estado")) <> ED_ANULADO Then
      
         If vFld(Rs("IdEntidad")) = 0 Then
            If vFld(Rs("RutEntidad")) <> "" And vFld(Rs("RutEntidad")) <> "0" Then
'               GridAnual.TextMatrix(i, C_RUT) = IIf(Val(GridAnual.TextMatrix(i, C_DOCIMPEXP)) = 0, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
               GridAnual.TextMatrix(i, C_RUT) = IIf(Val(GridAnual.TextMatrix(i, C_DOCIMPEXP)) = 0 And UCase(vFld(Rs("RutEntidad"))) <> RUT_VARIOS, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
'               If vFld(Rs("RutEntidad")) = gEmpresa.Rut Then
'                  GridAnual.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False)
'               Else
'                  GridAnual.TextMatrix(i, C_RUT) = IIf(Val(GridAnual.TextMatrix(i, C_DOCIMPEXP)) = 0 And UCase(vFld(Rs("RutEntidad"))) <> RUT_VARIOS, FmtCID(vFld(Rs("RutEntidad")), vFld(Rs("NotValidRut")) = False), vFld(Rs("RutEntidad")))
'               End If
               GridAnual.TextMatrix(i, C_NOMBRE) = vFld(Rs("NombreEntidad"), True)
            End If
         Else
            If vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0" Then
               GridAnual.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
               GridAnual.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
'               If vFld(Rs("IdDoc")) > 0 Then       'es importado
'                  GridAnual.TextMatrix(i, C_CONENTREL) = IIf(vFld(Rs("ConEntRel")) <> 0, "x", "")
'               End If
            End If
         End If
      
      Else
         GridAnual.TextMatrix(i, C_RUT) = ""
         GridAnual.TextMatrix(i, C_NOMBRE) = "NULO"
         
      End If
      
'      GridAnual.TextMatrix(i, C_IDENTREAL) = vFld(Rs("IdEntReal"))    ya no es necesaria esta columna ya que el SII recapacitó y en las ventas se usa el rut de receptor, en vez del emisor (que es siempre el RUT de la empresa actual)
      GridAnual.TextMatrix(i, C_FECHAOPER) = Format(vFld(Rs("FechaOperacion")), EDATEFMT)
      GridAnual.TextMatrix(i, C_LNGFECHAOPER) = vFld(Rs("FechaOperacion"))
      GridAnual.TextMatrix(i, C_LNGFECHAINGRESOLIBRO) = vFld(Rs("FechaIngresoLibro"))
      
      If vFld(Rs("FechaExigPago")) > 0 Then
         GridAnual.TextMatrix(i, C_FECHAEXIGPAGO) = Format(vFld(Rs("FechaExigPago")), EDATEFMT)
         GridAnual.TextMatrix(i, C_LNGFECHAEXIGPAGO) = vFld(Rs("FechaExigPago"))
      End If
      
      If IdxTipoDoc >= 0 Then
         EsRebaja = gTipoDoc(IdxTipoDoc).EsRebaja
      End If
      GridAnual.TextMatrix(i, C_ESREBAJA) = IIf(EsRebaja, 1, 0)
            
'      If EsRebaja Then                             'nota de crédito = valores negativos
'         GridAnual.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")) * -1, NEGNUMFMT)
'         GridAnual.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")) * -1, NEGNUMFMT)
'         GridAnual.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")) * -1, NEGNUMFMT)
'         GridAnual.TextMatrix(i, C_IVAIRREC) = Format(vFld(Rs("IVAIrrec")) * -1, NEGNUMFMT)
'         GridAnual.TextMatrix(i, C_OTROIMP) = Format(vFld(Rs("OtroImp")) * -1, NEGNUMFMT)
'         GridAnual.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")) * -1, NEGNUMFMT)
'
'      Else
         GridAnual.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")), NEGNUMFMT)
         GridAnual.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")), NEGNUMFMT)
         GridAnual.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")), NEGNUMFMT)
         GridAnual.TextMatrix(i, C_IVAIRREC) = Format(vFld(Rs("IVAIrrec")), NEGNUMFMT)
         GridAnual.TextMatrix(i, C_OTROIMP) = Format(vFld(Rs("OtroImp")), NEGNUMFMT)
         GridAnual.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NEGNUMFMT)
         
'      End If
      
      
      GridAnual.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      
      GridAnual.TextMatrix(i, C_PAGADO) = Format(vFld(Rs("Pagado")), NEGNUMFMT)
      
      'GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)
      
      '2690461
      Q2 = ""
      '2970116
      'Q2 = "select sum(afecto) as monto from documento "
      Q2 = "select sum(afecto + exento +OtroImp) as monto from documento "
      '2970116
      Q2 = Q2 & " Where tipoDoc = 3 and TipoLib = " & vFld(Rs("Tipolib"))
      Q2 = Q2 & " and IdDocAsoc = " & vFld(Rs("IdDoc"))

      Set Rs2 = OpenRs(DbMain, Q2)

      MontoNotaCred = 0

      If Rs2.EOF = False Then

        If vFld(Rs2("monto")) <> 0 Then
           MontoNotaCred = vFld(Rs2("monto"))
           GridAnual.TextMatrix(i, C_NOTACRED) = "x"
        Else
           GridAnual.TextMatrix(i, C_NOTACRED) = ""
        End If
      End If
      Call CloseRs(Rs2)
      'fin 2690461
      
      
       '2752418
      If vFld(Rs("OtroImp")) > 0 Then
      
        Q1 = "SELECT count(*) as resultado "
        Q1 = Q1 & "FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc) INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.idCuenta "
        Q1 = Q1 & " where numdoc = '" & vFld(Rs("NumDoc")) & "'"
        Q1 = Q1 & " and MovDocumento.debe = " & vFld(Rs("OtroImp"))
        Q1 = Q1 & " and Cuentas.codigo like '3*'  " ' num 3 corresponde a cuentas tipo 3 ejem:3010101
           
        Set Rs1 = OpenRs(DbMain, Q1)
        
        If Rs1.EOF = False Then
          If vFld(Rs1("resultado")) > 0 Then
          
              '2690461
             If (vFld(Rs("Afecto")) + vFld(Rs("Exento")) - MontoNotaCred) < vFld(Rs("Pagado")) Then

             GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento")) - MontoNotaCred), NEGNUMFMT)

             Else

             GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento"))), NEGNUMFMT)

             End If
             
             '2690461
          
             'GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(CDbl(vFld(Rs("MontoAfectaBaseImp")) + CDbl(vFld(Rs("OtroImp"))) + CDbl(vFld(Rs("IVAIrrec"))) + vFld(Rs("Exento"))), NEGNUMFMT)
          Else
             GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)
          End If
        End If
        
        Call CloseRs(Rs1)
      Else
      '2841464
           'If vFld(Rs("OtroImp")) < 0 Then
           ' Grid.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")) + vFld(Rs("OtroImp")), NEGNUMFMT)
    
           'Else
          '' GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)
           'End If
      'fin 2841464
      
      
         '2690461
             If (vFld(Rs("Afecto")) + vFld(Rs("Exento")) - MontoNotaCred) < vFld(Rs("Pagado")) Then

             GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format((vFld(Rs("Afecto")) - MontoNotaCred), NEGNUMFMT)

             Else

             GridAnual.TextMatrix(i, C_MONTOAFECTABASEIMP) = Format(vFld(Rs("MontoAfectaBaseImp")), NEGNUMFMT)

             End If
'
             '2690461
      
      
      End If

      
      
      
      If gEmpresa.Ano < 2020 Then
   '      If vFld(Rs("IdDoc")) = 0 Then       'NO es importado
            If vFld(Rs("ConEntRel")) <> 0 Then
               GridAnual.TextMatrix(i, C_CONENTREL) = "x"
            Else
               GridAnual.TextMatrix(i, C_CONENTREL) = ""
            End If
   '      End If
      Else
         If vFld(Rs("IdEntidad")) > 0 Then
            If ValidaEnt14D(vFld(Rs("IdEntidad"))) Then
               If vFld(Rs("ConEntRel")) <> 0 Then
                  GridAnual.TextMatrix(i, C_CONENTREL) = "x"
               Else
                  GridAnual.TextMatrix(i, C_CONENTREL) = ""
               End If
            End If
            
'         ElseIf vFld(Rs("IdEntReal")) > 0 Then
'            If ValidaEnt14D(vFld(Rs("IdEntReal"))) Then
'               If vFld(Rs("ConEntRel")) <> 0 Then
'                  GridAnual.TextMatrix(i, C_CONENTREL) = "x"
'               Else
'                  GridAnual.TextMatrix(i, C_CONENTREL) = ""
'               End If
'            End If
         End If
      End If

      If vFld(Rs("OperDevengada")) <> 0 Then
         GridAnual.TextMatrix(i, C_OPERDEVENGADA) = "x"
      Else
         GridAnual.TextMatrix(i, C_OPERDEVENGADA) = ""
      End If
     
      If vFld(Rs("PagoAPlazo")) <> 0 Then
         GridAnual.TextMatrix(i, C_PAGOAPLAZO) = "x"
      Else
         GridAnual.TextMatrix(i, C_PAGOAPLAZO) = ""
      End If
     
      GridAnual.TextMatrix(i, C_FECHAEXIGPAGO) = IIf(vFld(Rs("FechaExigPago")) > 0, Format(vFld(Rs("FechaExigPago")), EDATEFMT), "")
      
      Call CalcSaldo(i, GridAnual)
      
      GridAnual.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      GridAnual.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      GridAnual.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
               
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   If Row <= 0 Then
      Call FGrVRows(GridAnual)
      GridAnual.rows = GridAnual.rows + 1
      GridAnual.TopRow = GridAnual.FixedRows
      
      'Marco la columna Ordenada
      GridAnual.Row = 0
      GridAnual.Col = lOrdenSel
      Set GridAnual.CellPicture = FrmMain.Pc_Flecha
   
   Else
      GridAnual.Row = Row
      GridAnual.Col = C_TIPODOC
   
   End If
   
   Tx_CurrCell = ""
      
   Call CalcTot
   
   GridAnual.FlxGrid.Redraw = True
   
   If lOper = O_EDIT Then
      Call UnLockAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes), , , False)
   
      EditEnable = LockAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes))
   
      If EditEnable = False Then    'alguien más lo está editando, no podemos editarlo (esto se hace sólo una vez ya que en lOper = O_EDIT no se puede cambiar el mes)
         MsgBox1 "El Libro de Caja del mes de " & gNomMes(CbItemData(Cb_Mes)) & " se está editando en el equipo '" & IsLockedAction(DbMain, TOPERCAJA_LOCK + lTipoOper, CbItemData(Cb_Mes)) & "'. Sólo se abrirá de lectura.", vbInformation
         lEditEnabled = False
      End If
   
      Call EnableForm(Me, lEditEnabled)
      Call SetTxRO(Tx_CurrCell, True)
   End If
   
   Bt_List.Enabled = False
   
End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   Dim IdDoc As Long
   Dim ValPrevLine As Boolean
   Dim F1 As Long, F2 As Long
   Dim Msg As String
   Dim IdxTipoDoc As Integer
   Dim IdxTipoValLib As Integer
   

  
   If lEditEnabled = False Then
      Exit Sub
   End If
                           
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then           'es documento importado => no se edita
'      If (Col < C_PAGADO And Col <> C_FECHAOPER) Or Col = C_CONENTREL Then
'      If (Col < C_PAGADO And Col <> C_FECHAOPER) Then       '14 jul 2017 por solicitud de Claudio Villegas
      If Col <> C_OPERDEVENGADA Then                         '17 ago por solicitud de Claudio Villegas
         Exit Sub
      End If
   End If
                           
   If Grid.TextMatrix(Row, C_NUMDOC) = "" Then
      
      If gEmpresa.Franq14Ter = 0 Or gEmpresa.ObligaLibComprasVentas Then
         MsgBox1 "Empresa No acogida a Art. 14 TER u obligada a llevar Libro de Compras y Ventas." & vbCrLf & vbCrLf & "Debe ingresar los documentos en el Libro de Compras y Ventas.", vbInformation
         Exit Sub
      End If
      
'      If lHayDocsLibComprasVentas Then    'Importó documentos desde Libros de Compras y Venta => no puede agregar documentos de estos libros en forma manual, pero si otros ingresos y egresos
'         MsgBox1 "Dado que ya importó documentos desde los libros de Compras, Ventas y Retenciones, no puede ingresar nuevos documentos manualmente." & vbCrLf & vbCrLf & "Debe continuar ingresando los documentos en el Libro de Compras, Ventas y Retenciones, y traspasarlos hacia el Libro de Caja utilizando el botón para este efecto, ubicado en la esquina inferior derecha en esta misma ventana.", vbInformation
'         Exit Sub
'      End If

      If lHayDocsLibComprasVentas Then    'Importó documentos desde Libros de Compras y Venta => no puede agregar documentos de estos libros en forma manual, pero si otros ingresos y egresos
         If Not lMsgHayDocsLibCompasVentas Then
            If MsgBox1("Dado que ya importó documentos desde los libros de Compras, Ventas y Retenciones, no puede ingresar nuevos documentos de este tipo manualmente." & vbCrLf & vbCrLf & "Debe continuar ingresando los documentos en el Libro de Compras, Ventas y Retenciones, y traspasarlos hacia el Libro de Caja utilizando el botón para este efecto, ubicado en la esquina inferior derecha en esta misma ventana." & vbCrLf & vbCrLf & "Sin embargo si puede ingresar otros tipos de Ingresos y Egresos." & vbCrLf & vbCrLf & "¿Desea vlolver a ver este mensaje?", vbQuestion + vbYesNo) = vbNo Then
               Call SetIniString(gIniFile, "Msg", "LibCajaDocs", "1")
               lMsgHayDocsLibCompasVentas = True
            End If
         End If
      End If

      If Col <> C_NUMDOC Then
         MsgBox1 "Ingrese el número de documento antes de continuar.", vbExclamation + vbOKOnly
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
                      
      Grid.TextMatrix(Row, C_NUMLIN) = Row - Grid.FixedRows + 1
      If Row >= Grid.rows - 2 Then
         Grid.rows = Grid.rows + 1
      End If
      
      If gAppCode.Demo Then
         If Grid.rows > MAX_DOCDEMO Then
            MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
            Exit Sub
         End If
      End If

      
      Grid.TextMatrix(Row, C_IDTIPOOPER) = lTipoOper                 ' cuando se edita lTipoOper es ingreso o egreso, nunca "(todas)"
      Grid.TextMatrix(Row, C_TIPOOPER) = UCase(Left(gTipoOperCaja(lTipoOper), 1))
      Grid.TextMatrix(Row, C_IDESTADO) = ED_PENDIENTE
      
   'ElseIf Not ValidaEstadoEdit(Row) And Col <> C_DETALLE Then
   ElseIf Not ValidaEstadoEdit(Row) Then
      MsgBeep vbExclamation
      Exit Sub
   
   End If
   
   Grid.TxBox.MaxLength = 0
   
   IdxTipoDoc = GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   

   Select Case Col
   
       Case C_NUMDOC
     
         If Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
            Grid.TxBox.MaxLength = MAX_NUMDOCLEN
            EdType = FEG_Edit
         End If
         
         If lTipoOper = TOPERCAJA_INGRESO Then   'ventas
            If Row > Grid.FixedRows And Grid.TextMatrix(Row, C_TIPODOC) = "" And Grid.RowHeight(Row - 1) > 0 And Not lHayDocsLibComprasVentas Then
               Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
               Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
               Grid.TextMatrix(Row, C_IDTIPOLIB) = Grid.TextMatrix(Row - 1, C_IDTIPOLIB)
               Grid.TextMatrix(Row, C_DTE) = Grid.TextMatrix(Row - 1, C_DTE)
               Grid.TextMatrix(Row, C_TIPODOCEXT) = Grid.TextMatrix(Row - 1, C_TIPODOCEXT) & IIf(Grid.TextMatrix(Row - 1, C_DTE) <> "", " E", "")
               Grid.TextMatrix(Row, C_DOCIMPEXP) = Grid.TextMatrix(Row - 1, C_DOCIMPEXP)
     
            
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
                     MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     lMsgNotaCred = True
                  End If
                  
               End If
               
            End If
            
         End If
         
         EdType = FEG_Edit

      Case C_TIPODOC
      
         Grid.TxBox.MaxLength = 3
      
'         If lHayDocsLibComprasVentas Then
'            Grid.TextMatrix(Row, C_TIPODOC) = "OTR"
'            Grid.TextMatrix(Row, C_IDTIPODOC) = -1
        
         If Row > Grid.FixedRows And Grid.TextMatrix(Row, Col) = "" And Not lHayDocsLibComprasVentas Then
            Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
            Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
         End If
         
         EdType = FEG_Edit
               
      Case C_NUMDOCHASTA
           '2814014
         'If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTAEX Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLEXENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then      'venta sin documento
         If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLVENTAEX Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_BOLEXENTA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Or Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VALVENTAEX Then       'venta sin documento
            Grid.TxBox.MaxLength = MAX_NUMDOCLEN
            EdType = FEG_Edit
         End If
         'fin 2814014
               
      Case C_DTE
      
         If Trim(Grid.TextMatrix(Row, Col)) = "" And Grid.TextMatrix(Row, C_TIPODOC) <> "IMP" And Grid.TextMatrix(Row, C_TIPODOC) <> "FIC" And Grid.TextMatrix(Row, C_TIPODOC) <> "FIV" And Grid.TextMatrix(Row, C_TIPODOC) <> TDOC_VALEPAGOELECTR Then
            Grid.TextMatrix(Row, Col) = "x"
         Else
            Grid.TextMatrix(Row, Col) = ""
         End If
                  
         If Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) <> 0 And Val(Grid.TextMatrix(Row, C_IDTIPODOC)) Then
            Grid.TextMatrix(Row, C_TIPODOCEXT) = gTipoDoc(GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).Nombre & IIf(Grid.TextMatrix(Row, C_DTE) <> "", " E", "")
         Else
            Grid.TextMatrix(Row, C_TIPODOCEXT) = ""
         End If
         Call FGrModRow(Grid, Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
         lModifica = True
   
      Case C_OPERDEVENGADA, C_PAGOAPLAZO
         
         If Trim(Grid.TextMatrix(Row, Col)) = "" Then
            Grid.TextMatrix(Row, Col) = "x"
         Else
            Grid.TextMatrix(Row, Col) = ""
         End If
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
         lModifica = True
         
      Case C_CONENTREL
         
         If Trim(Grid.TextMatrix(Row, Col)) = "" Then
            
            If gEmpresa.Ano >= 2020 Then
               If ValidaEnt14D(Val(Grid.TextMatrix(Row, C_IDENTIDAD))) Then
                  Grid.TextMatrix(Row, Col) = "x"
               End If

            Else
               Grid.TextMatrix(Row, Col) = "x"
            End If
         Else
            Grid.TextMatrix(Row, Col) = ""
         End If
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
         lModifica = True
         
      Case C_RUT
         Grid.TxBox.MaxLength = 13
         EdType = FEG_Edit
         
      'Case C_NOMBRE
         'If Val(Grid.TextMatrix(Row, C_IDENTIDAD)) = 0 Then
         '   Grid.TxBox.MaxLength = 50
         '   EdType = FEG_Edit
         'End If
                       
      Case C_FECHAOPER
         
         Grid.TxBox.MaxLength = 10
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)) = 0 Then
         
            If lAno = Year(Now) And CbItemData(Cb_Mes) = month(Now) Then
               Grid.TextMatrix(Row, C_LNGFECHAOPER) = CLng(Now)
         
            Else
               Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), F1, F2)
               Grid.TextMatrix(Row, C_LNGFECHAOPER) = CLng(F2)
            End If
         End If
         
         Grid.TextMatrix(Row, C_FECHAOPER) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAOPER)), EDATEFMT)
         
         EdType = FEG_Edit
         
      Case C_EXENTO
      
         If IdxTipoDoc >= 0 Then
            If gTipoDoc(IdxTipoDoc).TieneExento <> 0 And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 Then
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
            End If
         End If
         
      Case C_AFECTO, C_IVA, C_OTROIMP
      
         If IdxTipoDoc >= 0 Then
      
            If gTipoDoc(IdxTipoDoc).TieneAfecto <> 0 And gTipoDoc(IdxTipoDoc).IngresarTotal = 0 Then
               
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
               
            End If
         
         End If
         
      Case C_IVAIRREC
      
         If IdxTipoDoc >= 0 And Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_COMPRAS Then
         
            IdxTipoValLib = FindTipoValLib(LIB_COMPRAS, LIBCOMPRAS_IVAIRREC1)
      
            If InStr(gTipoValLib(IdxTipoValLib).TipoDoc, "," & gTipoDoc(IdxTipoDoc).TipoDoc & ",") <> 0 Then
               
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
               
            End If
         
         End If
         
      Case C_TOTAL
            
         If IdxTipoDoc >= 0 Then
'            If Not gTipoDoc(IdxTipoDoc).TieneExento And Not gTipoDoc(IdxTipoDoc).TieneAfecto Then   'no se ingresa ni afecto ni exento => se ingresa el total (caso de boletas de venta y devoluciones con boleta)
            If gTipoDoc(IdxTipoDoc).IngresarTotal <> 0 Then   'no se ingresa ni afecto ni exento => se ingresa el total (caso de boletas de venta y devoluciones con boleta)
               Grid.TxBox.MaxLength = 15
               EdType = FEG_Edit
            End If
         End If
           
      Case C_PAGADO
      
         If IdxTipoDoc >= 0 Then
                     
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
                        
         End If
         
      Case C_DESCRIP
         Grid.TxBox.MaxLength = 100
         EdType = FEG_Edit
                  
      Case C_FECHAEXIGPAGO
      
         Grid.TxBox.MaxLength = 10
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO)) > 0 Then
            Grid.TextMatrix(Row, C_FECHAEXIGPAGO) = Format(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO), EDATEFMT)
         
         ElseIf vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)) > 0 Then
         
            If GetEntRelacionada(Val(Grid.TextMatrix(Row, C_IDENTIDAD))) And lTipoOper = TOPERCAJA_INGRESO Then   'si es Ent Relacionada y es ingreso, exigibilidad de pago contado
               Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = CLng(vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)))
            Else                                                                 ' se propone a 30 días
               Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = CLng(DateAdd("d", 30, vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER))))
            End If
            
            Grid.TextMatrix(Row, C_FECHAEXIGPAGO) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO)), EDATEFMT)
         End If
         
         EdType = FEG_Edit
               
   End Select

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim IdEnt As Long
   Dim Rc As Integer
   Dim Nombre As String
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
   Dim FEmision As Long
   Dim AuxTipoLib As Integer
   Dim IdxTipoDoc As Integer
   Dim EntRelacionada As Boolean
   Dim FechaIng As Long, FPago As Long
   Dim TipoDoc2 As Integer, TipoLib2 As Integer
   Dim FechaLey21 As Long
   
   FechaLey21 = DateSerial(2022, 3, 24)
      
   Action = vbOK
   
   Select Case Col
   
      Case C_NUMDOC, C_NUMDOCHASTA
         Value = Trim(Value)
         Grid.TextMatrix(Row, Col) = Value
                                           
'         If Ch_RepetirGlosa <> 0 And Row > Grid.FixedRows Then
'            Grid.TextMatrix(Row, C_DESCRIP) = Grid.TextMatrix(Row - 1, C_DESCRIP)
'         End If
            
      Case C_TIPODOC
                                                
         Value = Trim(Value)
         
         If Value <> "" Then
         
            TipoDoc = FindTipoDocLibCaja(Value, AuxTipoLib)    'obtiene TipoDoc y TipoLib a partir del diminutivo del TipoDoc
            
            If TipoDoc > 0 Then
            
               If lHayDocsLibComprasVentas And (AuxTipoLib = LIB_COMPRAS Or AuxTipoLib = LIB_VENTAS Or AuxTipoLib = LIB_RETEN) Then
                  If Not lMsgHayDocsLibCompasVentas Then
                     If MsgBox1("Dado que ya importó documentos desde los libros de Compras, Ventas y Retenciones, no puede ingresar nuevos documentos de este tipo manualmente." & vbCrLf & vbCrLf & "Debe continuar ingresando los documentos en el Libro de Compras, Ventas y Retenciones, y traspasarlos hacia el Libro de Caja utilizando el botón para este efecto, ubicado en la esquina inferior derecha en esta misma ventana." & vbCrLf & vbCrLf & "Sin embargo si puede ingresar otros tipos de Ingresos y Egresos." & vbCrLf & vbCrLf & "¿Desea vlolver a ver este mensaje?", vbQuestion + vbYesNo) = vbNo Then
                        Call SetIniString(gIniFile, "Msg", "LibCajaDocs", "1")
                        lMsgHayDocsLibCompasVentas = True
                     End If
                  Else
                     MsgBeep vbExclamation
                  End If
                  Grid.TextMatrix(Row, C_TIPODOC) = ""
                  Grid.TextMatrix(Row, C_IDTIPODOC) = 0
                  Value = ""
                  Action = vbRetry
               End If
               
               If (TipoDoc = LIBCAJA_OTROSINGINI And AuxTipoLib = LIB_CAJAING) Or (TipoDoc = LIBCAJA_OTROSEGRINI And AuxTipoLib = LIB_CAJAEGR) Then
                  If CbItemData(Cb_Mes) <> 1 Then    'estos tipos de ingresos/egresos sólo se pueden ingresar en enero
                     MsgBox1 "Sólo se permite ingresar este tipo de Ingresos/Egresos en el mes de enero.", vbExclamation
                     Grid.TextMatrix(Row, C_TIPODOC) = ""
                     Grid.TextMatrix(Row, C_IDTIPODOC) = 0
                     Action = vbCancel
                     Exit Sub
                  
                  Else
                     For i = Grid.FixedRows To Grid.rows - 1
                        
                        If Grid.RowHeight(i) > 0 Then

                           TipoDoc2 = Val(Grid.TextMatrix(i, C_IDTIPODOC))
                           TipoLib2 = Val(Grid.TextMatrix(i, C_IDTIPOLIB))
                           
                           If (TipoDoc2 = LIBCAJA_OTROSINGINI And TipoLib2 = LIB_CAJAING) Or (TipoDoc2 = LIBCAJA_OTROSEGRINI And TipoLib2 = LIB_CAJAEGR) Then
                              If i <> Row Then
                                 MsgBox1 "No se perimte ingresar más de un Ingreso/Egreso Inicial en el año.", vbExclamation
                                 Grid.TextMatrix(Row, C_TIPODOC) = ""
                                 Grid.TextMatrix(Row, C_IDTIPODOC) = 0
                                 Action = vbCancel
                                 Exit Sub
                              End If
                           End If
                           
                        End If
                     Next i
                              
                  End If
               End If
                     
               
               
               Grid.TextMatrix(Row, C_IDTIPODOC) = TipoDoc
               Grid.TextMatrix(Row, Col) = Value
               Grid.TextMatrix(Row, C_IDTIPOLIB) = AuxTipoLib
               IdxTipoDoc = GetTipoDoc(AuxTipoLib, TipoDoc)
               Grid.TextMatrix(Row, C_TIPODOCEXT) = gTipoDoc(IdxTipoDoc).Nombre & IIf(Grid.TextMatrix(Row, C_DTE) <> "", " E", "")
               Grid.TextMatrix(Row, C_DOCIMPEXP) = CInt(gTipoDoc(IdxTipoDoc).DocImpExp)
               Grid.TextMatrix(Row, C_ESREBAJA) = Int(gTipoDoc(IdxTipoDoc).EsRebaja)
               Call ActCambioTipoDoc(Row)
               
               Call CalcSaldo(Row, Grid)
               Call CalcTot
               
               If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_VENTASINDOC Then
                  Grid.TextMatrix(Row, C_NUMDOC) = ""
               End If
                                             
               If Value = "NCV" Then
                  If Not lMsgNotaCred Then
                     MsgBox1 "Recuerde que si está ingresando una nota de crédito por devolución de mercaderías o servicios resciliados, el plazo para rebajar el débito fiscal es de 3 meses según establece artículo 21 Ley de IVA", vbInformation + vbOKOnly
                     lMsgNotaCred = True
                  End If
               ElseIf lTipoOper = TOPERCAJA_INGRESO And Value <> "NCV" And Value <> "NDV" And Value <> "FCV" And Value <> "LFV" And Value <> "OIN" Then
                  Grid.TextMatrix(Row, C_RUT) = FmtCID(gEmpresa.Rut)
                  Grid.TextMatrix(Row, C_IDENTIDAD) = 0
                  Grid.TextMatrix(Row, C_NOMBRE) = gEmpresa.RazonSocial
               End If
                  
            Else
               MsgBox1 "Tipo de documento inválido. Presione el botón derecho del mouse para Ayuda.", vbExclamation + vbOKOnly
               Grid.TextMatrix(Row, C_TIPODOC) = ""
               Grid.TextMatrix(Row, C_IDTIPODOC) = 0
               Action = vbRetry
               
            End If
            
         Else
            Grid.TextMatrix(Row, C_IDTIPODOC) = ""
            Grid.TextMatrix(Row, Col) = ""
            Grid.TextMatrix(Row, C_IDTIPOLIB) = ""
            Grid.TextMatrix(Row, C_TIPODOCEXT) = ""
            Grid.TextMatrix(Row, C_DOCIMPEXP) = ""
            Grid.TextMatrix(Row, C_ESREBAJA) = ""
            Call ActCambioTipoDoc(Row)
            
            Call CalcSaldo(Row, Grid)
            Call CalcTot
         End If
         
      Case C_RUT
            
         If Trim(Value) = "" Then
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbOK
                  
         ElseIf Trim(Value) = "0-0" Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
         
         ElseIf (Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And DocExigeRut(Row)) And Not ValidCID(Value) Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
         
         ElseIf ValidCID(Value) Then
         
            IdEnt = GetIdEntidad(Trim(Value), Nombre, NotValidRut)
            
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
               
               If gEmpresa.Ano < 2020 Then
                  EntRelacionada = GetEntRelacionada(IdEnt)
                  If EntRelacionada Then
                     Grid.TextMatrix(Row, C_CONENTREL) = "x"
                  Else
                     Grid.TextMatrix(Row, C_CONENTREL) = ""
                  End If
               Else
                  If ValidaEnt14D(IdEnt) Then
                     Grid.TextMatrix(Row, C_CONENTREL) = "x"
                  Else
                     Grid.TextMatrix(Row, C_CONENTREL) = ""
                  End If
               End If
            End If
            
            Call ValidaNumDoc(Row)
         
         ElseIf Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) <> 0 Then
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
         Else
            Value = UCase(Trim(Value))
            If Value <> RUT_VARIOS Then
               MsgBox1 "RUT inválido.", vbExclamation
               Grid.TextMatrix(Row, C_RUT) = Value
               Grid.TextMatrix(Row, C_IDENTIDAD) = 0
               Grid.TextMatrix(Row, C_NOMBRE) = ""
               Action = vbRetry
            Else
               Grid.TextMatrix(Row, C_IDENTIDAD) = 0
               Grid.TextMatrix(Row, C_NOMBRE) = ""
            End If
         End If
         
      'Case C_NOMBRE
         'If Trim(Value) = "" Then
         '   MsgBox1 "Nombre o razón social inválido.", vbExclamation + vbOKOnly
         '   Action = vbCancel
         'End If
         
      Case C_EXENTO, C_AFECTO, C_IVA, C_OTROIMP
      
         If gTipoDoc(GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).EsRebaja Then
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

            Call CalcTotRow(Row, True)
         
         ElseIf Col = C_AFECTO Then
            
            Call CalcTotRow(Row, True)
         
         Else
            Call CalcTotRow(Row, False)
            
            If Col = C_IVA Then    'esto se hace por si el usuario cambia el valor y luego se arrepiente en función CalcTotRow
               Value = Grid.TextMatrix(Row, C_IVA)
            End If
            
         End If
                  
         Call CalcTot
                  
      Case C_TOTAL
         
         If vFmt(Value) < 0 Then
            MsgBox1 "Valor inválido.", vbExclamation + vbOKOnly
            Action = vbRetry
               
         ElseIf EsIngresoTotal(Row) Then
         
            Call CalcIngresoTotal(Row, Col, Value)
    
         End If
            
      Case C_FECHAOPER
         
         Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), FirstDay, LastDay)
         
         FEmision = GetDate(Value, "dmy")
         
         If lTipoOper = TOPERCAJA_INGRESO Then
            If FEmision >= FirstDay And FEmision <= LastDay Then
               Grid.TextMatrix(Row, C_LNGFECHAOPER) = FEmision
               
            ElseIf Grid.TextMatrix(Row, C_TIPODOC) = "NCV" Then    'para notas de crédito, permitimos ingreso fuera de plazo (solicitado el día 29/06/2011 por Victro Morales)
               'MsgBox1 "Recuerde verificar la rebaja Débito Fiscal según instrucciones del Articulo 21 Ley sobre Impuesto a las ventas y servicios", vbOKOnly + vbInformation
'               If FechaLey21 >= FEmision Then
'                    MsgBox1 "Recuerde verificar la rebaja Debito Fiscal según instrucciones de Ley sobre impuesto a las ventas y Servicios y Ley 21.398", vbOKOnly + vbInformation
'               Else
                    MsgBox1 "Recuerde verificar la rebaja Débito Fiscal según instrucciones del Articulo 21 Ley sobre Impuesto a las ventas y servicios", vbOKOnly + vbInformation
'               End If
               Grid.TextMatrix(Row, C_LNGFECHAOPER) = FEmision
            
            Else
               MsgBox1 "Fecha de emisión inválida.", vbExclamation + vbOKOnly
               Action = vbCancel
            End If
         
         Else
            If FEmision > LastDay Then
               MsgBox1 "Fecha de emisión inválida.", vbExclamation + vbOKOnly
               Action = vbCancel
            End If
               
            Grid.TextMatrix(Row, C_LNGFECHAOPER) = FEmision
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)) > 0 Then
            Value = Format(Grid.TextMatrix(Row, C_LNGFECHAOPER), EDATEFMT)
         End If
         
         If Grid.TextMatrix(Row, C_TIPODOC) = "IMP" Then
            'si el doc de importación no es del mismo mes, hay que mostrar advertencia
            If CbItemData(Cb_Mes) <> month(Val(Grid.TextMatrix(Row, C_LNGFECHAOPER))) Or gEmpresa.Ano <> Year(Val(Grid.TextMatrix(Row, C_LNGFECHAOPER))) Then
               MsgBox1 "Recuerde que según las normas de la Ley sobre Impuesto a las Ventas y Servicios, el IVA se puede aprovechar dentro del mismo periodo de emisión del documento.", vbInformation + vbOKOnly
            End If
         End If

         FechaIng = DateSerial(lAno, CbItemData(Cb_Mes), 1)
         Grid.TextMatrix(Row, C_LNGFECHAINGRESOLIBRO) = FechaIng
         
      'Case C_DESCRIP
         
       Case C_PAGADO
      
            
         If vFmt(Value) < 0 Then
            MsgBox1 "Valor inválido.", vbExclamation + vbOKOnly
            Action = vbRetry
         ElseIf vFmt(Value) > Abs(vFmt(Grid.TextMatrix(Row, C_TOTAL))) Then
            MsgBox1 "Valor pagado supera valor total del documento.", vbExclamation + vbOKOnly
            Action = vbRetry
         End If
         
         Value = Format(vFmt(Value), NEGNUMFMT)
                   
         Grid.TextMatrix(Row, Col) = Value
         
         Call CalcSaldo(Row, Grid)
         Call CalcTot
         
      Case C_FECHAEXIGPAGO
      
         Value = Trim(Value)
         If Value <> "" Then
         
            FPago = GetDate(Value, "dmy")
            FEmision = Val(Grid.TextMatrix(Row, C_LNGFECHAOPER))
   
            If FPago < FEmision Then
               MsgBox1 "Fecha Exigibilidad de Pago inválida.", vbExclamation + vbOKOnly
               Action = vbCancel
            Else
               Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = FPago
               Grid.TextMatrix(Row, C_FECHAEXIGPAGO) = Format(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO), EDATEFMT)
   
            End If
         
         Else
            Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = ""
            
         End If
         
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDLIBROCAJA, C_UPDATE)
      lModifica = True
   End If
   
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
   
'   If Fr_Opciones.Visible Then
'      Fr_Opciones.Visible = False
'   End If
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If

   Call OrdenaPorCol(Col)
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   Select Case Grid.Col
   
      Case C_FECHAOPER, C_FECHAEXIGPAGO
         Call KeyDate(KeyAscii)
      
      Case C_TIPODOC
         Call KeyUpper(KeyAscii)
         
      Case C_NUMDOC, C_NUMDOCHASTA
         
'         If (lTipoOper = TOPERCAJA_EGRESO And Grid.TextMatrix(Grid.Row, C_TIPODOC) = "IMP") Or (lTipoOper = TOPERCAJA_INGRESO And Grid.TextMatrix(Grid.Row, C_TIPODOC) = TDOC_VALEPAGOELECTR) Then
            Call KeyName(KeyAscii)
'         Else
'            Call KeyNum(KeyAscii)
'         End If
                     
      Case C_RUT
         Call KeyName(KeyAscii)
'         Call KeyUpper(KeyAscii)
         
      Case C_NOMBRE
         Call KeyName(KeyAscii)
         
      Case C_EXENTO, C_AFECTO, C_IVA, C_OTROIMP, C_TOTAL
         Call KeyNum(KeyAscii)
      
      Case C_DESCRIP
         Call KeyName(KeyAscii)
         
   End Select
   
End Sub
'Recibe la grilla que puedes ser la mensual o la anual
' "Grid" grilla mensual
' "GridAnual" grilla anual (Oculta)
Private Sub SaveGrid(Grid As FEd2Grid)
   Dim i As Integer
   Dim Lin As Integer
   Dim Rs As Recordset
   Dim Q1 As String, Q2 As String
   Dim IdLibroCaja As Long
   Dim j As Integer
   Dim EsRebaja As Boolean
   Dim NumDocVSD As String
   Dim FldArray(3) As AdvTbAddNew_t

   lIdLibroCaja = 0
   
   Lin = Grid.FixedRows
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_TIPOOPER) = "" Then    'ya terminó la lista de docs.
         Exit For
      End If
      
     If gAppCode.Demo Then
         If i - Grid.FixedRows >= MAX_DOCDEMO And Not W.InDesign Then
            MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
            Exit For
         End If
      End If
            
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
'         Set Rs = DbMain.OpenRecordset("LibroCaja")
'         Rs.AddNew
'
'         IdLibroCaja = vFld(Rs("IdLibroCaja"))
'         Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'         Rs.Fields("FechaCreacion") = CLng(Int(Now))
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
         
         IdLibroCaja = AdvTbAddNewMult(DbMain, "LibroCaja", "IdLibroCaja", FldArray)
         
         Grid.TextMatrix(i, C_IDLIBROCAJA) = IdLibroCaja
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
         
         If lIdLibroCaja = 0 Then   'selecciona el primero que se insertó
            lIdLibroCaja = IdLibroCaja
         End If
         
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE * FROM LibroCaja WHERE IdLibroCaja = " & Val(Grid.TextMatrix(i, C_IDLIBROCAJA))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdLibroCaja = " & Val(Grid.TextMatrix(i, C_IDLIBROCAJA))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "LibroCaja", Q1)
                                
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then 'Update
         Q1 = "UPDATE LibroCaja SET "
         Q1 = Q1 & "  TipoOper = " & Val(Grid.TextMatrix(i, C_IDTIPOOPER))
         
         If Grid.TextMatrix(i, C_TIPODOC) <> TDOC_VENTASINDOC Then     'venta sin documento
            Q1 = Q1 & ", NumDoc = '" & Trim(Grid.TextMatrix(i, C_NUMDOC)) & "'"
            Q1 = Q1 & ", NumDocHasta = '" & Trim(Grid.TextMatrix(i, C_NUMDOCHASTA)) & "'"
         Else
            NumDocVSD = GetNumDocVSDLibCaja(Val(Grid.TextMatrix(i, C_IDTIPOLIB)), Val(Grid.TextMatrix(i, C_IDTIPODOC)))
            Q1 = Q1 & ", NumDoc = '" & NumDocVSD & "'"
            Q1 = Q1 & ", NumDocHasta = '0'"
         End If
         
         Q1 = Q1 & ", TipoLib = " & Val(Grid.TextMatrix(i, C_IDTIPOLIB))
         Q1 = Q1 & ", TipoDoc = " & Val(Grid.TextMatrix(i, C_IDTIPODOC))
         Q1 = Q1 & ", DTE = " & IIf(Trim(Grid.TextMatrix(i, C_DTE)) <> "", -1, 0)
                        
         Q1 = Q1 & ", IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
         
         If lTipoOper = TOPERCAJA_EGRESO Then
            'por si acaso, ponemos la clasificación de la entidad
            Q2 = "UPDATE Entidades SET Clasif" & ENT_PROVEEDOR & " = 1 WHERE IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
            Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id
            Call ExecSQL(DbMain, Q2)
         Else
            'por si acaso, ponemos la clasificación de la entidad
            Q2 = "UPDATE Entidades SET Clasif" & ENT_CLIENTE & " = 1 WHERE IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
            Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id
            Call ExecSQL(DbMain, Q2)
         End If
         
         If Val(Grid.TextMatrix(i, C_IDENTIDAD)) = 0 Then
            Q1 = Q1 & ", RutEntidad ='" & IIf((Val(Grid.TextMatrix(i, C_DOCIMPEXP)) = 0 And DocExigeRut(i)), vFmtCID(Grid.TextMatrix(i, C_RUT)), ParaSQL(Grid.TextMatrix(i, C_RUT))) & "'"
            Q1 = Q1 & ", NombreEntidad = '" & ParaSQL(Left(Grid.TextMatrix(i, C_NOMBRE), 40)) & "'"
         End If
         
         Q1 = Q1 & ", FechaOperacion = " & Val(Grid.TextMatrix(i, C_LNGFECHAOPER))
         Q1 = Q1 & ", FechaIngresoLibro = " & Val(Grid.TextMatrix(i, C_LNGFECHAINGRESOLIBRO))
         
         Q1 = Q1 & ", Afecto = " & Abs(vFmt(Grid.TextMatrix(i, C_AFECTO)))
         Q1 = Q1 & ", IVA = " & Abs(vFmt(Grid.TextMatrix(i, C_IVA)))
         Q1 = Q1 & ", IVAIrrec = " & Abs(vFmt(Grid.TextMatrix(i, C_IVAIRREC)))
         Q1 = Q1 & ", Exento = " & Abs(vFmt(Grid.TextMatrix(i, C_EXENTO)))
         
         EsRebaja = gTipoDoc(GetTipoDoc(Val(Grid.TextMatrix(i, C_IDTIPOLIB)), Val(Grid.TextMatrix(i, C_IDTIPODOC)))).EsRebaja
  
         If EsRebaja Then
            Q1 = Q1 & ", OtroImp = " & vFmt(Grid.TextMatrix(i, C_OTROIMP)) * -1
         Else
            Q1 = Q1 & ", OtroImp = " & vFmt(Grid.TextMatrix(i, C_OTROIMP))
         End If
         
         Q1 = Q1 & ", Total = " & Abs(vFmt(Grid.TextMatrix(i, C_TOTAL)))
         Q1 = Q1 & ", Pagado = " & Abs(vFmt(Grid.TextMatrix(i, C_PAGADO)))
         Q1 = Q1 & ", Ingreso = " & Abs(vFmt(Grid.TextMatrix(i, C_INGRESO)))
         Q1 = Q1 & ", Egreso = " & Abs(vFmt(Grid.TextMatrix(i, C_EGRESO)))
         
         Q1 = Q1 & ", Descrip = '" & ParaSQL(Left(Grid.TextMatrix(i, C_DESCRIP), 50)) & "'"
         Q1 = Q1 & ", ConEntRel = " & IIf(Trim(Grid.TextMatrix(i, C_CONENTREL)) <> "", -1, 0)
         Q1 = Q1 & ", OperDevengada = " & IIf(Trim(Grid.TextMatrix(i, C_OPERDEVENGADA)) <> "", -1, 0)
         Q1 = Q1 & ", PagoAPlazo = " & IIf(Trim(Grid.TextMatrix(i, C_PAGOAPLAZO)) <> "", -1, 0)
         
         Q1 = Q1 & ", FechaExigPago = " & Val(Grid.TextMatrix(i, C_LNGFECHAEXIGPAGO))
         
         Q1 = Q1 & ", Estado = " & Val(Grid.TextMatrix(i, C_IDESTADO))
         
         Q1 = Q1 & " WHERE IdLibroCaja = " & Val(Grid.TextMatrix(i, C_IDLIBROCAJA))
         Call ExecSQL(DbMain, Q1)
                  
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
         Lin = Lin + 1
      End If
      
      Grid.TextMatrix(i, C_UPDATE) = ""     'lo limpiampos dado que esta función es invocada en Bt_DetDoc
      
   Next i

   lModifica = False
End Sub
Private Sub Form_Resize()

   Grid.Width = Me.Width - 200
   'Grid.Height = Me.Height - Grid.Top - GridTot.Height - Tx_CurrCell.Height - W.YCaption - W.yFrame * 2 - 100
   GridTot.Top = Grid.Top + Grid.Height + 60
   GridTot.Width = Grid.Width - W.xScroll
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   Tx_CurrCell.Width = Me.Width - Bt_Opciones.Width - Bt_ImportOIngEg.Width - Bt_ImportDocs.Width - 400
   Bt_ImportDocs.Top = GridTot.Top + GridTot.Height + 60
   Bt_ImportDocs.Left = Tx_CurrCell.Left + Tx_CurrCell.Width + 60
   Bt_ImportOIngEg.Top = GridTot.Top + GridTot.Height + 60
   Bt_ImportOIngEg.Left = Bt_ImportDocs.Left + Bt_ImportDocs.Width + 60
   Bt_Opciones.Top = Tx_CurrCell.Top
   Bt_Opciones.Left = Bt_ImportOIngEg.Left + Bt_ImportOIngEg.Width + 60
   Fr_Opciones.Top = Bt_Opciones.Top - Fr_Opciones.Height - 30
   Bt_ImportTotalDocs.Top = GridTot.Top + GridTot.Height + 480
   Bt_ImportTotalDocs.Left = Tx_CurrCell.Left + Tx_CurrCell.Width + 60
   Bt_ImportOIngEgAnual.Top = GridTot.Top + GridTot.Height + 480
   Bt_ImportOIngEgAnual.Left = Bt_ImportTotalDocs.Left + Bt_ImportTotalDocs.Width + 80

   Call FGrVRows(Grid, 1)
'
   Call FGrTotales(Grid, GridTot)

End Sub


Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   If Val(Grid.TextMatrix(Grid.Row, C_IDDOC)) > 0 Then
      If Grid.Col < C_PAGADO Or Grid.Col = C_DESCRIP Then
         Exit Sub
      End If
   End If
   
   If Button = vbRightButton Then
      If (Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC) Then
         Call PopupMenu(M_TipoDoc)
      End If
   End If

End Sub

Private Sub Grid_SelChange()
   Dim EdType As FlexEdGrid2.FEG2_EdType
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)

End Sub
Private Function valida() As Boolean
   Dim Row As Integer
   Dim ValLine As Boolean
   Dim Msg As String
   Dim DbName As String
   Dim i As Integer, j As Integer
   Dim TipoDoc1 As Integer, TipoLib1 As Integer
   Dim TipoDoc2 As Integer, TipoLib2 As Integer
   Dim HayIngEgIni As Boolean, IdLCajaIngEgIni As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TipoOper As String
   
   valida = False
   
   'vemos si las líneas están completas
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If LineaEnBlanco(Row) Then
         Exit For
      End If
   
      ValLine = IsValidLine(Row, Msg)
      ValLine = (ValLine Or Grid.RowHeight(Row) = 0 Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO)    'línea borrada o doc anulado
      
      If ValLine = False Then
         If Msg <> "" Then
            MsgBox1 "Línea " & Row - 1 & " inválida. " & Msg, vbExclamation + vbOKOnly
         End If
         Exit Function
      End If
      
   Next Row
   
   'vemos si se repite OII o OEI
   If CbItemData(Cb_Mes) = 1 Then    'sólo se permiten en enero y ya se validó arriba en IsValidLine
   
      For i = Grid.FixedRows To Grid.rows - 1
      
         If Grid.RowHeight(i) > 0 Then
         
            TipoDoc1 = Val(Grid.TextMatrix(i, C_IDTIPODOC))
            TipoLib1 = Val(Grid.TextMatrix(i, C_IDTIPOLIB))
            
            If (TipoDoc1 = LIBCAJA_OTROSINGINI And TipoLib1 = LIB_CAJAING) Or (TipoDoc1 = LIBCAJA_OTROSEGRINI And TipoLib1 = LIB_CAJAEGR) Then
            
               HayIngEgIni = True
               IdLCajaIngEgIni = Val(Grid.TextMatrix(i, C_IDLIBROCAJA))
            
               For j = Grid.FixedRows To Grid.rows - 1
                  If Grid.RowHeight(j) > 0 Then
                  
                     TipoDoc2 = Val(Grid.TextMatrix(j, C_IDTIPODOC))
                     TipoLib2 = Val(Grid.TextMatrix(j, C_IDTIPOLIB))
                     
                     If (TipoDoc2 = LIBCAJA_OTROSINGINI And TipoLib2 = LIB_CAJAING) Or (TipoDoc2 = LIBCAJA_OTROSEGRINI And TipoLib2 = LIB_CAJAEGR) Then
                        If i <> j Then
                           MsgBox1 "No se perimte ingresar más de un Ingreso/Egreso Inicial en el año.", vbExclamation
                           Exit Function
                        End If
                     End If
                  End If
               Next j
            End If
            
         End If
      Next i
      
   Else
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.RowHeight(i) > 0 Then
            TipoDoc1 = Val(Grid.TextMatrix(i, C_IDTIPODOC))
            TipoLib1 = Val(Grid.TextMatrix(i, C_IDTIPOLIB))
            
            If (TipoDoc1 = LIBCAJA_OTROSINGINI And TipoLib1 = LIB_CAJAING) Or (TipoDoc1 = LIBCAJA_OTROSEGRINI And TipoLib1 = LIB_CAJAEGR) Then
               MsgBox1 "Solo se perimte ingresar un Ingreso/Egreso Inicial en el mes de enero.", vbExclamation
               Exit Function
            End If
         End If
      Next i
   End If
   
   'veamos si hay otro ingreso o egreso inicial en otro mes
   If HayIngEgIni Then
      Q1 = "SELECT IdLibroCaja, FechaIngresoLibro, TipoOper "
      Q1 = Q1 & " FROM LibroCaja "
      Q1 = Q1 & " WHERE TipoDoc IN(" & LIBCAJA_OTROSINGINI & "," & LIBCAJA_OTROSEGRINI & ") AND TipoLib IN(" & LIB_CAJAING & "," & LIB_CAJAEGR & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         If IdLCajaIngEgIni <> vFld(Rs("IdLibroCaja")) Then
            TipoOper = IIf(vFld(Rs("TipoOper")) = TOPERCAJA_INGRESO, "Ingreso", "Egreso")
            MsgBox1 "Ya existe un " & TipoOper & " Inicial en el Libro de Caja de " & TipoOper & "s de " & gNomMes(month(vFld(Rs("FechaIngresoLibro")))) & "." & vbCrLf & vbCrLf & "Solo se permite un Ingreso/Egreso Inicial en el año.", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
      Call CloseRs(Rs)
   End If
         
   valida = True
   
End Function
Private Function LineaEnBlanco(ByVal Row As Integer) As Boolean
   Dim i As Integer
   
   LineaEnBlanco = True
   
   For i = C_NUMDOC To Grid.Cols - 1
   
      If Grid.TextMatrix(Row, i) <> "" And i <> C_ESTADO And i <> C_IDESTADO And i <> C_LNGFECHAEXIGPAGO And i <> C_FECHAEXIGPAGO And i <> C_UPDATE Then
         LineaEnBlanco = False
      End If
      
   Next i
     
   If LineaEnBlanco Then
      Grid.TextMatrix(Row, C_FECHAOPER) = ""
      Grid.TextMatrix(Row, C_TIPODOC) = ""
   End If
      
End Function
Private Sub ActCambioTipoDoc(ByVal Row As Integer)
   Dim IdxDoc As Integer
   Dim TipoLib As Integer
      
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   
   If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then

      If vFmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_IVA)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 Then
         
         IdxDoc = GetTipoDoc(TipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
         
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
                     
      End If
   
      Call CalcSaldo(Row, Grid)
   End If
   
End Sub

Private Sub CalcTot()
   Dim Tot(C_MONTOAFECTABASEIMP) As Double
   Dim i As Integer, j As Integer
   
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_SALDO) = "" Then
         Exit For
      End If
      If Grid.RowHeight(i) > 0 Then

         If Val(Grid.TextMatrix(i, C_IDTIPOOPER)) = TOPERCAJA_INGRESO Then
            Tot(C_AFECTO) = Tot(C_AFECTO) + vFmt(Grid.TextMatrix(i, C_AFECTO))
            Tot(C_EXENTO) = Tot(C_EXENTO) + vFmt(Grid.TextMatrix(i, C_EXENTO))
            Tot(C_IVA) = Tot(C_IVA) + vFmt(Grid.TextMatrix(i, C_IVA))
            Tot(C_OTROIMP) = Tot(C_OTROIMP) + vFmt(Grid.TextMatrix(i, C_OTROIMP))
            Tot(C_TOTAL) = Tot(C_TOTAL) + vFmt(Grid.TextMatrix(i, C_TOTAL))
            Tot(C_PAGADO) = Tot(C_PAGADO) + vFmt(Grid.TextMatrix(i, C_PAGADO))
            Tot(C_IVAIRREC) = Tot(C_IVAIRREC) + vFmt(Grid.TextMatrix(i, C_IVAIRREC))
            Tot(C_MONTOAFECTABASEIMP) = Tot(C_MONTOAFECTABASEIMP) + vFmt(Grid.TextMatrix(i, C_MONTOAFECTABASEIMP))
            
            '3289932
            vMontoBaseImpoIngreso = Tot(C_MONTOAFECTABASEIMP)
            '3289932
         Else
            Tot(C_AFECTO) = Tot(C_AFECTO) - vFmt(Grid.TextMatrix(i, C_AFECTO))
            Tot(C_EXENTO) = Tot(C_EXENTO) - vFmt(Grid.TextMatrix(i, C_EXENTO))
            Tot(C_IVA) = Tot(C_IVA) - vFmt(Grid.TextMatrix(i, C_IVA))
            Tot(C_OTROIMP) = Tot(C_OTROIMP) - vFmt(Grid.TextMatrix(i, C_OTROIMP))
            Tot(C_TOTAL) = Tot(C_TOTAL) - vFmt(Grid.TextMatrix(i, C_TOTAL))
            Tot(C_PAGADO) = Tot(C_PAGADO) - vFmt(Grid.TextMatrix(i, C_PAGADO))
            Tot(C_IVAIRREC) = Tot(C_IVAIRREC) - vFmt(Grid.TextMatrix(i, C_IVAIRREC))
            Tot(C_MONTOAFECTABASEIMP) = Tot(C_MONTOAFECTABASEIMP) - vFmt(Grid.TextMatrix(i, C_MONTOAFECTABASEIMP))
             '3289932
            Dim sumaEgreso As Double
            
            sumaEgreso = sumaEgreso + vFmt(Grid.TextMatrix(i, C_MONTOAFECTABASEIMP))
            vMontoBaseImpoEgreso = sumaEgreso
             '3289932
         End If
         Tot(C_INGRESO) = Tot(C_INGRESO) + vFmt(Grid.TextMatrix(i, C_INGRESO))
         Tot(C_EGRESO) = Tot(C_EGRESO) + vFmt(Grid.TextMatrix(i, C_EGRESO))
         If Trim(Grid.TextMatrix(i, C_TIPODOC)) <> "" Then
            Tot(C_SALDO) = vFmt(Grid.TextMatrix(i, C_SALDO))
         End If
      End If
      
   Next i
   
   GridTot.TextMatrix(0, C_AFECTO) = Format(Tot(C_AFECTO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_EXENTO) = Format(Tot(C_EXENTO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_IVA) = Format(Tot(C_IVA), NEGNUMFMT)
   GridTot.TextMatrix(0, C_OTROIMP) = Format(Tot(C_OTROIMP), NEGNUMFMT)
   GridTot.TextMatrix(0, C_PAGADO) = Format(Tot(C_PAGADO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_IVAIRREC) = Format(Tot(C_IVAIRREC), NEGNUMFMT)
   GridTot.TextMatrix(0, C_TOTAL) = Format(Tot(C_TOTAL), NEGNUMFMT)
   GridTot.TextMatrix(0, C_INGRESO) = Format(Tot(C_INGRESO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_EGRESO) = Format(Tot(C_EGRESO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(Tot(C_SALDO), NEGNUMFMT)
   GridTot.TextMatrix(0, C_MONTOAFECTABASEIMP) = Format(Tot(C_MONTOAFECTABASEIMP), NEGNUMFMT)
   
   '3289932
   If lOper = O_EDIT And lTipoOper = 1 Then
    vMontoBaseImpoIngreso = Tot(C_MONTOAFECTABASEIMP)
   ElseIf lOper = O_EDIT And lTipoOper = 2 Then
    vMontoBaseImpoEgreso = Tot(C_MONTOAFECTABASEIMP)
   End If
   '3289932
   
   
End Sub
Private Sub CalcSaldo(ByVal Row As Integer, Grid As FEd2Grid)

   If Val(Grid.TextMatrix(Row, C_IDTIPOOPER)) = TOPERCAJA_INGRESO Then
'      If Val(Grid.TextMatrix(Row, C_ESREBAJA)) <> 0 Then                     'Claudio Villegas solicitó este cambio a raiz de una consulta de una clienta especialista en 14 TER Fecha: 8 ago 2018
'         Grid.TextMatrix(Row, C_INGRESO) = ""
'         Grid.TextMatrix(Row, C_EGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'      Else
         Grid.TextMatrix(Row, C_INGRESO) = Grid.TextMatrix(Row, C_PAGADO)
         Grid.TextMatrix(Row, C_EGRESO) = ""
'      End If
   Else
'      If Val(Grid.TextMatrix(Row, C_ESREBAJA)) <> 0 Then
'         Grid.TextMatrix(Row, C_EGRESO) = ""
'         Grid.TextMatrix(Row, C_INGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'      Else
         Grid.TextMatrix(Row, C_EGRESO) = Grid.TextMatrix(Row, C_PAGADO)
         Grid.TextMatrix(Row, C_INGRESO) = ""
'      End If
   End If
   
   Grid.TextMatrix(Row, C_SALDO) = Format(vFmt(Grid.TextMatrix(Row - 1, C_SALDO)) + vFmt(Grid.TextMatrix(Row, C_INGRESO)) - vFmt(Grid.TextMatrix(Row, C_EGRESO)), NEGNUMFMT)
         
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
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadGrid
      
   Me.MousePointer = vbDefault
      
End Sub


Private Function GrabarParaContinuar(ByVal MsgOper As String) As Boolean
   Dim i As Integer
   Dim Upd As Boolean

   GrabarParaContinuar = False

   'vemos si el usuario hizo algún cambio
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_TIPOOPER) = "" Then
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
         Call SaveGrid(Grid)
      Else
         Exit Function
      End If
   End If
   
   GrabarParaContinuar = True

End Function

Private Sub SetOrderLst()
   Dim i As Integer
   Dim StrOrder As String

   StrOrder = " LibroCaja.TipoOper, LibroCaja.FechaOperacion, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.NumDoc"
   
   lOrdenGr(C_NUMLIN) = "LibroCaja.IdLibroCaja"
   lOrdenGr(C_TIPOOPER) = StrOrder
   lOrdenGr(C_NUMDOC) = "LibroCaja.NumDoc, LibroCaja.TipoDoc, LibroCaja.FechaOperacion "
   lOrdenGr(C_TIPODOC) = "LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, LibroCaja.FechaOperacion"
   lOrdenGr(C_DTE) = "LibroCaja.DTE, " & StrOrder
   lOrdenGr(C_TIPODOCEXT) = lOrdenGr(C_TIPODOC)
   lOrdenGr(C_RUT) = "iif( LibroCaja.IdEntidad = 0, LibroCaja.RutEntidad, Entidades.Rut)," & StrOrder
   lOrdenGr(C_NOMBRE) = "iif( LibroCaja.IdEntidad = 0, LibroCaja.NombreEntidad, Entidades.Nombre)," & StrOrder
   lOrdenGr(C_FECHAOPER) = "LibroCaja.FechaOperacion, LibroCaja.TipoOper, LibroCaja.TipoDoc, LibroCaja.NumDoc"
   lOrdenGr(C_AFECTO) = "LibroCaja.Afecto," & StrOrder
   lOrdenGr(C_IVA) = "LibroCaja.IVA," & StrOrder
   lOrdenGr(C_EXENTO) = "LibroCaja.Exento," & StrOrder
   lOrdenGr(C_OTROIMP) = "LibroCaja.OtroImp," & StrOrder
   lOrdenGr(C_TOTAL) = "LibroCaja.Total," & StrOrder
   lOrdenGr(C_PAGADO) = "LibroCaja.Pagado," & StrOrder
   lOrdenGr(C_DESCRIP) = "LibroCaja.Descrip," & StrOrder
   lOrdenGr(C_CONENTREL) = "LibroCaja.ConEntRel," & StrOrder
   lOrdenGr(C_OPERDEVENGADA) = "LibroCaja.OperDevengada," & StrOrder
   lOrdenGr(C_PAGOAPLAZO) = "LibroCaja.PagoAPlazo," & StrOrder
   lOrdenGr(C_FECHAEXIGPAGO) = "LibroCaja.FechaExigPago," & StrOrder
   lOrdenGr(C_INGRESO) = lOrdenGr(C_NUMLIN)
   lOrdenGr(C_EGRESO) = lOrdenGr(C_NUMLIN)
   lOrdenGr(C_SALDO) = lOrdenGr(C_NUMLIN)
   
   
   lOrdenGr(C_USUARIO) = "Usuarios.Usuario" & StrOrder
   
   lOrdenSel = C_TIPOOPER
End Sub

Private Function SetupPriv()
     
   If lOper = O_EDIT Or lOper = O_NEW Then
      lEditEnabled = True
   End If
   
   If lOper = O_EDIT Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         Call EnableForm(Me, False)
         lEditEnabled = False
      End If
   
   End If

End Function

Private Function ValidaEstadoEdit(ByVal Row As Integer) As Boolean
   
   ValidaEstadoEdit = lEditEnabled

End Function


Private Function IsValidLine(ByVal Row As Integer, Msg As String) As Boolean
   Dim ValLine As Boolean
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim FExigePago As Long, FRecep As Long
   Dim TipoDoc As Integer, TipoLib As Integer
   
   IsValidLine = False
   
   ValLine = True
   
   ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> "")
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el Tipo de Documento."
      Exit Function
   End If
   
   TipoDoc = Val(Grid.TextMatrix(Row, C_IDTIPODOC))
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   
   If (TipoDoc = LIBCAJA_OTROSINGINI And TipoLib = LIB_CAJAING) Or (TipoDoc = LIBCAJA_OTROSEGRINI And TipoLib = LIB_CAJAEGR) Then
      If CbItemData(Cb_Mes) <> 1 Then    'estos tipos de ingresos/egresos sólo se pueden ingresar en enero
         Msg = "Sólo se permite ingresar un Ingreso/Egreso Inicial en el mes de enero."
         Exit Function
      End If
   End If

   If Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> TDOC_VENTASINDOC And Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> TDOC_VALEPAGOELECTR Then    'venta sin documento o Vale Pago Electronico
      If Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> "IMP" Then
         ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_NUMDOC)) <> "")   'valor distinto de blanco
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
   
   If IsNumeric(Trim(Grid.TextMatrix(Row, C_NUMDOCHASTA))) Then
      ValLine = (ValLine And (Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) = 0 Or (Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) <> 0 And vFmt(Grid.TextMatrix(Row, C_NUMDOCHASTA)) >= vFmt(Grid.TextMatrix(Row, C_NUMDOC)))))
      If ValLine = False And Msg = "" Then
         Msg = "El rango de Números de Documentos es inválido."
         Exit Function
      End If
   End If
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAOPER)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la fecha de la operación."
      Exit Function
   End If
      
   Call FirstLastMonthDay(DateSerial(lAno, CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAOPER)) <= LastDay)
   If ValLine = False And Msg = "" Then
      Msg = "La fecha de la operación es posterior al último día de este mes."
      Exit Function
   End If
            
   
   'validamos sólo si no está borrada la línea y si la línea fue modificada
   If Not lMsgFechaErr And Grid.RowHeight(Row) > 0 And Grid.TextMatrix(Row, C_UPDATE) <> "" Then

      FRecep = Val(Grid.TextMatrix(Row, C_LNGFECHAINGRESOLIBRO))
      If Abs(DateDiff("m", FRecep, Val(Grid.TextMatrix(Row, C_LNGFECHAOPER)))) > 2 Then
         If MsgBox1("En la línea " & Grid.TextMatrix(Row, C_NUMLIN) & " la fecha de ingreso o recepción del documento es posterior a los dos meses siguientes de la fecha de emisión del mismo." & vbNewLine & vbNewLine & "¿Está seguro que desea almacenar esta información?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Function
         End If
         lMsgFechaErr = True
      End If

   End If
   
   If Val(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO)) > 0 And Val(Grid.TextMatrix(Row, C_LNGFECHAOPER)) > Val(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO)) Then
      Msg = "La fecha de la operación es posterior a la fecha exigibilidad pago."
      Exit Function
   End If
   
   ValLine = (ValLine And (Trim(Grid.TextMatrix(Row, C_RUT)) <> "" Or (Trim(Grid.TextMatrix(Row, C_RUT)) = "" And (Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO Or Not DocExigeRut(Row)))))
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el RUT."
      Exit Function
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) <> "" Then
      ValLine = ValLine And Not (Val(Grid.TextMatrix(Row, C_DOCIMPEXP)) = 0 And DocExigeRut(Row) And Not ValidCID(Grid.TextMatrix(Row, C_RUT)))
      If ValLine = False And Msg = "" Then
         Msg = "Debe ingresar un RUT válido para este tipo de documento."
         Exit Function
      End If
   End If
   
   If Not ValidaNumDoc(Row) Then
      Msg = "Documento inválido."
      Exit Function
   End If

'   If Grid.TextMatrix(Row, C_TIPODOC) = TDOC_MAQREGISTRADORA Then
'      If Trim(Grid.TextMatrix(Row, C_NUMFISCIMPR)) = "" Then
'         Msg = "Debe ingresar el número Fiscal de la Impresora."
'         Exit Function
'      End If
'
'      If Trim(Grid.TextMatrix(Row, C_NUMINFORMEZ)) = "" Then
'         Msg = "Debe ingresar el número de Informe ""Z""."
'         Exit Function
'      End If
'
'      If Val(Grid.TextMatrix(Row, C_NUMDOCHASTA)) = 0 Then
'         Msg = "Debe ingresar el número correlativo de vale-boleta final, según informe ""Z"", columna ""Num. Doc. Hasta""."
'         Exit Function
'      End If
'
'      If vFmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0 And vFmt(Grid.TextMatrix(Row, C_VENTASACUM)) = 0 Then
'         Msg = "Debe ingresar las ventas acumuladas según informe ""Z"", columna ""Ventas Acum. Informe Z""."
'         Exit Function
'      End If
'   End If
   
   
'   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
'      ValLine = (ValLine And (vfmt(Grid.TextMatrix(Row, C_EXENTO)) <> 0 Or (vfmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0 And vfmt(Grid.TextMatrix(Row, C_IVA)) <> 0)))
'      If ValLine = False And Msg = "" Then
'         Msg = "El total del documento es cero."
'         Exit Function
'      End If
'   End If


   If Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAC" Or Trim(Grid.TextMatrix(Row, C_TIPODOC)) = "FAV" Then    'factura de venta
      ValLine = (ValLine And vFmt(Grid.TextMatrix(Row, C_AFECTO)) <> 0)
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar el valor Afecto."
         Exit Function
      End If
   End If
     
   'si es factura de compra, n. cred de factura de compra o n. debito de factura de compra, debe tener IVA ret Parcial o Total
'   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
'      If Grid.TextMatrix(Row, C_TIPODOC) = "FCC" Or Grid.TextMatrix(Row, C_TIPODOC) = "NCF" Or Grid.TextMatrix(Row, C_TIPODOC) = "NDF" Or Grid.TextMatrix(Row, C_TIPODOC) = "FCV" Then
'
'         ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_OTROIMP)) = 0 Or (vFmt(Grid.TextMatrix(Row, C_OTROIMP)) <> 0 And Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0), False, True)))
'         If ValLine = False And Msg = "" Then
'            Msg = "Falta ingresar el detalle de IVA Retenido Parcial o Total. Utilice el botón 'Detalle documento seleccionado' para ingresar este dato."
'            Exit Function
'         End If
'      End If
'   End If
   
   'puede haber documentos con valor 0
   'If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
   '   ValLine = (ValLine And vfmt(Grid.TextMatrix(Row, C_TOTAL)) <> 0)
   'End If

   ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_DESCRIP)) <> "")
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la glosa de la operación."
      Exit Function
   End If

   IsValidLine = ValLine

End Function


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


Private Function ValidaNumDoc(ByVal Row As Integer) As Boolean
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
   Dim AuxTipoLib As Integer
  
   NumDoc = Trim(Grid.TextMatrix(Row, C_NUMDOC))
   TipoDoc = Val(Grid.TextMatrix(Row, C_IDTIPODOC))
   AuxTipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
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
      If Grid.TextMatrix(i, C_FECHAOPER) = "" Then
         Exit For
      End If
      
      If i <> Row And Grid.RowHeight(i) > 0 Then
         
         
         
         EqDoc = (TipoDoc = Val(Grid.TextMatrix(i, C_IDTIPODOC)) And NumDoc = Trim(Grid.TextMatrix(i, C_NUMDOC))) And DTE = (Grid.TextMatrix(i, C_DTE) <> "")
         If AuxTipoLib = LIB_VENTAS Then
            DocNotVal = EqDoc
         Else           ' LIB_COMPRAS o LIB_RETEN
            DocNotVal = EqDoc And IdEnt = Val(Grid.TextMatrix(i, C_IDENTIDAD))
         End If
         
         If DocNotVal Then
         
            If IdEnt = 0 Then  'no se ha ingresada la entidad, sólo se da un mensaje de advertencia
               If MsgBox1("¡Atención!" & vbNewLine & "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & ", sin entidad asociada, ya ha sido ingresado en este libro. Es posible que esté repetido." & vbNewLine & vbNewLine & "¿Desea verificar los datos antes de continuar?", vbQuestion + vbYesNo) = vbYes Then
                  Exit Function
               End If
               
            Else
             '2840454
                 If Grid.TextMatrix(i, C_PAGOAPLAZO) <> "x" Then
                    '3031507
                   If Grid.TextMatrix(Row, C_TIPODOC) <> "BOE" And Grid.TextMatrix(Row, C_TIPODOC) <> "BOV" And Grid.TextMatrix(Row, C_TIPODOC) <> "BEX" Then
                    MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en este libro.", vbExclamation + vbOKOnly
                    Exit Function
                   End If
                   '3031507
                 End If
                 'MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en este libro.", vbExclamation + vbOKOnly
                 'Exit Function
                 
              'fin 2840454
            End If
            
         End If
         
      End If
   Next i
      
   'ahora vemos si está en la DB, en otros meses. Esto ya no corre porque los documentos aparecen más de una vez por los ingresos percibidos (FCA 30 ago 2017)
      
'   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = '" & NumDoc & "' AND DTE = " & CInt(DTE)
'   Wh = " WHERE "
'   If AuxTipoLib = LIB_VENTAS Then
'      Wh = Wh & "TipoLib = " & AuxTipoLib & WhEquDoc
'   Else     ' LIB_COMPRAS o LIB_RETEN
'      Wh = Wh & "TipoLib = " & AuxTipoLib & WhEquDoc & " AND IdEntidad = " & IdEnt
'   End If
'   Wh = Wh & " AND Month(FEmision) <> " & CbItemData(Cb_Mes)
'
'   Q1 = "SELECT IdDoc, FEmision FROM Documento " & Wh
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'
'      If IdEnt = 0 Then  'no se ha ingresado al entidad, sólo se da un mensaje de advertencia
'         If MsgBox1("¡Atención!" & vbNewLine & "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & ", sin entidad asociada, ya ha sido ingresado en el libro del mes de " & gNomMes(Month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ". Es posible que esté repetido." & vbNewLine & vbNewLine & "¿Desea verificar antes de continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'            Call CloseRs(Rs)
'            Exit Function
'         End If
'
'      Else
'         MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en el libro del mes de " & gNomMes(Month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ".", vbExclamation + vbOKOnly
'         Call CloseRs(Rs)
'         Exit Function
'
'      End If
'
'   End If
'
'   Call CloseRs(Rs)
   
   'ahora vemos si está en el año anterior
      
   'se hace sólo en función Valida: al final se recorre todo el libro y se ve si ha sido ingresado el año anterior.
'   If lAno = gEmpresa.Ano Then
'      If Not ValidaNumDocAnoAnt(Row) Then
'         Exit Function
'      End If
'   End If
   
   ValidaNumDoc = True

End Function


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

   Call CalcSaldo(Row, Grid)
   
End Sub
Private Function EsIngresoTotal(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
   Dim TipoLib As Integer
      
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   EsIngresoTotal = False
   
   If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then
   
      Idx = GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
'      If Not gTipoDoc(Idx).TieneAfecto And Not gTipoDoc(Idx).TieneExento Then 'no se ingresa afecto y exento, sólo total
      If gTipoDoc(Idx).IngresarTotal <> 0 Then 'no se ingresa afecto y exento, sólo total
         EsIngresoTotal = True
      End If
      
   End If
   
End Function

Private Function EsDocExento(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
   Dim TipoLib As Integer
      
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   EsDocExento = False
   
   If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then
   
      Idx = GetTipoDoc(TipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
'      If gTipoDoc(Idx).TieneExento And Not gTipoDoc(Idx).TieneAfecto Then 'sólo exento
      If gTipoDoc(Idx).TieneExento <> 0 And gTipoDoc(Idx).TieneAfecto = 0 Then 'sólo exento
         EsDocExento = True
      End If
      
   End If
   
End Function
Private Function DocExigeRut(ByVal Row As Integer) As Boolean
   Dim Idx As Integer
   Dim TipoLib As Integer
      
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   DocExigeRut = False
   
   If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then
      
      Idx = GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
      DocExigeRut = gTipoDoc(Idx).ExigeRUT
      
   End If
   
End Function

Private Sub CalcIngresoTotal(ByVal Row As Integer, ByVal Col As Integer, Value As String)
   Dim IdxDoc As Long
   Dim IVA As Double, Exento As Double, Afecto As Double
   Dim TipoLib As Integer
      
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))
   
   If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Or TipoLib = LIB_RETEN Then
         
      IdxDoc = GetTipoDoc(TipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
      
      If gTipoDoc(IdxDoc).EsRebaja Then
         Value = Format(vFmt(Value) * -1, NEGNUMFMT)
      Else
         Value = Format(vFmt(Value), NEGNUMFMT)
      End If
      
      Grid.TextMatrix(Row, Col) = Value
        
      If InStr(LCase(gTipoDoc(IdxDoc).Nombre), "exent") > 0 Or InStr(LCase(gTipoDoc(IdxDoc).Nombre), "sin") > 0 Then
         
         'es exenta o venta SIN documento
         
         IVA = 0
         Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
         
         Exento = vFmt(Grid.TextMatrix(Row, C_TOTAL))
         Grid.TextMatrix(Row, C_EXENTO) = Format(Exento, NEGNUMFMT)
         Grid.TextMatrix(Row, C_AFECTO) = Format(0, NEGNUMFMT)
                        
      Else
      
         IVA = Round(vFmt(Grid.TextMatrix(Row, C_TOTAL)) * (1 - (1 / (1 + gIVA))))
         Grid.TextMatrix(Row, C_IVA) = Format(IVA, NEGNUMFMT)
         
         Afecto = vFmt(Grid.TextMatrix(Row, C_TOTAL)) - IVA
         Grid.TextMatrix(Row, C_AFECTO) = Format(Afecto, NEGNUMFMT)
         Grid.TextMatrix(Row, C_EXENTO) = Format(0, NEGNUMFMT)
         
      End If
         
      Call CalcTot
   End If
End Sub

Private Sub bt_Copy_Click()
   Clipboard.Clear
   Clipboard.SetText Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col)

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
   
   ValidCol = (Col = C_NUMDOC Or Col = C_NUMDOCHASTA Or Col = C_RUT Or Col = C_EXENTO Or Col = C_AFECTO Or Col = C_IVA Or Col = C_OTROIMP Or Col = C_FECHAOPER Or Col = C_DESCRIP)

   If Not ValidCol Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Not ValidaEstadoEdit(Row) Then
      MsgBox1 "Este documento no puede ser modificado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      If Col < C_PAGADO Or Col = C_DESCRIP Then
         Exit Sub
      End If
   End If
                           
   
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
         
         If gTipoDoc(GetTipoDoc(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_IDTIPODOC)))).EsRebaja Then
            DVal = DVal * -1
         End If
            
         Grid.TextMatrix(Row, Col) = Format(DVal, NEGNUMFMT)
         
         Call CalcTotRow(Row, True)
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         lModifica = True
         Call CalcTot
         
      End If
   
   Else
      
      If Col = C_FECHAOPER Then
         Grid.TextMatrix(Row, C_LNGFECHAOPER) = GetDate(Clipboard.GetText, "dmy")
         Grid.TextMatrix(Row, C_FECHAOPER) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)), EDATEFMT)
      ElseIf Col = C_FECHAEXIGPAGO Then
         Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = GetDate(Clipboard.GetText, "dmy")
         Grid.TextMatrix(Row, C_FECHAEXIGPAGO) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO)), EDATEFMT)
      Else
         Grid.TextMatrix(Row, Col) = Clipboard.GetText
      End If
      
      Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
      lModifica = True
   
   End If
      
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   'Call FGr2Clip(Grid, gTipoLib(lTipoLib) & vbTab & Cb_Mes & " " & Val(Cb_Ano))
   Clip = FGr2String(Grid, Me.Caption & vbTab & gTipoOperCaja(lTipoOper) & vbTab & Cb_Mes & " " & Val(Cb_Ano), False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTot)
   
   Clipboard.Clear
   Clipboard.SetText Clip

End Sub

Private Sub Bt_Duplicate_Click()
   Dim i As Integer
   Dim Row As Integer
   
   If Grid.TextMatrix(Grid.Row, C_TIPOOPER) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Row = 0
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_TIPOOPER) = "" Then
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

   For i = C_IDTIPOOPER To C_SALDO
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
   lModifica = True
   
   Call CalcTot
      
   Grid.Row = Row
   Grid.RowSel = Grid.Row
   Grid.FlxGrid.Col = C_NUMLIN
   Grid.ColSel = Grid.Cols - 1
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid, Grid.FlxGrid.Row, Grid.FlxGrid.RowSel, Grid.FlxGrid.Col, Grid.FlxGrid.ColSel)
   
   Set Frm = Nothing

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Frmu As FrmSalyTotLibCajas
   Dim Pag As Integer
   Dim FrmPrt As FrmPrtSetup
   Dim FrmSald As FrmSalyTotLibCajas
   Dim OldOrientacion As Integer
   
   If Bt_List.Enabled Then
      MsgBox1 "Debe presionar el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If SaveBeforePrint() = False Then
      Exit Sub
   End If
   
'   If Grid.ColWidth(C_NUMDOCHASTA) > 0 Then
'      If MsgBox1("En la impresión se utilizará letra pequeña para visualizar la mayor cantidad de datos posible." & vbNewLine & vbNewLine & "Si no utiliza la columna ""N° Doc. Hasta"", ocúltela seleccionando las ""Opciones de Vista/Edición""," & vbCrLf & vbCrLf & "en la parte inferior derecha de la ventana. De esta manera, el sistema realizará la impresión" & vbCrLf & vbCrLf & "con tamaño de letra normal." & vbNewLine & vbNewLine & "¿Desea continuar con la impresión?", vbQuestion + vbYesNo) = vbNo Then
'         Exit Sub
'      End If
'   End If
   
   lPapelFoliado = False
      
   lOrientacion = ORIENT_HOR
   
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, False, lInfoPreliminar, False) = vbOK Then
      
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      Set FrmPrt = Nothing
      
      Set Frm = New FrmPrintPreview
                  
      gPrtLibros.PermitirMasDe1Franja = True
      gPrtLibros.FixedCols = C_TIPODOCEXT + 1
      
      gPrtLibros.CallEndDoc = True
      
      Pag = gPrtLibros.PrtFlexGrid(Frm)
      
      'nuevoooo
      
        Set Frmu = New FrmSalyTotLibCajas
        If CbItemData(Cb_Mes) = 0 Then
         Frmu.GetMonth = False
        Else
         Frmu.GetMonth = True
        End If
        Frmu.Fecha = DateSerial(Me.Cb_Ano, CbItemData(Cb_Mes), 1)
        'FormSaldos (Frm)
        Call SetUpPrtGrid2(Frmu.GetGrid())
        'Call Frmu.FPrint(frm)
        Frm.NewPage
        Pag = gPrtLibros.PrtFlexGrid(Frm)
        Set Frmu = Nothing
        
      '************
      
      Me.MousePointer = vbDefault
      
      Set Frm.PrtControl = Bt_Print
      
       '** lo nuevo **
        
'      Set FrmSald = New FrmSalyTotLibCajas
'      FrmSald.Fecha = DateSerial(Me.Cb_Ano, CbItemData(Cb_Mes), 1)
'      FrmSald.FPrint
'      Call PrtResumenIVA(frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, gPrtLibros.TieneMasDe1Franja)
'
'      Me.MousePointer = vbDefault
'
'      Set frm.PrtControl = Bt_Print
      '**********
      
      Call Frm.FView(Caption)
      
      Set Frm = Nothing
            
      
      gPrtLibros.GrFontName = Grid.Font.Name
      gPrtLibros.GrFontSize = Grid.Font.Size
      gPrtLibros.TotFntBold = True
      gPrtLibros.PermitirMasDe1Franja = False

      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
   End If
      
   Call ResetPrtBas(gPrtLibros)
  
End Sub
Private Function SaveBeforePrint() As Boolean
   Dim i As Integer
   Dim Modif As Boolean

   SaveBeforePrint = False
   Modif = False
   
   If Bt_OK.visible Then
   
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_TIPOOPER) = "" Then
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
            Call SaveGrid(Grid)
         Else
            Exit Function
         End If
      End If
   
   End If

   SaveBeforePrint = True

End Function

Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   Dim Pag As Integer
   Dim Frmu As FrmSalyTotLibCajas
   
   If Bt_List.Enabled Then
      MsgBox1 "Debe presionar el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If

   If SaveBeforePrint() = False Then
      Exit Sub
   End If
   
'   If Grid.ColWidth(C_NUMDOCHASTA) > 0 Then
'      If MsgBox1("En la impresión se utilizará letra pequeña para visualizar la mayor cantidad de datos posible." & vbNewLine & vbNewLine & "Si no utiliza la columna ""N° Doc. Hasta"", ocúltela utilizando las ""Opciones de Vista / Edición"" que se preoveen en la parte inferior derecha de la ventana." & vbNewLine & vbNewLine & "Alternativamente, seleccione orientación del papel ""Horizontal"" cuando utilice el botón ""Imprimir""." & vbNewLine & vbNewLine & "De esta manera, el sistema realizará la impresión con de tamaño letra normal." & vbNewLine & vbNewLine & "¿Desea continuar con la impresión?", vbQuestion + vbYesNo) = vbNo Then
'         Exit Sub
'      End If
'   End If
   
'   lPapelFoliado = False
'
'   If Ch_LibOficial.Visible = True And Ch_LibOficial <> 0 Then
'
'      If QryLogImpreso(lLibOf, CbItemData(Cb_Mes), FDesde, FHasta, Fecha, Usuario) = True Then
'         If MsgBox1("El " & gLibroOficial(lLibOf) & " Oficial del mes de " & gNomMes(CbItemData(Cb_Mes)) & " ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
'            Exit Sub
'         End If
'      End If
'
'      lPapelFoliado = True
'   End If
'
'   If Ch_ViewMaqReg Then
'      lOrientacion = ORIENT_HOR
'   End If
   
   lOrientacion = ORIENT_HOR
 
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Call SetUpPrtGrid
      
      gPrtLibros.PermitirMasDe1Franja = True
      gPrtLibros.FixedCols = C_TIPODOCEXT + 1
      
      gPrtLibros.CallEndDoc = 0
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
        'nuevoooo

        Set Frmu = New FrmSalyTotLibCajas
        If CbItemData(Cb_Mes) = 0 Then
         Frmu.GetMonth = False
        Else
         Frmu.GetMonth = True
        End If
        Frmu.Fecha = DateSerial(Me.Cb_Ano, CbItemData(Cb_Mes), 1)
        'FormSaldos (Frm)
        Call SetUpPrtGrid2(Frmu.GetGrid())
        'Call Frmu.FPrint(frm)
        'Frm.NewPage
        gPrtLibros.CallEndDoc = -1
        'gPrtLibros.EsContinuacion = True
        'gPrtLibros.TieneMasDe1Franja = False
        OldOrientacion = Printer.Orientation
        Printer.NewPage
        Printer.Orientation = OldOrientacion
        Pag = gPrtLibros.PrtFlexGrid(Printer)
        Set Frmu = Nothing
        'gPrtLibros.TieneMasDe1Franja = True
      '************
            
'      If lPapelFoliado And Ch_LibOficial.Visible = True And Ch_LibOficial <> 0 Then
'         Call AppendLogImpreso(lLibOf, CbItemData(Cb_Mes))
'      End If
      
'      Call SetPrtNotas 'para reponer las notas que se sacaron en SetupPrtGrid
      
      gPrtLibros.PermitirMasDe1Franja = False
      
      gPrtLibros.GrFontName = Grid.Font.Name
      gPrtLibros.GrFontSize = Grid.Font.Size
      gPrtLibros.TotFntBold = True

      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
   End If
   
   Call ResetPrtBas(gPrtLibros)

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
   
   With Grid
    .ColWidth(C_CONENTREL) = 0
    .ColWidth(C_OPERDEVENGADA) = 0
    .ColWidth(C_PAGOAPLAZO) = 0
    .ColWidth(C_FECHAEXIGPAGO) = 0
    
    If lTipoOper = TOPERCAJA_INGRESO Then
        .ColWidth(C_INGRESO) = 1200
        .ColWidth(C_EGRESO) = 0
    ElseIf lTipoOper = TOPERCAJA_EGRESO Then
        .ColWidth(C_INGRESO) = 0
        .ColWidth(C_EGRESO) = 1200
    Else
        .ColWidth(C_INGRESO) = 1200
        .ColWidth(C_EGRESO) = 1200
    End If
    
   End With
   
   
   Set gPrtLibros.Grid = Grid.FlxGrid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = Me.Caption & " " & gTipoOperCaja(lTipoOper)
   
   FontTit(0).FontBold = True
      
   If lOper = O_EDIT Then
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
   
   ColWi(C_TIPODOC) = 0
   ColWi(C_DTE) = 0
   ColWi(C_NOMBRE) = 0
   
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.NTotLines = 1
   gPrtLibros.ColObligatoria = C_SALDO
   
   gPrtLibros.Obs = ""   'para que no ponga las notas
   
'   If Ch_LibOficial <> 0 Then
'      gPrtLibros.PrintFecha = False
'   End If
    With Grid
        .ColWidth(C_CONENTREL) = 1000
        .ColWidth(C_OPERDEVENGADA) = 1000
        .ColWidth(C_PAGOAPLAZO) = 900
        .ColWidth(C_FECHAEXIGPAGO) = 1100
        .ColWidth(C_INGRESO) = 1200
        .ColWidth(C_EGRESO) = 1200
   End With
   
End Sub

Private Sub SetUpPrtGrid2(Grid2 As MSFlexGrid)
   Dim i As Integer
   Dim ColWi(7) As Integer
   Dim ColWi2(7) As Integer
   Dim Total(7) As String
   'Dim ColWi(NCOLS) As Integer
   'Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(3) As String
   Dim FontTit(2) As FontDef_t
   Dim OldOrient As Integer
   Dim Mes As String
   Dim Idx As Integer
   
   With Grid2
   .Row = 0
   .Height = 0
   End With
   'Grid2.RemoveItem (Grid2.Row)
   Set gPrtLibros.Grid = Grid2
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = "SALDOS Y TOTALES LIBRO DE CAJA"

   FontTit(0).FontBold = True
      
'   If lOper = O_EDIT Then
'      Titulos(1) = Titulos(1) & gNomMes(CbItemData(Cb_Mes)) & " " & lAno
'   Else
'      Titulos(1) = Titulos(1) & Cb_Mes & " " & Val(Cb_Ano)
'   End If
   
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
      
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   Encabezados(1) = "                FLUJO DE INGRESOS Y EGRESOS                          MONTOS QUE AFECTAN LA BASE IMPONIBLE"
   gPrtLibros.Encabezados = Encabezados
   
   gPrtLibros.GrFontName = Grid.Font.Name
   gPrtLibros.GrFontSize = Grid.Font.Size
   
   For i = 0 To Grid2.Cols - 1
      ColWi(i) = Grid2.ColWidth(i)
      'Total(i) = GridTot.TextMatrix(1, i)
   Next i

   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.NTotLines = 0
   'gPrtLibros.Total = Total
   'gPrtLibros.NTotLines = 1
   'gPrtLibros.ColObligatoria = C_SALDO
   
   gPrtLibros.Obs = ""   'para que no ponga las notas
   
'   If Ch_LibOficial <> 0 Then
'      gPrtLibros.PrintFecha = False
'   End If
   
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

Private Sub Tx_Glosa_Change()
   Bt_List.Enabled = True

End Sub

Private Sub Tx_NumDoc_Change()
   Bt_List.Enabled = True

End Sub

Private Sub Tx_Rut_Change()
   Bt_List.Enabled = True

End Sub

Private Sub Tx_Valor_Change()
   Bt_List.Enabled = True

End Sub
Private Sub Tx_NumDoc_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
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

End Sub
Private Sub Cb_Entidad_Click()
      
   Cb_Nombre.Clear
   If CbItemData(Cb_Entidad) >= 0 Then
      Call SelCbEntidad(CbItemData(Cb_Entidad))
   Else
      Tx_Rut = ""
   End If
   
   Bt_List.Enabled = True

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
Private Sub Bt_CerrarOpt_Click()
   Fr_Opciones.visible = False
End Sub
Private Sub Bt_Opciones_Click()

   If lOper <> O_EDIT Then
      Fr_Opciones.Caption = "Opciones de Vista"
   End If
   
   Fr_Opciones.visible = Not Fr_Opciones.visible
   
End Sub

Private Sub Ch_ViewOper_Click()

   If Ch_ViewOper = 0 Then
      Grid.ColWidth(C_TIPOOPER) = 0
      Grid.TextMatrix(1, C_TIPOOPER) = "Oper."
      
   Else
      Grid.ColWidth(C_TIPOOPER) = 500
      Grid.TextMatrix(1, C_TIPOOPER) = "Oper."
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerLCajaOper", Abs(Ch_ViewOper.Value))
   gVarIniFile.VerLCajaOper = Abs(Ch_ViewOper.Value)

End Sub
Private Sub Ch_ViewDTE_Click()

   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
      Grid.TextMatrix(0, C_DTE) = ""
      Grid.TextMatrix(1, C_DTE) = ""
      
   Else
      Grid.ColWidth(C_DTE) = 400
      Grid.TextMatrix(0, C_DTE) = ""
      Grid.TextMatrix(1, C_DTE) = "DTE"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerLCajaDTE", Abs(Ch_ViewDTE.Value))
   gVarIniFile.VerLCajaDTE = Abs(Ch_ViewDTE.Value)

End Sub
Private Sub Ch_ViewNombre_Click()

   If Ch_ViewNombre = 0 Then
      Grid.ColWidth(C_NOMBRE) = 0
      Grid.TextMatrix(1, C_NOMBRE) = ""
      
   Else
      Grid.ColWidth(C_NOMBRE) = 2000
      Grid.TextMatrix(1, C_NOMBRE) = "Nombre"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerLCajaNombre", Abs(Ch_ViewNombre.Value))
   gVarIniFile.VerLCajaNombre = Abs(Ch_ViewNombre.Value)

End Sub

Private Sub Ch_ViewIVAIrrec_Click()

   If Ch_ViewIVAIrrec = 0 Then
      Grid.ColWidth(C_IVAIRREC) = 0
      Grid.TextMatrix(0, C_IVAIRREC) = ""
      Grid.TextMatrix(1, C_IVAIRREC) = ""
      
   Else
      Grid.ColWidth(C_IVAIRREC) = 1200
      Grid.TextMatrix(0, C_IVAIRREC) = "Afectas"
      Grid.TextMatrix(1, C_IVAIRREC) = "IVA No Recup."
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerLCajaIVAIrrec", Abs(Ch_ViewIVAIrrec.Value))
   gVarIniFile.VerLCajaIVAIrrec = Abs(Ch_ViewIVAIrrec.Value)

End Sub


Private Sub Ch_ViewOtrosImp_Click()

   If Ch_ViewOtrosImp = 0 Then
      Grid.ColWidth(C_OTROIMP) = 0
      Grid.TextMatrix(1, C_OTROIMP) = ""
      
   Else
      Grid.ColWidth(C_OTROIMP) = 1200
      Grid.TextMatrix(1, C_OTROIMP) = "Otros Imp."
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerLCajaOtrosImp", Abs(Ch_ViewOtrosImp.Value))
   gVarIniFile.VerLCajaOtrosImp = Abs(Ch_ViewOtrosImp.Value)

End Sub

Private Function InsertSaldoInicial(ByVal Row As Integer, ByVal FDesde As Long) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Saldo As Double, Ingresos As Double, Egresos As Double
   Dim IdTipoDoc As Integer, TipoOper As Integer, TipoLib As Integer
   
   InsertSaldoInicial = False
   
   If lOper <> O_VIEW Then
      Exit Function
   End If
   
   Ingresos = 0
   Egresos = 0
   
   'Esto no se hace así por las NCC, NCF, NCV, NCE, DVB
'   Q1 = "SELECT Sum(Ingreso) as Ingresos, Sum(Egreso) as Egresos "
'   Q1 = Q1 & " FROM LIbroCaja "
'   Q1 = Q1 & " WHERE FechaIngresoLibro < " & FDesde
'   Q1 = Q1 & " WHERE FechaIngresoLibro < " & FDesde
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Q1 = "SELECT Sum(Ingreso + Egreso) as Ingresos"
   Q1 = Q1 & " FROM LIbroCaja "
   Q1 = Q1 & " WHERE FechaIngresoLibro < " & FDesde
   Q1 = Q1 & " AND TipoOper = " & TOPERCAJA_INGRESO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      Ingresos = vFld(Rs("Ingresos"))
      
   End If
         
   Call CloseRs(Rs)
   
   Q1 = "SELECT Sum(Egreso + Ingreso) as Egresos "
   Q1 = Q1 & " FROM LIbroCaja "
   Q1 = Q1 & " WHERE FechaIngresoLibro < " & FDesde
   Q1 = Q1 & " AND TipoOper = " & TOPERCAJA_EGRESO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      Egresos = vFld(Rs("Egresos"))
      
   End If
         
   Call CloseRs(Rs)
   
   '3271845
   'Saldo = Ingresos - Egresos + gSaldoLibroCajaAnoAnt
   Saldo = ObtenerMontoAperturaInicial(FDesde)
   
   '2854314
   If Saldo = 0 Then
   'Saldo = ObtenerMontoAperturaInicial(FDesde)
   Saldo = Ingresos - Egresos + gSaldoLibroCajaAnoAnt
   End If
   'fin 2854314
   '3271845
   
   Grid.TextMatrix(Row, C_NUMLIN) = 1
   
   TipoOper = IIf(Saldo >= 0, TOPERCAJA_INGRESO, TOPERCAJA_EGRESO)
   Grid.TextMatrix(Row, C_IDTIPOOPER) = TipoOper
   Grid.TextMatrix(Row, C_TIPOOPER) = UCase(Left(gTipoOperCaja(TipoOper), 1))
   
   If TipoOper = TOPERCAJA_INGRESO Then
      TipoLib = LIB_CAJAING
   Else
      TipoLib = LIB_CAJAEGR
   End If
   
   IdTipoDoc = IIf(Saldo >= 0, LIBCAJA_OTROSING, LIBCAJA_OTROSEGR)
   Grid.TextMatrix(Row, C_IDTIPODOC) = IdTipoDoc
   Grid.TextMatrix(Row, C_TIPODOC) = GetDiminutivoDoc(TipoLib, IdTipoDoc)
   Grid.TextMatrix(Row, C_TIPODOCEXT) = gTipoDocCajaOtros(IdTipoDoc)
   
   Grid.TextMatrix(Row, C_TOTAL) = Format(Abs(Saldo), NUMFMT)
   Grid.TextMatrix(Row, C_PAGADO) = Format(Abs(Saldo), NUMFMT)
   
   Grid.TextMatrix(Row, C_DESCRIP) = "Saldo Inicial"
   
   If TipoOper = TOPERCAJA_INGRESO Then
      Grid.TextMatrix(Row, C_INGRESO) = Format(Abs(Saldo), NUMFMT)
   Else
      Grid.TextMatrix(Row, C_EGRESO) = Format(Abs(Saldo), NUMFMT)
   End If
       
      
   Grid.TextMatrix(Row, C_SALDO) = Format(Saldo, NEGNUMFMT)
      
   InsertSaldoInicial = True
      
End Function
Private Function GetMontoQueAfectaBaseImp(Grid As FEd2Grid) As Double
   Dim Row As Integer
   Dim Valor As Double
   Dim TipoComp As Integer
   Dim TipoAjuste As Integer
   Dim TipoOperCaja As Integer
   Dim LstCuentas As String, AuxCuentas As String
   Dim Q1 As String
   
   
   If lTipoOper = TOPERCAJA_INGRESO Then
      TipoComp = TC_INGRESO
      TipoAjuste = TAEC_AGREGADOS
      TipoOperCaja = TOPERCAJA_INGRESO
   Else
      TipoComp = TC_EGRESO
      TipoAjuste = TAEC_DEDUCCIONES
      TipoOperCaja = TOPERCAJA_EGRESO
   End If

   
   For Row = Grid.FixedRows To Grid.rows - 1
   
      Valor = 0
   
      If Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)) <> 0 Then
      
         If lTipoOper = TOPERCAJA_INGRESO Then
         
            If Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_VENTAS Then
            
               If Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)) > 0 Then
                  'Ingresos Percibidos del Giro
                  Valor = GetPercibidosPagados(lTipoOper, FTE_14DN3, Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)))
                  If Valor = 0 Then
                     'Ingresos devengados o percibidos con empresas relacionadas acogidas al régimen 14 Letra A.
                     Valor = GetPercibidosPagados(lTipoOper, FTE_14A, Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)))
                  End If
               End If
               
            ElseIf Valor = 0 And Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then           'Otros Ingresos
                       
               'Desarrollo de una actividad agrícola
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 8).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = Mid(AuxCuentas, 2)
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
   
               'Arriendo de bienes raices agrícolas
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 9).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               'Arriendo de bienes raices no agrícolas
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 10).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               'Intereses de depósitos o instrumentos financieros
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 11).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               'Mayor valor en el rescate de cuotas de FM o FI
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 12).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               'Participación en contratos de participación o cuentas de participación
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 13).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               'Ingresos percibidos por la enajenación de bienes depreciables
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 14).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               Valor = Valor + LoadValCuentasAjustes14D(TipoAjuste, 0, TipoComp, LstCuentas, Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Otros ingresos percibidos
               Valor = Valor + GetTotCta_CodF22_14Ter(651, "C", Val(Grid.TextMatrix(Row, C_IDCOMP))) + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 15, Val(Grid.TextMatrix(Row, C_IDCOMP))) + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 19, Val(Grid.TextMatrix(Row, C_IDCOMP)))
                              
            End If
            
   
         ElseIf lTipoOper = TOPERCAJA_EGRESO Then
         
            If Grid.TextMatrix(Row, C_IDTIPOLIB) = LIB_COMPRAS Or Grid.TextMatrix(Row, C_IDTIPOLIB) = LIB_RETEN Then
            
               'Existencias o Insumos del Negocio Pagados
               If Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)) > 0 Then
                  Valor = GetPercibidosPagados(lTipoOper, 0, Val(Grid.TextMatrix(Row, C_IDLIBROCAJA)))
               End If
            
            ElseIf Valor = 0 And Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then           'Otros Egresos
                                 
               'Remuneraciones pagadas
               Valor = Valor + GetTotCta_CodF22_14Ter(631, "D", Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Intereses pagados por préstamos
               Valor = Valor + GetTotCta_CodF22_14Ter(633, "D", Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Arriendos pagados
               Valor = Valor + GetTotCta_CodF22_14Ter(1140, "D", Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Partidas pagadas del inciso 1° del art 21, no afectas al I.U.
               Valor = Valor + GetTotCta_CodF22_14Ter(1144, "D", Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Honorarios pagados
               LstCuentas = ""
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 4).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = Mid(AuxCuentas, 2)
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If

               'Impuestos que no sean de la LIR
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 5).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
              
               'Gastos afectos al inciso 1°, del art 21 LIR
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 6).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Gastos afectos al inciso 3°, del art 21 LIR
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 7).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Pago de IDPC
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 9).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Pago de IDPC AT 2020 o anteriores que depuran REX
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 10).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Gastos asociados a INR
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 11).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Pago 30% ISFUT
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 12).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
                
               'Otras partidas pagadas del inciso 2° del art 21, distintos de los anteriores
               AuxCuentas = gCtasAjusteExtraCont(TipoAjuste, 13).LstCuentas
               If AuxCuentas <> "" Then
                  LstCuentas = LstCuentas & AuxCuentas
                  LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
               End If
               
               Valor = Valor + LoadValCuentasAjustes14D(TipoAjuste, 0, TipoComp, LstCuentas, Val(Grid.TextMatrix(Row, C_IDCOMP)))
               
               'Otras Deducciones a la  RLI
               Valor = Valor + GetTotCta_CodF22_14Ter(635, "D", Val(Grid.TextMatrix(Row, C_IDCOMP))) '  + GetValAjustesELC(TAEC_DEDUCCIONES, 17, Val(Grid.TextMatrix(Row, C_IDCOMP))) + GetValAjustesELC(TAEC_DEDUCCIONES, 5, Val(Grid.TextMatrix(Row, C_IDCOMP))) + GetValAjustesELC(TAEC_DEDUCCIONES, 15, Val(Grid.TextMatrix(Row, C_IDCOMP))) + GetValAjustesELC(TAEC_DEDUCCIONES, 8, Val(Grid.TextMatrix(Row, C_IDCOMP)))
           
            End If
         End If
         
         If Valor <> 0 Then
            
            '2690461
            If Val(ReplaceStr(Grid.TextMatrix(Row, C_MONTOAFECTABASEIMP), ".", "")) < Valor Then
                 Valor = Val(ReplaceStr(Grid.TextMatrix(Row, C_MONTOAFECTABASEIMP), ".", ""))
                 Grid.TextMatrix(Row, C_MONTOAFECTABASEIMP) = Format(Valor, NEGNUMFMT)
                 
            Else
                  Grid.TextMatrix(Row, C_MONTOAFECTABASEIMP) = Format(Valor, NEGNUMFMT)
            End If
            'fin 2690461
            
            'Grid.TextMatrix(row, C_MONTOAFECTABASEIMP) = Format(Valor, NEGNUMFMT)
            Q1 = "UPDATE LibroCaja SET MontoAfectaBaseImp = " & Valor & " WHERE IdLibroCaja = " & Grid.TextMatrix(Row, C_IDLIBROCAJA)
            Call ExecSQL(DbMain, Q1)
            
         End If
         
         
      End If
      
   Next Row
   
End Function

 '2802201
 Private Sub deleteCompraIngreso()
 Dim Q1 As String
 Dim Rs As Recordset
 Dim TmpTbl As String, TmpTbl2 As String, TmpTbl3 As String

 '2841617 se agrega mes de proceso a funcion deleteCompraIngreso
 Dim FirstDay As Long
   Dim LastDay As Long
 'fin '2841617
 
 TmpTbl3 = DbGenTmpName2(gDbType, "tmplibcaja_3")
 TmpTbl2 = DbGenTmpName2(gDbType, "tmplibcaja_2")
 TmpTbl = DbGenTmpName2(gDbType, "tmplibcaja_")
 
   'Q1 = "SELECT Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, Documento.TipoDoc, Documento.TipoLib, "
   'Q1 = Q1 & " IIf(Documento.NumDocHasta<>'' And Documento.NumDocHasta<>' ',Documento.NumDoc+'-'+Documento.NumDocHasta,Documento.NumDoc) AS numDoc,"
   'Q1 = Q1 & " Documento.IdEntidad,  Documento.FEmisionOri,  Documento.Afecto,  Documento.Exento, Documento.OtroImp, Documento.Total "
   'Q1 = Q1 & " FROM ((Documento INNER JOIN tmp_USE2DEVP6424_tmplibcaja__4304 ON Documento.IdDoc = tmp_USE2DEVP6424_tmplibcaja__4304.IdDoc) "
   'Q1 = Q1 & " INNER JOIN tmp_USE2DEVP6424_tmplibcaja_3_4304 ON Documento.IdDoc = tmp_USE2DEVP6424_tmplibcaja_3_4304.IdDoc) LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad  "
   'Q1 = Q1 & " WHERE (Documento.IdEmpresa)  = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   '2841617 se agrega mes de proceso a funcion deleteCompraIngreso
'   If CbItemData(Cb_Mes) > 0 Then
'      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), Mes, 1), FirstDay, LastDay)
'   Else
'      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
'      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
'   End If
   'fin
   
   
   Q1 = Q1 & "  SELECT  Documento.IdDoc, Documento.IdEmpresa, Documento.Ano, Documento.TipoDoc, Documento.TipoLib "
   Q1 = Q1 & ", iif(Documento.NumDocHasta <> '' AND Documento.NumDocHasta <> ' ' , Documento.NumDoc + '-' + Documento.NumDocHasta, Documento.NumDoc) AS numDoc "
   Q1 = Q1 & " FROM ((Documento INNER JOIN " & TmpTbl & " ON Documento.IdDoc = " & TmpTbl & ".IdDoc  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl) & " )"
   Q1 = Q1 & " INNER JOIN " & TmpTbl3 & " ON Documento.IdDoc = " & TmpTbl3 & ".IdDoc   "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl3) & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True)
   Q1 = Q1 & " WHERE (Documento.IdEmpresa)  = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   '2841617 se agrega mes de proceso a funcion deleteCompraIngreso
   'Q1 = Q1 & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
    'fin
    '2857886
    Q1 = Q1 & " and tipodoc = 3 and tipolib= 1 "
   'fin 2857886
   
   Set Rs = OpenRs(DbMain, Q1)
   
    Do While Rs.EOF = False
         Q1 = ""
         Q1 = " WHERE iddoc = " & vFld(Rs("IdDoc"))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call DeleteSQL(DbMain, "LibroCaja", Q1)

       Rs.MoveNext

   Loop
         
   Call CloseRs(Rs)

 End Sub
 'fin 2802201

'2854314
Private Function ObtenerMontoAperturaInicial(ByVal FDesde As Long) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   ObtenerMontoAperturaInicial = 0
   
   If lOper <> O_VIEW Then
      Exit Function
   End If

   Q1 = "SELECT Debe "
   Q1 = Q1 & " FROM (Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp) INNER JOIN CtasAjustesExCont ON  "
   Q1 = Q1 & " MovComprobante.IdCuenta = CtasAjustesExCont.IdCuenta "
   Q1 = Q1 & " WHERE Comprobante.Fecha =  " & FDesde
   Q1 = Q1 & " AND Tipo = " & TC_APERTURA
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND Comprobante.TipoAjuste = " & TAJUSTE_FINANCIERO
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      ObtenerMontoAperturaInicial = vFld(Rs("Debe"))
      
   End If
         
   Call CloseRs(Rs)
   
      
End Function

'fin 2854314


'2955019
Private Sub deleteVentaEgreso(ByVal Mes As Integer)
Dim Q1 As String
Dim Rs As Recordset
Dim Where As String
Dim FirstDay As Long
Dim LastDay As Long
Dim TipoDoc As Integer
Dim TipoLib As Integer
Dim lTipoOper As Integer
        
         lTipoOper = CbItemData(Cb_TipoOper)
         TipoDoc = CbItemData(Cb_TipoDoc)
         
         If TipoDoc > 0 Then
         
            If lTipoOper = TOPERCAJA_INGRESO Then        '(LIB_VENTAS y LIB_CAJAING)
               If TipoDoc > BASELIBCAJA_INGEGR Then
                  TipoLib = LIB_CAJAING
                  TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
               Else
                  TipoLib = LIB_VENTAS
               End If
               
               'Egresos
            ElseIf TipoDoc > BASELIBCAJA_INGEGR Then     '(LIB_COMPRAS, LIB_RETEN y LIB_CAJAEGR)
               TipoLib = LIB_CAJAEGR
               TipoDoc = TipoDoc - BASELIBCAJA_INGEGR
            ElseIf TipoDoc > BASELIBCAJA_RETEN Then
               TipoLib = LIB_RETEN
               TipoDoc = TipoDoc - BASELIBCAJA_RETEN
            Else
               TipoLib = LIB_COMPRAS
            End If
            
            Where = " AND LibroCaja.TipoLib = " & TipoLib & " AND LibroCaja.TipoDoc = " & TipoDoc
         End If
      





'   se agrega mes de proceso a funcion deleteCompraIngreso
   If CbItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), Mes, 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   'fin
   Q1 = ""
   Q1 = Q1 & "  SELECT  IdLibroCaja,IdDoc "
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " WHERE (FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   If lTipoOper > 0 Then
      Q1 = Q1 & " AND LibroCaja.TipoOper = " & lTipoOper
   End If
   Q1 = Q1 & Where
  
   Q1 = Q1 & " AND LibroCaja.IdDoc in (select Lib1.IdDoc  FROM  LibroCaja as Lib1 "
   Q1 = Q1 & "WHERE (Lib1.FechaIngresoLibro BETWEEN " & FirstDay & " AND " & LastDay & ")"
   If lTipoOper > 0 Then
      Q1 = Q1 & " AND LibroCaja.TipoOper = " & lTipoOper
   End If
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Lib1.Pagado  = 0 )"
   
   Set Rs = OpenRs(DbMain, Q1)
   
    Do While Rs.EOF = False
         Q1 = ""
         Q1 = " WHERE IdLibroCaja = " & vFld(Rs("IdLibroCaja"))
         Q1 = Q1 & " And iddoc = " & vFld(Rs("IdDoc"))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call DeleteSQL(DbMain, "LibroCaja", Q1)

      Rs.MoveNext

  Loop
         
   Call CloseRs(Rs)

End Sub
'fin 2955019
