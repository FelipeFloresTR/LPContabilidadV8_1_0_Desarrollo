VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmLibRetenciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Retenciones y  Honorarios"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   Icon            =   "FrmLibRetenciones.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Centralizar_Full 
      Caption         =   "&Centralizar"
      Height          =   675
      Left            =   11400
      Picture         =   "FrmLibRetenciones.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Generar un comprobante con todos los documentos con tick"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Fr_Opciones 
      Caption         =   "Opciones de Edición"
      Height          =   1395
      Left            =   9420
      TabIndex        =   53
      Top             =   6300
      Width           =   3135
      Begin VB.CheckBox Ch_AplicarRet3Porc 
         Caption         =   "Aplicar Retención 3% Prést. Sol."
         Height          =   375
         Left            =   300
         TabIndex        =   57
         Top             =   840
         Width           =   2715
      End
      Begin VB.CheckBox Ch_ViewDTE 
         Caption         =   "Ver  DTE"
         Height          =   195
         Left            =   300
         TabIndex        =   56
         Top             =   300
         Width           =   1020
      End
      Begin VB.CheckBox Ch_ViewSucursal 
         Caption         =   "Ver  Sucursal"
         Height          =   195
         Left            =   300
         TabIndex        =   55
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton Bt_CerrarOpt 
         Caption         =   "X"
         Height          =   195
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.CommandButton Bt_Opciones 
      Caption         =   "Opciones de Vista/Edición"
      Height          =   375
      Left            =   10260
      TabIndex        =   52
      Top             =   7740
      Width           =   2295
   End
   Begin VB.CommandButton Bt_ToRight 
      Height          =   375
      Left            =   6960
      Picture         =   "FrmLibRetenciones.frx":0521
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Siguente conjunto de registros"
      Top             =   7740
      Width           =   315
   End
   Begin VB.CommandButton Bt_ToLeft 
      Height          =   375
      Left            =   6540
      Picture         =   "FrmLibRetenciones.frx":082B
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Anterior conjunto de registros"
      Top             =   7740
      Width           =   315
   End
   Begin VB.CommandButton Bt_DelAll 
      Caption         =   "Eliminar Todo..."
      Height          =   315
      Left            =   8820
      TabIndex        =   15
      ToolTipText     =   "Eliminar todos los documentos en estado Pendiente"
      Top             =   7800
      Width           =   1275
   End
   Begin VB.CommandButton Bt_HlpImport 
      Caption         =   "?"
      Height          =   315
      Left            =   8460
      TabIndex        =   14
      ToolTipText     =   "Formato del archivo de importación"
      Top             =   7800
      Width           =   255
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Capturar..."
      Height          =   315
      Left            =   7440
      TabIndex        =   13
      ToolTipText     =   "Capturar documentos desde archivo de texto"
      Top             =   7800
      Width           =   1035
   End
   Begin VB.TextBox Tx_CurrCell 
      Height          =   315
      Left            =   0
      TabIndex        =   43
      Top             =   7800
      Width           =   6375
   End
   Begin VB.CommandButton Bt_ExitNewDoc 
      Caption         =   "&Nuevo Doc"
      Height          =   615
      Left            =   9900
      Picture         =   "FrmLibRetenciones.frx":0B35
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Pc_Cent 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   11880
      Picture         =   "FrmLibRetenciones.frx":0FD2
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   11580
      Picture         =   "FrmLibRetenciones.frx":1436
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Bt_Centralizar 
      Caption         =   "&Centralizar"
      Height          =   675
      Left            =   11340
      Picture         =   "FrmLibRetenciones.frx":14AD
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Generar un comprobante con todos los documentos con tick"
      Top             =   840
      Width           =   1095
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5715
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   10081
      Cols            =   35
      Rows            =   5
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
      Width           =   12555
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
         Left            =   4500
         Picture         =   "FrmLibRetenciones.frx":19C2
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Resumen Retención 3% Préstamo Solidario"
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox Cb_Sucursal 
         Height          =   315
         Left            =   8220
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   "Sucursales"
         Top             =   180
         Width           =   2055
      End
      Begin VB.Frame Fr_BtEdit 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   60
         TabIndex        =   41
         Top             =   180
         Width           =   3495
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
            Picture         =   "FrmLibRetenciones.frx":1DB8
            Style           =   1  'Graphical
            TabIndex        =   45
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
            Left            =   840
            Picture         =   "FrmLibRetenciones.frx":222C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Seleccionar Entidad"
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
            Picture         =   "FrmLibRetenciones.frx":26CA
            Style           =   1  'Graphical
            TabIndex        =   42
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
            Left            =   1320
            Picture         =   "FrmLibRetenciones.frx":2AB3
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Left            =   3060
            Picture         =   "FrmLibRetenciones.frx":2E93
            Style           =   1  'Graphical
            TabIndex        =   20
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
            Left            =   2220
            Picture         =   "FrmLibRetenciones.frx":328F
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Duplicar documento seleccionado"
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
            Picture         =   "FrmLibRetenciones.frx":36E1
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Anular documento seleccionado"
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
            Picture         =   "FrmLibRetenciones.frx":3B52
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Plan de Cuentas"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   11520
         TabIndex        =   30
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   10560
         TabIndex        =   29
         Top             =   180
         Width           =   915
      End
      Begin VB.Frame Fr_BtGen 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   3600
         TabIndex        =   40
         Top             =   180
         Width           =   3975
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
            Picture         =   "FrmLibRetenciones.frx":3F13
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Detalle documento seleccionado"
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
            Left            =   1440
            Picture         =   "FrmLibRetenciones.frx":4378
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
            Left            =   1860
            Picture         =   "FrmLibRetenciones.frx":481F
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
            Left            =   2280
            Picture         =   "FrmLibRetenciones.frx":4CD9
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
            Left            =   480
            Picture         =   "FrmLibRetenciones.frx":511E
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
            Left            =   3180
            Picture         =   "FrmLibRetenciones.frx":51C2
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
            Left            =   2760
            Picture         =   "FrmLibRetenciones.frx":5523
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
            Left            =   3600
            Picture         =   "FrmLibRetenciones.frx":58C1
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Calendario"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Lb_Sucursal 
         AutoSize        =   -1  'True
         Caption         =   "Suc.:"
         Height          =   195
         Left            =   7800
         TabIndex        =   49
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Fr_List 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   660
      Width           =   11175
      Begin VB.CheckBox Ch_CentralizacionFull 
         Caption         =   "Centralizacion full"
         Height          =   195
         Left            =   8160
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   540
         Width           =   2175
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1155
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   2595
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   5040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton Bt_List 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   9900
         Picture         =   "FrmLibRetenciones.frx":5CEA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   47
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc.:"
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   4
         Left            =   5820
         TabIndex        =   35
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   5
         Left            =   3180
         TabIndex        =   34
         Top             =   240
         Width           =   390
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7440
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   615
      Left            =   11340
      Picture         =   "FrmLibRetenciones.frx":6128
      Style           =   1  'Graphical
      TabIndex        =   10
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
Attribute VB_Name = "FrmLibRetenciones"
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
Const C_DTE = 5
Const C_NUMDOC = 6
Const C_FECHAEMIORI = 7
Const C_LNGFECHAEMIORI = 8
Const C_CHECK = 9
Const C_IDENTIDAD = 10
Const C_RUT = 11
Const C_NOMBRE = 12
Const C_DESCRIP = 13
Const C_IDSUCURSAL = 14
Const C_SUCURSAL = 15
Const C_HONORSINRET = 16
Const C_BRUTO = 17
Const C_IDCUENTA = 18
Const C_CODCUENTA = 19
Const C_CUENTA = 20
Const C_PIMPTO = 21
Const C_IDPIMPTO = 22
Const C_IMPTO = 23
Const C_IMP_IDCUENTA = 24
Const C_RET3PORC = 25
Const C_RET3PORC_IDCUENTA = 26
Const C_NETO = 27
Const C_NETO_IDCUENTA = 28
Const C_TIPORETEN = 29
Const C_IDTIPORETEN = 30
Const C_DETALLE = 31
Const C_FECHAVENC = 32
Const C_LNGFECHAVENC = 33
Const C_IDESTADO = 34
Const C_ESTADO = 35
Const C_USUARIO = 36
Const C_MOVEDITED = 37
Const C_IDCOMPCENT = 38
Const C_IDCOMPPAGO = 39
Const C_UPDATE = 40

Const NCOLS = C_UPDATE

Const O_VIEWLIBLEGAL = -1

Const MITEM_OTRA = "(Otra)..."
Const TX_DETALLE = ">>"

Dim lInLoad As Boolean

Dim lIdDoc As Long

Dim lOper As Integer
Dim lRc As Integer
Dim lMes As Integer
Dim lAno As Integer
Dim IRutEmp As Boolean

'Dim lCurReg As Long     'primer registro del ragno de registros de la página actual
'Dim lNumReg As Long     'cantidad de registros en la grilla
'Dim lToRightPressed As Boolean   'para indicar que se presionó Bt_ToRight

Dim lClsPaging As ClsPaging   'Clase de paginamiento

Dim lTipoLib As Integer

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Const MAX_CUENTAS = 50
Dim lCuentas(MAX_CUENTAS) As Long

Dim lCtaBruto As Cuenta_t
Dim lCtaHonSinRet As Cuenta_t

Const M_IDENTIDAD = 1
Const M_RUT = 2
Dim lcbNombre As ClsCombo

Dim lOrdenGr(C_UPDATE) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

'indican si se puede editar/adm/ingresar en el libro actual, de acuerdo a los privilegios y a si está o no locked por otro usuario
Dim lEditEnabled As Boolean         'editar (todo)
Dim lAdmDocsEnabled As Boolean      'administrar (botón Bt_Centralizar)
Dim lIngDocsEnabled As Boolean      'ingresar (botón Bt_ExitNewDoc)

Dim lSucursal As Long

Dim lMsgPagadoNoCent As Boolean         'para no desplegar más el mensaje de pagado pero no centralizado

#If DATACON = 1 Then
Dim lDbAnoAnt As Database              'base de datos año anterior
#End If

Dim lFNameLogImp As String             'archivo de log de importación


'2784017
Dim NumReg As Integer

Private Sub Bt_Centralizar_Full_Click()
Me.MousePointer = vbHourglass
Call CentralizarFull
Me.MousePointer = vbDefault
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   'Call FGr2Clip(Grid, "Libro de Retenciones" & vbTab & "Mes: " & Cb_Mes & " " & Val(Cb_Ano))
   Clip = FGr2String(Grid, "Libro de Retenciones" & vbTab & "Mes: " & Cb_Mes & " " & Val(Cb_Ano), False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTot)
   
   Clipboard.Clear
   Clipboard.SetText Clip
End Sub

Private Sub Bt_DelAll_Click()
   Dim Row As Integer
   
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
      End If
      
   Next Row
      
   Call FGrVRows(Grid, 2)
      
   Call CalcTot
   
   Me.MousePointer = vbDefault
   
   MsgBox1 "Si presiona el botón Cancelar, se anulará esta operación.", vbInformation

End Sub

Private Sub Bt_HlpImport_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewLibReten
   
   Set Frm = Nothing


End Sub

Private Sub Bt_Importar_Click()
   Call ImportFromFile

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim FrmRes As FrmResRet3Porc
   Dim PrtPage As Object
   Dim FrmPrt As FrmPrtSetup
   Dim OldOrientacion As Integer
   
   '**** Graba la retencion en caso que no la visualize antes 02-12-2021 *******
   Call SaveGrid
   '****************************************************************************
   
   lPapelFoliado = False
   
   Set Frm = Nothing
   
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar, False) = vbOK Then
      
      OldOrientacion = Printer.Orientation
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      Set FrmPrt = Nothing
      
      Set Frm = New FrmPrintPreview
                  
      gPrtLibros.CallEndDoc = False
      
      Call gPrtLibros.PrtFlexGrid(Frm)
            
      Set PrtPage = Nothing
      Set PrtPage = GetPrtPage(Frm)
      
       Set FrmRes = New FrmResRet3Porc
    
  
    '2779115
    If GridTot.TextMatrix(0, C_RET3PORC) > 0 And CbItemData(Cb_Mes) >= 9 And Val(Cb_Ano) >= 2021 Then
       Set PrtPage = NewPage(Frm)
       Call FrmRes.FPrtRes(Frm, CbItemData(Cb_Mes), Val(Cb_Ano))
        Set FrmRes = Nothing
    Else
         Set FrmRes = Nothing
    End If
        
            
      Set Frm.PrtControl = Bt_Print
      Me.MousePointer = vbDefault
   End If
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call ResetPrtBas(gPrtLibros)
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   Dim nFolio As Integer
   Dim Pag As Integer
   Dim PrtPage As Object
   
   lPapelFoliado = False
   
   If Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(LIBOF_RETEN, lMes, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(LIBOF_RETEN) & " Oficial del mes de " & gNomMes(lMes) & " ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
      
      lPapelFoliado = True
   End If
   
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Call SetUpPrtGrid
      
      gPrtLibros.CallEndDoc = False
      
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
      If Pag > 0 Then
         Dim FrmRes As FrmResRet3Porc
         
         Set PrtPage = Nothing
         Set PrtPage = GetPrtPage(Printer)
         'PrtPage.NewPage 'comentado por 2779115

         
         Set FrmRes = New FrmResRet3Porc
         'Call FrmRes.FPrtRes(Printer, CbItemData(Cb_Mes), Val(Cb_Ano)) 'comentado por 2779115
         'Set FrmRes = Nothing 'comentado por 2779115
         
       
         
    '2779115
    If GridTot.TextMatrix(0, C_RET3PORC) > 0 And CbItemData(Cb_Mes) >= 9 And Val(Cb_Ano) >= 2021 Then
       PrtPage.NewPage
       Call FrmRes.FPrtRes(Printer, CbItemData(Cb_Mes), Val(Cb_Ano))
        Set FrmRes = Nothing
    Else
         Set FrmRes = Nothing
    End If
         
         
      End If
      Printer.EndDoc
      
      
      If lPapelFoliado And Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(LIBOF_RETEN, lMes)
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, nFolio)
      
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
   Dim Encabezados(0) As String
   Dim FontTit(2) As FontDef_t
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = UCase(gTipoLib(lTipoLib))
   FontTit(0).FontBold = True
   
   If ItemData(Cb_Sucursal) <> 0 Then
      Titulos(1) = "Sucursal: " & Cb_Sucursal & " - "
   End If
   
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Titulos(1) = Titulos(1) & gNomMes(ItemData(Cb_Mes)) & " " & lAno
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
   
   ColWi(C_NUMLIN) = ColWi(C_NUMLIN) - 60
   ColWi(C_NUMDOC) = ColWi(C_NUMDOC) - 120
   ColWi(C_RUT) = ColWi(C_RUT) - 170
   ColWi(C_IMPTO) = ColWi(C_IMPTO) - 140
   ColWi(C_FECHA) = ColWi(C_FECHA) - 130
   
   ColWi(C_NOMBRE) = 2200
   If Grid.ColWidth(C_DTE) > 0 Then
      ColWi(C_NOMBRE) = ColWi(C_NOMBRE) - 330
   End If
   
   ColWi(C_DESCRIP) = 0
   ColWi(C_SUCURSAL) = 0
   ColWi(C_ESTADO) = 0
   ColWi(C_CODCUENTA) = 0
   ColWi(C_CUENTA) = 0
   ColWi(C_CHECK) = 0
   'ColWi(C_FECHAEMIORI) = 0
   ColWi(C_FECHA) = 0
   ColWi(C_FECHAVENC) = 0
   ColWi(C_USUARIO) = 0
   ColWi(C_DETALLE) = 0

   gPrtLibros.GrFontSize = 8
   gPrtLibros.GrFontName = "Arial"
   gPrtLibros.TotFntBold = False
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.NTotLines = 1
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
End Sub

Private Sub Bt_Resumen_Click()
   Dim Frm As FrmResRet3Porc
   
   'FCA - 15/10/2021
   If MsgBox1("Antes de ingresar al resumen de Retención 3% se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   If Not valida() Then
      Exit Sub
   End If
   
   Call SaveGrid
   
   Set Frm = New FrmResRet3Porc
   Call Frm.FView(CbItemData(Cb_Mes), Val(Cb_Ano))
   Set Frm = Nothing
   
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

Private Sub Form_Load()
   Dim Q1 As String
   Dim i As Integer
   Dim j As Integer
   
   lInLoad = True
   
   lOrientacion = ORIENT_VER
   lFNameLogImp = gImportPath & "\Log\ImpRet-" & Format(Now, "yyyymmdd") & ".log"
   
   Set lClsPaging = New ClsPaging
   
   Call lClsPaging.Init(Bt_ToLeft, Bt_ToRight)
      
   Bt_Importar.visible = gFunciones.ImportRetenciones
   Bt_HlpImport.visible = gFunciones.ImportRetenciones
   Bt_DelAll.visible = gFunciones.ImportRetenciones
   
   lTipoLib = LIB_RETEN
   
   Fr_Opciones.visible = False

   
   Ch_LibOficial.visible = False
   
   'leemos las variables globales para setear las opciones de vista
   
   Ch_ViewDTE = gVarIniFile.VerLibRetDTE
   Ch_ViewSucursal = gVarIniFile.VerLibRetSucursal
   Ch_AplicarRet3Porc = gVarIniFile.VerRet3Porc
   
   Bt_DetDoc.visible = gFunciones.DetDocReten
      
   Select Case lOper
   
      Case O_VIEWLIBLEGAL
         Me.Caption = gTipoLib(lTipoLib)
         Ch_ViewSucursal.visible = False
         Ch_LibOficial.visible = True
         
      Case O_VIEW
         Me.Caption = "Listar " & gTipoLib(lTipoLib)
         Ch_AplicarRet3Porc.Caption = "Ver Retención 3% Prést. Sol."
         
      Case O_EDIT
         Me.Caption = "Editar " & gTipoLib(lTipoLib)

      Case O_SELECT
         Me.Caption = "Seleccionar Documento del " & gTipoLib(lTipoLib)

   End Select
   
   Call SetUpGrid
   
   '3217885
    If gDbType = SQL_SERVER Then
   Ch_CentralizacionFull.visible = True
   Else
   Ch_CentralizacionFull.visible = False
   End If
   '3217885
   
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Call AddItem(Grid.CbList(C_SUCURSAL), " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales WHERE IdEmpresa = " & gEmpresa.id
   If lOper = O_EDIT Then
      Q1 = Q1 & " AND Vigente <> 0 "
   End If
   Q1 = Q1 & " ORDER BY Descripcion"
   Call FillCombo(Grid.CbList(C_SUCURSAL), DbMain, Q1, -1)
   
   Call AddItem(Cb_Sucursal, " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales WHERE IdEmpresa = " & gEmpresa.id
   If lOper = O_EDIT Then
      Q1 = Q1 & " AND Vigente <> 0 "
   End If
   Q1 = Q1 & " ORDER BY Descripcion"
   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)
   
   If lOper = O_VIEW Or lOper = O_VIEWLIBLEGAL Or lOper = O_SELECT Then
      Grid.Locked = True
      Bt_Cancel.Caption = "Cerrar"
      Bt_OK.visible = False
      Fr_BtEdit.visible = False
      Fr_BtGen.Left = Fr_BtEdit.Left
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
   
   Call LoadDefCuentas
   
   Call SetupPriv
      
   Call LoadGrid
   
   If lOper = O_EDIT Then    'debe estar después del SetupPriv
   
      ' si está fuera del rango de fechas sept 2021 - dic 2024, debe estar deshabilitado
      If Not ((Val(Cb_Ano) = 2021 And CbItemData(Cb_Mes) >= 9) Or (Val(Cb_Ano) >= 2022 And Val(Cb_Ano) <= 2024)) Then
         Ch_AplicarRet3Porc = 0
         Ch_AplicarRet3Porc.Enabled = False
      End If
   End If
   
   
   If Val(Cb_Ano) = gEmpresa.Ano Then
   
   Call SetupRet3Porc         'para que valide rango de fechas y setee la columna
   
   Call RecalcRet3Porc    'por si alguna entidad cambió de asignación de Retención 3%
   
   End If
   
    '2784017
  
   Dim vCol As Integer
   Dim vRow As Integer
   Dim Neto As Double
   Dim Impto As Double
   Dim TotAntesImpuesto As Double

    For vRow = Grid.FixedRows To NumReg + 1
        For vCol = Grid.FixedCols To Grid.Cols - 1

        Select Case vCol

            Case C_PIMPTO
            
            '2794091
            Impto = 0
            ' fin 2794091

        If vFmt(Grid.TextMatrix(vRow, C_BRUTO)) > 0 Then
              If Val(Grid.TextMatrix(vRow, C_IDPIMPTO)) > 0 Then
                 Impto = Round(vFmt(Grid.TextMatrix(vRow, C_BRUTO)) * gImpRet(Val(Grid.TextMatrix(vRow, C_IDPIMPTO))))
              End If
        End If
        
        '2858854
        If Val(Grid.TextMatrix(vRow, C_IDPIMPTO)) <> 3 Then
         Grid.TextMatrix(vRow, C_IMPTO) = Format(Impto, NUMFMT)
         
        End If
        'end 2858854
   
      TotAntesImpuesto = (vFmt(Grid.TextMatrix(vRow, C_BRUTO)) + vFmt(Grid.TextMatrix(vRow, C_HONORSINRET)))
   
        Neto = TotAntesImpuesto - vFmt(Grid.TextMatrix(vRow, C_IMPTO))
        
         Neto = Neto - vFmt(Grid.TextMatrix(vRow, C_RET3PORC))
        
        Grid.TextMatrix(vRow, C_NETO) = Format(Neto, NUMFMT)

       Call FGrModRow(Grid, vRow, FGR_U, C_IDDOC, C_UPDATE)
                  Call CalcTot

             End Select

         Next vCol

    Next vRow
''
    'fin 2784017
   
   
   'limpiamos los títulos de la grilla de las coliumnas que tienen ancho cero
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) = 0 Then
         For j = 0 To Grid.FixedRows - 1
            Grid.TextMatrix(j, i) = ""
         Next j
      End If
   Next i

   lInLoad = False
End Sub

Private Sub Form_Resize()
   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 900
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 120
   GridTot.Width = Grid.Width - 230
   Grid.LeftCol = Grid.FixedCols  'C_NUMLIN
   GridTot.LeftCol = GridTot.FixedCols 'C_NUMLIN
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   Tx_CurrCell.Width = Bt_ToLeft.Left - Tx_CurrCell.Left - 200
   Bt_ToLeft.Top = Tx_CurrCell.Top
   Bt_ToRight.Top = Tx_CurrCell.Top
   Bt_Importar.Top = Tx_CurrCell.Top
   Bt_HlpImport.Top = Tx_CurrCell.Top
   Bt_Opciones.Top = Tx_CurrCell.Top
   Bt_DelAll.Top = Tx_CurrCell.Top
   
   Fr_Opciones.Top = Bt_Opciones.Top - Fr_Opciones.Height - 30
   
   Call FGrVRows(Grid)
End Sub



Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCopy(KeyCode, Shift) Then
      Call bt_Copy_Click
   ElseIf KeyPaste(KeyCode, Shift) Then
      Call Bt_Paste_Click
   End If


End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub

Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.Cols = NCOLS + 1
   
   Grid.ColWidth(C_IDDOC) = 0
   
   If lOper = O_VIEW Then
      Grid.ColWidth(C_CHECK) = 300
   Else
      Grid.ColWidth(C_CHECK) = 0
   End If
      
   Grid.ColWidth(C_NUMLIN) = 500
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Grid.ColWidth(C_FECHA) = 350
   Else
      Grid.ColWidth(C_FECHA) = 800
   End If
   Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_IDTIPODOC) = 0
   Grid.ColWidth(C_NUMDOC) = 800
   Grid.ColWidth(C_DTE) = 400
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_NOMBRE) = 2000
   Grid.ColWidth(C_DESCRIP) = 2000
   Grid.ColWidth(C_IDSUCURSAL) = 0
   Grid.ColWidth(C_SUCURSAL) = 1500
   Grid.ColWidth(C_HONORSINRET) = 1050
   Grid.ColWidth(C_BRUTO) = 1050
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CODCUENTA) = Me.TextWidth(gFmtCodigoCta) + 250
   Grid.ColWidth(C_CUENTA) = 1550
   Grid.ColWidth(C_IDPIMPTO) = 0
   Grid.ColWidth(C_PIMPTO) = 600
   Grid.ColWidth(C_IMPTO) = 1050
   Grid.ColWidth(C_IMP_IDCUENTA) = 0
   Grid.ColWidth(C_RET3PORC) = 1050
   Grid.ColWidth(C_RET3PORC_IDCUENTA) = 0
   Grid.ColWidth(C_NETO) = 1050
   Grid.ColWidth(C_NETO_IDCUENTA) = 0
   Grid.ColWidth(C_IDTIPORETEN) = 0
   Grid.ColWidth(C_TIPORETEN) = 860
   Grid.ColWidth(C_FECHAEMIORI) = 780
   Grid.ColWidth(C_LNGFECHAEMIORI) = 0
   Grid.ColWidth(C_FECHAVENC) = 780
   Grid.ColWidth(C_LNGFECHAVENC) = 0
   
   If gFunciones.DetDocReten Then
      Grid.ColWidth(C_DETALLE) = 300
   Else
      Grid.ColWidth(C_DETALLE) = 0
   End If
   
   Grid.ColWidth(C_ESTADO) = 1000
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_USUARIO) = 1000
   Grid.ColWidth(C_MOVEDITED) = 0
   Grid.ColWidth(C_IDCOMPCENT) = 0
   Grid.ColWidth(C_IDCOMPPAGO) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
   End If
      
   Grid.ColAlignment(C_NUMLIN) = flexAlignRightCenter
   Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
   Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESCRIP) = flexAlignLeftCenter
   Grid.ColAlignment(C_SUCURSAL) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_DETALLE) = flexAlignCenterCenter

   Grid.ColAlignment(C_HONORSINRET) = flexAlignRightCenter
   Grid.ColAlignment(C_BRUTO) = flexAlignRightCenter
   Grid.ColAlignment(C_IMPTO) = flexAlignRightCenter
   Grid.ColAlignment(C_RET3PORC) = flexAlignRightCenter
   Grid.ColAlignment(C_NETO) = flexAlignRightCenter
   
   Grid.TextMatrix(1, C_NUMLIN) = "Línea"
   If lOper = O_EDIT Or lOper = O_VIEWLIBLEGAL Then
      Grid.TextMatrix(1, C_FECHA) = "Día"
   Else
      Grid.TextMatrix(1, C_FECHA) = "Fecha"
   End If
   Grid.TextMatrix(1, C_TIPODOC) = "TD"
   Grid.TextMatrix(1, C_NUMDOC) = "Nº Doc."
   Grid.TextMatrix(1, C_DTE) = "DTE"
   Grid.TextMatrix(1, C_RUT) = "Rut"
   Grid.TextMatrix(1, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(1, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(1, C_SUCURSAL) = "Sucursal"
   Grid.TextMatrix(1, C_BRUTO) = "Bruto"
   Grid.TextMatrix(0, C_HONORSINRET) = "Honorarios"
   Grid.TextMatrix(1, C_HONORSINRET) = "s/Retención"
   Grid.TextMatrix(1, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(1, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(1, C_PIMPTO) = "%Imp."
   Grid.TextMatrix(1, C_IMPTO) = "Impuesto"
   Grid.TextMatrix(0, C_RET3PORC) = "Retención 3%"
   Grid.TextMatrix(1, C_RET3PORC) = "Prést. Sol."
   Grid.TextMatrix(1, C_NETO) = "Neto"
   Grid.TextMatrix(1, C_TIPORETEN) = "Tipo"
   Grid.TextMatrix(0, C_FECHAEMIORI) = "Fecha"
   Grid.TextMatrix(1, C_FECHAEMIORI) = "Emisión"
   Grid.TextMatrix(0, C_FECHAVENC) = "Fecha"
   Grid.TextMatrix(1, C_FECHAVENC) = "Vencim."
   Grid.TextMatrix(1, C_ESTADO) = "Estado"
   Grid.TextMatrix(1, C_USUARIO) = "Usuario"
      
   If lOper = O_VIEWLIBLEGAL Then
   
      lOrientacion = ORIENT_VER
      
      Grid.ColWidth(C_NOMBRE) = 2900
      Grid.ColWidth(C_DESCRIP) = 0
      Grid.ColWidth(C_SUCURSAL) = 0
      Grid.ColWidth(C_ESTADO) = 0
      Grid.ColWidth(C_CODCUENTA) = 0
      Grid.ColWidth(C_CUENTA) = 0
      Grid.ColWidth(C_FECHA) = 0
      'Grid.ColWidth(C_FECHAEMIORI) = 0
      Grid.ColWidth(C_FECHAVENC) = 0
      Grid.ColWidth(C_USUARIO) = 0
      Grid.ColWidth(C_DETALLE) = 0

   End If
   
   If Grid.ColWidth(C_DETALLE) <> 0 Then
      Grid.Row = 1
      Grid.Col = C_DETALLE
      Grid.CellPictureAlignment = flexAlignCenterCenter
      Set Grid.CellPicture = FrmMain.Pc_Lupa
   End If
   
   If lOper = O_VIEW Then
      Grid.Row = 1
      Grid.Col = C_CHECK
      Set Grid.CellPicture = Pc_Cent
   End If
   
   If lOper = O_EDIT Then
      Grid.FixedCols = 2
      Grid.ColWidth(C_DESCRIP) = Grid.ColWidth(C_DESCRIP) + 400
   Else
      Grid.FixedCols = 6
   End If

   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub Bt_AnulaDoc_Click()
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_IDDOC) = "" Then
      MsgBox1 "Este documento acaba de ser ingresado. Elimínelo directamente con el botón que sigue.", vbExclamation + vbOKOnly
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
   
   If GetMesActual() <= 0 Then
      MsgBox1 "No hay mes abierto. No es posible generar el comprobante.", vbExclamation
      Exit Sub
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_IDDOC) = "" Then
         Exit For
      End If
      
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
         StrIdDoc = StrIdDoc & "," & Grid.TextMatrix(i, C_IDDOC)
      End If
   Next i
      
   If StrIdDoc = "" Then
      'MsgBox1 "No hay documentos marcados para centralizar.", vbExclamation + vbOKOnly
      
      'If MsgBox1("¿Está seguro que desea centralizar todos los documentos pendientes en un solo comprobante?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
      If MsgBox1("¿Está seguro que desea centralizar todos los documentos pendientes, y los pagados pero no centralizados, en un solo comprobante?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      Else
         'marcamos todos los pendientes
         For i = Grid.FixedRows To Grid.rows - 1
            If Grid.TextMatrix(i, C_IDDOC) = "" Then
               Exit For
            End If
            
            Grid.Row = i
            Grid.Col = C_CHECK
                        
            'If Grid.TextMatrix(i, C_IDESTADO) = ED_PENDIENTE Then
            If Grid.TextMatrix(i, C_IDESTADO) = ED_PENDIENTE Or (Val(Grid.TextMatrix(i, C_IDESTADO)) = ED_PAGADO And vFmt(Grid.TextMatrix(i, C_IDCOMPCENT)) = 0) Then
               StrIdDoc = StrIdDoc & "," & Grid.TextMatrix(i, C_IDDOC)
               Call FGrSetPicture(Grid, i, C_CHECK, Pc_Check, 0)
            End If
            
         Next i
         
         If StrIdDoc = "" Then
            MsgBox1 "No hay documentos para centralizar.", vbExclamation + vbOKOnly
            Exit Sub
         End If
         
      End If
      
   ElseIf MsgBox1("¿Está seguro que desea centralizar todos los documentos marcados en un solo comprobante?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
       Exit Sub
       
   End If
   
   StrIdDoc = Mid(StrIdDoc, 2)

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
  
End Sub

Private Sub bt_Copy_Click()
   Clipboard.Clear
   Clipboard.SetText Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col)

End Sub

Private Sub Bt_Cuentas_Click()
   Dim IdCuenta As Long
   Dim Descrip As String
   Dim Nombre As String
   Dim Frm As FrmPlanCuentas
   Dim Col As Integer
   Dim Row As Integer
   Dim Codigo As String
   
   Col = Grid.Col
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Set Frm = New FrmPlanCuentas
      Call FrmPlanCuentas.FEdit(False)
      Set Frm = Nothing
   
      Exit Sub
   End If
            
   If Grid.TextMatrix(Row, C_FECHA) = "" Or Col <> C_CODCUENTA Or Not ValidaEstadoEdit(Row) Then
      Set Frm = New FrmPlanCuentas
      Call FrmPlanCuentas.FEdit(False)
      Set Frm = Nothing
      
      Exit Sub
   End If
   
   If Grid.Col = C_CODCUENTA Then
      Call LoadCuentas(Row)
      If M_ItCuenta.Count > 1 Then
         Call PopupMenu(M_Cuenta, , Grid.FlxGrid.ColPos(Grid.Col) + Grid.Left + 200, Grid.FlxGrid.RowPos(Grid.Row) + Grid.Top + 100)
      Else
         MsgBox1 "No hay cuentas definidas para este item del libro. Defínalas en la configuración de la empresa.", vbExclamation + vbOKOnly
      End If
   End If

End Sub

Private Sub Bt_Cut_Click()
   Dim ValidCol As Boolean
   
   ValidCol = (Grid.Col = C_NUMDOC Or Grid.Col = C_BRUTO Or Grid.Col = C_HONORSINRET Or Grid.Col Or Grid.Col = C_DESCRIP)
   
   If Not ValidCol Then
      Exit Sub
   End If
   
   Clipboard.Clear
   Call Clipboard.SetText(Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col))
   
   Grid.TextMatrix(Grid.FlxGrid.Row, Grid.FlxGrid.Col) = ""
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
   
   Call CalcTot
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
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
      MsgBox1 "Este documento no se puede borrar, ya que ha sido centralizado en un comprobante.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If MsgBox1("¿Está seguro que desea borrar este documento?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_IDDOC, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot
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

   For i = C_FECHA To C_NETO_IDCUENTA
      Grid.TextMatrix(Row, i) = Grid.TextMatrix(Grid.Row, i)
   Next i

   If Grid.Row > Grid.FixedRows Then
      Grid.TextMatrix(Row, C_NUMLIN) = Grid.TextMatrix(Row - 1, C_NUMLIN) + 1
   Else
      Grid.TextMatrix(Row, C_NUMLIN) = 1
   End If
   
   Grid.TextMatrix(Row, C_IDESTADO) = ED_PENDIENTE
   Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(ED_PENDIENTE)
  
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
Private Sub Bt_OK_Click()

   If valida() = False Then
      Exit Sub
   End If
   
   Call SaveGrid
   
   lRc = vbOK
   
   Unload Me

End Sub

Private Sub Bt_Paste_Click()
   Dim Fmt As Integer
   Dim DVal As Double
   Dim ValidCol As Boolean
   Dim Action As Integer
   Dim Value As String
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.FlxGrid.Row
   Col = Grid.FlxGrid.Col
   
   ValidCol = (Col = C_NUMDOC Or Col = C_RUT Or Col = C_BRUTO Or Col = C_HONORSINRET Or Col = C_DESCRIP Or Col = C_FECHAEMIORI Or Col = C_FECHAVENC)

   If Not ValidCol Then
      MsgBeep vbExclamation
      Exit Sub
   End If
      
   If Not ValidaEstadoEdit(Row) Then
      MsgBox1 "Este documento no puede ser modificado.", vbExclamation + vbOKOnly
      Exit Sub
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
   
   If (Col = C_HONORSINRET Or Col = C_BRUTO) Then
      If DVal <> 0 Then
         Grid.TextMatrix(Row, Col) = Format(Abs(DVal), NUMFMT)
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)
         Call CalcTot
      End If
   ElseIf Grid.Col = C_FECHAEMIORI Then
      Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = GetDate(Clipboard.GetText, "dmy")
      Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)), SDATEFMT)
   ElseIf Col = C_FECHAVENC Then
      Grid.TextMatrix(Row, C_LNGFECHAVENC) = GetDate(Clipboard.GetText, "dmy")
      Grid.TextMatrix(Row, C_FECHAVENC) = Format(vFmt(Grid.TextMatrix(Row, C_LNGFECHAVENC)), SDATEFMT)
   Else   'descripción
      Grid.TextMatrix(Row, Col) = Clipboard.GetText
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
   
'   lCurReg = 1
'   lNumReg = 0
   
   lClsPaging.Clear
'2784017
   Impuesto3RetPorc
'fin 2784017
   Call LoadGrid
   
   
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
   
   MsgBox "Total calculado = " & Format(Tot, NUMFMT), vbInformation + vbOKOnly
   
End Sub
Private Sub Bt_ToLeft_Click()
   Dim i As Integer
   Dim Modif As Boolean

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
   Dim i As Integer
   Dim Modif As Boolean

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

   If Not lInLoad Then
      Call SetupRet3Porc
   End If
   
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub
Private Sub Cb_Ano_Click()

   If Not lInLoad Then
      Call SetupRet3Porc
   End If
   
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub

Private Sub Cb_TipoDoc_Click()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False
End Sub
Public Function FEdit(ByVal Mes As Integer, ByVal Ano As Integer, IdDoc As Long) As Integer

   lOper = O_EDIT
   lMes = Mes
   lAno = Ano
   Me.Show vbModal
   
   FEdit = lRc
   
   IdDoc = lIdDoc
   
End Function
Public Sub FView(Optional ByVal Mes As Integer = 0)

   lOper = O_VIEW
   lAno = gEmpresa.Ano
   
   lMes = Mes
   
   Me.Show vbModal
End Sub
Public Function FSelect(IdDoc As Long) As Integer

   lOper = O_SELECT
   lAno = gEmpresa.Ano
   
   Me.Show vbModal
   
   IdDoc = lIdDoc
   FSelect = lRc
   
End Function

Public Sub FViewLibroLeg(Optional ByVal Mes As Integer = 0, Optional ByVal Ano As Integer = 0)

   lOper = O_VIEWLIBLEGAL
   lMes = Mes
   lAno = Ano
      
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
   Dim Cb As ComboBox
   
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
            Cb_Mes.ListIndex = GetUltimoMesConMovs
         End If
      End If
           
   'Else           'O_EDIT
         
   '   Cb_Mes.AddItem gNomMes(MesActual)
   '   Cb_Mes.ItemData(Cb_Mes.NewIndex) = MesActual
   '   Cb_Mes.ListIndex = 0
            
   'End If
      
   PrefLen = Len("Libro de") + 1
   
   Call LoadTipoDoc
   
   
   '2759379 se agregan los años segun el formulario FrmSelLibDoc
   If lAno > 0 Then
    Cb_Ano.AddItem gEmpresa.Ano - 5
    Cb_Ano.AddItem gEmpresa.Ano - 4
    Cb_Ano.AddItem gEmpresa.Ano - 3
   End If
   

   Cb_Ano.AddItem gEmpresa.Ano - 2
   Cb_Ano.AddItem gEmpresa.Ano - 1
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex
   If lAno > 0 Then
      For i = 0 To Cb_Ano.ListCount - 1
         If Val(Cb_Ano.list(i)) = lAno Then
            Cb_Ano.ListIndex = i
            Exit For
         End If
      Next i
   End If
   
   'Call LoadTipoDoc

   Call AddItem(Cb_Estado, "(todos)", 0)
   For i = 1 To UBound(gEstadoDoc)
      Call AddItem(Cb_Estado, gEstadoDoc(i), i)
   Next i
   Cb_Estado.ListIndex = 0
   
   Call AddItem(Cb_Entidad, "", -1)
   For i = ENT_CLIENTE To ENT_OTRO
      Call AddItem(Cb_Entidad, gClasifEnt(i), i)
      
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir

   Set Cb = Grid.CbList(C_PIMPTO)
   
   '2784017
  Dim CurYear As Long
      CurYear = DateSerial(lAno, 1, 1)
   
   gImpRet(IMPRET_NAC) = ImpBolHono(CurYear)
   ' fin 2784017
   
   Cb.AddItem gImpRet(IMPRET_NAC) * 100 & "%"
   Cb.ItemData(Cb.NewIndex) = IMPRET_NAC
   Cb.AddItem gImpRet(IMPRET_EXT) * 100 & "%"
   Cb.ItemData(Cb.NewIndex) = IMPRET_EXT
   Cb.AddItem "Otro"
   Cb.ItemData(Cb.NewIndex) = IMPRET_OTRO
   
   Set Cb = Grid.CbList(C_TIPORETEN)
   Cb.AddItem "Honorarios"
   Cb.ItemData(Cb.NewIndex) = TR_HONORARIOS
   Cb.AddItem "Dieta"
   Cb.ItemData(Cb.NewIndex) = TR_DIETA
   Cb.AddItem "Otro"
   Cb.ItemData(Cb.NewIndex) = TR_OTRO
   
   
   
End Sub
Private Sub LoadCuentas(ByVal Row As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
      
   If lTipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
      Q1 = Q1 & " WHERE TipoLib = " & lTipoLib & " AND TipoValor ="
   
      If vFmt(Grid.TextMatrix(Row, C_BRUTO)) > 0 Then
         Q1 = Q1 & LIBRETEN_BRUTO
      ElseIf vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) > 0 Then
         Q1 = Q1 & LIBRETEN_HONORSINRET
      Else
         MsgBox1 "Ingrese el valor del documento antes de seleccionar la cuenta.", vbExclamation + vbOKOnly
         Exit Sub
      End If
            
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
Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   Dim IdDoc As Long
   Dim ValPrevLine As Boolean
   Dim F1 As Long, F2 As Long
   Dim Msg As String
   Dim IdxTipoDoc As Integer
                           
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
      
      Grid.TextMatrix(Row, C_IDPIMPTO) = IMPRET_NAC
      Grid.TextMatrix(Row, C_PIMPTO) = gImpRet(IMPRET_NAC) * 100 & "%"
      Grid.TextMatrix(Row, C_IDTIPORETEN) = TR_HONORARIOS
      Grid.TextMatrix(Row, C_TIPORETEN) = gTipoRetencion(TR_HONORARIOS)
            
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
   
   IdxTipoDoc = GetTipoDoc(lTipoLib, Val(Grid.TextMatrix(Row, C_IDTIPODOC)))
   
   Select Case Col
   
      Case C_FECHA
     
         Grid.TxBox.MaxLength = 2
         If month(Now) <> ItemData(Cb_Mes) Then
            Call FirstLastMonthDay(DateSerial(lAno, ItemData(Cb_Mes), 1), F1, F2)
            Grid.TextMatrix(Row, Col) = Day(F2)
         Else
            Grid.TextMatrix(Row, Col) = Day(Now)
         End If
         
         If Row > Grid.FixedRows And Grid.TextMatrix(Row, C_TIPODOC) = "" Then
            Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
            Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
            If Grid.TextMatrix(Row, C_TIPODOC) = "BRT" Then  'Boleta Retención a Terceros
               If IsNumeric(Grid.TextMatrix(Row - 1, C_NUMDOC)) Then
                  Grid.TextMatrix(Row, C_NUMDOC) = Val(Grid.TextMatrix(Row - 1, C_NUMDOC)) + 1
               End If
            End If
         End If
                  
         EdType = FEG_Edit
         
      Case C_TIPODOC
      
         Grid.TxBox.MaxLength = 3
      
         If Row > Grid.FixedRows And Grid.TextMatrix(Row, Col) = "" Then
            Grid.TextMatrix(Row, C_TIPODOC) = Grid.TextMatrix(Row - 1, C_TIPODOC)
            Grid.TextMatrix(Row, C_IDTIPODOC) = Grid.TextMatrix(Row - 1, C_IDTIPODOC)
         End If
         
         EdType = FEG_Edit
         
      Case C_NUMDOC
     
         Grid.TxBox.MaxLength = MAX_NUMDOCLEN
         EdType = FEG_Edit
         
      Case C_DTE
      
         If Trim(Grid.TextMatrix(Row, Col)) = "" Then
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

      Case C_RUT
         Grid.TxBox.MaxLength = 13
         EdType = FEG_Edit
         
      Case C_SUCURSAL
         EdType = FEG_List
      
      Case C_BRUTO
         If vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) = 0 Then
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
         Else
            MsgBox1 "Si ingresa Honorarios sin Retención, no puede ingresar valor bruto.", vbExclamation
         End If
         
      Case C_HONORSINRET
         If vFmt(Grid.TextMatrix(Row, C_BRUTO)) = 0 Then
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
         Else
            MsgBox1 "Si ingresa valor bruto, no puede ingresar Honorarios sin Retención.", vbExclamation
         End If
         
      Case C_CODCUENTA
            
         Grid.TxBox.MaxLength = 20
         EdType = FEG_Edit
   
      Case C_DESCRIP
         Grid.TxBox.MaxLength = 100
         EdType = FEG_Edit
         
      Case C_PIMPTO
         EdType = FEG_List
         
      Case C_IMPTO
         'If Val(Grid.TextMatrix(Row, C_IDPIMPTO)) = IMPRET_OTRO Then
            Grid.TxBox.MaxLength = 15
            EdType = FEG_Edit
         'End If
         
      Case C_TIPORETEN
         EdType = FEG_List
         
      Case C_FECHAEMIORI
         
         Grid.TxBox.MaxLength = 9
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
            Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(Grid.TextMatrix(Row, C_LNGFECHAEMIORI), SDATEFMT)
         Else
            Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(DateSerial(lAno, ItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA))), SDATEFMT)
         End If
         
         EdType = FEG_Edit
               
      Case C_FECHAVENC
         Grid.TxBox.MaxLength = 10
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAVENC)) > 0 Then
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
         ElseIf vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
            Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)), SDATEFMT)
         Else
            Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, DateSerial(lAno, CbItemData(Cb_Mes), Val(Grid.TextMatrix(Row, C_FECHA)))))
            Grid.TextMatrix(Row, C_FECHAVENC) = Format(Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)), SDATEFMT)
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
   Static EnAcceptVal As Boolean
   Dim IdCuenta As Long
   Dim Cod As String
   Dim DescCta As String
   Dim NombCta As String
   Dim UltimoNivel As Boolean
   Dim ColId As Integer
   Dim ColCta As Integer
   Dim NotValidRut As Boolean
   Dim AuxRut As String
   Dim TipoDoc As Integer
   Dim Es14TER As Boolean
   Dim FVenc30Dias As Long
   
   IRutEmp = False

   
   If EnAcceptVal Then
      Exit Sub
   End If
   
   EnAcceptVal = True
   
   Action = vbOK
   
   Es14TER = gEmpresa.Franq14Ter                                                    'Claudio Villegas - 24 ago 2017
   
   Value = Trim(Value)
   Value = ReplaceStr(Value, vbCr, "")
   Value = ReplaceStr(Value, vbLf, "")
   
    
   Select Case Col
   
      Case C_FECHA
     
         Call FirstLastMonthDay(DateSerial(lAno, ItemData(Cb_Mes), 1), FirstDay, LastDay)
         
         If Val(Value) < 1 Or Val(Value) > Day(LastDay) Then
            MsgBox1 "Día inválido.", vbExclamation + vbOKOnly
            Value = Day(LastDay)
            Action = vbCancel
         End If
         
         If Grid.TextMatrix(Row, C_FECHAEMIORI) = "" Then
         
            Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = CLng(DateSerial(lAno, ItemData(Cb_Mes), Val(Value)))
            
            If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
               Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(Grid.TextMatrix(Row, C_LNGFECHAEMIORI), SDATEFMT)
               
               'proponemos fecha de vencimiento a 30 días
'               If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) = 0 Then
'                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
'                  Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
'               End If
            
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
         
         If Val(Grid.TextMatrix(Row, C_IDSUCURSAL)) = 0 And ItemData(Cb_Sucursal) <> 0 Then
            Grid.TextMatrix(Row, C_IDSUCURSAL) = ItemData(Cb_Sucursal)
            Grid.TextMatrix(Row, C_SUCURSAL) = Cb_Sucursal
         End If
            
      Case C_TIPODOC
                                                
         Value = Trim(Value)
         
         If Value <> "" Then
         
            TipoDoc = FindTipoDoc(lTipoLib, Value)
            
            If TipoDoc > 0 Then
               Grid.TextMatrix(Row, C_IDTIPODOC) = TipoDoc
               Grid.TextMatrix(Row, Col) = Value
                              
            Else
               MsgBox1 "Tipo de documento inválido. Presione el botón derecho del mouse para Ayuda.", vbExclamation + vbOKOnly
               Grid.TextMatrix(Row, C_TIPODOC) = ""
               Grid.TextMatrix(Row, C_IDTIPODOC) = 0
               Action = vbRetry
               
            End If
            
         End If
         
      Case C_NUMDOC
         Value = Trim(Value)
     
      Case C_RUT
      
         If Trim(Value) = "" Then
            Action = vbOK
            
         ElseIf Trim(Value) = "0-0" Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
         ' ado 2913643 Tema 3 - 1
         ElseIf vFmtCID(Trim(Value)) = gEmpresa.Rut Then
            Grid.TextMatrix(Row, C_RUT) = ""
            IRutEmp = True
         Else
         
            IdEnt = GetIdEntidad(Trim(Value), Nombre, NotValidRut)
            
            If IdEnt <= 0 Then
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
               Grid.TextMatrix(Row, C_NOMBRE) = Nombre
               Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
            End If
            
            Call ValidaNumDoc(Row)
            Call CalcTotRow(Row, True) 'FCA 22-10-2021
         End If
      
      Case C_SUCURSAL
         Grid.TextMatrix(Row, C_IDSUCURSAL) = ItemData(Grid.CbList(C_SUCURSAL))
         
      Case C_BRUTO, C_HONORSINRET
        'feña
         If Value <> "" And Col = C_BRUTO Then
            'Value = IIf(Value = "", 0, Value)
            Value = Format(Value, NUMFMT)
            If Value = 0 Then
               MsgBox1 "El Valor Bruto debe ser mayor a 0.", vbInformation
               Value = ""
            End If
         End If
         Grid.TextMatrix(Row, Col) = Value
        
         If Val(Grid.TextMatrix(Row, C_IDCUENTA)) = 0 Then
            If Col = C_BRUTO Then
               Grid.TextMatrix(Row, C_IDCUENTA) = lCtaBruto.id
               Grid.TextMatrix(Row, C_CODCUENTA) = FmtCodCuenta(lCtaBruto.Codigo)
               Grid.TextMatrix(Row, C_CUENTA) = lCtaBruto.Descripcion
            Else
               Grid.TextMatrix(Row, C_IDCUENTA) = lCtaHonSinRet.id
               Grid.TextMatrix(Row, C_CODCUENTA) = FmtCodCuenta(lCtaHonSinRet.Codigo)
               Grid.TextMatrix(Row, C_CUENTA) = lCtaHonSinRet.Descripcion
            End If
         End If
         
         Call CalcTotRow(Row, True)     'también calcula Ret 3%
                           
         Call CalcTot
         
         If vFmt(Grid.TextMatrix(Row, C_IMPTO)) <> 0 Then
            Grid.TextMatrix(Row, C_IMP_IDCUENTA) = gCtasBas.IdCtaImpRet
         Else
            Grid.TextMatrix(Row, C_IMP_IDCUENTA) = 0
         End If
         
         'Se asigna siempre la cuenta porque puede ser que se elimine el check de aplicar Ret 3% y luego se vuelva a poner
         'En el SaveAll sólo se graba la cuenta si el Ret3Porc tiene valor <> 0
         Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA) = gCtasBas.IdCtaRet3Porc
         
         
         If vFmt(Grid.TextMatrix(Row, C_NETO)) <> 0 Then
            If Val(Grid.TextMatrix(Row, C_IDTIPORETEN)) = TR_HONORARIOS Then
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoHon
            Else
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoDieta
            End If
         Else
            Grid.TextMatrix(Row, C_NETO_IDCUENTA) = 0
         End If
         
      Case C_PIMPTO
         Grid.TextMatrix(Row, C_IDPIMPTO) = IMPRET_NAC
         If ItemData(Grid.CbList(C_PIMPTO)) > 0 Then
            Grid.TextMatrix(Row, C_IDPIMPTO) = ItemData(Grid.CbList(C_PIMPTO))
         End If
         
         Call CalcTotRow(Row, True)
                           
         Call CalcTot
         
      Case C_IMPTO
         Value = Format(Value, NUMFMT)
         Grid.TextMatrix(Row, Col) = Value

         Call CalcTotRow(Row, False)
         
         If vFmt(Grid.TextMatrix(Row, C_BRUTO)) <> 0 And Val(Grid.TextMatrix(Row, C_IDPIMPTO)) <> IMPRET_OTRO Then
            If Round(vFmt(Grid.TextMatrix(Row, C_IMPTO)) / vFmt(Grid.TextMatrix(Row, C_BRUTO)), 2) <> vFmt(Grid.TextMatrix(Row, C_PIMPTO)) Then
               MsgBox1 "El impuesto ingresado no corresponde al porcentaje seleccionado.", vbExclamation
            End If
         End If
                      
         Call CalcTot
         
         If vFmt(Grid.TextMatrix(Row, C_IMPTO)) <> 0 Then
            Grid.TextMatrix(Row, C_IMP_IDCUENTA) = gCtasBas.IdCtaImpRet
         Else
            Grid.TextMatrix(Row, C_IMP_IDCUENTA) = 0
         End If
         
         If vFmt(Grid.TextMatrix(Row, C_NETO)) <> 0 Then
            If Val(Grid.TextMatrix(Row, C_IDTIPORETEN)) = TR_HONORARIOS Then
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoHon
            Else
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoDieta
            End If
         Else
            Grid.TextMatrix(Row, C_NETO_IDCUENTA) = 0
         End If
        
      Case C_CODCUENTA
      
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
                  
         If IdCuenta = 0 Then
         
            If Value <> "" Then
               MsgBeep vbExclamation
               Action = vbCancel
            End If
            
         ElseIf UltimoNivel = False Then
            MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
            Action = vbCancel
            
'         ElseIf Not EsCuentaBasica(IdCuenta, Row) Then
'            MsgBox1 "Esta cuenta no es válida para este tipo de valor, de acuerdo a la configuración básica de la empresa.", vbExclamation + vbOKOnly
'            Action = vbCancel
            
         Else
         
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
            Value = Format(Cod, gFmtCodigoCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta
         
         End If
         
      Case C_TIPORETEN
         Grid.TextMatrix(Row, C_IDTIPORETEN) = ItemData(Grid.CbList(C_TIPORETEN))
         
         If vFmt(Grid.TextMatrix(Row, C_NETO)) <> 0 Then
            If Val(Grid.TextMatrix(Row, C_IDTIPORETEN)) = TR_HONORARIOS Then
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoHon
            Else
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoDieta
            End If
         Else
            Grid.TextMatrix(Row, C_NETO_IDCUENTA) = 0
         End If
         
         Call CalcTotRow(Row, False)     'también calcula Ret 3%
         Call CalcTot
      
         
      Case C_FECHAEMIORI
         
         Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = GetDate(Value, "dmy")
         
         If vFmt(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0 Then
            Value = Format(Grid.TextMatrix(Row, C_LNGFECHAEMIORI), SDATEFMT)
            
            'proponemos fecha de vencimiento a 30 días
            If Val(Grid.TextMatrix(Row, C_LNGFECHAVENC)) = 0 Then
'               Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
'               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
               
               If Es14TER Then    'pago contado
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)))
               Else        'Crédito 30 días
                  Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
               End If
               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
               
'            ElseIf MsgBox1("¿Desea ajustar la fecha de vencimiento a 30 días?", vbQuestion + vbYesNo) = vbYes Then
'               Grid.TextMatrix(Row, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI))))
'               Grid.TextMatrix(Row, C_FECHAVENC) = Format(Grid.TextMatrix(Row, C_LNGFECHAVENC), SDATEFMT)
            
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
      
   If Col = C_CHECK Then
   
      If Bt_Centralizar.visible And Bt_Centralizar.Enabled Then
         'If Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PENDIENTE Or Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_APROBADO Then
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
         
   ElseIf Col = C_DETALLE Then
      Call PostClick(Bt_DetDoc)
      
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
         
      Case C_NUMDOC
         Call KeyNum(KeyAscii)
         
      Case C_RUT
         Call KeyName(KeyAscii)
         Call KeyUpper(KeyAscii)
         
      Case C_NOMBRE
         Call KeyName(KeyAscii)
         
      Case C_BRUTO, C_HONORSINRET
         Call KeyNum(KeyAscii)
         
      Case C_FECHAEMIORI, C_FECHAVENC
         Call KeyDate(KeyAscii)
         
      Case C_DESCRIP
         Call KeyName(KeyAscii)
         
   End Select
   
End Sub

Private Sub CalcTotRow(ByVal Row As Integer, Optional ByVal RecalcImp As Boolean)
   Dim Neto As Double
   Dim Impto As Double
   Dim Ret3Porc As Double
   Dim TotAntesImpuesto As Double
   Dim FDesde As Long, FHasta As Long
   Dim Ret3 As Boolean
   Dim dt2021 As Long, dt2022 As Long, dt2023 As Long, dt2024 As Long
   
   
   Impto = 0
   
   If RecalcImp Then
      If vFmt(Grid.TextMatrix(Row, C_BRUTO)) > 0 Then
         If Val(Grid.TextMatrix(Row, C_IDPIMPTO)) > 0 Then
            Impto = Round(vFmt(Grid.TextMatrix(Row, C_BRUTO)) * gImpRet(Val(Grid.TextMatrix(Row, C_IDPIMPTO))))
         End If
      End If
      Grid.TextMatrix(Row, C_IMPTO) = Format(Impto, NUMFMT)
   End If
   
   
   TotAntesImpuesto = (vFmt(Grid.TextMatrix(Row, C_BRUTO)) + vFmt(Grid.TextMatrix(Row, C_HONORSINRET)))
   
   Neto = TotAntesImpuesto - vFmt(Grid.TextMatrix(Row, C_IMPTO))
   Grid.TextMatrix(Row, C_NETO) = Format(Neto, NUMFMT)

    
    Dim fechaemi As Long
    fechaemi = GetDate(Grid.TextMatrix(Row, C_FECHAEMIORI))
   Ret3Porc = 0
   If Grid.TextMatrix(Row, C_IDENTIDAD) <> "" Then
    Ret3 = GetEntRet3Porc(Grid.TextMatrix(Row, C_IDENTIDAD), FDesde, FHasta)
   End If
   'Se calcula la Retención 3% si y sólo si:
   ' - está marcado Aplicar Ret 3%
   ' - Está en rango de fechas
   ' - Es tipo Honorarios
   ' - la entidad tiene marcado Ret 3%
   ' - Se calcula sobre el valor de la boleta antes de impuesto      'FCA - 12/10/2021
   dt2021 = DateSerial(gEmpresa.Ano - 4, 12, 31)
   dt2022 = DateSerial(gEmpresa.Ano - 3, 12, 31)
   dt2023 = DateSerial(gEmpresa.Ano - 2, 12, 31)
   dt2024 = DateSerial(gEmpresa.Ano - 1, 12, 31)
   If FHasta = dt2021 Or FHasta = dt2022 Or FHasta = dt2023 Or FHasta = dt2024 Then
        FDesde = DateSerial(gEmpresa.Ano, 1, 1)
        FHasta = DateSerial(gEmpresa.Ano, 12, 31)
   End If
   If Ch_AplicarRet3Porc <> 0 And ((Val(Cb_Ano) = 2021 And CbItemData(Cb_Mes) >= 9) Or (Val(Cb_Ano) >= 2022 And Val(Cb_Ano) <= 2024)) And Grid.TextMatrix(Row, C_IDTIPORETEN) = TR_HONORARIOS And Ret3 Then
      If FDesde > 0 And FHasta > 0 And fechaemi >= FDesde And fechaemi <= FHasta Then
        Ret3Porc = TotAntesImpuesto * 0.03
      ElseIf FHasta = 0 Then
        Ret3Porc = TotAntesImpuesto * 0.03
      Else
        Ret3Porc = 0
      End If
   Else
      Ret3Porc = 0
   End If
   
   
   Grid.TextMatrix(Row, C_RET3PORC) = Format(Ret3Porc, NUMFMT)
   Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA) = gCtasBas.IdCtaRet3Porc   'FCA 21/10/2021
   
   Neto = Neto - Ret3Porc
   Grid.TextMatrix(Row, C_NETO) = Format(Neto, NUMFMT)
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDDOC, C_UPDATE)   'esto porque puede haber cambiado el seteo del Ret3Porc para la entidad
   
      
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
   
   If gCtasBas.IdCtaImpRet <= 0 Or gCtasBas.IdCtaNetoHon <= 0 Then
      MsgBox1 "No es posible ingresar documentos sin antes definir la configuración de las cuentas de Impuesto Retenido y Neto Honorarios.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If gCtasBas.IdCtaRet3Porc <= 0 And Ch_AplicarRet3Porc <> 0 And ((Val(Cb_Ano) = 2021 And CbItemData(Cb_Mes) >= 9) Or (Val(Cb_Ano) >= 2022 And Val(Cb_Ano) <= 2024)) Then
      MsgBox1 "No es posible ingresar documentos sin antes definir la configuración de la cuenta de Retención 3% Préstamo Solidario.", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   'vemos si las líneas están completas
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If LineaEnBlanco(Row) Then
         Exit For
      End If
      
      If vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) = 0 Then
        If Grid.TextMatrix(Row, C_BRUTO) = "" Then
          MsgBox1 "Valor del Bruto tiene que ser mayor a 0. Línea " & Row - 1, vbExclamation + vbOKOnly
          Exit Function
        End If
      End If
   
      If Grid.RowHeight(Row) > 0 And Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then    'no ha sido borrada
         ValLine = IsValidLine(Row, Msg)
         ValLine = (ValLine Or Grid.RowHeight(Row) = 0)   'línea borrada
         
         If ValLine = False Then
            MsgBox1 "Línea " & Row - 1 & " inválida. " & Msg, vbExclamation + vbOKOnly
            Exit Function
         End If
      End If
      
   Next Row
   
   'abrimos base de datos año anterior para ver algún doc ya fue ingresado en año anterior
   
#If DATACON = 1 Then       'Access
   
   If gEmpresa.TieneAnoAnt Then
      DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
      If ExistFile(DbName) Then
         'Call OpenDb(lDbAnoAnt, DbName)
         Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
         'hacemos CorrigeBase de año anterior por si las moscas
         Set AuxDb = DbMain
         Set DbMain = lDbAnoAnt
         Call CorrigeBase
         Set DbMain = AuxDb
      End If
   End If
   
#End If

   'vemos si no hay documentos repetidos
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If LineaEnBlanco(Row) Then
         Exit For
      End If
      
      If Grid.TextMatrix(Row, C_UPDATE) <> "" And Grid.RowHeight(Row) > 0 Then
         If Not ValidaNumDoc(Row) Then
            Call FGrSelRow(Grid, Row)
            Exit Function
         End If
      End If
   
   Next Row
   
#If DATACON = 1 Then       'Access
   
   If Not lDbAnoAnt Is Nothing Then
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
   Dim NotValidRut As Boolean
   Dim EditEnable As Boolean
   
   Grid.FlxGrid.Redraw = False
   
   If ItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), ItemData(Cb_Mes), 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
      If ItemData(Cb_TipoDoc) > 0 Then
         Where = Where & " AND Documento.TipoDoc = " & ItemData(Cb_TipoDoc)
      End If
      
      If ItemData(Cb_Estado) > 0 Then
         Where = Where & " AND Documento.Estado = " & ItemData(Cb_Estado)
      End If
      
      If Val(Tx_NumDoc) <> 0 Then
         Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
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
   End If
   
   If Row > 0 Then
      Where = Where & " AND IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC))
   End If

   If ItemData(Cb_Sucursal) > 0 Then
      Where = Where & " AND Documento.IdSucursal = " & ItemData(Cb_Sucursal)
   End If

   Q1 = "SELECT IdDoc, Documento.TipoDoc, NumDoc, NumDocHasta, DTE, Documento.IdEntidad, Documento.RutEntidad, Documento.MovEdited, "
   Q1 = Q1 & " Documento.NombreEntidad, Entidades.Rut, Entidades.Nombre, Entidades.NotValidRut, FEmision, FVenc, FEmisionOri, Exento, Documento.IdCompCent, Documento.IdCompPago,"
   Q1 = Q1 & " Afecto, IVA, OtroImp, PorcentRetencion, TipoRetencion, Total, Descrip, Documento.Estado, Documento.ValRet3Porc, "
   Q1 = Q1 & " IdCuentaExento, Usuarios.Usuario, Cuentas1.Codigo as CodCtaEx, Cuentas1.Descripcion as DescCtaEx, "
   Q1 = Q1 & " IdCuentaAfecto, Cuentas2.Codigo as CodCtaAf, Cuentas2.Descripcion as DescCtaAf, IdCuentaOtroImp, Documento.IdCuentaRet3Porc, IdCuentaTotal, "
   Q1 = Q1 & " Documento.IdSucursal, Sucursales.Descripcion as DescSucursal "
   Q1 = Q1 & " FROM ((((( Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON Documento.IdUsuario = Usuarios.IdUsuario )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas1") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaAfecto = Cuentas2.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas2") & " )"
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Sucursales", True, True) & " )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & Where
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   
   If Row = 0 Then
      Q1 = Q1 & SqlPaging(gDbType, lClsPaging.CurReg - 1, gPageNumReg)
   End If
   Set Rs = OpenRs(DbMain, Q1)

   If Row <= 0 Then
      Grid.rows = Grid.FixedRows
      i = Grid.FixedRows
   Else
      i = Row
   End If
   
   
   '2784017
   NumReg = Rs.RecordCount
   ' fin 2784017
   
   Do While Rs.EOF = False
      
      If Row <= 0 Then
         Grid.rows = Grid.rows + 1
      End If
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1 + (lClsPaging.CurReg - 1)
      If Row = 0 Then
         lClsPaging.NumReg = vFmt(Grid.TextMatrix(i, C_NUMLIN)) - (lClsPaging.CurReg - 1)
      End If
      
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(lTipoLib, vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_IDTIPODOC) = vFld(Rs("TipoDoc"))
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      If vFld(Rs("DTE")) <> 0 Then
         Grid.TextMatrix(i, C_DTE) = "x"
      End If
      Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      
      If vFld(Rs("Estado")) <> ED_ANULADO Then
      
         If vFld(Rs("IdEntidad")) = 0 Then

            If vFld(Rs("IdEntidad")) = 0 Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RutEntidad")))
               Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("NombreEntidad"), True)
            End If
            
         Else
            Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)
            Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
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
      
      Grid.TextMatrix(i, C_HONORSINRET) = Format(vFld(Rs("Exento")), NUMFMT)
      Grid.TextMatrix(i, C_BRUTO) = Format(vFld(Rs("Afecto")), NUMFMT)
      Grid.TextMatrix(i, C_IMPTO) = Format(vFld(Rs("OtroImp")), NUMFMT)
      Grid.TextMatrix(i, C_RET3PORC) = Format(vFld(Rs("ValRet3Porc")), NUMFMT)
      Grid.TextMatrix(i, C_NETO) = Format(vFld(Rs("Total")), NUMFMT)
                  
      Grid.TextMatrix(i, C_IDPIMPTO) = IMPRET_NAC
      If vFld(Rs("PorcentRetencion")) > 0 Then
         Grid.TextMatrix(i, C_IDPIMPTO) = vFld(Rs("PorcentRetencion"))
      End If
      If Val(Grid.TextMatrix(i, C_IDPIMPTO)) <> IMPRET_OTRO Then
         Grid.TextMatrix(i, C_PIMPTO) = gImpRet(Val(Grid.TextMatrix(i, C_IDPIMPTO))) * 100 & "%"
      Else
         Grid.TextMatrix(i, C_PIMPTO) = "Otro"
      End If
      
      Grid.TextMatrix(i, C_IDTIPORETEN) = TR_HONORARIOS
      If vFld(Rs("TipoRetencion")) > 0 And vFld(Rs("TipoRetencion")) <= UBound(gTipoRetencion) Then
         Grid.TextMatrix(i, C_IDTIPORETEN) = vFld(Rs("TipoRetencion"))
      End If
      
      Grid.TextMatrix(i, C_TIPORETEN) = gTipoRetencion(Val(Grid.TextMatrix(i, C_IDTIPORETEN)))
                        
      If vFld(Rs("MovEdited")) = 0 Then
         If vFld(Rs("Exento")) > 0 Then
            Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuentaExento"))
            Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("CodCtaEx")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("DescCtaEx"), True)
         Else
            Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuentaAfecto"))
            Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("CodCtaAf")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("DescCtaAf"), True)
         End If
      Else
         Grid.TextMatrix(i, C_IDCUENTA) = ""
         Grid.TextMatrix(i, C_CODCUENTA) = ""
         Grid.TextMatrix(i, C_CUENTA) = ""
      End If
      
      Grid.TextMatrix(i, C_IMP_IDCUENTA) = vFld(Rs("IdCuentaOtroImp"))
      Grid.TextMatrix(i, C_RET3PORC_IDCUENTA) = vFld(Rs("IdCuentaRet3Porc"))
      
      Grid.TextMatrix(i, C_NETO_IDCUENTA) = vFld(Rs("IdCuentaTotal"))
      
      If vFld(Rs("FEmisionOri")) > 0 Then
         Grid.TextMatrix(i, C_FECHAEMIORI) = Format(vFld(Rs("FEmisionOri")), SDATEFMT)
         Grid.TextMatrix(i, C_LNGFECHAEMIORI) = vFld(Rs("FEmisionOri"))
      End If
      
      'parche para docs ya ingresados, que tienen fecha emisión cero (ahora es obligatoria) y no están en estado pendiente (los de estado pendiente el usuario puede ingresarle la fecha)
      'le asignamos FEmision (que corresponde al día de recepción)
      If Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI)) = 0 And vFld(Rs("Estado")) <> ED_PENDIENTE Then
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
         Grid.TextMatrix(i, C_LNGFECHAVENC) = CLng(DateAdd("d", 30, vFld(Rs("FEmision"))))
         Grid.TextMatrix(i, C_FECHAVENC) = Format(vFmt(Grid.TextMatrix(i, C_LNGFECHAVENC)), SDATEFMT)
         Call FGrModRow(Grid, i, FGR_U, C_IDDOC, C_UPDATE)
      End If
     
      Grid.TextMatrix(i, C_DETALLE) = TX_DETALLE

      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(i, C_IDSUCURSAL) = vFld(Rs("IdSucursal"), True)
      Grid.TextMatrix(i, C_SUCURSAL) = vFld(Rs("DescSucursal"), True)
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"), True)
      Grid.TextMatrix(i, C_ESTADO) = gEstadoDoc(vFld(Rs("Estado")))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.TextMatrix(i, C_MOVEDITED) = IIf(vFld(Rs("MovEdited")) <> 0, -1, 0)
      Grid.TextMatrix(i, C_IDCOMPCENT) = vFld(Rs("IdCompCent"))
      Grid.TextMatrix(i, C_IDCOMPPAGO) = vFld(Rs("IdCompPago"))
      
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
      
      If vFld(Rs("MovEdited")) <> 0 Then
         Call FGrSetRowStyle(Grid, i, "FC", COLOR_AZULOSCURO)
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
   
   Call UnLockAction(DbMain, lTipoLib, , , , False)

   EditEnable = LockAction(DbMain, lTipoLib, ItemData(Cb_Mes))

   If EditEnable = False Then    'alguien más lo está editando, no podemos editarlo (esto se hace sólo una vez ya que en lOper = O_EDIT no se puede cambiar el mes)
      MsgBox1 "El " & gTipoLib(lTipoLib) & " del mes de " & gNomMes(ItemData(Cb_Mes)) & " se está editando en el equipo '" & IsLockedAction(DbMain, lTipoLib, ItemData(Cb_Mes)) & "'. Sólo se abrirá de lectura.", vbInformation
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
   Dim FldArray(3) As AdvTbAddNew_t
   Dim Descrip As String

   lIdDoc = 0
   
   Lin = Grid.FixedRows
   For i = Grid.FixedRows To Grid.rows - 1
        
            
            
      If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista de mov.
         Exit For
      End If
      
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
'         Set Rs = DbMain.OpenRecordset("Documento")
'         Rs.AddNew
'
'         IdDoc = vFld(Rs("IdDoc"))
'         Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'         Rs.Fields("FechaCreacion") = CLng(Int(Now))
'         Rs.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'         Rs.Update
'         Rs.Close
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
         
         IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
         
         Grid.TextMatrix(i, C_IDDOC) = IdDoc
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
         
         If lIdDoc = 0 Then   'selecciona el primero que se insertó
            lIdDoc = IdDoc
         End If
         
      End If
      
        '626924 opcion solo para cliente que no existe detalle de documentos
     If Grid.TextMatrix(i, C_UPDATE) = "" Then
     Grid.TextMatrix(i, C_UPDATE) = FGR_U
     Grid.TextMatrix(i, C_MOVEDITED) = 0
     End If
   '626924
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM Documento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "Documento", Q1)
         
'         Q1 = "DELETE FROM MovDocumento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "MovDocumento", Q1)
         
         '3133008
       
        Dim PathDbAnoAnt As String
        Dim ConnStr As String

        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If

   If gDbType = SQL_ACCESS Then
        PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")

        If ExistFile(PathDbAnoAnt) Then
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)

        Else
         ' Exit Sub
        End If
    End If

    Q1 = ""
    Q1 = "Update Documento Set FExported = null WHERE NumDoc = '" & Val(Grid.TextMatrix(i, C_NUMDOC)) & "'"
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
    Q1 = Q1 & " And TipoLib = " & lTipoLib
    Q1 = Q1 & " And TipoDoc = " & FindTipoDoc(lTipoLib, Grid.TextMatrix(i, C_TIPODOC))
    Q1 = Q1 & " And FEmisionOri = " & Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI))
    Q1 = Q1 & " And Total = " & Abs(vFmt(Grid.TextMatrix(i, C_NETO)))

    Call ExecSQL(DbAnoAnt, Q1)

    If gDbType = SQL_ACCESS Then
    Call CloseDb(DbAnoAnt)
    End If

    '3133008
                 
         
'         Q1 = "DELETE FROM LibroCaja WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "LibroCaja", Q1)

         
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then    'Update
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & lTipoLib
         Q1 = Q1 & ", TipoDoc = " & Val(Grid.TextMatrix(i, C_IDTIPODOC))
         Q1 = Q1 & ", NumDoc = '" & Trim(Grid.TextMatrix(i, C_NUMDOC)) & "'"
         Q1 = Q1 & ", DTE = " & IIf(Trim(Grid.TextMatrix(i, C_DTE)) <> "", -1, 0)
         Q1 = Q1 & ", IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
         Q1 = Q1 & ", EntRelacionada = " & Abs(GetEntRelacionada(Val(Grid.TextMatrix(i, C_IDENTIDAD))))
         Q1 = Q1 & ", IdSucursal = " & vFmt(Grid.TextMatrix(i, C_IDSUCURSAL))
         
         Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR
         
         'por si acaso, ponemos la clasificación de la entidad
         Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_PROVEEDOR & " = 1 WHERE IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD)))
         
         If Val(Grid.TextMatrix(i, C_IDENTIDAD)) = 0 Then
            Q1 = Q1 & ", RutEntidad ='" & vFmtCID(Grid.TextMatrix(i, C_RUT)) & "'"
            Q1 = Q1 & ", NombreEntidad = '" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
         End If
         
         Q1 = Q1 & ", FEmision = " & CLng(DateSerial(Val(Cb_Ano), ItemData(Cb_Mes), Val(Grid.TextMatrix(i, C_FECHA))))
         Q1 = Q1 & ", Exento = " & vFmt(Grid.TextMatrix(i, C_HONORSINRET))
         Q1 = Q1 & ", Afecto = " & vFmt(Grid.TextMatrix(i, C_BRUTO))
         If vFmt(Grid.TextMatrix(i, C_HONORSINRET)) > 0 Then
            Q1 = Q1 & ", IdCuentaExento = " & vFmt(Grid.TextMatrix(i, C_IDCUENTA))
            Q1 = Q1 & ", IdCuentaAfecto = 0"
         Else
            Q1 = Q1 & ", IdCuentaExento = 0"
            Q1 = Q1 & ", IdCuentaAfecto = " & vFmt(Grid.TextMatrix(i, C_IDCUENTA))
         End If
         Q1 = Q1 & ", PorcentRetencion = " & vFmt(Grid.TextMatrix(i, C_IDPIMPTO))
         Q1 = Q1 & ", TipoRetencion = " & vFmt(Grid.TextMatrix(i, C_IDTIPORETEN))
         Q1 = Q1 & ", OtroImp = " & vFmt(Grid.TextMatrix(i, C_IMPTO))
         Q1 = Q1 & ", IdCuentaOtroImp = " & vFmt(Grid.TextMatrix(i, C_IMP_IDCUENTA))
         
         Q1 = Q1 & ", ValRet3Porc = " & vFmt(Grid.TextMatrix(i, C_RET3PORC))
         
         If vFmt(Grid.TextMatrix(i, C_RET3PORC)) <> 0 Then
            Q1 = Q1 & ", IdCuentaRet3Porc = " & vFmt(Grid.TextMatrix(i, C_RET3PORC_IDCUENTA))
         Else
            Q1 = Q1 & ", IdCuentaRet3Porc = 0 "
         End If
         
         Q1 = Q1 & ", Total = " & vFmt(Grid.TextMatrix(i, C_NETO))
         Q1 = Q1 & ", IdCuentaTotal = " & vFmt(Grid.TextMatrix(i, C_NETO_IDCUENTA))
         Descrip = Left(ParaSQL(RemoveNoPrtChars(Grid.TextMatrix(i, C_DESCRIP), True)), 100)
         Q1 = Q1 & ", Descrip = '" & Descrip & "'"
         Q1 = Q1 & ", FEmisionOri = " & Val(Grid.TextMatrix(i, C_LNGFECHAEMIORI))
         Q1 = Q1 & ", FVenc = " & Val(Grid.TextMatrix(i, C_LNGFECHAVENC))
         Q1 = Q1 & ", Estado = " & Val(Grid.TextMatrix(i, C_IDESTADO))
         Q1 = Q1 & ", SaldoDoc = NULL "
         Q1 = Q1 & " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         If Val(Grid.TextMatrix(i, C_MOVEDITED)) = 0 Then
         
'            Q1 = "DELETE FROM MovDocumento WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
'            Call ExecSQL(DbMain, Q1)
            Q1 = " WHERE IdDoc = " & Val(Grid.TextMatrix(i, C_IDDOC))
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call DeleteSQL(DbMain, "MovDocumento", Q1)
            
            Call GenMovDocumento(i)
            
         End If
        
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> FGR_D Then  'Delete
         Lin = Lin + 1
      End If
      
      '2858854
      'Grid.TextMatrix(i, C_UPDATE) = ""     'lo limpiampos dado que esta función es invocada en Bt_DetDoc
       'fin 2858854
   Next i

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

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   Grid.Row = Grid.MouseRow
   Grid.Col = Grid.MouseCol
   
   If lOper <> O_EDIT And lOper <> O_NEW Then
      Exit Sub
   End If

   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Or Not ValidaEstadoEdit(Grid.Row) Then
      Exit Sub
   End If
   
   If Button = vbRightButton Then
      If (Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC) Then
         Call PopupMenu(M_TipoDoc)
      ElseIf Grid.Col = C_CODCUENTA Then
         Call Bt_Cuentas_Click
      End If
   End If
   
End Sub
Private Sub Grid_SelChange()
   Dim EdType As FlexEdGrid2.FEG2_EdType
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   'Feña
   If IRutEmp And Grid.TextMatrix(Grid.Row, C_RUT) <> "" Then
    'MsgBox1 "El Rut No puede ser el mismo que el de la Empresa que está trabajando.", vbInformation
    MsgBox1 "El Rut de la boleta No puede ser el mismo de la empresa declarante.", vbInformation
    Grid.TextMatrix(Grid.Row, C_RUT) = ""
   End If
   'Fin Feña
   
   If Grid.Col = C_FECHA And Grid.TextMatrix(Grid.Row, Grid.Col) = "" And (Grid.Row = Grid.FixedRows Or Grid.TextMatrix(Grid.Row - 1, Grid.Col) <> "") Then
      Call Grid_BeforeEdit(Grid.Row, Grid.Col, EdType)
   Else
      Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)
   End If

End Sub
Private Sub M_ItCuenta_Click(Index As Integer)
   Dim CodCta As String
   Dim DescCta As String
   Dim NombCta As String
   Dim IdCuenta As Long
   Dim Frm As FrmPlanCuentas
   Dim Row As Integer
   Dim Col As Integer

   Row = Grid.Row
   Col = Grid.Col

   If Col = C_CODCUENTA Then
      
      If M_ItCuenta(Index).Caption <> MITEM_OTRA Then
         Grid.TextMatrix(Row, C_IDCUENTA) = lCuentas(Index)
         Grid.TextMatrix(Row, C_CODCUENTA) = Left(M_ItCuenta(Index).Caption, InStr(M_ItCuenta(Index).Caption, " [") - 1)
         Grid.TextMatrix(Row, C_CUENTA) = Mid(M_ItCuenta(Index).Caption, InStr(M_ItCuenta(Index).Caption, "] ") + 2)
      
      Else
         Set Frm = New FrmPlanCuentas
         If Frm.FSelEdit(IdCuenta, CodCta, DescCta, NombCta, False) = vbOK Then
            
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
            Grid.TextMatrix(Row, C_CODCUENTA) = FmtCodCuenta(CodCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta
            
            Set Frm = Nothing
         Else
            Set Frm = Nothing
            Exit Sub
         End If
      
      End If
      
   Else
      Exit Sub
   
   End If
   
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)

End Sub

Private Function IsValidLine(ByVal Row As Integer, Msg As String) As Boolean
   Dim ValLine As Boolean
   Dim FirstDay As Long
   Dim LastDay As Long
   
   ValLine = True
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_FECHA)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar día de recepción."
      Exit Function
   End If
   
   ValLine = (ValLine And Trim(Grid.TextMatrix(Row, C_TIPODOC)) <> "")
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el Tipo de Documento."
      Exit Function
   End If
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_NUMDOC)) <> 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el número de documento."
      Exit Function
   End If
      
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) > 0)
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar la fecha de emisión."
      Exit Function
   End If
      
   Call FirstLastMonthDay(DateSerial(lAno, ItemData(Cb_Mes), 1), FirstDay, LastDay)
   
   ValLine = (ValLine And Val(Grid.TextMatrix(Row, C_LNGFECHAEMIORI)) <= LastDay)
   If ValLine = False And Msg = "" Then
      Msg = "La fecha de emisión es posterior al último día de este mes."
      Exit Function
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

      
   ValLine = (ValLine And (Trim(Grid.TextMatrix(Row, C_RUT)) <> "" Or (Trim(Grid.TextMatrix(Row, C_RUT)) = "" And Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_ANULADO)))
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el RUT."
      Exit Function
   End If
   
   ValLine = (ValLine And (Trim(Grid.TextMatrix(Row, C_TIPORETEN)) <> ""))
   If ValLine = False And Msg = "" Then
      Msg = "Falta ingresar el tipo de retención."
      Exit Function
   End If
   
   'validamos que el impuesto no sobrepase la base
   If vFmt(Grid.TextMatrix(Row, C_BRUTO)) <> 0 Then
      ValLine = (ValLine And vFmt(Grid.TextMatrix(Row, C_IMPTO)) <= vFmt(Grid.TextMatrix(Row, C_BRUTO)))
      If ValLine = False And Msg = "" Then
         Msg = "El Impuesto es mayor que el valor Bruto."
         Exit Function
      End If
   End If
   
'   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
'      ValLine = (ValLine And (Val(Grid.TextMatrix(Row, C_BRUTO)) > 0 Or (Val(Grid.TextMatrix(Row, C_HONORSINRET)) > 0 And Val(Grid.TextMatrix(Row, C_IMPTO)) > 0)))
'      If ValLine = False And Msg = "" Then
'         Msg = "Valor cero."
'         Exit Function
'      End If
'   End If

   If gEmpresa.Ano >= 2017 Then
      ValLine = ValLine And ((gEmpresa.Franq14Ter And Trim(Grid.TextMatrix(Row, C_DESCRIP)) <> "") Or Not gEmpresa.Franq14Ter)
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la descripción del documento." & vbCrLf & vbCrLf & "Ésta es obligatoria para el Libro de Caja, en empresas acogidas a 14TER."
         Exit Function
      End If
   End If

   If Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0 Then  'estas validaciones no tienen sentido si el documento ha sido editado en detalle,porque las cuentas se asignan en el detalle
   
      ValLine = (ValLine And (IIf((vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) <> 0 Or vFmt(Grid.TextMatrix(Row, C_BRUTO)) <> 0), Val(Grid.TextMatrix(Row, C_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta de honorarios sin retención."
         Exit Function
      End If
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_IMPTO)) <> 0, Val(Grid.TextMatrix(Row, C_IMP_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta definir la cuenta de Impuesto en la configuración inicial."
         Exit Function
      End If
      
      ValLine = (ValLine And (IIf(vFmt(Grid.TextMatrix(Row, C_NETO)) <> 0, vFmt(Grid.TextMatrix(Row, C_NETO_IDCUENTA)) <> 0, True)))
      If ValLine = False And Msg = "" Then
         Msg = "Falta ingresar la cuenta del Neto " & Grid.TextMatrix(Row, C_TIPORETEN) & " en la configuración inicial de la empresa."
         Exit Function
      End If
   
   End If
   
'   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
'      ValLine = (ValLine And vFmt(Grid.TextMatrix(Row, C_NETO)) > 0)
'      If ValLine = False And Msg = "" Then
'         Msg = "El valor neto es cero."
'         Exit Function
'      End If
'   End If

   


   IsValidLine = ValLine

End Function

Private Function LineaEnBlanco(ByVal Row As Integer) As Boolean
   Dim i As Integer
   
   
   LineaEnBlanco = True
   
   For i = C_NUMDOC To Grid.Cols - 1
   
      If Grid.TextMatrix(Row, i) <> "" And i <> C_ESTADO And i <> C_IDESTADO And i <> C_UPDATE And i <> C_FECHAEMIORI And i <> C_LNGFECHAEMIORI Then
         LineaEnBlanco = False
      End If
      
   Next i
     
   If LineaEnBlanco Then
      Grid.TextMatrix(Row, C_FECHA) = ""
      Grid.TextMatrix(Row, C_TIPODOC) = ""
   End If
      
End Function

Private Sub CalcTot()
   Dim Tot(NCOLS) As Double
   Dim i As Integer, j As Integer
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      For j = C_HONORSINRET To C_NETO
         If j <> C_PIMPTO Then
            Tot(j) = Tot(j) + vFmt(Grid.TextMatrix(i, j))
         End If
      Next j
   Next i
   
   GridTot.TextMatrix(0, C_NOMBRE) = "TOTAL"
   
   GridTot.TextMatrix(0, C_HONORSINRET) = Format(Tot(C_HONORSINRET), NUMFMT)
   GridTot.TextMatrix(0, C_BRUTO) = Format(Tot(C_BRUTO), NUMFMT)
   GridTot.TextMatrix(0, C_IMPTO) = Format(Tot(C_IMPTO), NUMFMT)
   GridTot.TextMatrix(0, C_NETO) = Format(Tot(C_NETO), NUMFMT)
   GridTot.TextMatrix(0, C_RET3PORC) = Format(Tot(C_RET3PORC), NUMFMT)
   
End Sub

Private Function ValidaNumDoc(ByVal Row As Integer) As Boolean
   Dim NumDoc As String
   Dim TipoDoc As Integer
   Dim IdEnt As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim DocNotVal As Boolean
   Dim EqDoc As Boolean
   Dim Wh As String, WhEquDoc As String
   Dim TipoDocBOH As Integer
   Dim DTE As Boolean
   Dim IdDoc As Long
  
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   NumDoc = Trim(Grid.TextMatrix(Row, C_NUMDOC))
   TipoDoc = Val(Grid.TextMatrix(Row, C_IDTIPODOC))
   IdEnt = Trim(Grid.TextMatrix(Row, C_IDENTIDAD))
   DTE = (Grid.TextMatrix(Row, C_DTE) <> "")
   
   ValidaNumDoc = False
   
   'veamos si faltan algunos datos
   If NumDoc = "" Or TipoDoc = 0 Then
      ValidaNumDoc = True
      Exit Function
   End If
   
   If IdEnt = 0 Then
      ValidaNumDoc = True
      Exit Function
   End If
   
   'primero vemos si está en la grilla
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      If i <> Row And Grid.RowHeight(i) > 0 Then
                  
         EqDoc = (TipoDoc = Val(Grid.TextMatrix(i, C_IDTIPODOC)) And NumDoc = Trim(Grid.TextMatrix(i, C_NUMDOC))) And DTE = (Grid.TextMatrix(i, C_DTE) <> "")
         DocNotVal = (Grid.TextMatrix(i, C_TIPODOC) = "BOH" And EqDoc And IdEnt = Val(Grid.TextMatrix(i, C_IDENTIDAD)))
         DocNotVal = DocNotVal Or (Grid.TextMatrix(i, C_TIPODOC) = "BRT" And EqDoc)
         
         If DocNotVal Then
            MsgBox1 "El documento " & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en este libro.", vbExclamation + vbOKOnly
            Exit Function
         End If
         
      End If
   Next i
      
   'ahora vemos si está en la DB, sólo si es nuevo en otros meses si es ACCESS o en este u otros meses si es SQL Server(pensando en el paginamiento)
   
   TipoDocBOH = FindTipoDoc(LIB_RETEN, "BOH")
        
   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = '" & NumDoc & "' AND Abs(DTE) = " & Abs(CInt(DTE))
   Wh = " WHERE TipoLib = " & LIB_RETEN & WhEquDoc
   Wh = Wh & " AND ((TipoDoc = " & TipoDocBOH & " AND IdEntidad = " & IdEnt & ") OR TipoDoc <> " & TipoDocBOH & ")"
   If gDbType = SQL_ACCESS Then    'en otros meses si es ACCESS
      Wh = Wh & " AND " & SqlMonthLng("FEmision") & " <> " & lMes
   End If
   
   Q1 = "SELECT IdDoc, FEmision FROM Documento " & Wh
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("IdDoc")) <> IdDoc Then
         MsgBox1 "El documento " & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & ".", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
      
   Call CloseRs(Rs)
   
   'ahora vemos si está en el año anterior
   If Not ValidaNumDocAnoAnt(Row) Then
      Exit Function
   End If

   
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
   Dim TipoDocBOH As Integer
   
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
   
   ValidaNumDocAnoAnt = False
   
   'veamos si faltan algunos datos
   If NumDoc = "" Or TipoDoc = 0 Then
      ValidaNumDocAnoAnt = True
      Exit Function
   End If
   
   
   TipoDocBOH = FindTipoDoc(LIB_RETEN, "BOH")
   
      
   WhEquDoc = " AND TipoDoc = " & TipoDoc & " AND NumDoc = '" & NumDoc & "' AND DTE = " & CInt(DTE)
   Wh = " WHERE TipoLib = " & LIB_RETEN & WhEquDoc
'   Wh = Wh & " AND ((TipoDoc = " & TipoDocBOH & " AND IdEntidad = " & IdEnt & ") OR TipoDoc <> " & TipoDocBOH & ")"      'FCA 27 may 2019
   Wh = Wh & " AND ((TipoDoc = " & TipoDocBOH & " AND Rut = '" & RutEnt & "') OR TipoDoc <> " & TipoDocBOH & ")"
   Wh = Wh & " AND Month(FEmision) <> " & lMes
   
'   Q1 = "SELECT IdDoc, FEmision FROM Documento " & Wh
   Q1 = "SELECT IdDoc, FEmision FROM Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad " & Wh
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano - 1
   
   If gEmprSeparadas Then
#If DATACON = 1 Then       'Access
      Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
      Set Rs = OpenRs(lDbAnoAnt, Q1)
#End If
   Else
      Set Rs = OpenRs(DbMain, Q1)
   End If
   
   If Rs.EOF = False Then
   
      MsgBox1 "El documento " & Grid.TextMatrix(Row, C_TIPODOC) & "-" & NumDoc & " de " & Grid.TextMatrix(Row, C_NOMBRE) & " ya ha sido ingresado en el libro del mes de " & gNomMes(month(vFld(Rs("FEmision")))) & " " & Year(vFld(Rs("FEmision"))) & ".", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Function
            
   End If
      
   Call CloseRs(Rs)
      
   ValidaNumDocAnoAnt = True

End Function

Private Sub GenMovDocumento(ByVal Row As Integer)
   Dim Q1 As String
   Dim QBase As String
   Dim i As Integer
   Dim Glosa As String, GlosaImp As String
   Dim TipoValLib As Integer
            
   QBase = "INSERT INTO MovDocumento"
   QBase = QBase & "(IdDoc, IdEmpresa, Ano, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg) "
   QBase = QBase & " VALUES(" & Grid.TextMatrix(Row, C_IDDOC) & "," & gEmpresa.id & "," & gEmpresa.Ano & ","
      
   Glosa = ParaSQL(Left(Trim(Grid.TextMatrix(Row, C_DESCRIP)), 50))
   GlosaImp = "[" & ParaSQL(Trim(Grid.TextMatrix(Row, C_TIPODOC))) & " " & ParaSQL(Trim(Grid.TextMatrix(Row, C_NUMDOC))) & "] " & ParaSQL(Trim(Grid.TextMatrix(Row, C_RUT))) & " " & ParaSQL(Trim(Grid.TextMatrix(Row, C_NOMBRE)))
   GlosaImp = Left(GlosaImp, 50)
   
   
   i = 1
   
   '2868501
   Dim vCcosto As Integer
   Dim vAreaNego As Integer
   'fin 2868501
                  
   'Bruto / Honorarios Sin Retención
   If vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) > 0 Or vFmt(Grid.TextMatrix(Row, C_BRUTO)) > 0 Then
      Q1 = QBase & i & ","                                              'Orden
      Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) & ","                  'IdCuenta
         
      If vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) > 0 Then
         Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) & ","      'Debe
         TipoValLib = LIBRETEN_HONORSINRET
      Else
         Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_BRUTO)) & ","            'Debe
         TipoValLib = LIBRETEN_BRUTO
      End If
      Q1 = Q1 & "0" & ","                                               'Haber
      Q1 = Q1 & "'" & Glosa & "',"                                      'Glosa
      Q1 = Q1 & TipoValLib & ","                                        'IdTipoValLib
      
       '2868501
'         If Grid.TextMatrix(Row, C_IDCUENTA) <> "" Then
                If CuentaCentro(Grid.TextMatrix(Row, C_IDCUENTA)) = True Then
                 vCcosto = CentroCosto()
                Else
                 vCcosto = 0
                End If
                
                If CuentaArea(Grid.TextMatrix(Row, C_IDCUENTA)) = True Then
                 vAreaNego = AreaNegocio()
                Else
                 vAreaNego = 0
                End If
'           Else
'           vCcosto = 0
'           vCcosto = 0
'          End If
          
            Q1 = Q1 & "0," & vCcosto & "," & vAreaNego & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
        
         
         'Q1 = Q1 & "0,0,0" & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
         
         'FIN 2868501
      
      
         
      Call ExecSQL(DbMain, Q1)
         
      i = i + 1
   End If
   
   
   
   'Impuesto
   If vFmt(Grid.TextMatrix(Row, C_IMPTO)) > 0 Then
      Q1 = QBase & i & ","                                           'Orden
      Q1 = Q1 & Grid.TextMatrix(Row, C_IMP_IDCUENTA) & ","           'IdCuenta
      Q1 = Q1 & "0" & ","                                            'Debe
      Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_IMPTO)) & ","            'Haber
      'Q1 = Q1 & "'" & Glosa & "',"                                   'Glosa
      Q1 = Q1 & "'" & GlosaImp & "',"                                   'Glosa
      Q1 = Q1 & LIBRETEN_IMPUESTO & ","                              'IdTipoValLib
      
      
       '2868501
         
         If CuentaCentro(Grid.TextMatrix(Row, C_IMP_IDCUENTA)) = True Then
            vCcosto = CentroCosto()
           Else
            vCcosto = 0
           End If
           
           If CuentaArea(Grid.TextMatrix(Row, C_IMP_IDCUENTA)) = True Then
            vAreaNego = AreaNegocio()
           Else
            vAreaNego = 0
           End If
                      
            Q1 = Q1 & "0," & vCcosto & "," & vAreaNego & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
         
         'Q1 = Q1 & "0,0,0" & ")"                                        'EsTotalDoc, IdCCosto, IdAreaNeg
         
         'FIN 2868501
      
      
   
      Call ExecSQL(DbMain, Q1)
      
      i = i + 1
   End If
   
   'Retención 3%
   If vFmt(Grid.TextMatrix(Row, C_RET3PORC)) > 0 Then
      Q1 = QBase & i & ","                                           'Orden
      Q1 = Q1 & Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA) & ","           'IdCuenta
      Q1 = Q1 & "0" & ","                                            'Debe
      Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_RET3PORC)) & ","            'Haber
      Q1 = Q1 & "'" & Glosa & "',"                                   'Glosa
      Q1 = Q1 & LIBRETEN_RET3PORC & ","                              'IdTipoValLib
      
      '2868501
         
           If CuentaCentro(Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA)) = True Then
            vCcosto = CentroCosto()
           Else
            vCcosto = 0
           End If
           
           If CuentaArea(Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA)) = True Then
            vAreaNego = AreaNegocio()
           Else
            vAreaNego = 0
           End If
                      
            Q1 = Q1 & "0," & vCcosto & "," & vAreaNego & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
                   
         
         'Q1 = Q1 & "0,0,0" & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
         
         'FIN 2868501
      
      Call ExecSQL(DbMain, Q1)
      
      i = i + 1
   End If
   'Neto
   If vFmt(Grid.TextMatrix(Row, C_NETO)) > 0 Then
      Q1 = QBase & i & ","                                           'Orden
      Q1 = Q1 & Grid.TextMatrix(Row, C_NETO_IDCUENTA) & ","          'IdCuenta
      Q1 = Q1 & "0" & ","                                            'Debe
      Q1 = Q1 & vFmt(Grid.TextMatrix(Row, C_NETO)) & ","             'Haber
      Q1 = Q1 & "'" & Glosa & "',"                                   'Glosa
      Q1 = Q1 & LIBRETEN_NETO & ","                                  'IdTipoValLib
      
      
      '2868501
         
          If CuentaCentro(Grid.TextMatrix(Row, C_NETO_IDCUENTA)) = True Then
            vCcosto = CentroCosto()
           Else
            vCcosto = 0
           End If
           
           If CuentaArea(Grid.TextMatrix(Row, C_NETO_IDCUENTA)) = True Then
            vAreaNego = AreaNegocio()
           Else
            vAreaNego = 0
           End If
           
           Q1 = Q1 & "1," & vCcosto & "," & vAreaNego & ")"                            'EsTotalDoc, IdCCosto, IdAreaNeg
                
         
         'FIN 2868501
      
      'Q1 = Q1 & "1,0,0" & ")"                                        'EsTotalDoc, IdCCosto, IdAreaNeg
      
      Call ExecSQL(DbMain, Q1)
   End If
  
  
End Sub

Private Sub Tx_NumDoc_Change()
   Bt_List.Enabled = True
   Bt_Centralizar.Enabled = False
   Bt_Sel.Enabled = False

End Sub

Private Sub Tx_NumDoc_KeyPress(KeyAscii As Integer)
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
   Else
      Call KeyCID(KeyAscii)
   End If
   
End Sub

Private Function EsCuentaBasica(ByVal IdCuenta, ByVal Row As Integer) As Boolean
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
  
      If vFmt(Grid.TextMatrix(Row, C_HONORSINRET)) > 0 Then
         Q1 = Q1 & LIBRETEN_HONORSINRET
      Else
         Q1 = Q1 & LIBRETEN_BRUTO
      End If
              
      Set Rs = OpenRs(DbMain, Q1)
            
      If Rs.EOF = False Then
         EsCuentaBasica = True
      End If
           
      Call CloseRs(Rs)
                          
   End If
   
End Function

Private Sub LoadDefCuentas()
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
         
            Case LIBRETEN_HONORSINRET
            
               If lCtaHonSinRet.id = 0 Then
                  lCtaHonSinRet.id = vFld(Rs("IdCuenta"))
                  lCtaHonSinRet.Codigo = vFld(Rs("Codigo"))
                  lCtaHonSinRet.Descripcion = FCase(vFld(Rs("Descripcion"), True))
               End If
               
            Case LIBRETEN_BRUTO
            
               If lCtaBruto.id = 0 Then
                  lCtaBruto.id = vFld(Rs("IdCuenta"))
                  lCtaBruto.Codigo = vFld(Rs("Codigo"))
                  lCtaBruto.Descripcion = FCase(vFld(Rs("Descripcion"), True))
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

   If Tx_Rut = "" Then
      Cb_Entidad.ListIndex = 0  'en blanco
      Exit Sub
   End If

   If Not MsgValidCID(Tx_Rut) Then
      Tx_Rut.SetFocus
      Exit Sub

   End If
      
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut) & "'"
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
   
   MousePointer = vbHourglass
      
   Tx_Rut = FmtCID(vFmtCID(Tx_Rut))
   MousePointer = vbDefault
   
End Sub
Private Sub cb_Nombre_Click()
   
   If lcbNombre.ListIndex >= 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT))
   End If
   
   Bt_List.Enabled = True
   Bt_Sel.Enabled = False
   Bt_Centralizar.Enabled = False

End Sub
Private Sub Cb_Entidad_Click()
      
   Cb_Nombre.Clear
   If ItemData(Cb_Entidad) >= 0 Then
      Call SelCbEntidad(ItemData(Cb_Entidad))
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
   
   lOrdenGr(C_FECHA) = "Documento.FEmision" & StrOrder
   lOrdenGr(C_TIPODOC) = "Documento.TipoDoc, Documento.NumDoc"
   lOrdenGr(C_NUMDOC) = "Documento.NumDoc, Documento.TipoDoc"
   lOrdenGr(C_RUT) = "Entidades.Rut" & StrOrder
   lOrdenGr(C_NOMBRE) = "Entidades.Nombre" & StrOrder
   lOrdenGr(C_DESCRIP) = "Documento.Descrip" & StrOrder
   lOrdenGr(C_SUCURSAL) = "Sucursales.Descripcion" & StrOrder
   lOrdenGr(C_HONORSINRET) = "Documento.Exento" & StrOrder
   lOrdenGr(C_BRUTO) = "Documento.Afecto" & StrOrder
   lOrdenGr(C_CODCUENTA) = "Cuentas1.Codigo, Cuentas2.Codigo" & StrOrder
   lOrdenGr(C_CUENTA) = "Cuentas1.Descripcion, Cuentas2.Descripcion" & StrOrder
   lOrdenGr(C_IMPTO) = "Documento.OtroImp" & StrOrder
   lOrdenGr(C_NETO) = "Documento.Total" & StrOrder
   lOrdenGr(C_FECHAEMIORI) = "Documento.FEmisionOri" & StrOrder
   lOrdenGr(C_FECHAVENC) = "Documento.FVenc" & StrOrder
   lOrdenGr(C_ESTADO) = "Documento.Estado" & StrOrder
   lOrdenGr(C_USUARIO) = "Usuarios.Usuario" & StrOrder

   If lOper = O_VIEWLIBLEGAL Then
      lOrdenSel = C_FECHAEMIORI
   Else
      lOrdenSel = C_FECHA
   End If
   
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   Dim i As Integer
   Dim Upd As Boolean
   
   If lOrdenGr(Col) = "" Then
      Exit Sub
   End If
   
   If GrabarParaContinuar("cambiar el ordenamiento de las columnas") = False Then
      Exit Sub
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

Private Sub Bt_SelEnt_Click_Old()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      Exit Sub
   End If
   
   Set Frm = New FrmEntidades
   If Frm.FSelect(Entidad) = vbOK Then
      Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
      Grid.TextMatrix(Row, C_RUT) = FmtCID(Entidad.Rut)
      Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
      
      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
   End If
   
   Set Frm = Nothing
   
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
   
   TipoEnt = ENT_PROVEEDOR
   
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
         
   Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
   Grid.TextMatrix(Row, C_RUT) = FmtCID(Entidad.Rut, Entidad.NotValidRut = False)
   Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
   
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
      
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

      End If
   
   Else
   
      If Bt_ExitNewDoc.visible = True And Not ChkPriv(PRV_ING_DOCS) Then
         Bt_ExitNewDoc.Enabled = False
         lIngDocsEnabled = False
      End If
         
      If Bt_Centralizar.visible = True And Not ChkPriv(PRV_ADM_DOCS) Then
         Bt_Centralizar.Enabled = False
         lAdmDocsEnabled = False
      End If
      
   End If
   
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If

End Function

Private Sub Ch_ViewDTE_Click()
   
   If Ch_ViewDTE = 0 Then
      Grid.ColWidth(C_DTE) = 0
      
   Else
      Grid.ColWidth(C_DTE) = 400
         
   End If

   Call SetIniString(gIniFile, "Opciones", "VerLibRetDTE", Abs(Ch_ViewDTE.Value))
   gVarIniFile.VerLibRetDTE = Abs(Ch_ViewDTE.Value)

End Sub
Private Sub Ch_AplicarRet3Porc_Click()

   If Not lInLoad And Ch_AplicarRet3Porc <> 0 Then

      If gCtasBas.IdCtaRet3Porc <= 0 Then
         MsgBox1 "No es posible aplicar la Retención del 3% Préstamo Solidario sin antes definir la configuración de la cuenta básica de Retención 3% Préstamo Solidario.", vbExclamation + vbOKOnly
         Ch_AplicarRet3Porc = 0
      End If

   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerRet3Porc", Abs(Ch_AplicarRet3Porc.Value))
   gVarIniFile.VerRet3Porc = Abs(Ch_AplicarRet3Porc.Value)
   
   Call SetupRet3Porc
   
   Call RecalcRet3Porc

End Sub

Private Function ValidaEstadoEdit(ByVal Row As Integer) As Boolean

   ValidaEstadoEdit = lEditEnabled And (Val(Grid.TextMatrix(Row, C_MOVEDITED)) = 0 And (Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_PENDIENTE Or (lAno < gEmpresa.Ano And Val(Grid.TextMatrix(Row, C_IDESTADO)) = ED_CENTRALIZADO)))

End Function
Private Sub Form_Unload(Cancel As Integer)
   Call UnLockAction(DbMain, lTipoLib, , , , False)
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
      
      lSucursal = ItemData(Cb_Sucursal)
   
   End If
   
   InClick = False

End Sub

Private Sub Ch_ViewSucursal_Click()
   
   If Ch_ViewSucursal = 0 Then
      Grid.ColWidth(C_SUCURSAL) = 0
      
   Else
      Grid.ColWidth(C_SUCURSAL) = 1500
         
   End If

   Call SetIniString(gIniFile, "Opciones", "VerLibRetSucursal", Abs(Ch_ViewSucursal.Value))
   gVarIniFile.VerLibRetSucursal = Abs(Ch_ViewSucursal.Value)

End Sub

Private Sub Bt_TipoDoc_Click()

   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Or Not ValidaEstadoEdit(Grid.Row) Then
      Exit Sub
   End If
   
'   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      If M_ItTipoDoc.Count > 1 Then
         Call PopupMenu(M_TipoDoc, , Grid.FlxGrid.ColPos(Grid.Col) + Grid.Left + 200, Grid.FlxGrid.RowPos(Grid.Row) + Grid.Top + 100)
      End If
'   End If
   
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
Private Sub M_ItTipoDoc_Click(Index As Integer)
   Dim Value As String
   Dim TipoDoc As Integer
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_FECHA) = "" Then
      Exit Sub
   End If
   
   TipoDoc = Cb_TipoDoc.ItemData(Index)
   
   If Grid.Col = C_TIPODOC Or Grid.Col = C_NUMDOC Then
      Value = GetDiminutivoDoc(lTipoLib, TipoDoc)
      Grid.TextMatrix(Grid.Row, C_TIPODOC) = Value
      Grid.TextMatrix(Grid.Row, C_IDTIPODOC) = TipoDoc
           
      Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_IDDOC, C_UPDATE)
   End If

End Sub
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
   
   If Val(Grid.TextMatrix(Row, C_FECHA)) = 0 Then
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
      Set Frm = New FrmDocLib
      Call Frm.FView(IdDoc)
      Set Frm = Nothing
      Exit Sub
   End If
      
   If Grid.TextMatrix(Row, C_UPDATE) <> "" Then
   
      If MsgBox1("Antes de ingresar al detalle de documento se grabarán las modificaciones realizadas en este libro." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If Not valida() Then
         Exit Sub
      End If
      
      Call SaveGrid
      
      IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   
   End If
   
   Set Frm = New FrmDocLib
   Call Frm.FEdit(IdDoc)
   Set Frm = Nothing
   
   Call LoadGrid(Row)
      
End Sub

Private Function ImportFromFile() As Boolean
   Dim fname As String
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ImpEnable As Boolean
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim i As Integer, l As Integer
   Dim j As Integer, p As Long
   Dim NRecErroneos As Long, StrNRecErroneos As String
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
   Dim HonSinRet As Double, Bruto As Double, PImp As Double, Impuesto As Double, Neto As Double
   Dim StrPImp As String, IdPImp As Integer
   Dim Row As Integer, r As Integer
   Dim TipoDoc As String
   Dim CodSuc As String, IdSucursal As Long, Sucursal As String
   Dim Descrip As String
   Dim Dt1 As Long, Dt2 As Long
   Dim IdCtaHonSinRet As Long, CodCtaHonSinRet As String, DescCtaHonSinRet As String
   Dim IdCtaBruto As Long, CodCtaBruto As String, DescCtaBruto As String
   Dim IdCtaImp As Long, IdCtaNeto As Long
   Dim Estado As Integer
   Dim IdTipoReten As Integer, TipoReten As String
   Dim AuxCodCta As String, AuxIdCta As Long, AuxDescCta As String
   Dim NomCta As String, UltNivel As Boolean
   Dim FldArray(5) As AdvTbAddNew_t
   Dim Ret3Porc As Double, ValidaRet3Porc As Boolean
 
   
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
   
   On Error GoTo 0
   
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
   
   ClasifEnt = ENT_PROVEEDOR
   
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
      ValidaRet3Porc = True
      
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
      IdTipoDoc = FindTipoDoc(LIB_RETEN, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Tipo de documento inválido o no corresponde al libro en edición. Valores perimtidos ""BOH"", ""BRT"".")
      End If
         
      
      'DTE
      TxtDTE = Trim(NextField2(Buf, p))
      DTE = IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)
      
      'NumDoc
      NumDoc = Trim(Trim(NextField2(Buf, p)))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "N° de documento inválido.")
      End If
      
      'Fecha emisión
      Aux = Trim(NextField2(Buf, p))
      DtEmi = ValFmtDate(Aux, False)
      If DtEmi = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha emisión inválida.")
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
      
      'Descrpción
      Descrip = Trim(NextField2(Buf, p))
      If Descrip = "NULO" Then
         Estado = ED_ANULADO
      End If
      
      'sucursal
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
      HonSinRet = vFmt(Trim(NextField2(Buf, p)))
      Bruto = vFmt(Trim(NextField2(Buf, p)))
      StrPImp = Trim(NextField2(Buf, p))
      PImp = vFmt(StrPImp)
      Impuesto = vFmt(Trim(NextField2(Buf, p)))
      Ret3Porc = vFmt(Trim(NextField2(Buf, p)))
      Neto = vFmt(Trim(NextField2(Buf, p)))
      
      'Cuentas Contables
      
      If HonSinRet < 0 Or Bruto < 0 Or Neto < 0 Or Impuesto < 0 Or PImp < 0 Or Ret3Porc < 0 Then
         CampoInvalido = CampoInvalido & "," & "Honorarios, Bruto, %Imp, Impuesto, Neto"
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Honorarios, Bruto, %Imp, Impuesto, Retención 3% y/o Neto inválido.")
      End If
      
      If HonSinRet > 0 And Bruto > 0 Then
         CampoInvalido = CampoInvalido & "," & "Honorarios, Bruto"
         Call AddLogImp(lFNameLogImp, fname, l, "No es posible ingresar Honorarios y valor Bruto para un mismo documento.")
      End If
                                                                              
      If (HonSinRet > 0 Or Bruto > 0) And Neto = 0 Then
         CampoInvalido = CampoInvalido & "," & "Neto"
         Call AddLogImp(lFNameLogImp, fname, l, "El valor Neto está en cero o es inválido.")
      End If
                                                                              
      Select Case LCase(StrPImp)
         Case gImpRet(IMPRET_NAC) * 100
            IdPImp = IMPRET_NAC
            StrPImp = StrPImp & "%"
         
         Case gImpRet(IMPRET_EXT) * 100
            IdPImp = IMPRET_EXT
            StrPImp = StrPImp & "%"
         
         Case "otro"
            IdPImp = IMPRET_OTRO
            StrPImp = "Otro"
         
         Case Else
            CampoInvalido = CampoInvalido & "," & "% Imp"
            Call AddLogImp(lFNameLogImp, fname, l, "Porcentaje de impuesto inválido. Valores perimitidos """ & gImpRet(IMPRET_NAC) * 100 & ", """ & gImpRet(IMPRET_EXT) * 100 & ", ""Otro"".")
            
      End Select
      
      TipoReten = Trim(NextField2(Buf, p))
     
      Select Case LCase(TipoReten)
         Case "honorarios"
            IdTipoReten = TR_HONORARIOS
            TipoReten = "Honorarios"
            
         Case "dieta"
            IdTipoReten = TR_DIETA
            TipoReten = "Dieta"
            
         Case "otro"
            IdTipoReten = TR_OTRO
            TipoReten = "Otro"
            
         Case Else
            CampoInvalido = CampoInvalido & "," & "% Imp"
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo retención inválido. Valores permitidos ""Honorarios"", ""Dieta"", ""Otro"".")
            
      End Select
      
      If Ret3Porc > 0 Then
         If gCtasBas.IdCtaRet3Porc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Falta definir la cuenta básica para Retención 3%'.")
         End If
         If IdTipoReten <> TR_HONORARIOS Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo retención no permite 'Retención 3% Préstamo Solidario'.")
         End If
      End If
                                                                             
      'código cuenta
      AuxCodCta = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      If AuxCodCta <> "" Then
         AuxIdCta = GetIdCuenta(NomCta, AuxCodCta, AuxDescCta, UltNivel)
         If AuxIdCta <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido")
         End If
      Else
         AuxIdCta = 0
         AuxDescCta = ""
      End If
      
      NomCta = ""
      
      'Cuenta Contable Default
                  
      If HonSinRet <> 0 Then
         If AuxIdCta > 0 Then
            IdCtaHonSinRet = AuxIdCta
            CodCtaHonSinRet = FmtCodCuenta(AuxCodCta)
            DescCtaHonSinRet = AuxDescCta
         Else
            IdCtaHonSinRet = lCtaHonSinRet.id
            CodCtaHonSinRet = FmtCodCuenta(lCtaHonSinRet.Codigo)
            DescCtaHonSinRet = lCtaHonSinRet.Descripcion
         End If
      Else
         IdCtaHonSinRet = 0
         CodCtaHonSinRet = ""
         DescCtaHonSinRet = ""
      End If
      
      If Bruto <> 0 Then
         If AuxIdCta > 0 Then
            IdCtaBruto = AuxIdCta
            CodCtaBruto = FmtCodCuenta(AuxCodCta)
            DescCtaBruto = AuxDescCta
         Else
            IdCtaBruto = lCtaBruto.id
            CodCtaBruto = FmtCodCuenta(lCtaBruto.Codigo)
            DescCtaBruto = lCtaBruto.Descripcion
         End If
      Else
         IdCtaBruto = 0
         CodCtaBruto = ""
         DescCtaBruto = ""
      End If
                                 
'      If HonSinRet <> 0 Then
'         IdCtaHonSinRet = lCtaHonSinRet.id
'         CodCtaHonSinRet = FmtCodCuenta(lCtaHonSinRet.Codigo)
'         DescCtaHonSinRet = lCtaHonSinRet.Descripcion
'      Else
'         IdCtaHonSinRet = 0
'         CodCtaHonSinRet = ""
'         DescCtaHonSinRet = ""
'      End If
'
'
'      If Bruto <> 0 Then
'         IdCtaBruto = lCtaBruto.id
'         CodCtaBruto = FmtCodCuenta(lCtaBruto.Codigo)
'         DescCtaBruto = lCtaBruto.Descripcion
'      Else
'         IdCtaBruto = 0
'         CodCtaBruto = ""
'         DescCtaBruto = ""
'      End If

                                    
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

               FldArray(0).FldName = "NotValidRut"
               FldArray(0).FldValue = 0
               FldArray(0).FldIsNum = True
               
               FldArray(1).FldName = "RUT"
               FldArray(1).FldValue = RutEnt
               FldArray(1).FldIsNum = False
                     
               FldArray(2).FldName = "IdEmpresa"
               FldArray(2).FldValue = gEmpresa.id
               FldArray(2).FldIsNum = True
                           
               FldArray(3).FldName = "Codigo"
               FldArray(3).FldValue = CodEnt
               FldArray(3).FldIsNum = False
               
               FldArray(4).FldName = "Nombre"
               FldArray(4).FldValue = NombEnt
               FldArray(4).FldIsNum = False
               
               FldArray(5).FldName = "Clasif" & ClasifEnt
               FldArray(5).FldValue = 1
               FldArray(5).FldIsNum = True
               
               IdEnt = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)

            End If
            
         End If
         
         'Esta validación tiene que estar después de validar la anetidad
         If Ret3Porc > 0 Then
            If Not GetEntRet3Porc(IdEnt, 0, 0) Then
               ValidaRet3Porc = False
               Call AddLogImp(lFNameLogImp, fname, l, "La entidad no tiene la característica de Retención 3%.")
            End If
         End If
         
        
         
               
         If ValidaRet3Porc = True Then
           'si no hay errores, ingresamos el registro a la grilla
            Grid.TextMatrix(Row, C_NUMLIN) = vFmt(Grid.TextMatrix(Row - 1, C_NUMLIN)) + 1
            Grid.TextMatrix(Row, C_FECHA) = Day(DtRec)
            Grid.TextMatrix(Row, C_IDTIPODOC) = IdTipoDoc
            Grid.TextMatrix(Row, C_TIPODOC) = TipoDoc
            
            Grid.TextMatrix(Row, C_DTE) = IIf(DTE <> 0, "x", "")
            Grid.TextMatrix(Row, C_NUMDOC) = NumDoc
            Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(DtEmi, SDATEFMT)
            Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = DtEmi
            Grid.TextMatrix(Row, C_RUT) = FmtRut(RutEnt)
            Grid.TextMatrix(Row, C_NOMBRE) = NombEnt
            Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
            Grid.TextMatrix(Row, C_DESCRIP) = Descrip
            Grid.TextMatrix(Row, C_IDSUCURSAL) = IdSucursal
            Grid.TextMatrix(Row, C_SUCURSAL) = Sucursal
                    
            Grid.TextMatrix(Row, C_HONORSINRET) = Format(HonSinRet, NUMFMT)
            If HonSinRet > 0 Then
               Grid.TextMatrix(Row, C_IDCUENTA) = IdCtaHonSinRet
               Grid.TextMatrix(Row, C_CODCUENTA) = CodCtaHonSinRet
               Grid.TextMatrix(Row, C_CUENTA) = DescCtaHonSinRet
            End If
            
            Grid.TextMatrix(Row, C_BRUTO) = Format(Bruto, NUMFMT)
            
            If Bruto > 0 Then
               Grid.TextMatrix(Row, C_IDCUENTA) = IdCtaBruto
               Grid.TextMatrix(Row, C_CODCUENTA) = CodCtaBruto
               Grid.TextMatrix(Row, C_CUENTA) = DescCtaBruto
            End If
               
            Grid.TextMatrix(Row, C_PIMPTO) = StrPImp
            Grid.TextMatrix(Row, C_IDPIMPTO) = IdPImp
            
            Grid.TextMatrix(Row, C_IMPTO) = Format(Impuesto, NUMFMT)
            Grid.TextMatrix(Row, C_RET3PORC) = Format(Ret3Porc, NUMFMT)
            Grid.TextMatrix(Row, C_NETO) = Format(Neto, NUMFMT)
            
            If vFmt(Grid.TextMatrix(Row, C_IMPTO)) <> 0 Then
               Grid.TextMatrix(Row, C_IMP_IDCUENTA) = gCtasBas.IdCtaImpRet
            Else
               Grid.TextMatrix(Row, C_IMP_IDCUENTA) = 0
            End If
            
            If vFmt(Grid.TextMatrix(Row, C_RET3PORC)) <> 0 Then
               Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA) = gCtasBas.IdCtaRet3Porc
            Else
               Grid.TextMatrix(Row, C_RET3PORC_IDCUENTA) = 0
            End If
            
            Grid.TextMatrix(Row, C_IDTIPORETEN) = IdTipoReten
            Grid.TextMatrix(Row, C_TIPORETEN) = TipoReten
            
            If vFmt(Grid.TextMatrix(Row, C_NETO)) <> 0 Then
               If IdTipoReten = TR_HONORARIOS Then
                  Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoHon
               Else
                  Grid.TextMatrix(Row, C_NETO_IDCUENTA) = gCtasBas.IdCtaNetoDieta
               End If
            Else
               Grid.TextMatrix(Row, C_NETO_IDCUENTA) = 0
            End If
            
            
            Grid.TextMatrix(Row, C_DETALLE) = TX_DETALLE
            Grid.TextMatrix(Row, C_FECHAVENC) = IIf(DtVenc <> 0, Format(DtVenc, SDATEFMT), "")
            Grid.TextMatrix(Row, C_LNGFECHAVENC) = DtVenc
            Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(Estado)
            Grid.TextMatrix(Row, C_IDESTADO) = Estado
            Grid.TextMatrix(Row, C_USUARIO) = gUsuario.Nombre
            Grid.TextMatrix(Row, C_UPDATE) = FGR_I
            
            Call CalcTotRow(Row, False)   'no recalcula Impuesto, deja el que viene
            Call CalcTot
            
            Row = Row + 1
            r = r + 1
         
            Grid.rows = Grid.rows + 1
         
         Else
            NRecErroneos = NRecErroneos + 1
         End If
         
      Else
         NRecErroneos = NRecErroneos + 1
      End If
      
NextRec:
   Loop

   Close #Fd
   
   Grid.FlxGrid.Redraw = True
   
   Me.MousePointer = vbDefault
   
   If NRecErroneos = 0 Then
      If r = 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
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
      ElseIf r > 1 Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
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

Private Sub Bt_CerrarOpt_Click()
   Fr_Opciones.visible = False
End Sub
Private Sub Bt_Opciones_Click()

   If lOper <> O_EDIT Then
      Fr_Opciones.Caption = "Opciones de Vista"
   End If
   
   Fr_Opciones.visible = Not Fr_Opciones.visible
   
End Sub

Private Sub SetupRet3Porc()

   If gVarIniFile.VerRet3Porc <> 0 Then
      If Val(Cb_Ano) >= 2021 And Val(Cb_Ano) <= 2024 Then
         Grid.ColWidth(C_RET3PORC) = 1050
         If Val(Cb_Ano) = 2021 And CbItemData(Cb_Mes) < 9 Then
            Grid.ColWidth(C_RET3PORC) = 0
         End If
      Else
         Grid.ColWidth(C_RET3PORC) = 0
      End If
   Else
      Grid.ColWidth(C_RET3PORC) = 0
   End If
   
   If Grid.ColWidth(C_RET3PORC) = 0 Then
      Grid.TextMatrix(0, C_RET3PORC) = ""
      Grid.TextMatrix(1, C_RET3PORC) = ""
      
   Else
      Grid.TextMatrix(0, C_RET3PORC) = "Retención 3%"
      Grid.TextMatrix(1, C_RET3PORC) = "Prést. Sol."
   End If

   DoEvents
   
   
   Call FGrTotales(Grid, GridTot)
End Sub
Private Function GetEntRet3Porc(ByVal IdEntidad As Long, ByRef FDesde As Long, ByRef FHasta As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   
   GetEntRet3Porc = 0
   
   Q1 = "SELECT Ret3Porc, FDesde3Porc, FHasta3Porc FROM Entidades WHERE IdEntidad = " & IdEntidad
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      FDesde = vFld(Rs("FDesde3Porc"))
      FHasta = vFld(Rs("FHasta3Porc"))
      GetEntRet3Porc = vFld(Rs("Ret3Porc"))
   End If
   
   Call CloseRs(Rs)
   
End Function

Private Sub RecalcRet3Porc()
   Dim i As Integer
   Dim Tot As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
   
      
      'Call CalcTotRow(i, true)
      Call CalcTotRow(i, False) ' 2782040
      
      Tot = Tot + vFmt(Grid.TextMatrix(i, C_RET3PORC))
   Next i

   GridTot.TextMatrix(0, C_RET3PORC) = Format(Tot, NUMFMT)
   
End Sub

'2784017
Private Sub Impuesto3RetPorc()

 Dim Cb As ComboBox

 '2784017
  Dim CurYear As Long
      CurYear = DateSerial(Cb_Ano, 1, 1)
   
   Set Cb = Grid.CbList(C_PIMPTO)
   
   gImpRet(IMPRET_NAC) = ImpBolHono(CurYear)
   ' fin 2784017
   
   Cb.AddItem gImpRet(IMPRET_NAC) * 100 & "%"
   Cb.ItemData(Cb.NewIndex) = IMPRET_NAC
   Cb.AddItem gImpRet(IMPRET_EXT) * 100 & "%"
   Cb.ItemData(Cb.NewIndex) = IMPRET_EXT
   Cb.AddItem "Otro"
   Cb.ItemData(Cb.NewIndex) = IMPRET_OTRO
   
   Set Cb = Grid.CbList(C_TIPORETEN)
   Cb.AddItem "Honorarios"
   Cb.ItemData(Cb.NewIndex) = TR_HONORARIOS
   Cb.AddItem "Dieta"
   Cb.ItemData(Cb.NewIndex) = TR_DIETA
   Cb.AddItem "Otro"
   Cb.ItemData(Cb.NewIndex) = TR_OTRO
   

End Sub

'2868501
Private Function CuentaCentro(vIdCuenta As String) As Boolean
 Dim Q1 As String
 Dim Rs As Recordset
 
 CuentaCentro = False

         Q1 = "SELECT count(*) as valor FROM Cuentas WHERE idCuenta =" & vIdCuenta & ""
         Q1 = Q1 & " AND Atrib5 = 1 and IdEmpresa = " & gEmpresa.id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            If vFld(Rs("valor")) > 0 Then
         
            CuentaCentro = True
            
            ElseIf vFld(Rs("valor")) = 0 Then
             CuentaCentro = False
            End If
         End If
         
         Call CloseRs(Rs)

End Function
'fin 2868501

Private Function CuentaArea(vIdCuenta As String) As Boolean
 Dim Q1 As String
 Dim Rs As Recordset
 
 CuentaArea = False

         Q1 = "SELECT count(*) as valor FROM Cuentas WHERE idCuenta =" & vIdCuenta & ""
         Q1 = Q1 & " And Atrib6 = 1 and IdEmpresa = " & gEmpresa.id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            If vFld(Rs("valor")) > 0 Then
         
            CuentaArea = True
            
            ElseIf vFld(Rs("valor")) = 0 Then
             CuentaArea = False
            End If
         End If
         
         Call CloseRs(Rs)

End Function
'fin 2868501


'2868501
Private Function CentroCosto() As Integer
 Dim Q1 As String
 Dim Rs As Recordset
 
 CentroCosto = 0

         Q1 = "SELECT IDCCOSTO,DESCRIPCION FROM CENTROCOSTO "
         Q1 = Q1 & " WHERE IDCCOSTO = 1 "
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            CentroCosto = vFld(Rs("IDCCOSTO"))
         End If
         
         Call CloseRs(Rs)

End Function
'fin 2868501

'2868501

Private Function AreaNegocio() As Integer
 Dim Q1 As String
 Dim Rs As Recordset
 
 AreaNegocio = 0

         Q1 = "SELECT IDAREANEGOCIO, DESCRIPCION FROM AREANEGOCIO "
         Q1 = Q1 & " WHERE IDAREANEGOCIO= 1 "
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            AreaNegocio = vFld(Rs("IDAREANEGOCIO"))
         End If
         
         Call CloseRs(Rs)

End Function

'fin 2868501


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
   
   Grid.FlxGrid.Redraw = False
   
   If ItemData(Cb_Mes) > 0 Then
      Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), ItemData(Cb_Mes), 1), FirstDay, LastDay)
   Else
      FirstDay = DateSerial(Val(Cb_Ano), 1, 1)
      LastDay = DateSerial(Val(Cb_Ano), 12, 31)
   End If
   
   If Fr_List.Enabled = True Then
      If ItemData(Cb_TipoDoc) > 0 Then
         Where = Where & " AND Documento.TipoDoc = " & ItemData(Cb_TipoDoc)
      End If
      
      If ItemData(Cb_Estado) > 0 Then
         Where = Where & " AND Documento.Estado = " & ItemData(Cb_Estado)
      End If
      
      If Val(Tx_NumDoc) <> 0 Then
         Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
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
   End If
   
'   If Row > 0 Then
'      Where = Where & " AND IdDoc=" & Val(Grid.TextMatrix(Row, C_IDDOC))
'   End If

   If ItemData(Cb_Sucursal) > 0 Then
      Where = Where & " AND Documento.IdSucursal = " & ItemData(Cb_Sucursal)
   End If

   Q1 = "SELECT IdDoc, Documento.TipoDoc, NumDoc, NumDocHasta, DTE, Documento.IdEntidad, Documento.RutEntidad, Documento.MovEdited, "
   Q1 = Q1 & " Documento.NombreEntidad, Entidades.Rut, Entidades.Nombre, Entidades.NotValidRut, FEmision, FVenc, FEmisionOri, Exento, Documento.IdCompCent, Documento.IdCompPago,"
   Q1 = Q1 & " Afecto, IVA, OtroImp, PorcentRetencion, TipoRetencion, Total, Descrip, Documento.Estado, Documento.ValRet3Porc, "
   Q1 = Q1 & " IdCuentaExento, Usuarios.Usuario, Cuentas1.Codigo as CodCtaEx, Cuentas1.Descripcion as DescCtaEx, "
   Q1 = Q1 & " IdCuentaAfecto, Cuentas2.Codigo as CodCtaAf, Cuentas2.Descripcion as DescCtaAf, IdCuentaOtroImp, Documento.IdCuentaRet3Porc, IdCuentaTotal, "
   Q1 = Q1 & " Documento.IdSucursal, Sucursales.Descripcion as DescSucursal "
   Q1 = Q1 & " FROM ((((( Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
   Q1 = Q1 & " LEFT JOIN Usuarios ON Documento.IdUsuario = Usuarios.IdUsuario )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas1") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaAfecto = Cuentas2.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas2") & " )"
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Sucursales", True, True) & " )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & lTipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & Where
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   
'   If Row = 0 Then
'      Q1 = Q1 & SqlPaging(gDbType, lClsPaging.CurReg - 1, gPageNumReg)
'   End If
   
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
