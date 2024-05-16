VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LPContabilidad - Fairware Ltda. - 2002"
   ClientHeight    =   7470
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   12420
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   12420
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Access 
      Height          =   540
      Index           =   3
      Left            =   1320
      TabIndex        =   20
      Top             =   6900
      Width           =   9735
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
         Left            =   60
         Picture         =   "FrmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Calculadora"
         Top             =   120
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
         Left            =   1200
         Picture         =   "FrmMain.frx":0C47
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Convertir moneda"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Equivalencia 
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
         Left            =   1620
         Picture         =   "FrmMain.frx":10CF
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Equivalencias"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Indices 
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
         Left            =   2040
         Picture         =   "FrmMain.frx":151F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Valores e Índices"
         Top             =   120
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
         Left            =   480
         Picture         =   "FrmMain.frx":1927
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Calendario"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Fr_Access 
      Height          =   7440
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Bt_Contrib14Ter 
         Caption         =   "Contrib. 14 D LIR ProPyme"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":1D65
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Listar, Centralizar o Buscar documentos"
         Top             =   6360
         Width           =   1155
      End
      Begin VB.CommandButton Bt_NewDoc 
         Caption         =   "Ingresar/Editar"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":21FD
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ingresar o Modificar Documento"
         Top             =   4560
         Width           =   1155
      End
      Begin VB.CommandButton Bt_LstDoc 
         Caption         =   "Listar Centralizar"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":27ED
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Listar, Centralizar o Buscar documentos"
         Top             =   5460
         Width           =   1155
      End
      Begin VB.CommandButton Bt_LstComp 
         Caption         =   "Listar / Editar"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":2DD2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Listado de Comprobantes"
         Top             =   3240
         Width           =   1155
      End
      Begin VB.CommandButton Bt_NewComprob 
         Caption         =   "Nuevo"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3356
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo Comprobante"
         Top             =   2340
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Plan 
         Caption         =   "Plan Cuentas"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3685
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Plan de Cuentas"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Emp 
         Caption         =   "Empresa"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":3B8B
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Datos Empresa"
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Documentos"
         ForeColor       =   &H00A67300&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   33
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Comprobantes"
         ForeColor       =   &H00A67300&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   2100
         Width           =   1155
      End
   End
   Begin VB.Frame Fr_Access 
      Height          =   7440
      Index           =   1
      Left            =   11100
      TabIndex        =   19
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Bt_Result 
         Caption         =   "Resultado"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":42B8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Estado de Resultado"
         Top             =   3780
         Width           =   1155
      End
      Begin VB.CommandButton Bt_InfAnalitico 
         Caption         =   "Inf. Analítico"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":48B5
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta de Saldos"
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Libros 
         Caption         =   "Libros"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":4DEB
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Libros"
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Balances 
         Caption         =   "Balances"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":53CB
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Balances"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CommandButton Bt_MantActivoFijo 
         Caption         =   "Mantención"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":59CC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Mantención de Activos Fijos"
         Top             =   5100
         Width           =   1155
      End
      Begin VB.CommandButton Bt_ContActFijo 
         Caption         =   "Rep. Control"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":5FF7
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Reporte de Control de Activo Fijo Tributario"
         Top             =   6000
         Width           =   1155
      End
      Begin VB.CommandButton Bt_InfoIFRS 
         Caption         =   "Informes IFRS"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":65B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Informes IFRS"
         Top             =   1980
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Activos Fijos"
         ForeColor       =   &H00A67300&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   40
         Top             =   4860
         Width           =   1155
      End
      Begin VB.Label La_demo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "DEMO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   39
         Top             =   6960
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame Fr_Top 
      Height          =   1275
      Left            =   1320
      TabIndex        =   22
      Top             =   0
      Width           =   9735
      Begin VB.Line Line1 
         BorderColor     =   &H00A67300&
         Index           =   5
         Visible         =   0   'False
         X1              =   0
         X2              =   12400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Lb_Mes 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   240
         Left            =   8460
         TabIndex        =   34
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Lb_Cierre 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(año cerrado)"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   8505
         TabIndex        =   31
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   10
         Left            =   6900
         TabIndex        =   29
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Lb_Año 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   240
         Left            =   9000
         TabIndex        =   28
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Lb_Tel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "341 5788       205 4335"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   7800
         TabIndex        =   27
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Lb_Dir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El Belloto 3942, P1"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   1140
         TabIndex        =   26
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Lb_RUT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "77.049.060-K"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   1140
         TabIndex        =   25
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUT: "
         ForeColor       =   &H00A67300&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Lb_Empresa 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   345
         Left            =   4830
         TabIndex        =   32
         Top             =   180
         Width           =   75
      End
   End
   Begin VB.PictureBox Pc_Access 
      DrawStyle       =   5  'Transparent
      Height          =   5595
      Left            =   1300
      Picture         =   "FrmMain.frx":6A9A
      ScaleHeight     =   5535
      ScaleWidth      =   9705
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   9765
      Begin VB.Frame Fr_Invisible 
         Caption         =   "Invisibles"
         Height          =   1875
         Left            =   5280
         TabIndex        =   35
         Top             =   420
         Visible         =   0   'False
         Width           =   3855
         Begin FlexEdGrid3.FEd3Grid FEd3Grid1 
            Height          =   435
            Left            =   2580
            TabIndex        =   43
            Top             =   300
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
            Cols            =   2
            Rows            =   2
            FixedCols       =   1
            FixedRows       =   1
            ScrollBars      =   3
            AllowUserResizing=   0
            HighLight       =   1
            SelectionMode   =   0
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   -1  'True
            Locked          =   0   'False
         End
         Begin VB.Timer Tmr_ChkActive 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   2700
            Top             =   900
         End
         Begin VB.PictureBox Pc_Nota 
            AutoSize        =   -1  'True
            Height          =   135
            Left            =   2280
            Picture         =   "FrmMain.frx":165CF
            ScaleHeight     =   75
            ScaleWidth      =   75
            TabIndex        =   42
            Top             =   300
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Timer Tm_ChkUsr 
            Interval        =   60000
            Left            =   2220
            Top             =   900
         End
         Begin VB.Timer Tmr_Chk 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   1680
            Top             =   900
         End
         Begin VB.PictureBox Pc_Flecha 
            AutoSize        =   -1  'True
            Height          =   150
            Left            =   1920
            Picture         =   "FrmMain.frx":16637
            ScaleHeight     =   90
            ScaleWidth      =   135
            TabIndex        =   38
            Top             =   300
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.PictureBox Pc_Lupa 
            AutoSize        =   -1  'True
            Height          =   270
            Left            =   1500
            Picture         =   "FrmMain.frx":166A5
            ScaleHeight     =   210
            ScaleWidth      =   210
            TabIndex        =   37
            Top             =   300
            Visible         =   0   'False
            Width           =   270
         End
         Begin FlexEdGrid2.FEd2Grid FEd2Grid1 
            Height          =   495
            Left            =   180
            TabIndex        =   36
            Top             =   300
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   873
            Cols            =   2
            Rows            =   2
            FixedCols       =   1
            FixedRows       =   1
            ScrollBars      =   3
            AllowUserResizing=   0
            HighLight       =   1
            SelectionMode   =   0
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   -1  'True
            Locked          =   0   'False
         End
         Begin MSComDlg.CommonDialog Cm_ComDlg 
            Left            =   420
            Top             =   900
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog Cm_PrtDlg 
            Left            =   1020
            Top             =   900
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Im_Orden 
            Height          =   270
            Left            =   180
            Picture         =   "FrmMain.frx":16A1A
            Top             =   900
            Width           =   75
         End
         Begin VB.Image Im_Down 
            BorderStyle     =   1  'Fixed Single
            Height          =   105
            Left            =   1200
            Picture         =   "FrmMain.frx":16D71
            Top             =   300
            Visible         =   0   'False
            Width           =   150
         End
      End
      Begin VB.Label Lb_Version 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   375
         Left            =   1440
         TabIndex        =   44
         Top             =   1980
         Width           =   2925
      End
   End
   Begin VB.PictureBox Pc_SQLServer 
      AutoSize        =   -1  'True
      Height          =   5610
      Left            =   1300
      Picture         =   "FrmMain.frx":16DFF
      ScaleHeight     =   5550
      ScaleWidth      =   9750
      TabIndex        =   45
      Top             =   1320
      Width           =   9810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00A67300&
      Index           =   4
      Visible         =   0   'False
      X1              =   0
      X2              =   12400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu M_Empresa 
      Caption         =   "&Empresa"
      Begin VB.Menu M_SelEmp 
         Caption         =   "&Seleccionar..."
      End
      Begin VB.Menu M_EditEmp 
         Caption         =   "&Modificar..."
      End
      Begin VB.Menu Sep_Emp2 
         Caption         =   "-"
      End
      Begin VB.Menu M_Salir 
         Caption         =   "&Salir            Alt+F4"
      End
   End
   Begin VB.Menu M_Base 
      Caption         =   "De&finiciones"
      Begin VB.Menu M_PlanCtas 
         Caption         =   "&Plan de Cuentas"
         Begin VB.Menu M_Plan 
            Caption         =   "&Ingresar / Modificar..."
         End
         Begin VB.Menu M_LstCuentas 
            Caption         =   "&Listado de Cuentas..."
         End
      End
      Begin VB.Menu Sep_Datos2 
         Caption         =   "-"
      End
      Begin VB.Menu M_EntRel 
         Caption         =   "&Entidades Relacionadas..."
      End
      Begin VB.Menu Sep_Datos3 
         Caption         =   "-"
      End
      Begin VB.Menu M_AreasNeg 
         Caption         =   "Áreas de &Negocio..."
      End
      Begin VB.Menu M_CentrosGestion 
         Caption         =   "&Centros de Gestión..."
      End
      Begin VB.Menu M_Sucursales 
         Caption         =   "&Sucursales..."
      End
      Begin VB.Menu Sep_Datos4 
         Caption         =   "-"
      End
      Begin VB.Menu M_TipoDocs 
         Caption         =   "&Tipos de Documentos..."
      End
   End
   Begin VB.Menu M_Comprob 
      Caption         =   "&Comprobantes"
      Begin VB.Menu M_NewComprob 
         Caption         =   "&Nuevo..."
      End
      Begin VB.Menu M_ListComprob 
         Caption         =   "&Listar/Editar..."
      End
      Begin VB.Menu M_EditCompAp 
         Caption         =   "Modificar Comprobante de Apertura"
         Begin VB.Menu M_EditCompApFin 
            Caption         =   "Financiero..."
         End
         Begin VB.Menu M_EditCompApTrib 
            Caption         =   "Tributario..."
         End
      End
      Begin VB.Menu Sep_Compr1 
         Caption         =   "-"
      End
      Begin VB.Menu M_PrtComprob 
         Caption         =   "&Impresión masiva..."
      End
      Begin VB.Menu M_ImpComprobantes 
         Caption         =   "Importar Comprobantes..."
      End
      Begin VB.Menu Sep_Compr2 
         Caption         =   "-"
      End
      Begin VB.Menu M_LstCompTipo 
         Caption         =   "Comprobantes &tipo..."
      End
      Begin VB.Menu Sep_Compr3 
         Caption         =   "-"
      End
      Begin VB.Menu M_Renum 
         Caption         =   "Renumerar Comprobantes..."
      End
      Begin VB.Menu M_Auditoria 
         Caption         =   "Informe de Auditoría de Comprobantes..."
      End
      Begin VB.Menu M_MANT_PERCE 
         Caption         =   "Mantenedor Percepciones"
      End
   End
   Begin VB.Menu M_Docs 
      Caption         =   "&Documentos"
      Begin VB.Menu M_NewDoc 
         Caption         =   "&Ingresar / Modificar..."
      End
      Begin VB.Menu M_LstDocs 
         Caption         =   "&Listar / Centralizar..."
      End
      Begin VB.Menu Sep_D1 
         Caption         =   "-"
      End
      Begin VB.Menu M_LibroCaja 
         Caption         =   "Libro de Caja"
         Begin VB.Menu M_EditLibCaja 
            Caption         =   "Ingresar Libro de Caja..."
         End
         Begin VB.Menu M_ViewLibCaja 
            Caption         =   "Listar Libro de Caja..."
         End
      End
      Begin VB.Menu M_LibroIngEg 
         Caption         =   "Libro de Ingresos y Egresos..."
      End
      Begin VB.Menu SepLibCaja 
         Caption         =   "-"
      End
      Begin VB.Menu M_ResLibAux 
         Caption         =   "&Resumen de Libros Auxiliares..."
      End
      Begin VB.Menu M_TraspODToODF 
         Caption         =   "Traspaso de OD a ODF"
      End
      Begin VB.Menu M_ImpOtrosDocs 
         Caption         =   "Importar Otros Documentos..."
      End
      Begin VB.Menu M_RecalcSaldos 
         Caption         =   "Re&calcular saldos..."
      End
   End
   Begin VB.Menu MActFijo 
      Caption         =   "&Activo Fijo"
      Begin VB.Menu M_ConfigActFijo 
         Caption         =   "Configuración Activos Fijos Financieros..."
      End
      Begin VB.Menu M_ActFijo 
         Caption         =   "Mantención Activos Fijos..."
      End
      Begin VB.Menu Sep_ActFijo 
         Caption         =   "-"
      End
      Begin VB.Menu M_InfActFijo 
         Caption         =   "Control Activo Fijo Tributario..."
      End
      Begin VB.Menu M_RepActFijoIFRS 
         Caption         =   "Control Activo Fijo Financiero (IFRS)"
      End
      Begin VB.Menu Sep_ActFijo2 
         Caption         =   "-"
      End
      Begin VB.Menu M_AFImportFile 
         Caption         =   "Importar Activos Fijos desde Archivo..."
      End
      Begin VB.Menu M_ReimportActFijo 
         Caption         =   "Traer Activos Fijos año anterior..."
      End
      Begin VB.Menu Sep_ActFijo3 
         Caption         =   "-"
      End
      Begin VB.Menu M_ManActFijo 
         Caption         =   "Manual Módulo Activo Fijo..."
      End
   End
   Begin VB.Menu M_Informes 
      Caption         =   "Info&rmes"
      Begin VB.Menu M_Libros 
         Caption         =   "&Libros..."
      End
      Begin VB.Menu M_Balances 
         Caption         =   "&Balances (bajo norma antigua)"
         Begin VB.Menu M_BalComprob 
            Caption         =   "&Comprobación y Saldos..."
         End
         Begin VB.Menu M_BalTrib 
            Caption         =   "&General 8 Columnas..."
         End
         Begin VB.Menu SepBal1 
            Caption         =   "-"
         End
         Begin VB.Menu M_BalClasif 
            Caption         =   "C&lasificado..."
         End
         Begin VB.Menu M_BalClasifANeg 
            Caption         =   "Clasificado por Área de Negocio..."
         End
         Begin VB.Menu M_BalClasifCCosto 
            Caption         =   "Clasificado por Centro de Costo..."
         End
         Begin VB.Menu SepClasif 
            Caption         =   "-"
         End
         Begin VB.Menu M_BalClasifComp 
            Caption         =   "Clasificado Comparativo..."
         End
         Begin VB.Menu M_BalClasifEjec 
            Caption         =   "Clasificado Ejecutivo..."
         End
      End
      Begin VB.Menu Sep_Info1 
         Caption         =   "-"
      End
      Begin VB.Menu M_InfoAnalit 
         Caption         =   "Informe &Analítico"
         Begin VB.Menu M_InfAnalitEnt 
            Caption         =   "por &Entidad..."
         End
         Begin VB.Menu M_InfAnalitCta 
            Caption         =   "por &Cuenta..."
         End
      End
      Begin VB.Menu M_InfoAnalitODF 
         Caption         =   "Informe &Analitíco ODF"
         Begin VB.Menu M_InfAnalitEntODF 
            Caption         =   "por &Entidad..."
         End
         Begin VB.Menu M_InfAnalitCtaODF 
            Caption         =   "por &Cuenta..."
         End
      End
      Begin VB.Menu M_AuditLibrosContables 
         Caption         =   "Auditoría de Libros Contables..."
      End
      Begin VB.Menu M_InfoOtrosDocs 
         Caption         =   "Informe &Otros Documentos..."
      End
      Begin VB.Menu Sep_Info2 
         Caption         =   "-"
      End
      Begin VB.Menu M_EstadoRes 
         Caption         =   "&Estado de Resultado"
         Begin VB.Menu M_ResClasificado 
            Caption         =   "&Clasificado..."
         End
         Begin VB.Menu M_ResClasificadoANeg 
            Caption         =   "&Clasificado por Área de Negocio..."
         End
         Begin VB.Menu M_ResClasificadoCCosto 
            Caption         =   "&Clasificado por Centro de Costo..."
         End
         Begin VB.Menu SepMResClas 
            Caption         =   "-"
         End
         Begin VB.Menu M_ResMensual 
            Caption         =   "&Mensual..."
         End
         Begin VB.Menu SepMRes 
            Caption         =   "-"
         End
         Begin VB.Menu M_ResComparativo 
            Caption         =   "Co&mparativo Mes Anterior..."
         End
         Begin VB.Menu M_ResCompPeriodo 
            Caption         =   "Comparativo Periodo Anterior..."
         End
      End
      Begin VB.Menu M_CapitalPropio 
         Caption         =   "Capital Propio..."
         Begin VB.Menu M_CapitalPropioTrib 
            Caption         =   "Capital Propio Tributario General..."
         End
         Begin VB.Menu M_CapitalPropioSimpl 
            Caption         =   "Capital Propio Simplificado"
            Begin VB.Menu M_CapitalPropioSimplDet 
               Caption         =   "General..."
               Index           =   1
            End
            Begin VB.Menu M_CapitalPropioSimplDet 
               Caption         =   "Variación del Año..."
               Index           =   2
            End
         End
      End
      Begin VB.Menu Sep_Info3 
         Caption         =   "-"
      End
      Begin VB.Menu M_CalcRazFin 
         Caption         =   "&Razones Financieras..."
      End
      Begin VB.Menu Sep_Info4 
         Caption         =   "-"
      End
      Begin VB.Menu M_RepPagoPlazo 
         Caption         =   "Reporte de Pagos a Plazo..."
      End
      Begin VB.Menu M_ResVPE 
         Caption         =   "Resumen Vales Pago Electrónico..."
      End
      Begin VB.Menu M_ResSupermercado 
         Caption         =   "Resumen Supermercados y/o Com. Similares..."
      End
      Begin VB.Menu Sep_Info5 
         Caption         =   "-"
      End
      Begin VB.Menu M_LstInfoImp 
         Caption         =   "Listado de Libros Impresos..."
      End
      Begin VB.Menu M_OtrosInformes 
         Caption         =   "&Otros Informes"
         Begin VB.Menu M_LstChequesEmit 
            Caption         =   "Listado de Cheques &Emitidos..."
         End
         Begin VB.Menu M_LstChequesAnula 
            Caption         =   "Listado de Cheques &Anulados..."
         End
         Begin VB.Menu M_LstChequesaFecha 
            Caption         =   "Listado de Cheques a &Fecha..."
         End
         Begin VB.Menu SepOInf1 
            Caption         =   "-"
         End
         Begin VB.Menu M_InfoVenc 
            Caption         =   "Informe de Vencimientos"
            Begin VB.Menu M_InfoVenc30 
               Caption         =   "a 30 días..."
            End
            Begin VB.Menu M_InfoVenc60 
               Caption         =   "a 60 días..."
            End
            Begin VB.Menu M_InfoVenc90 
               Caption         =   "a 90 días..."
            End
         End
      End
   End
   Begin VB.Menu M_InfoIFRS 
      Caption         =   "&IFRS"
      Begin VB.Menu M_ConfigIFRS 
         Caption         =   "Configurar Códigos IFRS..."
      End
      Begin VB.Menu Sep_IFRS1 
         Caption         =   "-"
      End
      Begin VB.Menu M_IFRS_EstFin 
         Caption         =   "Estado de Situación Financiera Clasificado..."
      End
      Begin VB.Menu M_IFRS_EstRes 
         Caption         =   "Estado de Resultados por Función..."
      End
      Begin VB.Menu Sep_InfoIFRS1 
         Caption         =   "-"
      End
      Begin VB.Menu M_IFRS_BalEjec 
         Caption         =   "Estado de Situación Financiera Ejecutivo..."
      End
      Begin VB.Menu M_IFRS_BalTrib 
         Caption         =   "Balance General 8 Columnas..."
      End
   End
   Begin VB.Menu M_Concil 
      Caption         =   "C&onciliación"
      Begin VB.Menu M_ProcConcil 
         Caption         =   "&Proceso de Conciliación..."
      End
      Begin VB.Menu M_IngCartola 
         Caption         =   "&Ingresar o Importar Cartolas..."
      End
      Begin VB.Menu Sep_Con1 
         Caption         =   "-"
      End
      Begin VB.Menu M_InfoResConcil 
         Caption         =   "Resumen de Conciliación..."
      End
      Begin VB.Menu M_InfoCartBanc 
         Caption         =   "Cartolas Bancarias..."
      End
      Begin VB.Menu M_InfoConciliacion 
         Caption         =   "Informe de Conciliación Bancaria..."
      End
      Begin VB.Menu Sep_Con2 
         Caption         =   "-"
      End
      Begin VB.Menu M_ManConciliacion 
         Caption         =   "Manual Conciliación Bancaria..."
      End
   End
   Begin VB.Menu M_Procesos 
      Caption         =   "&Procesos"
      Begin VB.Menu M_AbrirCerrarMes 
         Caption         =   "Abrir / Cerrar Mes..."
      End
      Begin VB.Menu Sep_P1 
         Caption         =   "-"
      End
      Begin VB.Menu M_CalcPropIVA 
         Caption         =   "Calcular Proporcionalidad de IVA CF..."
      End
      Begin VB.Menu Sep_PropIVA 
         Caption         =   "-"
      End
      Begin VB.Menu M_ExpF29 
         Caption         =   "Exportar a HR-IVA F 29..."
      End
      Begin VB.Menu M_ExpFUT 
         Caption         =   "Exportar a HR-FUT..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu M_ExpHRCertif 
         Caption         =   "Exportar a HR-Certificados"
         Begin VB.Menu M_ExpHRDJ1879 
            Caption         =   "DJ 1879..."
         End
         Begin VB.Menu M_ExpHRDJ1847 
            Caption         =   "DJ 1847..."
         End
         Begin VB.Menu M_ExpHRDJ1923 
            Caption         =   "DJ 1923 (Sección B)..."
         End
         Begin VB.Menu M_ExpHRDJ1924B 
            Caption         =   "DJ 1924 - Sección B"
         End
         Begin VB.Menu M_ExpHRDJ1924C 
            Caption         =   "DJ 1924 - Sección C"
         End
      End
      Begin VB.Menu M_ExpF22 
         Caption         =   "Exportar a HR-Form 22..."
      End
      Begin VB.Menu M_ExpHR_RABbase 
         Caption         =   "Exportar a HR-RAB"
         Begin VB.Menu M_ExpHR_RAB 
            Caption         =   "Exportar a HR-RAB - Resultado s/Balance..."
         End
         Begin VB.Menu M_ExpHR_RAB_RLI 
            Caption         =   "Exportar HR-RAB RLI - Ajustes RLI..."
         End
         Begin VB.Menu M_ExpHR_RetirosDividendos 
            Caption         =   "Exportar Retiros/Dividendos..."
         End
         Begin VB.Menu M_ExpHR_RAD 
            Caption         =   "Exportar 14D HR RAD..."
         End
         Begin VB.Menu M_ExpHR_RADPERC 
            Caption         =   "Exportar Percepciones"
         End
      End
      Begin VB.Menu M_ImpF29Av 
         Caption         =   "Importar desde HR IVA..."
      End
      Begin VB.Menu Sep_ExpImpLib 
         Caption         =   "-"
      End
      Begin VB.Menu M_ExpLibAux 
         Caption         =   "Exportar Libros Auxiliares en Sucursal..."
      End
      Begin VB.Menu M_ImpLibAux 
         Caption         =   "Importar Libros Auxiliares desde Sucursal..."
      End
      Begin VB.Menu M_ExpEntidades 
         Caption         =   "Exportar Entidades..."
      End
      Begin VB.Menu M_LibElectCompras 
         Caption         =   "Generar Libro Electrónico de Compras..."
      End
      Begin VB.Menu M_ImpRegSII 
         Caption         =   "Importar Registros SII"
         Begin VB.Menu M_LibCompSII 
            Caption         =   "Compras..."
         End
         Begin VB.Menu M_LibVentasSII 
            Caption         =   "Ventas..."
         End
         Begin VB.Menu M_LibRetenSII 
            Caption         =   "Retenciones"
         End
      End
      Begin VB.Menu M_ImpFacturacion 
         Caption         =   "Importar desde Facturación..."
      End
      Begin VB.Menu M_ExpFacturacion 
         Caption         =   "Exportar Libro Compras para Facturación Electrónica..."
      End
      Begin VB.Menu Sep_PF1 
         Caption         =   "-"
      End
      Begin VB.Menu M_ImpRemu 
         Caption         =   "Importar desde Remuneraciones..."
      End
      Begin VB.Menu Sep_Remu 
         Caption         =   "-"
      End
      Begin VB.Menu M_PrtHojasTimb 
         Caption         =   "&Foliar Hojas para Timbraje..."
      End
      Begin VB.Menu Sep_P3 
         Caption         =   "-"
      End
      Begin VB.Menu M_Periodo 
         Caption         =   "&Período Contable"
         Begin VB.Menu M_CerrarPer 
            Caption         =   "&Cerrar..."
         End
         Begin VB.Menu M_ReabrirPer 
            Caption         =   "&Reabrir..."
         End
      End
      Begin VB.Menu M_ContEmpresa 
         Caption         =   "&Control Empresa..."
      End
   End
   Begin VB.Menu M_ConfigTop 
      Caption         =   "Confi&guración"
      Begin VB.Menu M_Config 
         Caption         =   "&Configuración Inicial..."
      End
      Begin VB.Menu M_ActConfig 
         Caption         =   "&Actualizar Configuración..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu Sep_O1 
         Caption         =   "-"
      End
      Begin VB.Menu M_ConfigHojasTimbraje 
         Caption         =   "Información de Folios Timbraje..."
      End
      Begin VB.Menu Sep_O2 
         Caption         =   "-"
      End
      Begin VB.Menu M_Monedas 
         Caption         =   "&Monedas"
         Begin VB.Menu M_Equivalencias 
            Caption         =   "&Equivalencias..."
         End
         Begin VB.Menu M_ConfigMonedas 
            Caption         =   "&Configuración..."
         End
      End
      Begin VB.Menu M_Indices 
         Caption         =   "&Valores e Índices..."
      End
      Begin VB.Menu Sep_DefRazones 
         Caption         =   "-"
      End
      Begin VB.Menu M_ConfigCtasFUT 
         Caption         =   "Configurar Cuentas para FUT..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu M_ConfigCtasAjustes 
         Caption         =   "Configurar Cuentas Ajustes Extra-contables-14 TER A)..."
      End
      Begin VB.Menu M_ConfigCtasAjustesRLI 
         Caption         =   "Configurar  Cuentas Ajustes Extra - Contables RLI HR RAB..."
      End
      Begin VB.Menu M_ConfigRemu 
         Caption         =   "Configurar Traspaso Remuneraciones..."
      End
      Begin VB.Menu M_DefRazones 
         Caption         =   "Configurar Razones Financieras..."
      End
      Begin VB.Menu M_SepConfigPrtCheque 
         Caption         =   "-"
      End
      Begin VB.Menu M_ConfigPrtCheque 
         Caption         =   "Configurar Impresión de Cheques..."
      End
   End
   Begin VB.Menu M_Sistema 
      Caption         =   "&Sistema"
      Begin VB.Menu M_NuevaInstancia 
         Caption         =   "Nueva Instancia..."
      End
      Begin VB.Menu M_Seguridad 
         Caption         =   "&Seguridad"
         Begin VB.Menu M_CambiarClave 
            Caption         =   "&Cambiar clave..."
         End
      End
      Begin VB.Menu Sep_Sis2 
         Caption         =   "-"
      End
      Begin VB.Menu M_SetupPrt 
         Caption         =   "Preparar &Impresora..."
      End
      Begin VB.Menu Sep_Sis3 
         Caption         =   "-"
      End
      Begin VB.Menu M_MantDB 
         Caption         =   "&Mantención Base Datos"
         Begin VB.Menu M_Reparar 
            Caption         =   "&Reparar..."
         End
         Begin VB.Menu M_Compactar 
            Caption         =   "&Compactar..."
         End
         Begin VB.Menu M_Unlock 
            Caption         =   "Desbloquear procesos..."
         End
         Begin VB.Menu M_RelinkearTblBasicas 
            Caption         =   "Relinkear tablas básicas..."
         End
         Begin VB.Menu M_RevDetDoc 
            Caption         =   "Revisar cuadratura detalle Documentos..."
         End
      End
   End
   Begin VB.Menu MH__Help 
      Caption         =   "A&yuda"
      Begin VB.Menu MH_Tutorial 
         Caption         =   "Tutorial de Uso del Sistema..."
      End
      Begin VB.Menu MH_ManualesDeUso 
         Caption         =   "Manuales de Uso..."
      End
      Begin VB.Menu MH_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MH_HlpBackup 
         Caption         =   "Ayuda Respaldo..."
      End
      Begin VB.Menu MH_RepErr 
         Caption         =   "Reporte de problema..."
      End
      Begin VB.Menu MH_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MH_About 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu M_FormToolsInt 
         Caption         =   "FormTools..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BT_14TER = 6
Const BT_14D = 7

Dim FrmActivate As Boolean

Private Sub Bt_Balances_Click()
   Dim Frm As FrmSelBalances
   
   Set Frm = New FrmSelBalances
   Frm.Show vbModeless
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

Private Sub Bt_Conciliacion_Click(Index As Integer)
   Call M_ProcConcil_Click
End Sub

Private Sub Bt_ContActFijo_Click()
   Dim Frm As FrmSelRepActFijo
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT * FROM MovActivoFijo"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
      Call MsgBox1("No hay activos fijos ingresados en el sistema.", vbExclamation + vbOKOnly)
      Call CloseRs(Rs)
      Exit Sub
   End If
   
   Call CloseRs(Rs)
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmSelRepActFijo
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault
End Sub

Private Sub Bt_Contrib14Ter_Click()
   Dim Frm2020 As FrmSel14DProPyme
   Dim Frm As FrmSelLib14ter
   Dim Msg As Integer
      
   If gEmpresa.Ano >= 2020 Then
   
      Msg = Val(GetIniString(gIniFile, "Msg", "14DEntRelacionadas", "0"))

      If Msg = 0 Then
         MsgBox1 "Recuerde verificar la Franquicia Tributaria de las Entidades Relacionadas", vbInformation
         Call SetIniString(gIniFile, "Msg", "14DEntRelacionadas", "1")
      End If

      Set Frm2020 = New FrmSel14DProPyme
      Frm2020.Show vbModal
      Set Frm = Nothing
   Else
      Set Frm = New FrmSelLib14ter
      Call Frm.FView
      Set Frm = Nothing
   End If
   
   
End Sub

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   
   Set Frm = New FrmConverMoneda
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Emp_Click()

   Call M_EditEmp_Click
   
End Sub

Private Sub bt_Equivalencia_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_ImprTimb_Click(Index As Integer)
'   Dim Frm As FrmFoliacion
'
'   Set Frm = New FrmFoliacion
'   Frm.Show vbModal
'   Set Frm = Nothing
   
   Call M_PrtHojasTimb_Click
   
End Sub

Private Sub Bt_Indices_Click()
   Call M_Indices_Click
End Sub

Private Sub Bt_InfAnalitico_Click()
   Dim Frm As FrmSelInfAnalit
   
   Set Frm = New FrmSelInfAnalit
   Frm.Show vbModeless
   Set Frm = Nothing
   
End Sub

Private Sub Bt_InfoIFRS_Click()
   Dim Frm As FrmSelInfIFRS
   
   Set Frm = New FrmSelInfIFRS
   Frm.Show vbModeless
   Set Frm = Nothing

End Sub

Private Sub Bt_Libros_Click()
   Call M_Libros_Click
End Sub

Private Sub Bt_LstComp_Click()
   Call M_PrtComprob_Click
End Sub

Private Sub Bt_LstDoc_Click()
   Call M_LstDocs_Click
End Sub

Private Sub Bt_MantActivoFijo_Click()
   Call M_ActFijo_Click
End Sub

Private Sub Bt_NewComprob_Click()
   Call M_NewComprob_Click
End Sub

Private Sub Bt_NewDoc_Click()
   Call M_NewDoc_Click
End Sub

Private Sub Bt_Plan_Click()
   Call M_Plan_Click
End Sub
Private Sub Bt_Result_Click()
   Dim Frm As FrmSelEstRes
   
   Set Frm = New FrmSelEstRes
   Frm.Show vbModeless
   Set Frm = Nothing
End Sub


Private Sub Form_Activate()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Frm As FrmConfig
   Dim PlanVacio As Boolean
   Dim Rc As Integer
   Dim vSaldos As String
   
'   Debug.Print 566 * Screen.TwipsPerPixelX
'   Debug.Print 551 * Screen.TwipsPerPixelY

'14520904
' Me.MousePointer = vbHourglass
'vSaldos = "0"
'
'   Q1 = ""
'   Q1 = Q1 & " SELECT Codigo From ParamEmpresa where Tipo = 'SALDOS' and Valor ='" & ParaSQL(W.Version) & "'"
'   Q1 = Q1 & " and  IdEmpresa = " & gEmpresa.id
'   Q1 = Q1 & " and  Ano = " & gEmpresa.Ano
'
'  Set Rs = OpenRs(DbMain, Q1)
'   If Not Rs.EOF Then
'      vSaldos = vFld(Rs(0))
'   End If
'   Call CloseRs(Rs)
'
'  If vSaldos = "0" Then
'   'Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
'   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
'  End If
' Me.MousePointer = vbDefault
 '14520904
 
   M_FormToolsInt.visible = W.InDesign
   M_CapitalPropioSimplDet.Item(1).visible = IIf(gEmpresa.Ano > 2020, False, True)
   Call AddDebug("FrmMain_Activate: Antes de ExitDemo - FrmActivate=" & FrmActivate)
   
   If ExitDemo() Then
      Unload Me
   End If
   
   Call AddDebug("FrmMain_Activate: Antes de FrmActivate")
   Call InitBaseImponible14D
   Call InitPercepciones
   
   If FrmActivate = True Then
      Exit Sub
   End If
        
   Call AddDebug("FrmMain_Activate: Después de FrmActivate")
        
        
   Lb_Version = "Versión " & App.Major & "." & App.Minor & " " & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
   
 #If DATACON = 2 Then       'SQL Server o MySQL   Nueva interfaz
'   Lb_VersionSQL = "Versión " & App.Major & "." & App.Minor & " " & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
#End If

   FrmActivate = True
      
   Call AddDebug("FrmMain_Activate: Antes de Select")
   
   Q1 = "SELECT IdCuenta FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   PlanVacio = (Rs.EOF = True)
   Call CloseRs(Rs)
   
   Call AddDebug("FrmMain_Activate: Después de Select")
   
   If PlanVacio Then    'no hay cuentas definidas
      'Suponemos que es el primer año de la empresa y es la primera vez que se abre
      If GetDatosEmpHR(gEmpresa.Rut) Then
         Dim FrmEmp As FrmEmpresa
         
         Call AddDebug("FrmMain_Activate: Después de GetDatosEmpHR")
         Set FrmEmp = New FrmEmpresa
         Rc = FrmEmp.FEdit(gEmpresa.id, True)
         Set FrmEmp = Nothing
         
         Call AddDebug("FrmMain_Activate: Después de New FrmEmpresa")
         If Rc = vbOK Then
            Call FillDatosEmp
         End If
         
      End If
      
      MsgBox1 "Se ha detectado que no está definido el Plan de Cuentas para esta empresa." & vbNewLine & vbNewLine & "La ventana de Configuración Inicial le permitirá definir el Plan de Cuentas y otros elementos básicos para trabajar con el sistema.", vbInformation + vbOKOnly
      Set Frm = New FrmConfig
      Frm.Show vbModal
      Set Frm = Nothing
      
      MsgBox1 "Recuerde configurar las Razones Financieras para esta empresa, utilizando la opción que provee el sistema, bajo el menú 'Configuración'", vbOKOnly + vbInformation

   End If

   M_DefRazones.visible = gFunciones.RazFinancieras
   Sep_DefRazones.visible = gFunciones.RazFinancieras Or gFunciones.ExpFUT
   M_CalcRazFin.visible = gFunciones.RazFinancieras
   
   M_ConfigCtasFUT.visible = gFunciones.ExpFUT
   'M_ExpFUT.visible = gFunciones.ExpFUT
   
   M_ExpHRCertif.visible = gFunciones.ExpHRCertificados
   M_ExpF22.visible = gFunciones.ExpHRForm22
   
   M_InfActFijo.visible = gFunciones.ActivoFijo
   M_ActFijo.visible = gFunciones.ActivoFijo
   Sep_ActFijo.visible = gFunciones.ActivoFijo
   
   M_OtrosInformes.visible = gFunciones.OtrosInformes
   
   M_ExpLibAux.visible = gFunciones.ExpImpLibrosAux
   M_ImpLibAux.visible = gFunciones.ExpImpLibrosAux
   Sep_ExpImpLib.visible = gFunciones.ExpImpLibrosAux
     
'   M_ConfigPrtCheque.Enabled = gFunciones.PrtCheque And Not gAppCode.Demo
'   M_SepConfigPrtCheque.Enabled = gFunciones.PrtCheque And Not gAppCode.Demo
'
'   M_ConfigRemu.Enabled = gFunciones.ImportRemu And Not gAppCode.Demo
'   M_ConfigRemu.Enabled = gFunciones.ImportRemu And Not gAppCode.Demo
'
'   M_ImpRemu.Enabled = gFunciones.ImportRemu And Not gAppCode.Demo
'   M_ImpRemu.Enabled = gFunciones.ImportRemu And Not gAppCode.Demo

   M_ConfigIFRS.visible = gFunciones.IFRS
   Sep_IFRS1.visible = gFunciones.IFRS
   M_InfoIFRS.visible = gFunciones.IFRS
   M_IFRS_BalEjec.visible = gFunciones.IFRS_Ejecutivo
   M_IFRS_BalTrib.visible = gFunciones.IFRS_BalanceTributario
   Sep_InfoIFRS1.visible = gFunciones.IFRS_Ejecutivo Or gFunciones.IFRS_BalanceTributario
   
   M_ImpComprobantes.visible = gFunciones.ImportComprobantes
   M_Auditoria.visible = gFunciones.AuditoriaInterna
   
   M_CalcPropIVA.visible = gFunciones.ProporcionalidadIVA
   Sep_PropIVA.visible = gFunciones.ProporcionalidadIVA
   
'   M_ExpSII.Visible = gFunciones.ExpLibCompVentasSII
   
   M_ConfigActFijo.visible = gFunciones.ActFijoFinanciero
   M_RepActFijoIFRS.visible = gFunciones.RepActFijoFinanciero
   M_LibroCaja.visible = gFunciones.LibroCaja
   M_LibroIngEg.visible = gFunciones.LibroCaja
   'Sep_LibroCaja.visible = gFunciones.LibroCaja

   Bt_Contrib14Ter.visible = gFunciones.LibroCaja
   M_RepPagoPlazo.visible = gFunciones.DocCuotas
   
#If DATACON = 2 Then       'SQL Server o MySQL
   M_Reparar.visible = False
   M_Compactar.visible = False
   M_RelinkearTblBasicas.visible = False
   M_ImpF29Av.visible = False
   M_ExpLibAux.visible = False
   M_ImpLibAux.visible = False
#End If
   
'   M_ConfigCtasAjustes.Visible = gFunciones.OtrosIngEgresos
   
'   Call SetMainSQLServer   Nueva Interfaz
      
   Call AddDebug("FrmMain_Activate: nos vamos OK")
   
   Call ShowMsgBackup
   
   Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
   
   
   
    
End Sub

Private Sub Form_Load()
   Dim DbName As String
   Set gFrmMain = Me
   
   On Error Resume Next

   Call AddDebug("FrmMain_Load: Antes de FEd2Grid1")

   FEd2Grid1.TextMatrix(0, 0) = "$1#2¿P" ' No borrar
   FEd2Grid1.TextMatrix(0, 0) = "" ' No borrar
   
   FEd3Grid1.TextMatrix(0, 0) = "$7#3?F#" ' No borrar
   FEd3Grid1.TextMatrix(0, 0) = "" ' No borrar
   
   If ERR Then
      MsgErr "La versión del objeto FlexEdGrid2 no corresponde."
   End If
   
   Set gPrtDlg = Me.Cm_PrtDlg
   
   Call AddDebug("FrmMain_Load: Antes de FillDatosEmp")
   
   Call FillDatosEmp
   
   Call SetupPriv
   
   La_demo(0).visible = gAppCode.Demo
   La_demo(1).visible = gAppCode.Demo
   
   Tmr_Chk.Enabled = (gAppCode.Demo = False)
   
   Tmr_ChkActive.Enabled = Not gFwChkActive

   If gEmpresa.Ano < 2020 Then
      Bt_Contrib14Ter.Caption = "Contribuyentes 14TER Let. A"
   Else
      Bt_Contrib14Ter.Caption = "Contrib. 14 D LIR ProPyme"
   End If
   
   If gDbType = SQL_ACCESS Then
      Pc_SQLServer.visible = False
      Pc_Access.visible = True
   Else
      Pc_SQLServer.visible = True
      Pc_Access.visible = False
      M_RevDetDoc.visible = False         'esta opción no tiene sentido en SQL Server dado que no se daña la base de datos
   End If
   
   #If DATACON = 1 Then       'Access

   If gEmpresa.TieneAnoAnt Then

         DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
         If ExistFile(DbName) Then
         
            Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
            Call CorrigeBase
            Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)

         End If
   End If

#End If
    
   Call AddDebug("FrmMain_Load: nos vamos OK")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not DbMain Is Nothing Then
      Call ContUnregisterPc(1)
   End If
   Call CloseDb(DbMain)
   Call CheckRs(True)
   Cancel = True
   End
   
End Sub



Private Sub M_ActConfig_Click()
   Dim Rc As Integer
   Rc = MsgBox1("¿Desea volver a leer la configuración de la empresa desde la base de datos, para actualizarla en esta sesión, por posibles cambios realizados por el administrador?", vbYesNo + vbDefaultButton1 + vbQuestion)
   
   If Rc = vbYes Then
      
      Me.MousePointer = vbHourglass
      Call ReadDatosBasEmpresa
      Me.MousePointer = vbDefault
      
      MsgBox1 "Configuración actualizada con éxito.", vbInformation + vbOKOnly
   End If
   
End Sub

Private Sub M_ActFijo_Click()
   Dim Frm As FrmLstActFijo
   
   If Not gFunciones.ActivoFijo Then
      Exit Sub
   End If
         
   Call MsgLey21210("Estimado Usuario debido a que se publicó la Ley 21.210 Moderniza Legislación Tributaria D.O. 24.02.2020, este módulo de activo fijo sufrirá modificaciones que saldrán en próximas versiones del sistema")
         
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstActFijo
   Call Frm.FEdit
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_AFImportFile_Click()
   Dim Frm As FrmImpActFijoFile
   
   Set Frm = New FrmImpActFijoFile
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_Auditoria_Click()
'   Dim Frm As FrmAuditoria
'
'   Set Frm = New FrmAuditoria
'   Call Frm.FView
'   Set Frm = Nothing

    Dim Frm As FrmSelSeguimiento
    
    Set Frm = New FrmSelSeguimiento
    Frm.Show vbModal
    Set Frm = Nothing
   
End Sub

Private Sub M_BalClasifANeg_Click()
   Dim Frm As FrmBalClasifDesglo
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalClasifDesglo
   Frm.FViewBalClasifDesglo ("AREANEG")
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_BalClasifCCosto_Click()
   Dim Frm As FrmBalClasifDesglo
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalClasifDesglo
   Frm.FViewBalClasifDesglo ("CCOSTO")
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_BalClasifComp_Click()
   Dim Frm As FrmBalClasifCompar
   
   If Not gEmpresa.TieneAnoAnt Then
      MsgBox1 "Esta empresa no tiene año anterior en el sistema. No se puede generar el reporte.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalClasifCompar
   Frm.FViewBalClasif
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_BalClasifEjec_Click()
   Dim Frm As FrmBalClasifEjec
   Dim FrmIFRS As FrmBalEjecIFRS
   
   Me.MousePointer = vbHourglass

   If gPlanCuentas = "IFRS" Then
      Set FrmIFRS = New FrmBalEjecIFRS
      FrmIFRS.FView
      Set FrmIFRS = Nothing
   Else
      Set Frm = New FrmBalClasifEjec
      Frm.FViewBalClasif
      Set Frm = Nothing
   End If
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_CalcPropIVA_Click()
   Dim Frm As FrmPropIVA
   
   Set Frm = New FrmPropIVA
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_CapitalPropioSimplDet_Click(Index As Integer)
   Dim Frm As Form
   
   'Call MsgBox1("Este reporte aún no está disponible.", vbExclamation + vbOKOnly)
   
   If gEmpresa.TipoContrib = 0 Then
      MsgBox1 "Debe seleccionar el Tipo de Contribuyente en la ventana Empresa antes de continuar.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Select Case Index
   
      Case 1
         Set Frm = New FrmCapPropioSimpl
         Call Frm.FView(CPS_TIPOINFO_GENERAL)
         Set Frm = Nothing
      
      Case 2
         Set Frm = New FrmCapPropioSimpl
         Call Frm.FView(CPS_TIPOINFO_VARANUAL)
         Set Frm = Nothing
         
   End Select
   
   Me.MousePointer = vbDefault
End Sub

Private Sub M_CapitalPropioTrib_Click()
   Dim Frm As FrmCapitalPropio
   
   'Call MsgBox1("Este reporte aún no está disponible.", vbExclamation + vbOKOnly)
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmCapitalPropio
   Call Frm.FView
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_ConfigActFijo_Click()
   Dim Frm As FrmConfigActFijoIFRS
   
   Set Frm = New FrmConfigActFijoIFRS
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ConfigCtasAjustes_Click()
   Dim Frm As FrmConfigCtasAjustes
   
   Call MsgLey21210
   
   Set Frm = New FrmConfigCtasAjustes
   Frm.Show vbModal
   Set Frm = Nothing
End Sub


Private Sub M_ConfigCtasAjustesRLI_Click()
   Dim Frm As FrmConfigCtasAjustesRLI
   
   Set Frm = New FrmConfigCtasAjustesRLI
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_ConfigIFRS_Click()
   Dim Frm As FrmConfigCodIFRS

   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmConfigCodIFRS
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_ConfigRemu_Click()
   Dim Frm As FrmConfigRemu

   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmConfigRemu
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault


End Sub

Private Sub M_ConfigPrtCheque_Click()
   Dim Frm As FrmConfigCheque
   
   Set Frm = New FrmConfigCheque
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_EditCompApFin_Click()
   Dim Frm As FrmComprobante
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdComp As Long
   
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_FINANCIERO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      IdComp = Rs(0)
      
      MsgBox1 "Recuerde que para modificar el Comprobante, este debe llevarse a estado Pendiente.", vbInformation
      
      Set Frm = New FrmComprobante
      Call Frm.FEdit(IdComp, False)
      Set Frm = Nothing
      
   Else     'esto no debiera ocurrir nunca
   
      MsgBox1 "No existe comprobante de apertura", vbExclamation
      
   End If
   
   Call CloseRs(Rs)
   
End Sub
Private Sub M_EditCompApTrib_Click()
   Dim Frm As FrmComprobante
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdComp As Long
   
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      IdComp = Rs(0)
      
      MsgBox1 "Recuerde que para modificar el Comprobante, este debe llevarse a estado Pendiente.", vbInformation
      
      Set Frm = New FrmComprobante
      Call Frm.FEdit(IdComp, False)
      Set Frm = Nothing
      
   Else     'esto no debiera ocurrir nunca
   
      MsgBox1 "No existe comprobante de apertura", vbExclamation
      
   End If
   
   Call CloseRs(Rs)
   
End Sub
Private Sub M_EditLibCaja_Click()
   Dim Frm As FrmSelLibCaja

   Call MsgLey21210

   If gEmpresa.Franq14Ter = 0 And gEmpresa.Ano < 2020 Then
      MsgBox1 "Empresa no acogida a Franquicia Artículo 14 TER, no lleva Libro de Caja.", vbInformation
      Exit Sub
   End If

   Set Frm = New FrmSelLibCaja
   Call Frm.FSelectOper
   Set Frm = Nothing

End Sub

Private Sub M_ExpEntidades_Click()
   Dim Frm As FrmExpEntidades
   
   Set Frm = New FrmExpEntidades
   Frm.Show vbModal
   Set Frm = Nothing
  
End Sub

Private Sub M_ExpF22_Click()
   Dim Frm As FrmExpF22
               
   Set Frm = New FrmExpF22
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ExpFacturacion_Click()
   Dim Frm As FrmLibElectCompras
   
   Set Frm = New FrmLibElectCompras
   Frm.FGenLibComprasAcepta
   Set Frm = Nothing
 
End Sub


Private Sub M_ExpHR_RAB_Click()
   Dim Frm As FrmExpDJAnual
   
   Set Frm = New FrmExpDJAnual
   Frm.ExpHRRAB
   Set Frm = Nothing

End Sub

Private Sub M_ExpHR_RAB_RLI_Click()
   Dim Frm As FrmAjustesExtraLibCajaRLI
   
   Set Frm = New FrmAjustesExtraLibCajaRLI
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ExpHR_RAD_Click()
   Dim Frm As FrmExpDJAnual
      
   Set Frm = New FrmExpDJAnual
   Frm.ExpHRRAD
   Set Frm = Nothing

End Sub

Private Sub M_ExpHR_RADPERC_Click()
Dim Frm As FrmExpPercepciones

   Set Frm = New FrmExpPercepciones
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ExpHR_RetirosDividendos_Click()
   Dim Frm As FrmExpDJAnual
      
   Set Frm = New FrmExpDJAnual
   Frm.ExpRetirosDividendos
   Set Frm = Nothing

End Sub

Private Sub M_ExpHRDJ1847_Click()
   Dim Frm As FrmExpDJAnual
   
   If gEmpresa.Ano < 2017 Then
      MsgBox1 "Esta Declaración Jurada está habilitada a partir del año 2017 (AT 2018)", vbOK + vbInformation
      Exit Sub
   End If
   
   Set Frm = New FrmExpDJAnual
   Frm.Exp1847
   Set Frm = Nothing
   
End Sub

Private Sub M_ExpHRDJ1879_Click()
   Dim Frm As FrmExpHRCertif
   
'   If gDbType = SQL_ACCESS Then
   
      Set Frm = New FrmExpHRCertif
      Frm.Show vbModal
      Set Frm = Nothing
      
'   Else
'      MsgBox1 "Esta funcionalidad aún no está disponible para la versión SQL Server.", vbInformation
      
'   End If


End Sub
Private Sub M_ExpHRDJ1923_Click()
   Dim Frm As FrmExpDJAnual
   
   If gEmpresa.Ano < 2017 Then
      MsgBox1 "Esta Declaración Jurada está habilitada a partir del año 2017 (AT 2018)", vbOK + vbInformation
      Exit Sub
   End If
      
   
   Set Frm = New FrmExpDJAnual
   Frm.Exp1923
   Set Frm = Nothing

End Sub

Private Sub M_ExpHRDJ1924B_Click()
   Dim Frm As FrmExpDJAnual
   
   If gEmpresa.Ano < 2017 Then
      MsgBox1 "Esta Declaración Jurada está habilitada a partir del año 2017 (AT 2018)", vbOK + vbInformation
      Exit Sub
   End If
   
   If gEmpresa.Ano >= 2020 Then
      MsgBox1 "Esta Declaración Jurada no está habilitada a partir del año 2020 (AT 2021)", vbOK + vbInformation
      Exit Sub
   End If
   
   Set Frm = New FrmExpDJAnual
   Frm.Exp1924B
   Set Frm = Nothing

End Sub
Private Sub M_ExpHRDJ1924C_Click()
   Dim Frm As FrmExpDJAnual
   
   If gEmpresa.Ano < 2017 Then
      MsgBox1 "Esta Declaración Jurada está habilitada a partir del año 2017 (AT 2018)", vbOK + vbInformation
      Exit Sub
   End If
   
   If gEmpresa.Ano >= 2020 Then
      MsgBox1 "Esta Declaración Jurada no está habilitada a partir del año 2020 (AT 2021)", vbOK + vbInformation
      Exit Sub
   End If
   
   Set Frm = New FrmExpDJAnual
   Frm.Exp1924C
   Set Frm = Nothing

End Sub

Private Sub M_ExpLibAux_Click()
   Dim Frm As FrmImpExpLib
      
      
   If gDbType = SQL_ACCESS Then
      Set Frm = New FrmImpExpLib
      Call Frm.FExport
      Set Frm = Nothing
      
   Else
      MsgBox1 "Esta funcionalidad sólo está habilitada para versión Access", vbExclamation
      
   End If

End Sub

'Private Sub M_ExpSII_Click()
'   Dim Frm As FrmExpLibSII
'
'   Set Frm = New FrmExpLibSII
'   Frm.Show vbModal
'   Set Frm = Nothing
'
'End Sub

Private Sub M_FormToolsInt_Click()
   Dim Frm As FrmIntTools
   
   Set Frm = New FrmIntTools
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_IFRS_BalEjec_Click()
   Dim Frm As FrmBalEjecIFRS
   
   Set Frm = New FrmBalEjecIFRS
   Call Frm.FView
   Set Frm = Nothing

End Sub

Private Sub M_IFRS_BalTrib_Click()
   Dim Frm As FrmBalTributarioIFRS
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalTributarioIFRS
   Call Frm.FView(0)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_IFRS_EstFin_Click()
   Dim Frm As FrmLstInformeIFRS
   
   Set Frm = New FrmLstInformeIFRS
   Call Frm.FView(IFRS_ESTFIN)
   Set Frm = Nothing
   
End Sub

Private Sub M_IFRS_EstRes_Click()
   Dim Frm As FrmLstInformeIFRS
   
   Set Frm = New FrmLstInformeIFRS
   Call Frm.FView(IFRS_ESTRES)
   Set Frm = Nothing

End Sub

Private Sub M_ImpComprobantes_Click()
   Dim Frm As FrmImpComp
   
   If ValidaIngresoComp() Then
      
      Set Frm = New FrmImpComp
      Frm.Show vbModal
      Set Frm = Nothing
      
   End If

End Sub

Private Sub M_ImpFacturacion_Click()
   Dim Frm As FrmImpFacturacion
   
   Set Frm = New FrmImpFacturacion
   Frm.Show vbModal
   Set Frm = Nothing
   
   
End Sub

Private Sub M_ImpLibAux_Click()
   Dim Frm As FrmImpExpLib
      
   If gDbType = SQL_ACCESS Then
      Set Frm = New FrmImpExpLib
      Call Frm.FImport
      Set Frm = Nothing
      
   Else
      MsgBox1 "Esta funcionalidad sólo está habilitada para versión Access", vbExclamation
      
   End If


End Sub

Private Sub M_ImpOtrosDocs_Click()
'   Dim Frm As FrmImpOtrosDocs
'
'   Set Frm = New FrmImpOtrosDocs
'   Frm.Show vbModal
'   Set Frm = Nothing
   
   Dim Frm As FrmSelImpoOD

   Set Frm = New FrmSelImpoOD
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ImpRemu_Click()
   Dim Frm As FrmImportRemu
   
'   If gDbType = SQL_SERVER Then
'      MsgBox1 "Esta opción aún no está disponible para versión SQL Server", vbInformation
'      Exit Sub
'   End If
      
   Set Frm = New FrmImportRemu
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_InfActFijo_Click()
   Dim Frm As FrmRepActivoFijo
   
   If gMaxCred33 < 0 Then    'el usuario no ha ingresado el Max Cred 33 bis
      If gMaxUTMCred33_Pesos = 0 Then
         If MsgBox1("No se ha ingresado el valor de la UTM. Este valor se utiliza para calcular el máximo para Crédito Art. 33 bis", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      Else
         If MsgBox1("Revise si el último valor de la UTM y del IPC ingresados en el sistema están actualizados.", vbInformation + vbOKCancel) = vbCancel Then
            Exit Sub
         End If
         If MsgBox1("Verifique la correcta aplicación del porcentaje del Crédito por Activo Fijo, según instrucciones del Artículo 33 bis Ley de Renta." & vbCrLf & vbCrLf & "Para esto, ingrese a la Configuración Inicial, botón Configurar Impuestos (Menú Configuración).", vbInformation + vbOKCancel) = vbCancel Then
            Exit Sub
         End If
      End If
   End If
   
   Set Frm = New FrmRepActivoFijo
   Me.MousePointer = vbHourglass
   Call Frm.FView
   Me.MousePointer = vbDefault
   Set Frm = Nothing

End Sub

Private Sub M_AreasNeg_Click()
   Dim Frm As FrmAreaNeg
   Dim AreaNeg As AreaNeg_t
   
   Set Frm = New FrmAreaNeg
   Call Frm.FEdit
   Set Frm = Nothing
   
End Sub


Private Sub M_BalClasif_Click()
   Dim Frm As FrmBalClasif
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalClasif
   Frm.FViewBalClasif
   Set Frm = Nothing
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_BalComprob_Click()
   Dim Frm As FrmBalComprobacion
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalComprobacion
   Call Frm.FView(0)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_BalTrib_Click()
   Dim Frm As FrmBalTributario

   Me.MousePointer = vbHourglass
   Set Frm = New FrmBalTributario
   Frm.FView (0)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_CambiarClave_Click()
   Dim Frm As FrmCambioClave
   
   Set Frm = New FrmCambioClave
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_CentrosGestion_Click()
   Dim Frm As FrmCentrosCosto
   
   Set Frm = New FrmCentrosCosto
   Call Frm.FEdit
   Set Frm = Nothing
   
End Sub

Private Sub M_AbrirCerrarMes_Click()
   Dim Frm As FrmEstadoMeses
   Dim MesActual As Integer
   
   Set Frm = New FrmEstadoMeses
   
   Frm.Show vbModal
   
   Set Frm = Nothing
   
   MesActual = GetMesActual()
   
   If MesActual > 0 And MesActual <= 12 Then
      Lb_Mes = Left(gNomMes(MesActual), 3)
   Else
      Lb_Mes = ""
   End If
      
End Sub

Private Sub M_CerrarPer_Click()
   Dim Frm As FrmCierreAnual
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Esta período ya ha sido cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   'por si acaso, recalculamos saldos para que no haya problemas al llevar
   'los docs al añoi siguiente
   Me.MousePointer = vbHourglass
   DoEvents
   
   ' 15 feb 2020
   If MsgBox1("¿Desea recalcular los saldos de TODOS los documentos?" & vbCrLf & vbCrLf & "Si elige NO, sólo se recalculan los saldos de los documentos modificados.", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
      Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") ")
   End If
   
   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano, False)
   Call RecalcSaldosFulle(gEmpresa.id, gEmpresa.Ano, False)
   Me.MousePointer = vbDefault
      
   Set Frm = New FrmCierreAnual
   Frm.Show vbModal
   Set Frm = Nothing
   
   Lb_Cierre.visible = gEmpresa.FCierre <> 0
End Sub

Private Sub M_Compactar_Click()
#If DATACON = 1 Then       'Access

   Dim ConnStr As String

   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   
   'ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
   If CompactDb2(DbMain, True, gEmpresa.ConnStr) = 0 Then 'no hubo error
      Call IniEmpresa
   Else
      MsgBox1 "Problemas al tratar de compactar la base de datos.", vbExclamation + vbOKOnly
   End If
   
   Me.MousePointer = vbDefault
   
#End If

End Sub

Private Sub M_Config_Click()
   Dim Frm As FrmConfig
   
   Set Frm = New FrmConfig
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub


Private Sub M_ConfigCtasFUT_Click()
'   Dim Frm As FrmConfigFUT
'
'   Me.MousePointer = vbHourglass
'
'   Set Frm = New FrmConfigFUT
'   Frm.Show vbModal
'   Set Frm = Nothing
'
'   Me.MousePointer = vbDefault

End Sub

Private Sub M_ConfigHojasTimbraje_Click()
   Dim Frm As FrmFoliacion

   Set Frm = New FrmFoliacion
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_ConfigMonedas_Click()
   Dim Frm As FrmMonedas
   
   Set Frm = New FrmMonedas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub


Private Sub M_EditComprob_Click()
   Dim Frm As FrmLstComp
   
   Set Frm = New FrmLstComp
   Call Frm.FView
   Set Frm = Nothing
   
End Sub

Private Sub M_ContEmpresa_Click()
   Dim Frm As FrmContEmpresa
      
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmContEmpresa
   FrmContEmpresa.Show vbModal
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_DefRazones_Click()
   Dim Frm As FrmRazones
      
   Set Frm = New FrmRazones
   Call Frm.FConfigParam
   Set Frm = Nothing
   
End Sub

Private Sub M_EditEmp_Click()
   Dim Frm As FrmEmpresa
   
   Set Frm = New FrmEmpresa
   MousePointer = vbHourglass
   
   If Frm.FEdit(gEmpresa.id) = vbOK Then
      Lb_Dir = gEmpresa.Direccion
      Lb_Tel = gEmpresa.Telefono
   End If
   
   MousePointer = vbDefault
   Set Frm = Nothing
   
   Call EnableCertif

   
End Sub

Private Sub M_EntRel_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmEntidades
   Call Frm.FEdit
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_Equivalencias_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Frm.FEdit (0)
   Set Frm = Nothing
   
End Sub

Private Sub M_ExpF29_Click()
   Dim Frm As FrmExpF29
   
'   If gDbType = SQL_ACCESS Then
   
'      If gLinkF22 = False Then
'         MsgBox1 "No se encontraron los archivos correspondientes al producto HR-IVA Estándar en " & vbLf & W.AppPath & "\..\PAR", vbExclamation
'         Exit Sub
'      End If
      
      Set Frm = New FrmExpF29
      Frm.Show vbModal
      Set Frm = Nothing
      
'   Else
'      MsgBox1 "Esta funcionalidad aún no está disponible para la versión SQL Server.", vbInformation
'
'   End If
   
   
End Sub

'Private Sub M_ExpFUT_Click()
'   Dim Frm As FrmExpFUT
   
'   MsgBox1 "Esta opción está actualmente en desarrollo.", vbInformation
'   Exit Sub
   
'   Set Frm = New FrmExpFUT
'   Frm.Show vbModal
'   Set Frm = Nothing

'End Sub

Private Sub M_ImpF29Av_Click()
   Dim Frm As FrmImportF29
      
   If gDbType = SQL_ACCESS Then
      Set Frm = New FrmImportF29
      Frm.Show vbModal
      Set Frm = Nothing
      
   Else
      MsgBox1 "Esta funcionalidad sólo está habilitada para versión Access", vbExclamation
      
   End If
   
End Sub

Private Sub M_Indices_Click()
   Dim Frm As FrmIPC
   
   Set Frm = New FrmIPC
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_InfActFijoIFRS_Click()

End Sub

Private Sub M_InfAnalitCtaODF_Click()
   Dim Frm As FrmInfAnaliticoFulle
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfAnaliticoFulle
   Call Frm.FViewPorCuenta(0)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
End Sub

Private Sub M_InfAnalitEnt_Click()
   Dim Frm As FrmInfAnalitico
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfAnalitico
   Call Frm.FViewPorEntidad(0)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub
Private Sub M_InfAnalitCta_Click()
   Dim Frm As FrmInfAnalitico
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfAnalitico
   Call Frm.FViewPorCuenta(0)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_InfAnalitEntODF_Click()
   Dim Frm As FrmInfAnaliticoFulle
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfAnaliticoFulle
   Call Frm.FViewPorEntidad(0)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
End Sub

Private Sub M_InfoCartBanc_Click()
   Dim Frm As FrmResCartolas
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResCartolas
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_InfoConciliacion_Click()
   Dim Frm As FrmInfConciliacion
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfConciliacion
   Frm.Show vbModeless
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_InfoResConcil_Click()
   Dim Frm As FrmResInfConcil
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResInfConcil
   Frm.Show vbModeless
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub


Private Sub M_InfoVenc30_Click()
   Dim Frm As FrmInfoVencim
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmInfoVencim
   Frm.FView (30)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_InfoVenc60_Click()
   Dim Frm As FrmInfoVencim
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmInfoVencim
   Frm.FView (60)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_InfoVenc90_Click()
   Dim Frm As FrmInfoVencim
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmInfoVencim
   Frm.FView (90)
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_IngCartola_Click()
   Dim Frm As FrmImpCartola
   
   Set Frm = New FrmImpCartola
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_LibCompSII_Click()
   Dim Frm As FrmImpLibComprasSII

   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If ValidaIngresoDoc() = False Then
      Exit Sub
   End If
   
   Set Frm = New FrmImpLibComprasSII
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_LibElectCompras_Click()
   Dim Frm As FrmLibElectCompras
   
   Set Frm = New FrmLibElectCompras
   Frm.FGenLibComprasSII
   Set Frm = Nothing
 
End Sub

Private Sub M_AuditLibrosContables_Click()
   Dim Frm As FrmAuditLibContables
   
   Set Frm = New FrmAuditLibContables
   Call Frm.FView
   Set Frm = Nothing

End Sub

Private Sub M_LibRetenSII_Click()
   Dim Frm As FrmImpLibRetencionesSII

   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   If ValidaIngresoDoc() = False Then
      Exit Sub
   End If

   Set Frm = New FrmImpLibRetencionesSII
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub M_LibroIngEg_Click()
   Dim Frm As FrmLibIngEg

   Call MsgLey21210

   If gEmpresa.Ano < 2020 Then
      If gEmpresa.Franq14Ter = 0 Then
         MsgBox1 "Empresa no acogida a Franquicia Artículo 14 TER, no lleva Libro de Caja ni Libro de Ingresos y Egresos.", vbInformation
         Exit Sub
      ElseIf gEmpresa.ObligaLibComprasVentas Then
         MsgBox1 "Empresa acogida a Franquicia Artículo 14 TER y obligada a llevar Libro de Compras y Ventas según la Ley de IVA, no lleva Libro de Ingresos y Egresos.", vbInformation
         Exit Sub
      End If
   End If

   Set Frm = New FrmLibIngEg
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub M_Libros_Click()
   Dim Frm As FrmSelLibros
   
   Set Frm = New FrmSelLibros
   Call Frm.FSelectMes
   Set Frm = Nothing

End Sub

Private Sub M_LibVentasSII_Click()
   Dim Frm As FrmImpLibVentasSII

   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If ValidaIngresoDoc() = False Then
      Exit Sub
   End If
   
   Set Frm = New FrmImpLibVentasSII
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_ListComprob_Click()
   Dim Frm As FrmLstComp
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstComp
   Call Frm.FPrint
   Set Frm = Nothing

   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_LstChequesaFecha_Click()
   Dim Frm As FrmLstDoc
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstDoc
   Call Frm.FView(LIB_OTROS, FindTipoDoc(LIB_OTROS, "CHF"), 0, month(Now), Year(Now))
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_LstChequesAnula_Click()
   Dim Frm As FrmLstDoc
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstDoc
   Call Frm.FView(LIB_OTROS, FindTipoDoc(LIB_OTROS, "CHE"), ED_ANULADO, month(Now), Year(Now))
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub
Private Sub M_LstChequesEmit_Click()
   Dim Frm As FrmLstDoc
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstDoc
   Call Frm.FView(LIB_OTROS, FindTipoDoc(LIB_OTROS, "CHE"), ED_PENDIENTE, month(Now), Year(Now))
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_LstCompTipo_Click()
   Dim Frm As FrmLstCompTipo
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstCompTipo
   Frm.Show vbModal
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_LstCuentas_Click()
   Dim Frm As FrmLstPlanCuentas
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstPlanCuentas
   Frm.Show vbModal
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_LstDocs_Click()
   Dim Frm As Form
   Dim TipoLib As Integer
   Dim Mes As Integer
   Dim Año As Integer
   
   Set Frm = New FrmSelLibDocs
   
   Call Frm.FSelect(TipoLib, True)
   
   Set Frm = Nothing
   
End Sub

Private Sub M_LstInfoImp_Click()
   Dim Frm As FrmLstLibImpresos
    
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstLibImpresos
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_ManActFijo_Click()
   Dim Rc As Long
   Dim Buf As String
   
'   MsgBox1 "Actualmente el manual de Activo Fijo se encuentra en proceso de actualización.", vbInformation
   
   MousePointer = vbHourglass
   DoEvents

   Buf = gAppPath & "\Manual_Modulo_Activo_Fijo_2020.pdf"
   Rc = ExistFile(Buf)

   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual del Módulo de Activo Fijo, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual del Módulo de Activo Fijo." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub M_ManConciliacion_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_Conciliacion_Bancaria.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual de Conciliación Bancaria, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual de Conciliación Bancaria." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub M_MANT_PERCE_Click()
Dim Frm As FrmMantPercepciones

   Set Frm = New FrmMantPercepciones
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub M_NewComprob_Click()
   Dim Frm As FrmComprobante
    
   If ValidaIngresoComp() Then
      
      Me.MousePointer = vbHourglass
      
      Set Frm = New FrmComprobante
      Call Frm.FNew(False)
      Set Frm = Nothing
      
      Me.MousePointer = vbDefault
      
   End If
   
End Sub

Private Sub M_NewDoc_Click()
   Dim Frm As Form
   Dim TipoLib As Integer
   Dim Mes As Integer
   Dim Año As Integer
   Dim IdDoc As Long
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If ValidaIngresoDoc() = False Then
      Exit Sub
   End If
      
   Set Frm = New FrmSelLibDocs
   Call Frm.FSelectMes(TipoLib, Mes, Año, False)
   
'   If Frm.FSelectMes(TipoLib, Mes, Año, False) = vbOK Then
'      Set Frm = Nothing
'
'      If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Then
'
'         If gCtasBas.IdCtaIVACred <= 0 Or gCtasBas.IdCtaIVADeb <= 0 Then
'            MsgBox1 "No es posible ingresar documentos a los Libros de Compras y Ventas sin antes definir la configuración de las cuentas de IVA y Otros Impuestos." & vbNewLine & vbNewLine & "Utilice el botón ""Definir Cuentas Básicas"" provisto en el menú ""Configuración Inicial"".", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'
'         Me.MousePointer = vbHourglass
'
'         Set Frm = New FrmCompraVenta
'         Call Frm.FEdit(TipoLib, Mes, Año, IdDoc)
'
'         Me.MousePointer = vbDefault
'
'      ElseIf TipoLib = LIB_RETEN Then
'
'         If gCtasBas.IdCtaImpRet <= 0 Or gCtasBas.IdCtaNetoHon <= 0 Then
'            MsgBox1 "No es posible ingresar documentos al Libro de Retenciones sin antes definir la configuración de las cuentas de Impuesto Retenido y Neto Retención." & vbNewLine & vbNewLine & "Utilice el botón ""Definir Cuentas Básicas"" provisto en el menú ""Configuración Inicial"".", vbExclamation + vbOKOnly
'            Exit Sub
'         End If
'
'         Me.MousePointer = vbHourglass
'
'         Set Frm = New FrmLibRetenciones
'         Call Frm.FEdit(Mes, Año, IdDoc)
'
'         Me.MousePointer = vbDefault
'
'      Else
'         Me.MousePointer = vbHourglass
'
'         Set Frm = New FrmLstDoc
'         Call Frm.FEdit(TipoLib, Mes, Año, True)
'
'         Me.MousePointer = vbDefault
'
'      End If
'   End If
   
   Set Frm = Nothing
End Sub


' pam: Nueva Intancia
Private Sub M_NuevaInstancia_Click()
   Dim Key As Long

   If MsgBox1("¿ Desea ejecutar otra instancia de esta aplicación ?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
      Exit Sub
   End If
   
   Key = GenInstanceKey()

   Call ShellExecute(Me.hWnd, "open", W.AppPath & "\" & App.EXEName, " /i=" & Key, W.AppPath, SW_SHOW)

End Sub

Private Sub M_Plan_Click()
   Dim Frm As FrmPlanCuentas
   
   MousePointer = vbHourglass
   Set Frm = New FrmPlanCuentas
   Call Frm.FEdit
   Set Frm = Nothing
   MousePointer = vbDefault
   
End Sub

Private Sub M_ProcConcil_Click()
   Dim Frm As FrmConciliacion
   
   Me.MousePointer = vbHourglass
   Set Frm = New FrmConciliacion
   Call Frm.FEdit
   Set Frm = Nothing
   Me.MousePointer = vbDefault

End Sub

Private Sub M_PrtComprob_Click()
   Dim Frm As FrmLstComp
      
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstComp
   Call Frm.FPrint
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub


Private Sub M_PrtHojasTimb_Click()
   Dim Frm As FrmPrtFoliacion
   
   Set Frm = New FrmPrtFoliacion
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_CalcRazFin_Click()
   Dim Frm As FrmCalcRazones
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmCalcRazones
   Frm.Show vbModal
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub M_ReabrirPer_Click()
   Dim Q1 As String
   
   If gEmpresa.FCierre = 0 Then
      MsgBox1 "Este año ya está abierto.", vbExclamation
      Exit Sub
   End If
      
   If MsgBox1("¿Está seguro que desea volver a abrir el año " & gEmpresa.Ano & "?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   Q1 = "UPDATE EmpresasAno SET FCierre=0"
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
      
   Call ExecSQL(DbMain, Q1)
   
   gEmpresa.FCierre = 0
    
   Lb_Cierre.visible = gEmpresa.FCierre <> 0
   
End Sub

Private Sub M_RecalcSaldos_Click()
   Dim WhLib As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim n As Long
   
   
   WhLib = " Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ") "

   Set Rs = OpenRs(DbMain, "SELECT Count(*) FROM Documento WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Not Rs.EOF Then
      n = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
   If MsgBox1("Esta operación recalcula los saldos de TODOS los documentos (" & n & " documentos)." & vbNewLine & vbNewLine & "Puede tomar un poco de tiempo ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   'asignamos SaldoDoc = NULL para los docs de compras, ventas y retenciones para que los recalcule TOTOS
   Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
   'asignamos SaldoDoc = Total para Otros Documentos, por el error que se asignaron todos en NULL
'   WhLib = " Documento.TipoLib =" & LIB_OTROS
'   Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = Total WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
   Call RecalcSaldosFulle(gEmpresa.id, gEmpresa.Ano)
    
   Me.MousePointer = vbDefault
   
   MsgBox1 "Cálculo de saldos vigentes finalizado.", vbInformation
   
End Sub

Private Sub M_ReimportActFijo_Click()
   Dim Frm As FrmImpActFijos
   
   Set Frm = New FrmImpActFijos
   Frm.Show vbModal
   Set Frm = Nothing
   

End Sub
Private Sub M_RelinkearTblBasicas_Click()

   If MsgBox1("ADVERTENCIA: Para ejecutar esta operación nadie más tiene que estar trabajando en el sistema." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo) = vbNo Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   DoEvents

   Call LinkMdbAdm(True)
   
   MsgBox1 "Proceso finalizado.", vbInformation
   MousePointer = vbDefault
   
End Sub

Private Sub M_Renum_Click()
   Dim Frm As FrmRenum
   
   Set Frm = New FrmRenum
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_RepActFijoIFRS_Click()
   Dim Frm As FrmRepActFijoIFRS
   
   Set Frm = New FrmRepActFijoIFRS
   Call Frm.FView
   Set Frm = Nothing
End Sub

Private Sub M_Reparar_Click()
#If DATACON = 1 Then       'Access

   Dim DbPath As String
   
   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   DbPath = DbMain.Name
   
   Call CloseDb(DbMain)
   
   If RepairDb(DbPath) Then
      Call IniEmpresa
      Me.MousePointer = vbDefault
   Else
      Unload Me
      End
   End If
#End If

End Sub

Private Sub M_RepPagoPlazo_Click()
   Dim Frm As FrmLstDocCuotas
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmLstDocCuotas
   Call Frm.FView
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

   
End Sub

Private Sub M_ResClasificado_Click()
   Dim Frm As FrmBalClasif
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasif
   Call Frm.FViewEstResultClasif
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub


Private Sub M_ResClasificadoANeg_Click()
   Dim Frm As FrmBalClasifDesglo
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasifDesglo
   Call Frm.FViewEstResultClasifDesglo("AREANEG")
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_ResClasificadoCCosto_Click()
   Dim Frm As FrmBalClasifDesglo
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasifDesglo
   Call Frm.FViewEstResultClasifDesglo("CCOSTO")
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault


End Sub

Private Sub M_ResComparativo_Click()
   Dim Frm As FrmBalClasif
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasif
   Call Frm.FViewEstResultComparativo
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_ResCompPeriodo_Click()
   Dim Frm As FrmBalClasifCompar
   
   If Not gEmpresa.TieneAnoAnt Then
      MsgBox1 "Esta empresa no tiene año anterior en el sistema. No se puede generar el reporte.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasifCompar
   Call Frm.FViewEstResultClasif
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_ResLibAux_Click()
   Dim Frm As FrmResLibAux
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResLibAux
   Call Frm.FView
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub M_ResMensual_Click()
   Dim Frm As FrmBalClasif
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmBalClasif
   Call Frm.FViewEstResultMensual
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub


Private Sub M_ResSupermercado_Click()
   Dim Frm As FrmResDocs
   
   Set Frm = New FrmResDocs
   Call Frm.FViewSupermercado
   Set Frm = Nothing

End Sub

Private Sub M_ResVPE_Click()
   Dim Frm As FrmResDocs
   
   Set Frm = New FrmResDocs
   Call Frm.FViewVPE
   Set Frm = Nothing
   
End Sub

Private Sub M_RevDetDoc_Click()
   Dim Frm As FrmRevDetDocs
   
   Set Frm = New FrmRevDetDocs
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_Salir_Click()
   Unload Me
End Sub


Private Sub M_SelEmp_Click()
   Dim Rut As String
   Dim Frm As FrmSelEmpresas
   Dim Rc As Integer
   Dim BoolIniEmpresa As Boolean
#If DATACON = 1 Then
   Dim DbMainOld As Database
#End If
   Dim gCurEmp As Empresa_t
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PlanVacio As Boolean
   Dim FrmConfig As FrmConfig
   
   
   BoolIniEmpresa = False
      
   gCurEmp = gEmpresa
      
   Do While BoolIniEmpresa = False
      Set Frm = New FrmSelEmpresas
      Rc = Frm.FSelect
      Set Frm = Nothing
      
      If Rc = vbOK Then
      
#If DATACON = 1 Then
         If gEmprSeparadas Then
            Set DbMainOld = DbMain     'db de la empresa actual
            Set DbMain = Nothing       'para que no la cierre
            
            'pipe2
            'Call CloseDb(DbMain)
            
         End If
#End If

         BoolIniEmpresa = IniEmpresa()
         
#If DATACON = 1 Then
         If gEmprSeparadas Then
            If BoolIniEmpresa = False Then   'falló, dejamos la Db actual
               Set DbMain = DbMainOld
               gEmpresa = gCurEmp
            Else
               Call CloseDb(DbMainOld)       'abrió otra db, cerramos la anterior
            End If
         End If
#End If
      
      Else
         BoolIniEmpresa = True
      End If
      
   Loop
   
   If Rc = vbOK Then
      Call FillDatosEmp
      Call SetPrtData
   End If
   
   Q1 = "SELECT IdCuenta FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   PlanVacio = (Rs.EOF = True)
   Call CloseRs(Rs)
   
   If PlanVacio Then    'no hay cuentas definidas
      
      'Suponemos que es el primer año de la empresa y es la primera vez que se abre
      If GetDatosEmpHR(gEmpresa.Rut) Then
         Dim FrmEmp As FrmEmpresa
         
         Set FrmEmp = New FrmEmpresa
         Rc = FrmEmp.FEdit(gEmpresa.id, True)
         Set FrmEmp = Nothing
         
         If Rc = vbOK Then
            Call FillDatosEmp
         End If
         
      End If
      
      MsgBox1 "Se ha detectado que no está definido el Plan de Cuentas para esta empresa." & vbNewLine & vbNewLine & "La ventana de Configuración Inicial le permitirá definir el Plan de Cuentas y otros elementos básicos para trabajar con el sistema.", vbInformation + vbOKOnly
      Set FrmConfig = New FrmConfig
      FrmConfig.Show vbModal
      Set FrmConfig = Nothing
      
      MsgBox1 "Recuerde configurar las Razones Financieras para esta empresa, " & vbCrLf & vbCrLf & "utilizando la opción que provee el sistema, bajo el menú 'Configuración'", vbOKOnly + vbInformation
      
   End If
  
   Call SetupPriv

#If DATACON = 1 Then

   If gEmpresa.Ano < 2020 Then
      Bt_Contrib14Ter.Caption = "Contribuyentes 14TER Let. A"
   Else
      Bt_Contrib14Ter.Caption = "Contrib. 14 D LIR ProPyme"
   End If

#Else

' Nueva Interfaz
'   If gEmpresa.Ano < 2020 Then
'      Pc_Boton(BT_14TER).Visible = True
'      Pc_Boton(BT_14D).Visible = False
'   Else
'      Pc_Boton(BT_14TER).Visible = False
'      Pc_Boton(BT_14D).Visible = True
'   End If
   
#End If



End Sub

Private Sub M_SetupPrt_Click()
   Dim CurrPrt As String
   Dim Rc As Integer
   
   If PrepararPrt(Cm_PrtDlg) Then
   
      Call SetIniString(gIniFile, "Config", "Printer", Printer.DeviceName)
   Else
      Call FindPrinter(GetIniString(gIniFile, "Config", "Printer"), True)
    
   End If
   
   'CurrPrt = Printer.DeviceName
   'Set Printer = FindPrinter(CurrPrt)

End Sub
Private Sub FillDatosEmp()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Msg14Quater As Integer

   'Primero Chequeo
   Lb_RUT = FmtCID(gEmpresa.Rut)
   Lb_Dir = gEmpresa.Direccion
   Lb_Tel = gEmpresa.Telefono
   Lb_Empresa = gEmpresa.NombreCorto
   Lb_Año = gEmpresa.Ano
   Lb_Mes = Left(gNomMes(GetMesActual), 3)
   
   Me.Caption = gEmpresa.NombreCorto & " - " & gEmpresa.Ano & " - " & gLexContab & " - " & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
   
   If gAppCode.Demo Then
      Me.Caption = Me.Caption & " - D" & "E" & "M" & "O"
   End If
   
   If gNuevaInstancia Then
      Me.Caption = Me.Caption & " [R]"
   End If
  
   Lb_Cierre.visible = gEmpresa.FCierre <> 0
   
   'Franquicia 14 Bis y 14 quarter
   If gEmpresa.Ano >= 2017 Then
   
      Msg14Quater = Val(GetIniString(gIniFile, "Msg", "14quater", "0"))
      
      If Msg14Quater = 0 Then
      
         Q1 = "SELECT Franq14bis, Franq14quater "
         Q1 = Q1 & " FROM Empresa"
         Q1 = Q1 & " WHERE Id = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Set Rs = OpenRs(DbMain, Q1)
      
         If vFld(Rs("Franq14bis")) <> 0 Or vFld(Rs("Franq14quater")) Then
            If MsgBox1("Recuerde que a contar de Enero 2017, los Regímenes establecidos en los arts. 14 bis y 14 quater se encuentran  derogados (Ley 20.780 de 2014)." & vbCrLf & vbCrLf & "Actualice la información en Empresa\Tipo de Contribuyente." & vbCrLf & vbCrLf & "(En caso de estar acogido a Contabilidad Completa a contar de enero 2017 ignore este mensaje)." & vbCrLf & vbCrLf & "¿Desea volver a ver este mensaje?", vbInformation + vbYesNo) = vbNo Then
               Call SetIniString(gIniFile, "Msg", "14quater", "1")
            End If
         End If
         Call CloseRs(Rs)
         
      End If
   End If
   
End Sub
Private Sub M_Sucursales_Click()
   Dim Frm As FrmSucursales
   
   Set Frm = New FrmSucursales
   Call Frm.FEdit
   Set Frm = Nothing

End Sub

Private Sub M_TipoDocs_Click()
   Dim Frm As FrmTipoDocs
   
   Set Frm = New FrmTipoDocs
   Frm.Show vbModal
   Set Frm = Nothing
   

End Sub

Private Sub M_TraerChequesAnAntOld_Click()

   If DB_MSSQL Then
      MsgBox1 "Por ahora esta opción no está habilitada para SQL Server", vbInformation
   End If
   
   Me.MousePointer = vbHourglass
   
   Call TraerOtrosDocsAprobados(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True)
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_TraspODToODF_Click()
Dim Frm As FrmTrapasoODToODF
   
   Set Frm = New FrmTrapasoODToODF
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub M_Unlock_Click()
   Dim Frm As FrmUnlock
   
   Set Frm = New FrmUnlock
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_ViewLibCaja_Click()
   Dim Frm As FrmLibCaja

   Call MsgLey21210

   If gEmpresa.Franq14Ter = 0 And gEmpresa.Ano < 2020 Then
      MsgBox1 "Empresa no acogida a Franquicia Artículo 14 TER, no lleva Libro de Caja.", vbInformation
      Exit Sub
   End If

   Set Frm = New FrmLibCaja
   Call Frm.FView
   Set Frm = Nothing

End Sub

Private Sub MH_About_Click()
   Dim Frm As FrmAbout
   
   Set Frm = New FrmAbout
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub
Private Sub SetupPriv()

   If Not ChkPriv(PRV_ING_COMP) Then
      M_NewComprob.Enabled = False
      M_NewComprob.Enabled = False
      M_ImpComprobantes.Enabled = False
   End If
   
   If Not ChkPriv(PRV_ADM_COMP) Then
      M_Renum.Enabled = False
      M_Auditoria.Enabled = False
   End If
      
   If Not ChkPriv(PRV_ING_DOCS) Then
      M_NewDoc.Enabled = False
      M_CalcPropIVA.Enabled = False
   End If
   
   If Not ChkPriv(PRV_VER_INFO) Then
      M_Libros.Enabled = False
      M_Balances.Enabled = False
      M_InfoAnalit.Enabled = False
      M_EstadoRes.Enabled = False
      M_CapitalPropio.Enabled = False
      M_ActFijo.Enabled = False
      Bt_Libros.Enabled = False
      Bt_Balances.Enabled = False
      Bt_InfAnalitico.Enabled = False
      Bt_Result.Enabled = False
      M_CalcRazFin.Enabled = False
      M_OtrosInformes.Enabled = False
      M_InfActFijo.Enabled = False
      M_RepActFijoIFRS.Enabled = False
      M_InfoIFRS.Enabled = False
      Bt_InfoIFRS.Enabled = True
   End If
   
   If Not ChkPriv(PRV_ADM_CONCIL) Then
      M_Concil.Enabled = False
   End If
   
   If Not ChkPriv(PRV_ADM_EMPRESA) Then
      M_CerrarPer.Enabled = False
      M_ReabrirPer.Enabled = False
      M_ExpF29.Enabled = False
      M_ExpF22.Enabled = False
      M_ImpF29Av.Enabled = False
      M_ExpHRCertif.Enabled = False
      M_ExpLibAux.Enabled = False
      M_ImpLibAux.Enabled = False
      M_ExpEntidades.Enabled = False
      M_ImpFacturacion.Enabled = False
      M_ExpFacturacion.Enabled = False
      M_ImpRemu.Enabled = False
   Else
      Call EnableCertif
   End If
   
   
   If Not ChkPriv(PRV_ADM_SIS) Then
      M_MantDB.Enabled = False
   End If
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      M_Config.Enabled = False
      M_ConfigCtasFUT.Enabled = False
      M_ConfigRemu.Enabled = False
      M_DefRazones.Enabled = False
      M_ConfigCtasAjustes.Enabled = False
      M_TipoDocs.Enabled = False
   End If
   
   If Not ChkPriv(PRV_ADM_ACTFIJOS) Then
      M_ConfigActFijo.Enabled = False
      M_ActFijo.Enabled = False
      Bt_MantActivoFijo.Enabled = False
      M_AFImportFile.Enabled = False
      M_ReimportActFijo.Enabled = False
   End If

      
End Sub
Private Function EnableCertif()

   If ChkPriv(PRV_ADM_EMPRESA) Then
      M_ExpHRDJ1879.Enabled = True
      M_ExpHRDJ1847.Enabled = True
      M_ExpHRDJ1923.Enabled = True
      M_ExpHRDJ1923.visible = True
      M_ExpHRDJ1924B.Enabled = True
      M_ExpHRDJ1924B.visible = True
      M_ExpHRDJ1924C.Enabled = True
      M_ExpHRDJ1924C.visible = True
      M_ExpF22.Enabled = True
      M_ExpHR_RAB.Enabled = True
      M_ExpHR_RAB_RLI.Enabled = True
      M_ExpHR_RAD.Enabled = True
      'M_ExpHR_RADPERC.Enabled = True
      M_ExpHR_RetirosDividendos.Enabled = True
      M_ConfigCtasAjustes.Enabled = gEmpresa.Franq14Ter Or gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp
      M_ConfigCtasAjustesRLI.Enabled = True
            
      If gEmpresa.Franq14Ter Then
         M_ExpHRDJ1847.Enabled = False
         M_ExpHRDJ1923.Enabled = False
      End If
   
      If gEmpresa.SemiIntegrado Then
         M_ExpHRDJ1923.Enabled = False
      End If
   
      If gEmpresa.RentaAtribuida Or gEmpresa.SemiIntegrado Then
         M_ExpHRDJ1924B.Enabled = False
         M_ExpHRDJ1924C.Enabled = False
      End If
      
      If gEmpresa.SocProfSegCat Then
         M_ExpHRDJ1847.Enabled = False
         M_ExpHRDJ1923.Enabled = False
         M_ExpHRDJ1924B.Enabled = False
         M_ExpHRDJ1924C.Enabled = False
      End If
      
      If gEmpresa.Ano >= 2020 Then
         M_ExpHRDJ1923.visible = False
         M_ExpHRDJ1924B.visible = False
         M_ExpHRDJ1924C.visible = False
      End If
      
      If gEmpresa.TipoContrib = CONTRIB_SAABIERTA Or gEmpresa.TipoContrib = CONTRIB_SACERRADA Or gEmpresa.TipoContrib = CONTRIB_SPORACCION Then
         M_ExpHR_RetirosDividendos.Caption = "Exportar archivo Dividendos..."
      Else
         M_ExpHR_RetirosDividendos.Caption = "Exportar archivo Retiros..."
      End If
      
      If gEmpresa.Ano >= 2020 Then
         M_ExpHRDJ1847.Enabled = False
         If gEmpresa.R14ASemiIntegrado Then
            M_ExpHRDJ1847.Enabled = True
         End If
      End If
            
      If gEmpresa.ProPymeGeneral Then
         M_ExpHRDJ1847.Enabled = False
         M_ExpHR_RAB.Enabled = False
      End If
            
      If gEmpresa.ProPymeTransp Or gEmpresa.RentasPresuntas Or gEmpresa.RentaEfectiva Or gEmpresa.NoSujetoArt14 Then
         M_ExpHRDJ1847.Enabled = False
         M_ExpHR_RAB.Enabled = False
         M_ExpHR_RAB_RLI.Enabled = False
      End If
            
      If Not (gEmpresa.ProPymeTransp Or gEmpresa.ProPymeGeneral) Then
         M_ExpHR_RAD.Enabled = False
         'M_ExpHR_RADPERC.Enabled = False
      End If
      
'      If Not (gEmpresa.ProPymeTransp Or gEmpresa.ProPymeGeneral Or gEmpresa.R14ASemiIntegrado) Then
'         M_ExpHR_RADPERC.Enabled = False
'      End If
         
      If gEmpresa.RegimenOtro Then
         M_ExpHRDJ1847.Enabled = False
      End If
            
      If gEmpresa.Ano >= 2020 Then
'         M_ExpF22.Enabled = False
'         M_ExpHR_RAB.Enabled = False
'         M_ExpHR_RAB_RLI.Enabled = False
'         M_ExpHR_RetirosDividendos.Enabled = False
'         M_ConfigCtasAjustes.Enabled = False
'         M_ConfigCtasAjustesRLI.Enabled = False
         
         M_ExpHR_RABbase.Caption = ReplaceStr(M_ExpHR_RABbase.Caption, "RAB", "RAD")
         M_ExpHR_RAB.Caption = ReplaceStr(M_ExpHR_RAB.Caption, "RAB", "RAD")
         M_ExpHR_RAB_RLI.Caption = ReplaceStr(M_ExpHR_RAB_RLI.Caption, "RAB", "RAD")
         M_ConfigCtasAjustes.Caption = "Configurar Cuentas Ajustes 14 D LIR..."
         M_ConfigCtasAjustesRLI.Caption = ReplaceStr(M_ConfigCtasAjustesRLI.Caption, "RAB", "RAD")
         
      Else
         M_ExpHR_RAD.Enabled = False
         'M_ExpHR_RADPERC.Enabled = False
         M_ExpHR_RABbase.Caption = ReplaceStr(M_ExpHR_RABbase.Caption, "RAD", "RAB")
         M_ExpHR_RAB.Caption = ReplaceStr(M_ExpHR_RAB.Caption, "RAD", "RAB")
         M_ExpHR_RAB_RLI.Caption = ReplaceStr(M_ExpHR_RAB_RLI.Caption, "RAD", "RAB")
         M_ConfigCtasAjustes.Caption = "Configurar Cuentas Ajustes Extra-contables-14 TER A)..."
         M_ConfigCtasAjustesRLI.Caption = ReplaceStr(M_ConfigCtasAjustesRLI.Caption, "RAD", "RAB")
      
      End If
      
      If gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
         M_CapitalPropioSimpl.Enabled = True
         M_ExpHR_RAB_RLI.Enabled = False
      Else
         M_CapitalPropioSimpl.Enabled = False
      End If
      
   End If
   
End Function
Private Sub MH_DesInscr_Click()

   If MsgBox1("Al desinscribir este equipo el programa funcionará en modo demo." & vbLf & "¿Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   Call FwUnRegister

End Sub

Private Sub MH_HlpBackup_Click()
   Dim Frm As FrmBackup
   
   Set Frm = New FrmBackup
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub


Private Sub MH_ManualesDeUso_Click()
   Dim URLManuales As String
   
   URLManuales = "http://www.hyperrenta.cl/?page_id=656"      'enviado por Nicolás Catrin el 13 ago 2018

   Call ShellExecute(Me.hWnd, "open", URLManuales, "", "", 1)

End Sub

Private Sub MH_RepErr_Click()
   Dim Frm As FrmRepError
   
   Set Frm = New FrmRepError
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MH_Tutorial_Click()
   Dim URLTutotial As String
   
'   URLTutotial = "https://www.youtube.com/watch?v=6QbGXDyE_ys"

'   URLTutotial = "https://goo.gl/zWvkSz"      'enviado por Nicolás Catrin el 13 ago 2018

   'URLTutotial = "https://www.youtube.com/playlist?list=PLz2bqn2pcxA_N0WeLm7KW-r-3laBU9I-u"   'Enviado por Nicolás Catrín 4 oct 2018
    URLTutotial = "https://centrodesoluciones.thomsonreuters.cl/home"

   Call ShellExecute(Me.hWnd, "open", URLTutotial, "", "", 1)

End Sub

Private Sub M_InfoOtrosDocs_Click()
   Dim Frm As FrmInfAnalitico
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmInfAnalitico
   Call Frm.FViewOtrosDocs(0)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub


Private Sub Pc_Boton_Click(Index As Integer)
   Select Case Index
   
      Case 0
         Call Bt_Emp_Click
         
      Case 1
         Call Bt_Plan_Click
         
      Case 2
         Call Bt_NewComprob_Click
         
      Case 3
         Call Bt_LstComp_Click
         
      Case 4
         Call Bt_NewDoc_Click
         
      Case 5
         Call Bt_LstDoc_Click
         
      Case 6, 7
         Call Bt_Contrib14Ter_Click
                  
      Case 8
         Call Bt_Libros_Click
         
      Case 9
         Call Bt_Balances_Click
         
      Case 10
         Call Bt_InfoIFRS_Click
         
      Case 11
         Call Bt_InfAnalitico_Click
        
      Case 12
         Call Bt_Result_Click
         
      Case 13
         Call Bt_MantActivoFijo_Click

      Case 14
         Call Bt_ContActFijo_Click

      Case 15
         Call Bt_Calc_Click
       
      Case 16
         Call Bt_Calendar_Click
         
      Case 17
         Call Bt_ConvMoneda_Click
        
      Case 18
         Call bt_Equivalencia_Click
         
      Case 19
         Call Bt_Indices_Click
         
   End Select
   
End Sub

Private Sub Tm_ChkUsr_Timer()
   Dim Usr As String

   DbMainDate = GetDbNow(DbMain)

   Usr = ContRegisteredUsr()
   If Usr <> "" Then
      Call CloseDb(DbMain)
      If Usr = "." Then
         MsgBox1 "El administrador ha desconectado este usuario, se cerrará esta aplicación.", vbCritical
      Else
         MsgBox1 "El usuario " & Usr & " inició una sesión en este equipo, se cerrará esta aplicación.", vbCritical
      End If
      End
   End If

End Sub


Private Sub Tmr_Chk_Timer()

   Tmr_Chk.Enabled = Not CheckVersion(Me, False)

End Sub

Private Sub Tmr_ChkActive_Timer()
   Dim Dmo As String
   Static Tm As Double, Last As Double
   Static nChk As Long
      
   If Now < Last Or Now > Tm Then
      Tmr_ChkActive.Enabled = False
      
      nChk = nChk + 1
      
      Call FwChkActive(1)
      
      If gAppCode.Demo Then
         Dmo = "D" & "E" & "M" & "O" & " " & "-" & " "
         If Left(Me.Caption, Len(Dmo)) <> Dmo Then
            Me.Caption = Dmo & Me.Caption
         End If

         La_demo(0).visible = True
         La_demo(1).visible = True

      End If
      
      If gAppCode.Msg <> "" And (gAppCode.MinMsg > 0 Or nChk = 1) Then
         MsgBox1 gAppCode.Msg, vbInformation
      End If

      If gAppCode.MinMsg > 0 Then
         
         Tm = Now + TimeSerial(0, gAppCode.MinMsg, 0)
      Else
         Tm = Now + TimeSerial(0, 60, 0)
      End If
            
      Tmr_ChkActive.Enabled = (Not gFwChkActive Or gAppCode.MinMsg > 0)
   End If

   Last = Now

End Sub

Private Sub SetMainSQLServer()

   If gDbType = SQL_SERVER Then
   
      Me.Height = 9225
      Fr_Access(1).visible = False
      Fr_Access(2).visible = False
      Fr_Access(3).visible = False
'      Fr_Top.Left = 0
'      Fr_Top.Width = Me.Width
      Fr_Top.BorderStyle = 0
      Fr_Top.BackColor = COLOR_LIGHTBLUE
      
      Me.BackColor = COLOR_LIGHTBLUE
      Line1(4).visible = True
      Line1(5).visible = True
      
#If DATACON = 2 Then       'SQL Server o MySQL
'      FR_SQLServer(0).Visible = True        Nueva Interfaz

''      FR_SQLServer(1).Visible = True
'
''      Pc_Logo.Visible = True
''      Fr_Top.Left = Pc_Logo.Width - 100

'      If gEmpresa.Ano < 2020 Then
'         Pc_Boton(BT_14TER).Visible = True
'         Pc_Boton(BT_14D).Visible = False
'      Else
'         Pc_Boton(BT_14TER).Visible = False
'         Pc_Boton(BT_14D).Visible = True
'      End If
#End If

   Else
      Me.Height = 8190

  End If
  
End Sub
