VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRepActFijoIFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Activo Fijo Financiero (Modalidad IFRS)"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13920
   Icon            =   "FrmRepActFijoIFRS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   13920
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Opciones 
      Height          =   2535
      Left            =   9540
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CheckBox Ch_ViewRevalorizacion 
         Caption         =   "Ver Revalorizaciï¿½n"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   1020
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewValorRazonable 
         Caption         =   "Ver Valor Razonable"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   300
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewFactor 
         Caption         =   "Ver Factor"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   660
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewPjeAmortizacion 
         Caption         =   "Ver Pje. Amortizaciï¿½n"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   2100
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewValorInicial 
         Caption         =   "Ver Valor Inicial"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1740
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewFechaCompra 
         Caption         =   "Ver Fecha Compra"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   1380
         Width           =   1920
      End
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   8940
      Width           =   7155
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7755
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   13679
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   13875
      Begin VB.PictureBox Pc_HdCheck 
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   4680
         Picture         =   "FrmRepActFijoIFRS.frx":000C
         ScaleHeight     =   150
         ScaleWidth      =   150
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.PictureBox Pc_Check 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   4080
         Picture         =   "FrmRepActFijoIFRS.frx":0371
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Bt_CrearComprobante 
         Caption         =   "Crear Comprobante"
         Height          =   315
         Left            =   7680
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Bt_Opciones 
         Caption         =   "Opciones de Vista"
         Height          =   315
         Left            =   9540
         TabIndex        =   11
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton Bt_ViewRes 
         Caption         =   "Ver Resumen"
         Height          =   315
         Left            =   11160
         TabIndex        =   18
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Bt_FechaHasta 
         Caption         =   "?"
         Height          =   315
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   215
      End
      Begin VB.TextBox Tx_FechaHasta 
         Height          =   315
         Left            =   6240
         TabIndex        =   9
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_VerActivoFijo 
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
         Left            =   120
         Picture         =   "FrmRepActFijoIFRS.frx":03E8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ver/editar detalle de Activo Fijo seleccionado"
         Top             =   180
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
         Left            =   1020
         Picture         =   "FrmRepActFijoIFRS.frx":07E6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   12420
         TabIndex        =   19
         Top             =   180
         Width           =   1215
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
         Left            =   600
         Picture         =   "FrmRepActFijoIFRS.frx":0CA0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
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
         Left            =   1500
         Picture         =   "FrmRepActFijoIFRS.frx":1147
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copiar Excel"
         Top             =   180
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
         Left            =   3420
         Picture         =   "FrmRepActFijoIFRS.frx":158C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Calendario"
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
         Left            =   2580
         Picture         =   "FrmRepActFijoIFRS.frx":19B5
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   3000
         Picture         =   "FrmRepActFijoIFRS.frx":1D53
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Calculadora"
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
         Left            =   2040
         Picture         =   "FrmRepActFijoIFRS.frx":20B4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   0
         Left            =   5700
         TabIndex        =   21
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8580
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmRepActFijoIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const C_IDACTFIJO = 0
'Const C_IDCOMPFICHA = 1
'Const C_ACTFIJO = 2
'Const C_COMPONENTE = 3
'Const C_CUENTA = 4
'Const C_VALRAZONABLE = 5
'Const C_FACTOR = 6
'Const C_REVALORIZACION = 7
'Const C_FECHACOMPRA = 8
'Const C_FECHADISPONIBLE = 9
'Const C_VALORINICIAL = 10
'Const C_VALORBIEN = 11
'Const C_DEPACUM = 12
'Const C_VALORLIBRO = 13
'Const C_PJEAMORTIZACION = 14
'Const C_VALORRESIDUAL = 15
'Const C_VALDEPRECIAR = 16
'Const C_VIDAUTILTOTAL = 17
'Const C_VIDAUTILYADEP = 18
'Const C_VIDAUTILRESIDUAL = 19
'Const C_VIDAUTILADEP = 20
'Const C_DEPMENSUAL = 21
'Const C_DEPPERIODO = 22
'Const C_VALLIBROANTESREVAL = 23
'Const C_REVALDETERIORO = 24
'Const C_VALLIBRODESPREVAL = 25
'Const C_RESERVAACUMANTERIOR = 26
'Const C_RESERVAPERIODO = 27
'Const C_PERDIDAPERIODO = 28
'Const C_OTRASDIF = 29
'Const C_RESEVAACUMSGTE = 30
'Const C_OBLIGATORIA = 31
'Const C_FMT = 32
'
'
'Const NCOLS = C_FMT


'2861591
Const C_IDACTFIJO = 0
Const C_IDCOMPFICHA = 1
Const C_ACTFIJO = 2
Const C_COMPONENTE = 3
Const C_SELECT = 4
Const C_CUENTA = 5
Const C_VALRAZONABLE = 6
Const C_FACTOR = 7
Const C_REVALORIZACION = 8
Const C_FECHACOMPRA = 9
Const C_FECHADISPONIBLE = 10
Const C_VALORINICIAL = 11
Const C_VALORBIEN = 12
Const C_DEPACUM = 13
Const C_VALORLIBRO = 14
Const C_PJEAMORTIZACION = 15
Const C_VALORRESIDUAL = 16
Const C_VALDEPRECIAR = 17
Const C_VIDAUTILTOTAL = 18
Const C_VIDAUTILYADEP = 19
Const C_VIDAUTILRESIDUAL = 20
Const C_VIDAUTILADEP = 21
Const C_DEPMENSUAL = 22
Const C_DEPPERIODO = 23
Const C_VALLIBROANTESREVAL = 24
Const C_REVALDETERIORO = 25
Const C_VALLIBRODESPREVAL = 26
Const C_RESERVAACUMANTERIOR = 27
Const C_RESERVAPERIODO = 28
Const C_PERDIDAPERIODO = 29
Const C_OTRASDIF = 30
Const C_RESEVAACUMSGTE = 31
Const C_OBLIGATORIA = 32
Const C_FMT = 33

Const NCOLS = C_FMT
'2861591

Dim lFecha As Long

Dim lViewRes As Boolean
 
'2861591
Dim vIdCuenta As String
Dim vCuenta As String
Dim vDescCuenta As String
Dim vDebe As Boolean
Dim vHaber As Boolean
Dim vValorRazonable As Double
Dim vValorInicial As Double
Dim vValorDelBien As Double
Dim vValorLibro As Double
Dim vValorResidual As Double
Dim vValorDeprecial As Double
Dim vValorLibroAntRevalor As Double
Dim vValorLibroDespRevalo As Double
'2861591

Public Sub FView()
   Me.Show vbModal
End Sub
 
Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
    
   Grid.Cols = NCOLS + 1
   '2861591
   'Grid.rows = 10
   Grid.rows = 11
   '2861591
   Grid.FixedRows = 3
   
   Call FGrSetup(Grid)
   
   Grid.FixedCols = C_COMPONENTE + 1
   
   Call SetupGridRes(False)
      
   Grid.ColAlignment(C_ACTFIJO) = flexAlignLeftCenter
   Grid.ColAlignment(C_COMPONENTE) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHACOMPRA) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHADISPONIBLE) = flexAlignRightCenter
   
   For i = C_VALORINICIAL To Grid.Cols - 1
      Grid.ColAlignment(i) = flexAlignRightCenter
   Next i
         
   Call FGrTotales(Grid, GridTot)

   Call FGrVRows(Grid)
    
End Sub
Private Sub SetupGridRes(ByVal ViewRes As Boolean)
   Dim i As Integer

   Grid.ColWidth(C_IDACTFIJO) = 0
   Grid.ColWidth(C_IDCOMPFICHA) = 0
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   
      '2861591
   Grid.ColWidth(C_SELECT) = 300
   Grid.Row = 0
   Grid.Col = C_SELECT
   'Set Grid.CellPicture = Pc_Prt
   Set Grid.CellPicture = Pc_HdCheck
   Grid.CellPictureAlignment = flexAlignCenterCenter
   '2861591
   
   If Not ViewRes Then
      
      Grid.ColWidth(C_ACTFIJO) = 2000
      Grid.ColWidth(C_COMPONENTE) = 2000
      Grid.ColWidth(C_CUENTA) = 2000
      
      Grid.ColWidth(C_VALRAZONABLE) = 1200
      Grid.ColWidth(C_FACTOR) = 800
      Grid.ColWidth(C_REVALORIZACION) = 1300
      
      Grid.ColWidth(C_FECHACOMPRA) = 900
      Grid.ColWidth(C_FECHADISPONIBLE) = 900
     
      For i = C_VALORINICIAL To NCOLS
         Grid.ColWidth(i) = 1200
      Next i
   
      Grid.ColWidth(C_OBLIGATORIA) = 0
      Grid.ColWidth(C_FMT) = 0
   
      Grid.TextMatrix(1, C_ACTFIJO) = ""
      Grid.TextMatrix(2, C_ACTFIJO) = "Módulo"
      
      Grid.TextMatrix(1, C_COMPONENTE) = ""
      Grid.TextMatrix(2, C_COMPONENTE) = "Componente"
      
      Grid.TextMatrix(1, C_CUENTA) = ""
      Grid.TextMatrix(2, C_CUENTA) = "Cuenta"
      
      Grid.TextMatrix(0, C_FECHACOMPRA) = ""
      Grid.TextMatrix(1, C_FECHACOMPRA) = "Fecha"
      Grid.TextMatrix(2, C_FECHACOMPRA) = "Compra"
      
      Grid.TextMatrix(0, C_FECHADISPONIBLE) = "Fecha"
      Grid.TextMatrix(1, C_FECHADISPONIBLE) = "Disponible"
      Grid.TextMatrix(2, C_FECHADISPONIBLE) = "p/Utilizar"
      
      Grid.TextMatrix(0, C_VALRAZONABLE) = ""
      Grid.TextMatrix(1, C_VALRAZONABLE) = "Valor"
      Grid.TextMatrix(2, C_VALRAZONABLE) = "Razonable"
      
      Grid.TextMatrix(0, C_FACTOR) = ""
      Grid.TextMatrix(1, C_FACTOR) = ""
      Grid.TextMatrix(2, C_FACTOR) = "Factor"
      
      Grid.TextMatrix(0, C_REVALORIZACION) = ""
      Grid.TextMatrix(1, C_REVALORIZACION) = ""
      Grid.TextMatrix(2, C_REVALORIZACION) = "Revalorización"
      
      Grid.TextMatrix(0, C_VALORINICIAL) = ""
      Grid.TextMatrix(1, C_VALORINICIAL) = "Valor"
      Grid.TextMatrix(2, C_VALORINICIAL) = "Inicial"
      
      Grid.TextMatrix(0, C_VALORBIEN) = ""
      Grid.TextMatrix(1, C_VALORBIEN) = "Valor"
      Grid.TextMatrix(2, C_VALORBIEN) = "del Bien"
      
      Grid.TextMatrix(0, C_DEPACUM) = ""
      Grid.TextMatrix(1, C_DEPACUM) = "Depreciación"
      Grid.TextMatrix(2, C_DEPACUM) = "Acumulada"
      
      Grid.TextMatrix(0, C_VALORLIBRO) = ""
      Grid.TextMatrix(1, C_VALORLIBRO) = "Valor"
      Grid.TextMatrix(2, C_VALORLIBRO) = "Libro"
      
      Grid.TextMatrix(0, C_PJEAMORTIZACION) = ""
      Grid.TextMatrix(1, C_PJEAMORTIZACION) = "Porcentaje"
      Grid.TextMatrix(2, C_PJEAMORTIZACION) = "Amortización"
      
      Grid.TextMatrix(0, C_VALORRESIDUAL) = ""
      Grid.TextMatrix(1, C_VALORRESIDUAL) = "Valor"
      Grid.TextMatrix(2, C_VALORRESIDUAL) = "Residual"
      
      Grid.TextMatrix(0, C_VALDEPRECIAR) = ""
      Grid.TextMatrix(1, C_VALDEPRECIAR) = "Valor"
      Grid.TextMatrix(2, C_VALDEPRECIAR) = "a Depreciar"
      
      Grid.TextMatrix(0, C_VIDAUTILTOTAL) = ""
      Grid.TextMatrix(1, C_VIDAUTILTOTAL) = "Vida Útil"
      Grid.TextMatrix(2, C_VIDAUTILTOTAL) = "Total"
      
      Grid.TextMatrix(0, C_VIDAUTILYADEP) = ""
      Grid.TextMatrix(1, C_VIDAUTILYADEP) = "Vida Útil"
      Grid.TextMatrix(2, C_VIDAUTILYADEP) = "ya Depreciada"
   
      Grid.TextMatrix(0, C_VIDAUTILRESIDUAL) = ""
      Grid.TextMatrix(1, C_VIDAUTILRESIDUAL) = "Vida Útil"
      Grid.TextMatrix(2, C_VIDAUTILRESIDUAL) = "Residual"
      
      Grid.TextMatrix(0, C_VIDAUTILADEP) = ""
      Grid.TextMatrix(1, C_VIDAUTILADEP) = "Vida Útil"
      Grid.TextMatrix(2, C_VIDAUTILADEP) = "a Depreciar"
      
      Grid.TextMatrix(0, C_DEPMENSUAL) = ""
      Grid.TextMatrix(1, C_DEPMENSUAL) = "Depreciación"
      Grid.TextMatrix(2, C_DEPMENSUAL) = "Mensual"
      
      Grid.TextMatrix(0, C_DEPPERIODO) = ""
      Grid.TextMatrix(1, C_DEPPERIODO) = "Depreciación"
      Grid.TextMatrix(2, C_DEPPERIODO) = "del Periodo"
      
      Grid.TextMatrix(0, C_VALLIBROANTESREVAL) = "Valor Libro"
      Grid.TextMatrix(1, C_VALLIBROANTESREVAL) = "antes de"
      Grid.TextMatrix(2, C_VALLIBROANTESREVAL) = "Revalorización"
      
      Grid.TextMatrix(0, C_REVALDETERIORO) = "Reajuste de"
      Grid.TextMatrix(1, C_REVALDETERIORO) = "Revalorización"
      Grid.TextMatrix(2, C_REVALDETERIORO) = "y/o Deterioro"
      
      Grid.TextMatrix(0, C_REVALDETERIORO) = "Reajuste de"
      Grid.TextMatrix(1, C_REVALDETERIORO) = "Revalorización"
      Grid.TextMatrix(2, C_REVALDETERIORO) = "y/o Deterioro"
      
      Grid.TextMatrix(0, C_VALLIBRODESPREVAL) = "Valor Libro"
      Grid.TextMatrix(1, C_VALLIBRODESPREVAL) = "después de"
      Grid.TextMatrix(2, C_VALLIBRODESPREVAL) = "Revalorización"
      
      Grid.TextMatrix(0, C_RESERVAACUMANTERIOR) = "Reserva"
      Grid.TextMatrix(1, C_RESERVAACUMANTERIOR) = "Acumulada"
      Grid.TextMatrix(2, C_RESERVAACUMANTERIOR) = "Anterior"
      
      Grid.TextMatrix(0, C_RESERVAPERIODO) = "Reserva"
      Grid.TextMatrix(1, C_RESERVAPERIODO) = "del"
      Grid.TextMatrix(2, C_RESERVAPERIODO) = "Periodo"
      
      Grid.TextMatrix(0, C_PERDIDAPERIODO) = "Pérdida"
      Grid.TextMatrix(1, C_PERDIDAPERIODO) = "del"
      Grid.TextMatrix(2, C_PERDIDAPERIODO) = "Periodo"
      
      Grid.TextMatrix(0, C_OTRASDIF) = ""
      Grid.TextMatrix(1, C_OTRASDIF) = "Otras"
      Grid.TextMatrix(2, C_OTRASDIF) = "Diferencias"
      
      Grid.TextMatrix(0, C_RESEVAACUMSGTE) = "Reserva"
      Grid.TextMatrix(1, C_RESEVAACUMSGTE) = "Acumulada"
      Grid.TextMatrix(2, C_RESEVAACUMSGTE) = "Siguiente"
   
   Else
   
      For i = C_VALRAZONABLE To Grid.Cols - 1
         Grid.ColWidth(i) = 0
         Grid.TextMatrix(0, i) = ""
         Grid.TextMatrix(1, i) = ""
         Grid.TextMatrix(2, i) = ""
      Next i
            
      Grid.ColWidth(C_VALRAZONABLE) = 1200
      Grid.ColWidth(C_REVALORIZACION) = 1200
      Grid.ColWidth(C_VALORINICIAL) = 1200
      Grid.ColWidth(C_DEPPERIODO) = 1200
      Grid.ColWidth(C_VALLIBROANTESREVAL) = 1200
      Grid.ColWidth(C_VALLIBRODESPREVAL) = 1200
      Grid.ColWidth(C_RESERVAPERIODO) = 1200
      Grid.ColWidth(C_PERDIDAPERIODO) = 1200
      Grid.ColWidth(C_RESEVAACUMSGTE) = 1200
      
      Grid.TextMatrix(0, C_VALRAZONABLE) = ""
      Grid.TextMatrix(1, C_VALRAZONABLE) = "Valor"
      Grid.TextMatrix(2, C_VALRAZONABLE) = "Razonable"
      
      Grid.TextMatrix(0, C_REVALORIZACION) = ""
      Grid.TextMatrix(1, C_REVALORIZACION) = ""
      Grid.TextMatrix(2, C_REVALORIZACION) = "Revalorización"
      
      Grid.TextMatrix(0, C_VALORINICIAL) = ""
      Grid.TextMatrix(1, C_VALORINICIAL) = "Valor"
      Grid.TextMatrix(2, C_VALORINICIAL) = "Inicial"
      
      Grid.TextMatrix(0, C_DEPPERIODO) = ""
      Grid.TextMatrix(1, C_DEPPERIODO) = "Depreciación"
      Grid.TextMatrix(2, C_DEPPERIODO) = "del Periodo"
   
      Grid.TextMatrix(0, C_VALLIBROANTESREVAL) = "Valor Libro"
      Grid.TextMatrix(1, C_VALLIBROANTESREVAL) = "antes de"
      Grid.TextMatrix(2, C_VALLIBROANTESREVAL) = "Revalorización"
      
      Grid.TextMatrix(0, C_VALLIBRODESPREVAL) = "Valor Libro"
      Grid.TextMatrix(1, C_VALLIBRODESPREVAL) = "después de"
      Grid.TextMatrix(2, C_VALLIBRODESPREVAL) = "Revalorización"
      
      Grid.TextMatrix(0, C_RESERVAPERIODO) = "Reserva"
      Grid.TextMatrix(1, C_RESERVAPERIODO) = "del"
      Grid.TextMatrix(2, C_RESERVAPERIODO) = "Periodo"
            
      Grid.TextMatrix(0, C_PERDIDAPERIODO) = "Pérdida"
      Grid.TextMatrix(1, C_PERDIDAPERIODO) = "del"
      Grid.TextMatrix(2, C_PERDIDAPERIODO) = "Periodo"
      
      Grid.TextMatrix(0, C_RESEVAACUMSGTE) = "Reserva"
      Grid.TextMatrix(1, C_RESEVAACUMSGTE) = "Acumulada"
      Grid.TextMatrix(2, C_RESEVAACUMSGTE) = "Siguiente"
  
   End If
   
   Call FGrTotales(Grid, GridTot)
       
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

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   'Call LP_FGr2Clip(Grid, Me.Caption & vbTab & "Año: " & gEmpresa.Ano)
   Clip = LP_FGr2String(Grid, Me.Caption & vbTab & "Año: " & gEmpresa.Ano, False, C_OBLIGATORIA)
   
   If Clip <> "" Then
      Clip = Clip & FGr2String(GridTot)
      
      Clipboard.Clear
      Clipboard.SetText Clip
   End If

End Sub


'2861591
Private Sub Bt_CrearComprobante_Click()
Dim Frm As FrmConfigCompActFijo
Dim FrmComp As FrmComprobante

 Dim Suma As Double

   Dim rowGrid As Integer
   Dim ColGrid As Integer
   Dim NLin As Integer
   Dim AuxRow As Integer
   Dim AuxCol As Integer
   Dim FirstRow As Integer
   Dim LastRow As Integer

   Suma = 0


Set Frm = New FrmConfigCompActFijo

Me.MousePointer = vbHourglass
   Dim Total As Double
   Dim Row As Integer
   Dim Col As Integer

   Row = Grid.Row

   If Row < Grid.FixedRows Then
      Exit Sub
   End If

     vValorRazonable = 0
     vValorInicial = 0
     vValorDelBien = 0
     vValorLibro = 0
     vValorResidual = 0
     vValorDeprecial = 0
     vValorLibroAntRevalor = 0
     vValorLibroDespRevalo = 0

    For rowGrid = Grid.FixedRows To Grid.rows - 1

      NLin = NLin + 1

      Grid.Row = rowGrid
      Grid.Col = C_SELECT

 If Grid.CellPicture <> 0 Then


   If Val(Grid.TextMatrix(rowGrid, C_IDACTFIJO)) > 0 Then

     vValorRazonable = vValorRazonable + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALRAZONABLE)))
     vValorInicial = vValorInicial + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORINICIAL)))
     vValorDelBien = vValorDelBien + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORBIEN)))
     vValorLibro = vValorLibro + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORLIBRO)))
     vValorResidual = vValorResidual + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORRESIDUAL)))
     vValorDeprecial = vValorDeprecial + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALDEPRECIAR)))
     vValorLibroAntRevalor = vValorLibroAntRevalor + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALLIBROANTESREVAL)))
     vValorLibroDespRevalo = vValorLibroDespRevalo + Val(vFmt(Grid.TextMatrix(rowGrid, C_VALLIBRODESPREVAL)))

'    Call Frm.FSelect2(vIdCuenta, vCuenta, vDescCuenta, Val(vFmt(Grid.TextMatrix(rowGrid, C_VALRAZONABLE))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORINICIAL))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORBIEN))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORLIBRO))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALORRESIDUAL))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALDEPRECIAR))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALLIBROANTESREVAL))), Val(vFmt(Grid.TextMatrix(rowGrid, C_VALLIBRODESPREVAL))))

    End If
 End If

    Next rowGrid

    Call Frm.FSelect2(vIdCuenta, vCuenta, vDescCuenta, vValorRazonable, vValorInicial, vValorDelBien, vValorLibro, vValorResidual, vValorDeprecial, vValorLibroAntRevalor, vValorLibroDespRevalo)

      Me.MousePointer = vbDefault


End Sub
'2861591

Private Sub Bt_Opciones_Click()
   Fr_Opciones.visible = Not Fr_Opciones.visible

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   
'   If Not lViewRes Then
'      If MsgBox1("Este informe debe imprimirse en hoja oficio. Asegúrese de tener configurada la impresora para tal efecto." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
'         Exit Sub
'      End If
'   End If
   
   Me.MousePointer = vbHourglass
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Pag = gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   If Pag >= 0 Then
      Call Frm.FView(Caption)
   End If
   
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
   Call ResetPrtBas(gPrtReportes)


End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
      
'   If Not lViewRes Then
'      If MsgBox1("Este informe debe imprimirse en hoja oficio. Asegúrese de tener configurada la impresora para tal efecto." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
'         Exit Sub
'      End If
'   End If
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
   Call ResetPrtBas(gPrtReportes)

End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim NumWi As Integer
   
   Printer.Orientation = ORIENT_HOR
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   Titulos(1) = "Año " & gEmpresa.Ano
   gPrtReportes.Titulos = Titulos
         
   gPrtReportes.GrFontName = "Arial"
   gPrtReportes.GrFontSize = 8 'Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   For i = 0 To Grid.Cols - 1
      If ColWi(i) > 0 Then
         ColWi(i) = ColWi(i) * 0.92
      End If
   Next i
      
   gPrtReportes.GrFontSize = 7
   gPrtReportes.GrFontName = "Arial"
               
   Total(C_ACTFIJO) = "  Total"
   For i = C_VALORINICIAL To Grid.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 1

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing


End Sub

Private Sub Bt_VerActivoFijo_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmActivoFijo
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDACTFIJO)) > 0 Then
      Set Frm = New FrmActivoFijo
      If Frm.FEdit(Val(Grid.TextMatrix(Row, C_IDACTFIJO))) = vbOK Then
         Me.MousePointer = vbHourglass
         Call LoadAll
         Me.MousePointer = vbDefault
      End If
      Set Frm = Nothing
   End If

End Sub


Private Sub Bt_ViewRes_Click()
   lViewRes = Not lViewRes
   
   Call SetupGridRes(lViewRes)
   
   If lViewRes Then
      Bt_ViewRes.Caption = "Ver Completo"
   Else
      Bt_ViewRes.Caption = "Ver Resumen"
   End If

End Sub

Private Sub Ch_ViewFactor_Click()

   If Ch_ViewFactor = 0 Then
      Grid.ColWidth(C_FACTOR) = 0
      Grid.TextMatrix(0, C_FACTOR) = ""
      Grid.TextMatrix(1, C_FACTOR) = ""
      Grid.TextMatrix(2, C_FACTOR) = ""
      
   Else
      Grid.ColWidth(C_FACTOR) = 800
      Grid.TextMatrix(0, C_FACTOR) = ""
      Grid.TextMatrix(1, C_FACTOR) = ""
      Grid.TextMatrix(2, C_FACTOR) = "Factor"
   
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerFactor", Abs(Ch_ViewFactor.Value))
   gVarIniFile.VerFactor = Abs(Ch_ViewFactor.Value)


End Sub

Private Sub Ch_ViewFechaCompra_Click()

   If Ch_ViewFechaCompra = 0 Then
      Grid.ColWidth(C_FECHACOMPRA) = 0
      Grid.TextMatrix(0, C_FECHACOMPRA) = ""
      Grid.TextMatrix(1, C_FECHACOMPRA) = ""
      Grid.TextMatrix(2, C_FECHACOMPRA) = ""
      
   Else
      Grid.ColWidth(C_FECHACOMPRA) = 900
      Grid.TextMatrix(0, C_FECHACOMPRA) = ""
      Grid.TextMatrix(1, C_FECHACOMPRA) = "Fecha"
      Grid.TextMatrix(2, C_FECHACOMPRA) = "Compra"
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerFechaCompra", Abs(Ch_ViewFechaCompra.Value))
   gVarIniFile.VerFechaCompra = Abs(Ch_ViewFechaCompra.Value)

End Sub

Private Sub Ch_ViewPjeAmortizacion_Click()

   If Ch_ViewPjeAmortizacion = 0 Then
      Grid.ColWidth(C_PJEAMORTIZACION) = 0
      Grid.TextMatrix(0, C_PJEAMORTIZACION) = ""
      Grid.TextMatrix(1, C_PJEAMORTIZACION) = ""
      Grid.TextMatrix(2, C_PJEAMORTIZACION) = ""
      
   Else
      Grid.ColWidth(C_PJEAMORTIZACION) = 1200
      Grid.TextMatrix(0, C_PJEAMORTIZACION) = ""
      Grid.TextMatrix(1, C_PJEAMORTIZACION) = "Porcentaje"
      Grid.TextMatrix(2, C_PJEAMORTIZACION) = "Amortización"
   
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerPjeAmortizacion", Abs(Ch_ViewPjeAmortizacion.Value))
   gVarIniFile.VerPjeAmortizacion = Abs(Ch_ViewPjeAmortizacion.Value)

End Sub

Private Sub Ch_ViewRevalorizacion_Click()

   If Ch_ViewRevalorizacion = 0 Then
      Grid.ColWidth(C_REVALORIZACION) = 0
      Grid.TextMatrix(0, C_REVALORIZACION) = ""
      Grid.TextMatrix(1, C_REVALORIZACION) = ""
      Grid.TextMatrix(2, C_REVALORIZACION) = ""
      
   Else
      Grid.ColWidth(C_REVALORIZACION) = 1200
      Grid.TextMatrix(0, C_REVALORIZACION) = ""
      Grid.TextMatrix(1, C_REVALORIZACION) = ""
      Grid.TextMatrix(2, C_REVALORIZACION) = "Revalorización"
     
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerRevalorizacion", Abs(Ch_ViewRevalorizacion.Value))
   gVarIniFile.VerRevalorizacion = Abs(Ch_ViewRevalorizacion.Value)

End Sub

Private Sub Ch_ViewValorInicial_Click()

   If Ch_ViewValorInicial = 0 Then
      Grid.ColWidth(C_VALORINICIAL) = 0
      Grid.TextMatrix(0, C_VALORINICIAL) = ""
      Grid.TextMatrix(1, C_VALORINICIAL) = ""
      Grid.TextMatrix(2, C_VALORINICIAL) = ""
      
   Else
      Grid.ColWidth(C_VALORINICIAL) = 1200
      Grid.TextMatrix(0, C_VALORINICIAL) = ""
      Grid.TextMatrix(1, C_VALORINICIAL) = "Valor"
      Grid.TextMatrix(2, C_VALORINICIAL) = "Inicial"
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerValorInicial", Abs(Ch_ViewValorInicial.Value))
   gVarIniFile.VerValorInicial = Abs(Ch_ViewValorInicial.Value)


End Sub

Private Sub Ch_ViewValorRazonable_Click()

   If Ch_ViewValorRazonable = 0 Then
      Grid.ColWidth(C_VALRAZONABLE) = 0
      Grid.TextMatrix(0, C_VALRAZONABLE) = ""
      Grid.TextMatrix(1, C_VALRAZONABLE) = ""
      Grid.TextMatrix(2, C_VALRAZONABLE) = ""
      
   Else
      Grid.ColWidth(C_VALRAZONABLE) = 1200
      Grid.TextMatrix(0, C_VALRAZONABLE) = ""
      Grid.TextMatrix(1, C_VALRAZONABLE) = "Valor"
      Grid.TextMatrix(2, C_VALRAZONABLE) = "Razonable"
   
   
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerValorRazonable", Abs(Ch_ViewValorRazonable.Value))
   gVarIniFile.VerValorRazonable = Abs(Ch_ViewValorRazonable.Value)

End Sub

Private Sub Form_Load()
   
   Fr_Opciones.visible = False
   
   Call BtFechaImg(Bt_FechaHasta)
   lFecha = DateSerial(gEmpresa.Ano, 12, 31)
   Call SetTxDate(Tx_FechaHasta, lFecha)
   
   Call SetUpGrid
   
   Ch_ViewFechaCompra = gVarIniFile.VerFechaCompra
   Ch_ViewValorInicial = gVarIniFile.VerValorInicial
   Ch_ViewPjeAmortizacion = gVarIniFile.VerPjeAmortizacion
   Ch_ViewFactor = gVarIniFile.VerFactor
   Ch_ViewValorRazonable = gVarIniFile.VerValorRazonable
   Ch_ViewRevalorizacion = gVarIniFile.VerRevalorizacion
   
   Call Ch_ViewFechaCompra_Click
   Call Ch_ViewValorInicial_Click
   Call Ch_ViewPjeAmortizacion_Click
   Call Ch_ViewFactor_Click
   Call Ch_ViewValorRazonable_Click
   Call Ch_ViewRevalorizacion_Click
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
   DoEvents
   
   If gEmpresa.TieneAnoAnt Then
      MsgBox1 "ATENCIÓN: " & vbCrLf & vbCrLf & "Recuerde actualizar el Valor Razonable para aquellos activos fijos que vienen del año anterior, si los hay.", vbInformation + vbOKOnly
   End If
End Sub

Private Sub LoadAll(Optional ByVal Row As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Aux As Double
   Dim FechaHasta As Long, FechaDisponible As Long
   Dim NoExisteValRazonable As Boolean
   Dim Total(NCOLS + 1)
   Dim k As Integer
   Dim Factor As Double, Valor As Double, Valor1 As Double
   Dim Fecha As Long
   
   Fecha = GetTxDate(Tx_FechaHasta)
   
   Q1 = "SELECT MovActivoFijo.IdActFijo, ActFijoCompsFicha.IdCompFicha, MovActivoFijo.Descrip, AFComponentes.NombComp, Cuentas.Descripcion as Cuenta  "
   Q1 = Q1 & ", MovActivoFijo.Fecha as FechaCompra, ActFijoFicha.FechaDisponible, ActFijoCompsFicha.ValorRazonable_31_12 as ValorRazonable, NoExisteValRazonable "
   Q1 = Q1 & ", ActFijoCompsFicha.ValorCompra, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.PjeAmortizacion, ActFijoCompsFicha.ValorResidual  "
   Q1 = Q1 & ", ActFijoCompsFicha.VidaUtil, ActFijoCompsFicha.DepAcumuladaAnoAnt, ActFijoCompsFicha.VidaUtilYaDep, ActFijoCompsFicha.ReservaAcumAnt, OtrasDiferencias  "
   Q1 = Q1 & " FROM (((MovActivoFijo INNER JOIN ActFijoFicha ON MovActivoFijo.IdActFijo = ActFijoFicha.IdActFijo "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovActivoFijo", "ActFijoFicha") & " )"
   Q1 = Q1 & " INNER JOIN ActFijoCompsFicha ON MovActivoFijo.IdActFijo = ActFijoCompsFicha.IdActFijo "
   Q1 = Q1 & JoinEmpAno(gDbType, "ActFijoCompsFicha", "MovActivoFijo") & " )"
   Q1 = Q1 & " LEFT JOIN AFComponentes ON ActFijoCompsFicha.IdGrupo = AFComponentes.IdGrupo AND ActFijoCompsFicha.IdComp = AFComponentes.IdComp)"
   Q1 = Q1 & " LEFT JOIN Cuentas ON MovActivoFijo.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovActivoFijo")
   
   Q1 = Q1 & " WHERE  MovActivoFijo.IdEmpresa = " & gEmpresa.id & " AND MovActivoFijo.Ano = " & gEmpresa.Ano
   '3332911
   Q1 = Q1 & "  AND (FechaVentaBaja = 0 OR FechaVentaBaja IS NULL OR FechaVentaBaja > " & Fecha & ")"
   '3332911
   If Row <> 0 Then
      Q1 = Q1 & " AND MovActivoFijo.IdActFijo = " & Grid.TextMatrix(Row, C_IDACTFIJO) & " AND ActFijoCompsFicha.IdCompFicha = " & Grid.TextMatrix(Row, C_IDCOMPFICHA)
   End If
   
   Q1 = Q1 & " ORDER BY MovActivoFijo.Descrip, AFComponentes.NombComp "

   Set Rs = OpenRs(DbMain, Q1)
   
   If Row = 0 Then
      Grid.rows = Grid.FixedRows
      i = Grid.rows
   End If
   
   Do While Not Rs.EOF
   
      If Row = 0 Then
         Grid.rows = Grid.rows + 1
      Else
         i = Row
      End If
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If

      NoExisteValRazonable = vFld(Rs("NoExisteValRazonable"))
      Grid.TextMatrix(i, C_IDACTFIJO) = vFld(Rs("IdActFijo"))
      Grid.TextMatrix(i, C_IDCOMPFICHA) = vFld(Rs("IdCompFicha"))
      If Grid.TextMatrix(i - 1, C_IDACTFIJO) <> vFld(Rs("IdActFijo")) Then
         Grid.TextMatrix(i, C_ACTFIJO) = vFld(Rs("Descrip"))
         Call FGrSetColStyle(Grid, C_ACTFIJO, "B")
         Grid.TextMatrix(i, C_FMT) = "FCELL"
      End If
      Grid.TextMatrix(i, C_COMPONENTE) = vFld(Rs("NombComp"))
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Cuenta"))
      
      Grid.TextMatrix(i, C_VALRAZONABLE) = IIf(NoExisteValRazonable, "", Format(vFld(Rs("ValorRazonable")), NUMFMT))
      
      Grid.TextMatrix(i, C_FECHACOMPRA) = IIf(vFld(Rs("FechaCompra")) > 0, Format(vFld(Rs("FechaCompra")), SDATEFMT), "")
      Grid.TextMatrix(i, C_FECHADISPONIBLE) = IIf(vFld(Rs("FechaDisponible")) > 0, Format(vFld(Rs("FechaDisponible")), SDATEFMT), "")
      Grid.TextMatrix(i, C_VALORINICIAL) = Format(vFld(Rs("ValorCompra")), NUMFMT)
      Grid.TextMatrix(i, C_VALORBIEN) = Format(vFld(Rs("ValorBien")), NUMFMT)
      
      Grid.TextMatrix(i, C_DEPACUM) = Format(vFld(Rs("DepAcumuladaAnoAnt")), NUMFMT)            'arrastra año anterior
      Aux = vFld(Rs("ValorBien")) - vFld(Rs("DepAcumuladaAnoAnt"))
      If Aux >= 0 Then
         Grid.TextMatrix(i, C_VALORLIBRO) = Format(Aux, NUMFMT)
      End If
      
      Grid.TextMatrix(i, C_PJEAMORTIZACION) = Format(vFld(Rs("PjeAmortizacion")) * 100, DBLFMT1)
      Grid.TextMatrix(i, C_VALORRESIDUAL) = Format(vFld(Rs("ValorResidual")), NUMFMT)
      
      If Grid.TextMatrix(i, C_VALORLIBRO) <> "" Then
         Grid.TextMatrix(i, C_VALDEPRECIAR) = Format(Int(vFmt(Grid.TextMatrix(i, C_VALORLIBRO)) * Grid.TextMatrix(i, C_PJEAMORTIZACION) / 100 - vFmt(Grid.TextMatrix(i, C_VALORRESIDUAL))), NUMFMT)
      End If
      
      Grid.TextMatrix(i, C_VIDAUTILTOTAL) = Format(vFld(Rs("VidaUtil")), NUMFMT)
      Grid.TextMatrix(i, C_VIDAUTILYADEP) = Format(vFld(Rs("VidaUtilYaDep")), NUMFMT)        'arrastra del año anterior
      Grid.TextMatrix(i, C_VIDAUTILRESIDUAL) = Format(vFld(Rs("VidaUtil")) - vFld(Rs("VidaUtilYaDep")), NUMFMT)
      
      FechaHasta = GetTxDate(Tx_FechaHasta)
      FechaDisponible = vFld(Rs("FechaDisponible"))
            
      If Year(FechaDisponible) = Year(FechaHasta) Then    'fecha disponible es de este año
         If gAFMesCompleto Then
             Grid.TextMatrix(i, C_VIDAUTILADEP) = month(FechaHasta) - month(FechaDisponible) + 1
         Else
            Grid.TextMatrix(i, C_VIDAUTILADEP) = month(FechaHasta) - month(FechaDisponible) + IIf(Day(FechaDisponible) > 15, 0, 1)
         End If
      Else
         Grid.TextMatrix(i, C_VIDAUTILADEP) = month(FechaHasta)
      End If
      
      If vFmt(Grid.TextMatrix(i, C_VIDAUTILRESIDUAL)) < Grid.TextMatrix(i, C_VIDAUTILADEP) Then
         Grid.TextMatrix(i, C_VIDAUTILADEP) = Grid.TextMatrix(i, C_VIDAUTILRESIDUAL)
      End If
         
      If vFmt(Grid.TextMatrix(i, C_VIDAUTILRESIDUAL)) > 0 Then
         Grid.TextMatrix(i, C_DEPMENSUAL) = Format(vFmt(Grid.TextMatrix(i, C_VALDEPRECIAR)) / vFmt(Grid.TextMatrix(i, C_VIDAUTILRESIDUAL)), NUMFMT)
         
         'Grid.TextMatrix(i, C_DEPPERIODO) = Format(vFmt(Grid.TextMatrix(i, C_DEPMENSUAL)) * vFmt(Grid.TextMatrix(i, C_VIDAUTILADEP)), NUMFMT)   'diferencias de redondeo
         Grid.TextMatrix(i, C_DEPPERIODO) = Format(vFmt(Grid.TextMatrix(i, C_VALDEPRECIAR)) / vFmt(Grid.TextMatrix(i, C_VIDAUTILRESIDUAL)) * vFmt(Grid.TextMatrix(i, C_VIDAUTILADEP)), NUMFMT)
      End If
      
      
      Grid.TextMatrix(i, C_VALLIBROANTESREVAL) = Format(vFmt(Grid.TextMatrix(i, C_VALORBIEN)) - vFmt(Grid.TextMatrix(i, C_DEPPERIODO)) - vFmt(Grid.TextMatrix(i, C_DEPACUM)), NUMFMT)
      
      If NoExisteValRazonable Or vFmt(Grid.TextMatrix(i, C_VALRAZONABLE)) = 0 Or vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL)) = 0 Then
         Grid.TextMatrix(i, C_FACTOR) = ""
         Factor = 0
         Grid.TextMatrix(i, C_REVALORIZACION) = ""
      Else
         Factor = vFmt(Grid.TextMatrix(i, C_VALRAZONABLE)) / vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL))
         Grid.TextMatrix(i, C_FACTOR) = Format(Factor, DBLFMT2) & "%"    'se redondea a 2 decimales
         Valor = Round(vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL)) * Factor, 0)
         
         Grid.TextMatrix(i, C_REVALORIZACION) = Format(Valor - vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL)), NUMFMT)
      End If
      
      
      If NoExisteValRazonable Then
         Grid.TextMatrix(i, C_REVALDETERIORO) = ""
      ElseIf vFmt(Grid.TextMatrix(i, C_VALRAZONABLE)) >= 0 Then
         Grid.TextMatrix(i, C_REVALDETERIORO) = Grid.TextMatrix(i, C_REVALORIZACION)
      End If
      
      Grid.TextMatrix(i, C_VALLIBRODESPREVAL) = Format(vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL)) + vFmt(Grid.TextMatrix(i, C_REVALDETERIORO)), NUMFMT)
     
      Grid.TextMatrix(i, C_RESERVAACUMANTERIOR) = Format(vFld(Rs("ReservaAcumAnt")), NUMFMT)     'arrastra año anterior
      
      Grid.TextMatrix(i, C_RESERVAPERIODO) = IIf(vFmt(Grid.TextMatrix(i, C_REVALDETERIORO)) > 0, Grid.TextMatrix(i, C_REVALDETERIORO), 0)
      
      If vFmt(Grid.TextMatrix(i, C_REVALDETERIORO)) > 0 Then
         Grid.TextMatrix(i, C_RESERVAPERIODO) = Grid.TextMatrix(i, C_REVALDETERIORO)
         Grid.TextMatrix(i, C_PERDIDAPERIODO) = 0
      Else
         Grid.TextMatrix(i, C_RESERVAPERIODO) = 0
         Grid.TextMatrix(i, C_PERDIDAPERIODO) = Format(Abs(vFmt(Grid.TextMatrix(i, C_REVALDETERIORO))), NUMFMT)
      End If
         
      Grid.TextMatrix(i, C_OTRASDIF) = Format(vFld(Rs("OtrasDiferencias")), NUMFMT)
      
      Grid.TextMatrix(i, C_RESEVAACUMSGTE) = Format(vFmt(Grid.TextMatrix(i, C_RESERVAACUMANTERIOR)) + vFmt(Grid.TextMatrix(i, C_RESERVAPERIODO)) - vFmt(Grid.TextMatrix(i, C_PERDIDAPERIODO)) + vFmt(Grid.TextMatrix(i, C_OTRASDIF)), NUMFMT)
      
      Grid.TextMatrix(i, C_OBLIGATORIA) = "."
      
      If Row = 0 Then
         For k = C_VALORINICIAL To Grid.Cols - 1
            Total(k) = Total(k) + vFmt(Grid.TextMatrix(i, k))
         Next k
      End If
      
      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   
   If Row <> 0 Then
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_OBLIGATORIA) = "" Then
            Exit For
         End If
         
         For k = C_VALORINICIAL To Grid.Cols - 1
            Total(k) = Total(k) + vFmt(Grid.TextMatrix(i, k))
         Next k
      Next i
   End If

   For k = C_VALORINICIAL To Grid.Cols - 1
      GridTot.TextMatrix(0, k) = Format(Total(k), NUMFMT)
   Next k
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Factor As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_IDCOMPFICHA) = "" Then
         Exit For
      End If
      
      If vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL)) <> 0 Then
         Factor = vFmt(Grid.TextMatrix(i, C_VALRAZONABLE)) / vFmt(Grid.TextMatrix(i, C_VALLIBROANTESREVAL))
      Else
         Factor = 0
      End If

      Q1 = "UPDATE ActFijoCompsFicha SET "
      Q1 = Q1 & "  Factor = " & str(Factor)
      Q1 = Q1 & ", Revalorizacion = " & vFmt(Grid.TextMatrix(i, C_REVALORIZACION))
      Q1 = Q1 & ", DepAcum = " & vFmt(Grid.TextMatrix(i, C_DEPACUM))
      Q1 = Q1 & ", DepPeriodo = " & vFmt(Grid.TextMatrix(i, C_DEPPERIODO))
      Q1 = Q1 & ", VidaUtilDep = " & vFmt(Grid.TextMatrix(i, C_VIDAUTILADEP))
      Q1 = Q1 & ", ReservaAcum = " & vFmt(Grid.TextMatrix(i, C_RESEVAACUMSGTE))
      Q1 = Q1 & "  WHERE IdCompFicha = " & Grid.TextMatrix(i, C_IDCOMPFICHA)
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)
      
   Next i
      

End Sub
Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - Tx_CurrCell.Height - 600
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 180
   GridTot.Width = Grid.Width - 230
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   
   Call FGrVRows(Grid)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SaveAll
End Sub

Private Sub Grid_Click()
   If Fr_Opciones.visible Then
      Fr_Opciones.visible = False
   End If
End Sub

Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As Form
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
  
   'Call PostClick(Bt_VerActivoFijo)
      
   '2861591
  If Col = C_SELECT Then

      If Val(Grid.TextMatrix(Row, C_IDACTFIJO)) <> 0 Then
         If Grid.CellPicture = 0 Then
            Call FGrSetPicture(Grid, Row, Col, Pc_Check, 0)
            '2861591
            Bt_CrearComprobante.visible = True
            '2861591
         Else
            Set Grid.CellPicture = LoadPicture()
         End If
      End If

   Else
'2861591
      
   If Val(Grid.TextMatrix(Row, C_IDACTFIJO)) > 0 Then
      If Col = C_VALRAZONABLE Then
         Set Frm = New FrmAFCompsFicha
         If Frm.FEdit(Val(Grid.TextMatrix(Row, C_IDACTFIJO)), Val(Grid.TextMatrix(Row, C_IDCOMPFICHA)), True) = vbOK Then
            Me.MousePointer = vbHourglass
            Call LoadAll(Row)
            Me.MousePointer = vbDefault
         End If
         Set Frm = Nothing
      Else
         Set Frm = New FrmAFFicha
         If Frm.FEdit(Val(Grid.TextMatrix(Row, C_IDACTFIJO))) = vbOK Then
            Me.MousePointer = vbHourglass
            Call LoadAll
            Me.MousePointer = vbDefault
         End If
         Set Frm = Nothing
      End If
   End If
   
   '2861591
   End If
   '2861591

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub Bt_FechaHasta_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Frm.TxSelDate(Tx_FechaHasta) = vbOK Then
      lFecha = GetTxDate(Tx_FechaHasta)
      
      If lFecha < DateSerial(gEmpresa.Ano, 1, 1) Or lFecha > DateSerial(gEmpresa.Ano, 12, 31) Then
         MsgBox1 "Fecha inválida.", vbExclamation
         Call SetTxDate(Tx_FechaHasta, DateSerial(gEmpresa.Ano, 12, 31))
         lFecha = DateSerial(gEmpresa.Ano, 12, 31)
         Exit Sub
      End If

      If lFecha < DateSerial(gEmpresa.Ano, 12, 31) Then
         MsgBox1 "Asegúrese de asignar correctamente los meses a depreciar en el año actual, en la ventana de mantención de cada Activo Fijo.", vbInformation
      End If
      Me.MousePointer = vbHourglass
      Call LoadAll
      Me.MousePointer = vbDefault
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Tx_FechaHasta_GotFocus()
   Call DtGotFocus(Tx_FechaHasta)
End Sub
Private Sub Tx_FechaHasta_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_FechaHasta) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_FechaHasta)
   
   Call DtLostFocus(Tx_FechaHasta)
   
   lFecha = GetTxDate(Tx_FechaHasta)
   
   If lFecha < DateSerial(gEmpresa.Ano, 1, 1) Or lFecha > DateSerial(gEmpresa.Ano, 12, 31) Then
      MsgBox1 "Fecha inválida.", vbExclamation
      lFecha = DateSerial(gEmpresa.Ano, 12, 31)
      Call SetTxDate(Tx_FechaHasta, lFecha)
      Exit Sub
   End If
   
   If lFecha < DateSerial(gEmpresa.Ano, 12, 31) Then
      MsgBox1 "Asegúrese de asignar correctamente los meses a depreciar en el año actual, en la ventana de mantención de cada Activo Fijo.", vbInformation
   End If
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
End Sub

Private Sub Tx_FechaHasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub
