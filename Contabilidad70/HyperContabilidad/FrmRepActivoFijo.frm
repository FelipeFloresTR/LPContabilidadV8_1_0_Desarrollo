VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRepActivoFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Activo Fijo Tributario"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "FrmRepActivoFijo.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   11880
      TabIndex        =   34
      Top             =   0
      Width           =   8475
      Begin VB.ComboBox Cb_CGestion 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   180
         Width           =   1875
      End
      Begin VB.ComboBox Cb_CNegocio 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   180
         Width           =   1875
      End
      Begin VB.ComboBox Cb_Cuenta 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "C.Gestion:"
         Height          =   195
         Left            =   5640
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area Neg.:"
         Height          =   195
         Left            =   2880
         TabIndex        =   38
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Fr_Opciones 
      Height          =   3555
      Left            =   7620
      TabIndex        =   33
      Top             =   480
      Width           =   2415
      Begin VB.CheckBox Ch_ViewFechaProy 
         Caption         =   "Ver Fecha Proyecto"
         Height          =   195
         Left            =   300
         TabIndex        =   25
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewNombreProy 
         Caption         =   "Ver Nombre Proyecto"
         Height          =   195
         Left            =   300
         TabIndex        =   24
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewPatenteRol 
         Caption         =   "Ver Patente/Rol"
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewTipoDepHist 
         Caption         =   "Ver Tipo Dep, Histórica"
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewTipoDep 
         Caption         =   "Ver Tipo Dep, Actual"
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewValCompraHist 
         Caption         =   "Ver Valor Compra Hist."
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   240
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewCredArt33 
         Caption         =   "Ver Crédito Art. 33 bis"
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   600
         Width           =   1920
      End
      Begin VB.CheckBox Ch_ViewFVenta 
         Caption         =   "Ver  Fecha Venta/Baja"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Ch_ViewFUtiliz 
         Caption         =   "Ver  Fecha Utilización"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.CommandButton Bt_TopeUTM 
      Caption         =   "650 UTM"
      Height          =   315
      Left            =   10500
      TabIndex        =   2
      ToolTipText     =   "Asignar al tope  650 UTM"
      Top             =   7380
      Width           =   1095
   End
   Begin VB.TextBox Tx_MaxCred33 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9180
      TabIndex        =   1
      Top             =   7380
      Width           =   1155
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   7380
      Width           =   7155
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   11139
      _Version        =   393216
      Rows            =   25
      Cols            =   31
      FixedRows       =   3
      FixedCols       =   7
      WordWrap        =   -1  'True
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   28
      Top             =   0
      Width           =   11835
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
         Left            =   4740
         Picture         =   "FrmRepActivoFijo.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Valores e Índices"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Opciones 
         Caption         =   "Opciones de Vista"
         Height          =   315
         Left            =   7620
         TabIndex        =   16
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton Bt_ViewRes 
         Caption         =   "Ver Resumen"
         Height          =   315
         Left            =   9240
         TabIndex        =   26
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   215
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   5940
         TabIndex        =   14
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
         Picture         =   "FrmRepActivoFijo.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1920
         Picture         =   "FrmRepActivoFijo.frx":0812
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10500
         TabIndex        =   27
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
         Left            =   1500
         Picture         =   "FrmRepActivoFijo.frx":0CCC
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2400
         Picture         =   "FrmRepActivoFijo.frx":1173
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   4320
         Picture         =   "FrmRepActivoFijo.frx":15B8
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   3480
         Picture         =   "FrmRepActivoFijo.frx":19E1
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   3900
         Picture         =   "FrmRepActivoFijo.frx":1D7F
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   2940
         Picture         =   "FrmRepActivoFijo.frx":20E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerDoc 
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
         Left            =   960
         Picture         =   "FrmRepActivoFijo.frx":2184
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Detalle documento seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerComp 
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
         Left            =   540
         Picture         =   "FrmRepActivoFijo.frx":25F8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7020
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label Lb_Tope 
      AutoSize        =   -1  'True
      Caption         =   "Tope Cred. art. 33 bis"
      Height          =   195
      Left            =   7560
      TabIndex        =   32
      Top             =   7440
      Width           =   1530
   End
End
Attribute VB_Name = "FrmRepActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDACTFIJO = 0
Const C_TOTDEP = 1
Const C_NODEP = 2
Const C_IDDOC = 3
Const C_IDCOMP = 4
Const C_NDOC = 5
Const C_CANTIDAD = 6
Const C_DESC = 7
Const C_PATENTEROL = 8
Const C_NOMBREPROY = 9
Const C_FECHAPROY = 10
Const C_FECHA = 11
Const C_VALINIT = 12
Const C_VAL31DICANOANT = 13            'valor a reajustar
Const C_FACTACT = 14
Const C_VALREAJUSTADO = 15
Const C_TIENECREDART33 = 16
Const C_CREDART33 = 17
Const C_TIENECREDART33ANOINIT = 18
Const C_VALCRED33 = 19
Const C_VALREAJCRED = 20
Const C_TIPODEPHIST = 21
Const C_DEPACUMHIST = 22
Const C_TIPODEP = 23
Const C_DEPACUMACT = 24
Const C_FECHAUTIL = 25
Const C_VALDEPRECIAR = 26
Const C_FECHAVENTA = 27
Const C_VU_TOTAL = 28
Const C_VU_YADEP = 29
Const C_VU_DISPONRESID = 30
Const C_VU_ADEPRECIAR = 31
Const C_VU_RESIDUAL = 32
Const C_DEPMENSUAL = 33
Const C_DEPPERIODO = 34
Const C_DEPACUMULADAANO = 35
Const C_VALLIBRO = 36
Const C_LNGFECHA = 37
Const C_LNGFECHAUTIL = 38
Const C_LNGFECHAVENTA = 39
Const C_DEPLEY21210 = 40
Const C_DEPLEY21256 = 41
Const C_FMT = 42
Const C_OBLIGATORIA = 43
Const C_RESCCMMACTFIJO = 44
Const C_RESCCMMPERIODO = 45
Const C_RESCCMMDEPACUM = 46
Const C_RESDEPEJERCICIO = 47
Const C_RESCREDART33 = 48


Const NCOLS = C_RESCREDART33

Dim lFecha As Long
Dim lMsgUTM As Boolean
Dim lModTope As Boolean

Dim lViewRes As Boolean

Dim lMsgIPC As Boolean
Dim lMsgIPCCompra As Boolean


Private Sub SetUpGrid()
   Dim i As Integer
    
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid)
   
   Grid.FixedCols = C_DESC + 1
   
   Call SetupGridRes(False)
      
   Grid.ColAlignment(C_FECHA) = flexAlignLeftCenter
   Grid.ColAlignment(C_NDOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_CANTIDAD) = flexAlignRightCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   
   For i = C_VALINIT To Grid.Cols - 1
      Grid.ColAlignment(i) = flexAlignRightCenter
   Next i
         
   Grid.ColAlignment(C_TIPODEP) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODEPHIST) = flexAlignLeftCenter
        
      
   Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
    
End Sub
Private Sub SetupGridRes(ByVal ViewRes As Boolean)
   Dim i As Integer

   If Not ViewRes Then
      Grid.ColWidth(C_IDACTFIJO) = 0
      Grid.ColWidth(C_FECHA) = 800
      Grid.ColWidth(C_LNGFECHA) = 0
      Grid.ColWidth(C_IDDOC) = 0
      Grid.ColWidth(C_IDCOMP) = 0
      Grid.ColWidth(C_NODEP) = 0
      Grid.ColWidth(C_TOTDEP) = 0
      Grid.ColWidth(C_NDOC) = 2000
      Grid.ColWidth(C_CANTIDAD) = 500
      Grid.ColWidth(C_DESC) = 2500
      Grid.ColWidth(C_PATENTEROL) = 1500
      Grid.ColWidth(C_NOMBREPROY) = 2500
      Grid.ColWidth(C_FECHAPROY) = 800
      Grid.ColWidth(C_VALINIT) = 1200   '0
      Grid.ColWidth(C_VAL31DICANOANT) = 1200
      Grid.ColWidth(C_FACTACT) = 1100
      Grid.ColWidth(C_VALREAJUSTADO) = 1200
      Grid.ColWidth(C_TIENECREDART33) = 0
      Grid.ColWidth(C_CREDART33) = 1200
      Grid.ColWidth(C_TIENECREDART33ANOINIT) = 0
      Grid.ColWidth(C_VALCRED33) = 0
      Grid.ColWidth(C_VALREAJCRED) = 1200
      Grid.ColWidth(C_TIPODEPHIST) = 2210
      Grid.ColWidth(C_DEPACUMHIST) = 1200
      Grid.ColWidth(C_TIPODEP) = 2210
      Grid.ColWidth(C_DEPLEY21210) = 0
      Grid.ColWidth(C_DEPLEY21256) = 0
      Grid.ColWidth(C_DEPACUMACT) = 1200
      Grid.ColWidth(C_FECHAUTIL) = 800
      Grid.ColWidth(C_LNGFECHAUTIL) = 0
      Grid.ColWidth(C_VALDEPRECIAR) = 1200
      Grid.ColWidth(C_FECHAVENTA) = 900
      Grid.ColWidth(C_LNGFECHAVENTA) = 0
      Grid.ColWidth(C_VU_TOTAL) = 1100
      Grid.ColWidth(C_VU_YADEP) = 1100
      Grid.ColWidth(C_VU_DISPONRESID) = 1100
      Grid.ColWidth(C_VU_ADEPRECIAR) = 1100
      Grid.ColWidth(C_VU_RESIDUAL) = 1100
      Grid.ColWidth(C_DEPMENSUAL) = 1200
      Grid.ColWidth(C_DEPPERIODO) = 2145  '1200
      Grid.ColWidth(C_DEPACUMULADAANO) = 1200
      Grid.ColWidth(C_VALLIBRO) = 1200
      Grid.ColWidth(C_FMT) = 0
      Grid.ColWidth(C_OBLIGATORIA) = 0
   
      
      Grid.TextMatrix(1, C_FECHA) = "Fecha"
      Grid.TextMatrix(2, C_FECHA) = "Compra"
      
      Grid.TextMatrix(1, C_NDOC) = "Documento o"
      Grid.TextMatrix(2, C_NDOC) = "Comprobante"
      
      Grid.TextMatrix(2, C_CANTIDAD) = "Cant."
      Grid.TextMatrix(2, C_DESC) = "Descripción"
      
      Grid.TextMatrix(1, C_PATENTEROL) = "Patente,"
      Grid.TextMatrix(2, C_PATENTEROL) = "ROL o Inscr."
      
      Grid.TextMatrix(1, C_NOMBREPROY) = "Nombre"
      Grid.TextMatrix(2, C_NOMBREPROY) = "Proyecto"
      
      Grid.TextMatrix(1, C_FECHAPROY) = "Fecha"
      Grid.TextMatrix(2, C_FECHAPROY) = "Proyecto"
      
      Grid.TextMatrix(0, C_VALINIT) = "Valor Neto"
      Grid.TextMatrix(1, C_VALINIT) = "de Compra"
      Grid.TextMatrix(2, C_VALINIT) = "Histórico"
      
      Grid.TextMatrix(0, C_VAL31DICANOANT) = ""
      Grid.TextMatrix(1, C_VAL31DICANOANT) = "Valor a"
      Grid.TextMatrix(2, C_VAL31DICANOANT) = "Reajustar"
      
      Grid.TextMatrix(0, C_FACTACT) = "Factor"
      Grid.TextMatrix(1, C_FACTACT) = "Actualización"
      Grid.TextMatrix(2, C_FACTACT) = "Periodo"    '" "
      
      Grid.TextMatrix(1, C_VALREAJUSTADO) = "Valor"
      Grid.TextMatrix(2, C_VALREAJUSTADO) = "Reajustado"
         
      Grid.TextMatrix(0, C_CREDART33) = "Crédito"
      Grid.TextMatrix(1, C_CREDART33) = "Art. 33 bis"
      Grid.TextMatrix(2, C_CREDART33) = ""               'gCredArt33 * 100 & "%"
      
      Grid.TextMatrix(0, C_VALREAJCRED) = "Valor"
      Grid.TextMatrix(1, C_VALREAJCRED) = "Reajustado"
      Grid.TextMatrix(2, C_VALREAJCRED) = "Neto"
      
      Grid.TextMatrix(0, C_TIPODEPHIST) = "Tipo"
      Grid.TextMatrix(1, C_TIPODEPHIST) = "Depreciación"
      Grid.TextMatrix(2, C_TIPODEPHIST) = "Histórica"
      
      Grid.TextMatrix(0, C_DEPACUMHIST) = "Depreciación"
      Grid.TextMatrix(1, C_DEPACUMHIST) = "Acumulada"
      Grid.TextMatrix(2, C_DEPACUMHIST) = "Histórica"
      
      Grid.TextMatrix(1, C_FECHAUTIL) = "Fecha"
      Grid.TextMatrix(2, C_FECHAUTIL) = "Utiliz."
      
      Grid.TextMatrix(0, C_TIPODEP) = "Tipo"
      Grid.TextMatrix(1, C_TIPODEP) = "Depreciación"
      Grid.TextMatrix(2, C_TIPODEP) = "Actual"
      
      Grid.TextMatrix(0, C_DEPACUMACT) = "Depreciación"
      Grid.TextMatrix(1, C_DEPACUMACT) = "Acumulada"
      Grid.TextMatrix(2, C_DEPACUMACT) = "Actualizada"
      
      Grid.TextMatrix(0, C_VALDEPRECIAR) = "Valor Libro"
      Grid.TextMatrix(1, C_VALDEPRECIAR) = "Actualizado"
      Grid.TextMatrix(2, C_VALDEPRECIAR) = "a Depreciar"
      
      Grid.TextMatrix(1, C_FECHAVENTA) = "Fecha"
      Grid.TextMatrix(2, C_FECHAVENTA) = "Venta/Baja"
         
      Grid.TextMatrix(0, C_VU_TOTAL) = "Vida Útil"
      Grid.TextMatrix(1, C_VU_TOTAL) = "Total"
      Grid.TextMatrix(2, C_VU_TOTAL) = "(meses)"
      
      Grid.TextMatrix(0, C_VU_YADEP) = "Vida Útil ya"
      Grid.TextMatrix(1, C_VU_YADEP) = "Depreciada"
      Grid.TextMatrix(2, C_VU_YADEP) = "(meses)"
      
      Grid.TextMatrix(0, C_VU_DISPONRESID) = "Vida Útil"
      Grid.TextMatrix(1, C_VU_DISPONRESID) = "Disp. Residual"
      Grid.TextMatrix(2, C_VU_DISPONRESID) = "(meses)"
      
      Grid.TextMatrix(0, C_VU_ADEPRECIAR) = "Vida Útil a"
      Grid.TextMatrix(1, C_VU_ADEPRECIAR) = "Depreciar"
      Grid.TextMatrix(2, C_VU_ADEPRECIAR) = "(meses)"
      
      Grid.TextMatrix(0, C_VU_RESIDUAL) = "Vida Útil"
      Grid.TextMatrix(1, C_VU_RESIDUAL) = "Residual"
      Grid.TextMatrix(2, C_VU_RESIDUAL) = "(meses)"
      
      Grid.TextMatrix(1, C_DEPMENSUAL) = "Depreciación"
      Grid.TextMatrix(2, C_DEPMENSUAL) = "Mensual"
      
      Grid.TextMatrix(1, C_DEPPERIODO) = "Depreciación"
      Grid.TextMatrix(2, C_DEPPERIODO) = "Periodo"
      
      Grid.TextMatrix(0, C_DEPACUMULADAANO) = "Depreciación"
      Grid.TextMatrix(1, C_DEPACUMULADAANO) = "Acumulada"
      Grid.TextMatrix(2, C_DEPACUMULADAANO) = "a " & gEmpresa.Ano
      
      Grid.TextMatrix(0, C_VALLIBRO) = "Valor"
      Grid.TextMatrix(1, C_VALLIBRO) = "Libro"
      Grid.TextMatrix(2, C_VALLIBRO) = gEmpresa.Ano
      
      Grid.TextMatrix(0, C_FMT) = "" ' "          .FMT"
      
      
      'resumen

      For i = C_RESCCMMACTFIJO To C_RESCREDART33
         Grid.ColWidth(i) = 0
         Grid.TextMatrix(0, i) = ""
         Grid.TextMatrix(1, i) = ""
         Grid.TextMatrix(2, i) = ""
      Next i

   Else
   
      Grid.ColWidth(C_NDOC) = 0
      For i = C_VALINIT To C_RESCCMMACTFIJO
      
         If i <> C_TIPODEP Then
            Grid.ColWidth(i) = 0
            Grid.TextMatrix(0, i) = ""
            Grid.TextMatrix(1, i) = ""
            Grid.TextMatrix(2, i) = ""
         End If
         
      Next i
      
      Grid.ColWidth(C_VALINIT) = 1200
      Grid.ColWidth(C_RESCCMMACTFIJO) = 1200
      Grid.ColWidth(C_RESCCMMPERIODO) = 1200
      Grid.ColWidth(C_RESCCMMDEPACUM) = 1200
      Grid.ColWidth(C_RESDEPEJERCICIO) = 1200
      Grid.ColWidth(C_RESCREDART33) = 1200
      '2861733 tema 3
      Grid.ColWidth(C_VAL31DICANOANT) = 1200
      Grid.ColWidth(C_FACTACT) = 1100
      '2861733 tema 3
   
      'resumen
      Grid.TextMatrix(0, C_VALINIT) = "Valor Neto"
      Grid.TextMatrix(1, C_VALINIT) = "de Compra"
      Grid.TextMatrix(2, C_VALINIT) = "Histórico"
      
      Grid.TextMatrix(0, C_RESCCMMACTFIJO) = "Corrección"
      Grid.TextMatrix(1, C_RESCCMMACTFIJO) = "Monetaria"
      Grid.TextMatrix(2, C_RESCCMMACTFIJO) = "Act. Fijo"
      
      Grid.TextMatrix(0, C_RESCCMMPERIODO) = "Corrección"
      Grid.TextMatrix(1, C_RESCCMMPERIODO) = "Monetaria"
      Grid.TextMatrix(2, C_RESCCMMPERIODO) = "Periodo"
      
      Grid.TextMatrix(0, C_RESCCMMDEPACUM) = "Corrección"
      Grid.TextMatrix(1, C_RESCCMMDEPACUM) = "Monetaria"
      Grid.TextMatrix(2, C_RESCCMMDEPACUM) = "Dep. Acum."
         
      Grid.TextMatrix(0, C_RESDEPEJERCICIO) = ""
      Grid.TextMatrix(1, C_RESDEPEJERCICIO) = "Depreciación"
      Grid.TextMatrix(2, C_RESDEPEJERCICIO) = "Periodo"
         
      Grid.TextMatrix(0, C_RESCREDART33) = "Crédito"
      Grid.TextMatrix(1, C_RESCREDART33) = "Art. 33 bis"
      Grid.TextMatrix(2, C_RESCREDART33) = ""                  'gCredArt33 * 100 & "%"
            
      '2861733 tema 3
      Grid.TextMatrix(0, C_VAL31DICANOANT) = ""
      Grid.TextMatrix(1, C_VAL31DICANOANT) = "Valor a"
      Grid.TextMatrix(2, C_VAL31DICANOANT) = "Reajustar"

      Grid.TextMatrix(0, C_FACTACT) = "Factor"
      Grid.TextMatrix(1, C_FACTACT) = "Actualización"
      Grid.TextMatrix(2, C_FACTACT) = "Periodo"    '" "
      '2861733 tema 3
            
      Grid.TextMatrix(0, C_FMT) = "" ' "          .FMT"
            
   End If
          
   Call FGrTotales(Grid, GridTot)
       
End Sub
Private Sub bt_Cerrar_Click()
   
   Unload Me
End Sub


Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Call LP_FGr2Clip(Grid, Me.Caption & vbTab & "Año: " & gEmpresa.Ano)
'   Clip = LP_FGr2String(Grid, Me.Caption & vbTab & "Año: " & gEmpresa.Ano, False, C_OBLIGATORIA)
'
'   If Clip <> "" Then
'      Clip = Clip & FGr2String(GridTot)
'
'      Clipboard.Clear
'      Clipboard.SetText Clip
'   End If
End Sub

Private Sub Bt_Indices_Click()
   Dim Frm As FrmIPC
   
   Set Frm = New FrmIPC
   Frm.Show vbModal
   Set Frm = Nothing
   
   Call CalcDep

End Sub

Private Sub Bt_Opciones_Click()
   Fr_Opciones.visible = Not Fr_Opciones.visible

End Sub

Private Sub Bt_TopeUTM_Click()
   Dim Q1 As String
   
   Q1 = "UPDATE ParamEmpresa SET Valor = -1 WHERE Tipo='MAXCRED33'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   gMaxCred33 = -1
   
   Call LoadAll
   
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

'2861733
Private Sub Cb_CGestion_Click()
If Cb_CGestion.ListIndex < 0 Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
End Sub
'2861733
'2861733
Private Sub Cb_CNegocio_Click()
 If Cb_CNegocio.ListIndex < 0 Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub
'2861733

Private Sub Cb_Cuenta_Click()
   
   If Cb_Cuenta.ListIndex < 0 Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub


Private Sub Ch_ViewCredArt33_Click()
   
   If Ch_ViewCredArt33 = 0 Then
      Grid.ColWidth(C_CREDART33) = 0
      Grid.TextMatrix(0, C_CREDART33) = ""
      Grid.TextMatrix(1, C_CREDART33) = ""
      Grid.TextMatrix(2, C_CREDART33) = ""
      
   Else
      Grid.ColWidth(C_CREDART33) = 1200
      Grid.TextMatrix(0, C_CREDART33) = "Crédito"
      Grid.TextMatrix(1, C_CREDART33) = "Art. 33 bis"
      Grid.TextMatrix(2, C_CREDART33) = ""            'gCredArt33 * 100 & "%"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerCredArt33", Abs(Ch_ViewCredArt33.Value))
   gVarIniFile.VerCredArt33 = Abs(Ch_ViewCredArt33.Value)
   
   Call FGrTotales(Grid, GridTot)
   
End Sub

Private Sub Ch_ViewFechaProy_Click()

   If Ch_ViewFechaProy = 0 Then
      Grid.ColWidth(C_FECHAPROY) = 0
      Grid.TextMatrix(0, C_FECHAPROY) = ""
      Grid.TextMatrix(1, C_FECHAPROY) = ""
      Grid.TextMatrix(2, C_FECHAPROY) = ""
      
   Else
      Grid.ColWidth(C_FECHAPROY) = 1500
      Grid.TextMatrix(0, C_FECHAPROY) = ""
      Grid.TextMatrix(1, C_FECHAPROY) = "Fecha"
      Grid.TextMatrix(2, C_FECHAPROY) = "Proyecto"
        
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerFechaProy", Abs(Ch_ViewFechaProy.Value))
   gVarIniFile.VerFechaProy = Abs(Ch_ViewFechaProy.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewFUtiliz_Click()

   If Ch_ViewFUtiliz = 0 Then
      Grid.ColWidth(C_FECHAUTIL) = 0
      Grid.TextMatrix(0, C_FECHAUTIL) = ""
      Grid.TextMatrix(1, C_FECHAUTIL) = ""
      Grid.TextMatrix(2, C_FECHAUTIL) = ""
      
   Else
      Grid.ColWidth(C_FECHAUTIL) = 800
      Grid.TextMatrix(1, C_FECHAUTIL) = "Fecha"
      Grid.TextMatrix(2, C_FECHAUTIL) = "Utiliz."
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerFUtiliz", Abs(Ch_ViewFUtiliz.Value))
   gVarIniFile.VerFUtiliz = Abs(Ch_ViewFUtiliz.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewFVenta_Click()

   If Ch_ViewFVenta = 0 Then
      Grid.ColWidth(C_FECHAVENTA) = 0
      Grid.TextMatrix(0, C_FECHAVENTA) = ""
      Grid.TextMatrix(1, C_FECHAVENTA) = ""
      Grid.TextMatrix(2, C_FECHAVENTA) = ""
      
   Else
      Grid.ColWidth(C_FECHAVENTA) = 900
      Grid.TextMatrix(1, C_FECHAVENTA) = "Fecha"
      Grid.TextMatrix(2, C_FECHAVENTA) = "Venta/Baja"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerFVenta", Abs(Ch_ViewFVenta.Value))
   gVarIniFile.VerFVenta = Abs(Ch_ViewFVenta.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewNombreProy_Click()

   If Ch_ViewNombreProy = 0 Then
      Grid.ColWidth(C_NOMBREPROY) = 0
      Grid.TextMatrix(0, C_NOMBREPROY) = ""
      Grid.TextMatrix(1, C_NOMBREPROY) = ""
      Grid.TextMatrix(2, C_NOMBREPROY) = ""
      
   Else
      Grid.ColWidth(C_NOMBREPROY) = 1500
      Grid.TextMatrix(0, C_NOMBREPROY) = ""
      Grid.TextMatrix(1, C_NOMBREPROY) = "Nombre,"
      Grid.TextMatrix(2, C_NOMBREPROY) = "Proyecto"
        
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerNombreProy", Abs(Ch_ViewNombreProy.Value))
   gVarIniFile.VerNombreProy = Abs(Ch_ViewNombreProy.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewPatenteRol_Click()

   If Ch_ViewPatenteRol = 0 Then
      Grid.ColWidth(C_PATENTEROL) = 0
      Grid.TextMatrix(0, C_PATENTEROL) = ""
      Grid.TextMatrix(1, C_PATENTEROL) = ""
      Grid.TextMatrix(2, C_PATENTEROL) = ""
      
   Else
      Grid.ColWidth(C_PATENTEROL) = 1500
      Grid.TextMatrix(0, C_PATENTEROL) = ""
      Grid.TextMatrix(1, C_PATENTEROL) = "Patente,"
      Grid.TextMatrix(2, C_PATENTEROL) = "ROL o Inscr."
        
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerPatenteRol", Abs(Ch_ViewPatenteRol.Value))
   gVarIniFile.VerPatenteRol = Abs(Ch_ViewPatenteRol.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewTipoDep_Click()

   If Ch_ViewTipoDep = 0 Then
      Grid.ColWidth(C_TIPODEP) = 0
      Grid.TextMatrix(0, C_TIPODEP) = ""
      Grid.TextMatrix(1, C_TIPODEP) = ""
      Grid.TextMatrix(2, C_TIPODEP) = ""
      
   Else
      Grid.ColWidth(C_TIPODEP) = 2210
      Grid.TextMatrix(0, C_TIPODEP) = "Tipo"
      Grid.TextMatrix(1, C_TIPODEP) = "Depreciación"
      Grid.TextMatrix(2, C_TIPODEP) = "Actual"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerTipoDep", Abs(Ch_ViewTipoDep.Value))
   gVarIniFile.VerTipoDep = Abs(Ch_ViewTipoDep.Value)

   Call FGrTotales(Grid, GridTot)

End Sub

Private Sub Ch_ViewTipoDepHist_Click()

   If Ch_ViewTipoDepHist = 0 Then
      Grid.ColWidth(C_TIPODEPHIST) = 0
      Grid.TextMatrix(0, C_TIPODEPHIST) = ""
      Grid.TextMatrix(1, C_TIPODEPHIST) = ""
      Grid.TextMatrix(2, C_TIPODEPHIST) = ""
      
   Else
      Grid.ColWidth(C_TIPODEPHIST) = 2210
      Grid.TextMatrix(0, C_TIPODEPHIST) = "Tipo"
      Grid.TextMatrix(1, C_TIPODEPHIST) = "Depreciación"
      Grid.TextMatrix(2, C_TIPODEPHIST) = "Histórica"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerTipoDepHist", Abs(Ch_ViewTipoDepHist.Value))
   gVarIniFile.VerTipoDepHist = Abs(Ch_ViewTipoDepHist.Value)
   
   Call FGrTotales(Grid, GridTot)
   

End Sub

Private Sub Ch_ViewValCompraHist_Click()

   If Ch_ViewValCompraHist = 0 Then
      Grid.ColWidth(C_VALINIT) = 0
      Grid.TextMatrix(0, C_VALINIT) = ""
      Grid.TextMatrix(1, C_VALINIT) = ""
      Grid.TextMatrix(2, C_VALINIT) = ""
      
   Else
      Grid.ColWidth(C_VALINIT) = 1200
      Grid.TextMatrix(0, C_VALINIT) = "Valor Neto"
      Grid.TextMatrix(1, C_VALINIT) = "de Compra"
      Grid.TextMatrix(2, C_VALINIT) = "Histórico"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerValCompraHist", Abs(Ch_ViewValCompraHist.Value))
   gVarIniFile.VerValCompraHist = Abs(Ch_ViewValCompraHist.Value)
   
   Call FGrTotales(Grid, GridTot)
End Sub

Private Sub Form_Load()
   Dim Q1 As String

   Fr_Opciones.visible = False
   
   Call BtFechaImg(Bt_Fecha)
   lFecha = DateSerial(gEmpresa.Ano, 12, 31)
   Call SetTxDate(Tx_Fecha, lFecha)
   
   Call CbAddItem(Cb_Cuenta, " ", 0)
   Q1 = "SELECT Descripcion, IdCuenta FROM Cuentas WHERE Atrib" & ATRIB_ACTIVOFIJO & "<> 0 AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(Cb_Cuenta, DbMain, Q1, 0, True)
   Cb_Cuenta.ListIndex = -1    'para que no entre al la función Cb_Cuenta_Click
   
    '2861733 tema 2
   Q1 = "SELECT Descripcion,idCCosto FROM CentroCosto WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbAddItem(Cb_CGestion, " ", 0)
   Call FillCombo(Cb_CGestion, DbMain, Q1, 0, True)
   Cb_CGestion.ListIndex = -1

   Q1 = "SELECT Descripcion, idAreaNegocio FROM AreaNegocio WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbAddItem(Cb_CNegocio, " ", 0)
   Call FillCombo(Cb_CNegocio, DbMain, Q1, 0, True)
   Cb_CNegocio.ListIndex = -1
   '2861733 tema 2

   

   Call SetUpGrid
   
   Ch_ViewValCompraHist = gVarIniFile.VerValCompraHist
   Ch_ViewCredArt33 = gVarIniFile.VerCredArt33
   Ch_ViewFVenta = gVarIniFile.VerFVenta
   Ch_ViewFUtiliz = gVarIniFile.VerFUtiliz
   Ch_ViewTipoDep = gVarIniFile.VerTipoDep
   Ch_ViewTipoDepHist = gVarIniFile.VerTipoDepHist
   Ch_ViewPatenteRol = gVarIniFile.VerPatenteRol
   Ch_ViewNombreProy = gVarIniFile.VerNombreProy
   Ch_ViewFechaProy = gVarIniFile.VerFechaProy
   
   Call Ch_ViewValCompraHist_Click
   Call Ch_ViewCredArt33_Click
   Call Ch_ViewFVenta_Click
   Call Ch_ViewFUtiliz_Click
   Call Ch_ViewTipoDep_Click
   Call Ch_ViewTipoDepHist_Click
   Call Ch_ViewPatenteRol_Click
   Call Ch_ViewNombreProy_Click
   Call Ch_ViewFechaProy_Click
   
   Bt_TopeUTM.Caption = gMaxUTMCred33 & " UTM"
   Bt_TopeUTM.ToolTipText = "Asignar al tope" & gMaxUTMCred33 & " UTM"

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
   DoEvents
   
   If gEmpresa.Ano >= 2012 Then
      MsgBox1 "ATENCIÓN: " & vbCrLf & vbCrLf & "Recuerde configurar % Crédito de Art. 33 bis dentro de la " & vbCrLf & "Configuración Inicial - Configurar Impuestos", vbInformation + vbOKOnly
   End If

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
   Lb_Tope.Top = Tx_CurrCell.Top
   Tx_MaxCred33.Top = Tx_CurrCell.Top
   Bt_TopeUTM.Top = Tx_CurrCell.Top
   'Tx_CurrCell.Width = GridTot.Width
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

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
Public Function FView()
   Me.Show vbModal
   
End Function

Private Function LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim j As Integer
   Dim Total(NCOLS) As Double
   Dim SubTotal(NCOLS) As Double
   Dim IdCuenta As Long
   Dim Row As Integer
   Dim TotCred33 As Double
   Dim Fecha As Long
   Dim ValMoneda As Double
   Dim mDiff As Integer, DtCompra As Long
   
   Fecha = GetTxDate(Tx_Fecha)
   
   If gMaxCred33 >= 0 Then
      Tx_MaxCred33 = Format(gMaxCred33, NUMFMT)
      
   Else
      If Fecha < DateSerial(gEmpresa.Ano, 12, 31) Then
         If GetValMoneda("UTM", ValMoneda, Fecha, True) = True Then
            Tx_MaxCred33 = Format(ValMoneda * 650, NUMFMT)
         End If
      Else
         Tx_MaxCred33 = Format(gMaxUTMCred33_Pesos, NUMFMT)
      End If
      
'      If Not lMsgUTM Then    'se hace en FrmMain antes de llamar a este form
'         If gMaxUTMCred33_Pesos = 0 Then
'            MsgBox1 "No se ha ingresado el valor de la UTM. Este valor se utiliza para calcular el máximo para Crédito Art. 33 bis", vbExclamation + vbOKOnly
'         Else
'            MsgBox1 "Revise si el último valor de la UTM ingresado en el sistema está actualizado.", vbInformation
'         End If
'         lMsgUTM = True
'      End If
   End If
      
   
   Q1 = "SELECT IdActFijo, MovActivoFijo.IdDoc, MovActivoFijo.IdComp, IdMovComp, TipoMovAF, MovActivoFijo.Fecha, MovActivoFijo.FechaUtilizacion, MovActivoFijo.FechaVentaBaja "
   Q1 = Q1 & ", MovActivoFijo.Cantidad, MovActivoFijo.Descrip, MovActivoFijo.PatenteRol, MovActivoFijo.NombreProy, MovActivoFijo.FechaProy "
   Q1 = Q1 & ", MovActivoFijo.Neto, MovActivoFijo.IVA, Cred4Porc, Cred4PorcAnoInit, NoDepreciable, ValCred33, ValReajustadoNetoAnt "
   Q1 = Q1 & ", VidaUtil, DepNormal, DepAcelerada, DepInstant, DepDecimaparte, DepDecimaParte2, TipoDep "
   Q1 = Q1 & ", DepNormalHist, DepAceleradaHist, DepInstantHist, DepDecimaParteHist, DepDecimaParte2Hist, TipoDepHist, DepAcumHist "
   Q1 = Q1 & ", TipoDepLey21210, TipoDepLey21210Hist, DepLey21256, DepLey21256Hist , DepLey21256Hist, MovActivoFijo.IdCuenta "
   Q1 = Q1 & ", Documento.NumDoc, Documento.TipoLib, Documento.TipoDoc, Entidades.Rut, Entidades.Nombre "
   Q1 = Q1 & ", Comprobante.Correlativo, Comprobante.Tipo "
   Q1 = Q1 & ", Cuentas.Codigo, Cuentas.Descripcion as DescCta, TotalmenteDepreciado "
   Q1 = Q1 & "  FROM ((((MovActivoFijo "
   Q1 = Q1 & "  LEFT JOIN Documento ON MovActivoFijo.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovActivoFijo") & " )"
   Q1 = Q1 & "  LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True) & " )"
   Q1 = Q1 & "  LEFT JOIN Cuentas ON MovActivoFijo.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovActivoFijo") & " )"
   Q1 = Q1 & "  LEFT JOIN Comprobante ON MovActivoFijo.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovActivoFijo") & " )"
   Q1 = Q1 & "  LEFT JOIN MovComprobante ON MovActivoFijo.IdMovComp = MovComprobante.IdMov "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "MovActivoFijo")
   Q1 = Q1 & "  WHERE MovActivoFijo.Fecha <=" & Fecha
   Q1 = Q1 & "  AND (FechaVentaBaja = 0 OR FechaVentaBaja IS NULL OR FechaVentaBaja > " & Fecha & ")"
   Q1 = Q1 & "  AND MovActivoFijo.IdEmpresa = " & gEmpresa.id & " AND MovActivoFijo.Ano = " & gEmpresa.Ano
   If CbItemData(Cb_Cuenta) > 0 Then
      Q1 = Q1 & "  AND MovActivoFijo.IdCuenta = " & CbItemData(Cb_Cuenta)
   End If
   '2861733
   If CbItemData(Cb_CNegocio) > 0 Then
      Q1 = Q1 & "  AND MovActivoFijo.IdAreaNeg = " & CbItemData(Cb_CNegocio)
   End If
   '2861733
   '2861733
   If CbItemData(Cb_CGestion) > 0 Then
      Q1 = Q1 & "  AND MovActivoFijo.idCCosto = " & CbItemData(Cb_CGestion)
   End If
   '2861733
   Q1 = Q1 & "  ORDER BY Cuentas.Descripcion, MovActivoFijo.Fecha, MovActivoFijo.Descrip"

   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   
   Grid.rows = Grid.FixedRows
   Row = Grid.FixedRows
   
   IdCuenta = 0
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If
      
      If IdCuenta <> vFld(Rs("IdCuenta")) Then
      
'         If IdCuenta <> 0 Then   'ponemos total cuenta anterior (se elimina porque si no, no pone total de activos fijos que no tienen cuenta asociada FCA 19/12/17)
         If Row > Grid.FixedRows Then              'para que no ponga una línea de total sin ningún activo fijo antes
            Call FGrSetRowStyle(Grid, Row, "B")
            Grid.TextMatrix(Row, C_FMT) = "B"
            Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
            
            Grid.TextMatrix(Row, C_DESC) = "Total"
            Grid.TextMatrix(Row, C_VALINIT) = Format(SubTotal(C_VALINIT), NUMFMT)
            
            For j = 0 To NCOLS
               SubTotal(j) = 0
            Next j
                        
            Grid.rows = Grid.rows + 2
            Row = Row + 2
            Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"

         End If
         
         'nueva cuenta
         Grid.TextMatrix(Row, C_FECHA) = ""   ' "Cuenta"
         Grid.TextMatrix(Row, C_NDOC) = FmtCodCuenta(vFld(Rs("Codigo")))
         Grid.TextMatrix(Row, C_DESC) = vFld(Rs("DescCta"), True)
         Call FGrSetRowStyle(Grid, Row, "B")
         Grid.TextMatrix(Row, C_FMT) = "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         
         IdCuenta = vFld(Rs("IdCuenta"))
         Grid.rows = Grid.rows + 1
         Row = Row + 1
         
      End If

      'datos activo fijo
      Grid.TextMatrix(Row, C_IDACTFIJO) = vFld(Rs("IdActFijo"))
      Grid.TextMatrix(Row, C_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
      Grid.TextMatrix(Row, C_LNGFECHA) = vFld(Rs("Fecha"))
      DtCompra = vFmt(Grid.TextMatrix(Row, C_LNGFECHA))

      Grid.TextMatrix(Row, C_FECHAUTIL) = IIf(vFld(Rs("FechaUtilizacion")) > 0, Format(vFld(Rs("FechaUtilizacion")), SDATEFMT), "")
      Grid.TextMatrix(Row, C_LNGFECHAUTIL) = vFld(Rs("FechaUtilizacion"))
      
      Grid.TextMatrix(Row, C_NODEP) = vFld(Rs("NoDepreciable"))
      Grid.TextMatrix(Row, C_TOTDEP) = vFld(Rs("TotalmenteDepreciado"))
      
      If Year(vFld(Rs("Fecha"))) = gEmpresa.Ano Then
         Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("IdDoc"))
         Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("IdComp"))
         
         If vFld(Rs("IdDoc")) > 0 Then
            Grid.TextMatrix(Row, C_NDOC) = "[" & GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc")) & "]   " & FmtCID(vFld(Rs("Rut"))) & " " & vFld(Rs("Nombre"), True)
         ElseIf vFld(Rs("IdComp")) > 0 Then
            Grid.TextMatrix(Row, C_NDOC) = "Comprobante " & UCase(Left(gTipoComp(vFld(Rs("Tipo"))), 1)) & "-" & vFld(Rs("Correlativo"))
         End If
      End If
      
      Grid.TextMatrix(Row, C_CANTIDAD) = Format(vFld(Rs("Cantidad")), NUMFMT)
      
      'ADO 2747741 Quita doble comilla y salto de linea de la descripcion
      Grid.TextMatrix(Row, C_DESC) = Replace(Replace(vFld(Rs("Descrip")), Chr(34), ""), Chr(10), " ")
      
      Grid.TextMatrix(Row, C_PATENTEROL) = vFld(Rs("PatenteRol"))
      Grid.TextMatrix(Row, C_NOMBREPROY) = vFld(Rs("NombreProy"))
      Grid.TextMatrix(Row, C_FECHAPROY) = IIf(vFld(Rs("FechaProy")) <> 0, Format(vFld(Rs("FechaProy")), SDATEFMT), "")
      
      Grid.TextMatrix(Row, C_VALINIT) = Format(vFld(Rs("Neto")), NUMFMT)
      
      If Not IsNull(Rs("ValReajustadoNetoAnt")) Then
         Grid.TextMatrix(Row, C_VAL31DICANOANT) = Format(vFld(Rs("ValReajustadoNetoAnt")), NUMFMT)
      End If
            
      Grid.TextMatrix(Row, C_TIENECREDART33) = IIf(vFld(Rs("Cred4Porc")) <> 0, "1", "0")
      Grid.TextMatrix(Row, C_TIENECREDART33ANOINIT) = IIf(vFld(Rs("Cred4PorcAnoInit")) <> 0, "1", "0")
      Grid.TextMatrix(Row, C_VALCRED33) = vFld(Rs("ValCred33"))
      
      If vFld(Rs("TipoDepLey21210Hist")) <> 0 Then
         Grid.TextMatrix(Row, C_TIPODEPHIST) = gTipoDepLey21210Str(vFld(Rs("TipoDepLey21210Hist")))
      End If
      
      If vFld(Rs("DepLey21256Hist")) <> 0 Then
         Grid.TextMatrix(Row, C_TIPODEPHIST) = "Inst.e Inmed. - " & gTipoDepLey21256Str
      End If
      
      If vFld(Rs("TipoDepHist")) <> 0 And vFld(Rs("DepLey21256Hist")) = 0 Then
         If Grid.TextMatrix(Row, C_TIPODEPHIST) = "" Then
            Grid.TextMatrix(Row, C_TIPODEPHIST) = gTipoDepStr(vFld(Rs("TipoDepHist")))
         Else
            Grid.TextMatrix(Row, C_TIPODEPHIST) = Grid.TextMatrix(Row, C_TIPODEPHIST) & " + " & gTipoDepStr(vFld(Rs("TipoDepHist")))
         End If
      End If
      
      Grid.TextMatrix(Row, C_DEPACUMHIST) = Format(vFld(Rs("DepAcumHist")), NUMFMT)
      
      If vFld(Rs("TipoDepLey21210")) <> 0 Then
         Grid.TextMatrix(Row, C_TIPODEP) = gTipoDepLey21210Str(vFld(Rs("TipoDepLey21210")))
         Grid.TextMatrix(Row, C_DEPLEY21210) = vFld(Rs("TipoDepLey21210"))
      End If
      
      If vFld(Rs("DepLey21256")) <> 0 Then
         Grid.TextMatrix(Row, C_TIPODEP) = "Inst.e Inmed. - " & gTipoDepLey21256Str
      End If
      
      If vFld(Rs("TipoDep")) <> 0 And vFld(Rs("DepLey21256")) = 0 Then
         If Grid.TextMatrix(Row, C_TIPODEP) = "" Then
            Grid.TextMatrix(Row, C_TIPODEP) = gTipoDepStr(vFld(Rs("TipoDep")))
         Else
            Grid.TextMatrix(Row, C_TIPODEP) = Grid.TextMatrix(Row, C_TIPODEP) & " + " & gTipoDepStr(vFld(Rs("TipoDep")))
         End If
      End If
     
      If vFld(Rs("FechaVentaBaja")) <> 0 Then
         Grid.TextMatrix(Row, C_FECHAVENTA) = Format(vFld(Rs("FechaVentaBaja")), SDATEFMT)
         Grid.TextMatrix(Row, C_LNGFECHAVENTA) = vFld(Rs("FechaVentaBaja"))
      End If
      
      Grid.TextMatrix(Row, C_VU_TOTAL) = Format(vFld(Rs("VidaUtil")), NUMFMT)
      
      Select Case vFld(Rs("TipoDepHist"))
         Case DEP_NORMAL
            Grid.TextMatrix(Row, C_VU_YADEP) = Format(vFld(Rs("DepNormalHist")), NUMFMT)
         Case DEP_ACELERADA
            Grid.TextMatrix(Row, C_VU_YADEP) = Format(vFld(Rs("DepAceleradaHist")), NUMFMT)
         Case DEP_INSTANTANEA
            Grid.TextMatrix(Row, C_VU_YADEP) = Format(vFld(Rs("DepInstantHist")), NUMFMT)
         Case DEP_DECIMAPARTE
            Grid.TextMatrix(Row, C_VU_YADEP) = Format(vFld(Rs("DepDecimaParteHist")), NUMFMT)
         Case DEP_DECIMAPARTE2
            Grid.TextMatrix(Row, C_VU_YADEP) = Format(vFld(Rs("DepDecimaParte2Hist")), NUMFMT)
      End Select
      
      If vFld(Rs("FechaVentaBaja")) = 0 Or vFld(Rs("FechaVentaBaja")) > lFecha Then
         Grid.TextMatrix(Row, C_VU_DISPONRESID) = Format(vFmt(Grid.TextMatrix(Row, C_VU_TOTAL)) - vFmt(Grid.TextMatrix(Row, C_VU_YADEP)), NUMFMT)
      Else
         Grid.TextMatrix(Row, C_VU_DISPONRESID) = 0
      End If
      
      If vFld(Rs("FechaVentaBaja")) = 0 Or vFld(Rs("FechaVentaBaja")) > lFecha Then
         
         If vFld(Rs("TipoDepLey21210")) = DEP_LEY21210_ARAUCANIA Or vFld(Rs("DepLey21256")) <> 0 Then
            Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("VidaUtil")), NUMFMT)
         Else
         
            Select Case vFld(Rs("TipoDep"))
               Case DEP_NORMAL
                  Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("DepNormal")), NUMFMT)
                  
                  If Year(DtCompra) = gEmpresa.Ano Then
                     mDiff = DateDiff("m", DtCompra, lFecha) + 1
                  Else
                     mDiff = DateDiff("m", DateSerial(gEmpresa.Ano, 1, 1), lFecha) + 1
                  End If
                  If mDiff < vFmt(Grid.TextMatrix(Row, C_VU_ADEPRECIAR)) Then
                     Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(mDiff, NUMFMT)
                  End If
                     
               Case DEP_ACELERADA
                  Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("DepAcelerada")), NUMFMT)
               Case DEP_INSTANTANEA
                  Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("DepInstant")), NUMFMT)
               Case DEP_DECIMAPARTE
                  Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("DepDecimaParte")), NUMFMT)
               Case DEP_DECIMAPARTE2
                  Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = Format(vFld(Rs("DepDecimaParte2")), NUMFMT)
            End Select
            
         End If
         
      Else
         Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = 0
         
      End If
      
      ' **** ADO 2747741 Tema 2
      If Year(vFld(Rs("FechaUtilizacion"))) > gEmpresa.Ano Then
        Grid.TextMatrix(Row, C_VU_ADEPRECIAR) = 0
        
        Q1 = "UPDATE MovActivoFijo "
        Q1 = Q1 & " Set DepNormal = 0 "
        Q1 = Q1 & " Where IdActFijo = " & vFld(Rs("IdActFijo")) & " And IdEmpresa = " & gEmpresa.id & " And Ano = " & gEmpresa.Ano
        Call ExecSQL(DbMain, Q1)
        
      End If
      '********** Fin *************
      
      
      If vFld(Rs("FechaVentaBaja")) = 0 Or vFld(Rs("FechaVentaBaja")) > lFecha Then
         Grid.TextMatrix(Row, C_VU_RESIDUAL) = Format(vFmt(Grid.TextMatrix(Row, C_VU_DISPONRESID)) - vFmt(Grid.TextMatrix(Row, C_VU_ADEPRECIAR)), NUMFMT)
      Else
         Grid.TextMatrix(Row, C_VU_RESIDUAL) = 0
      End If
      
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      
      SubTotal(C_VALINIT) = SubTotal(C_VALINIT) + vFld(Rs("Neto"))
      Total(C_VALINIT) = Total(C_VALINIT) + vFld(Rs("Neto"))
      
      Row = Row + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   'ponemos el total de la última cuenta
   
   If Grid.rows > Grid.FixedRows Then
      If IdCuenta <> 0 Or (IdCuenta = 0 And Grid.TextMatrix(Grid.FixedRows, C_FECHA) <> "") Then
         
         Grid.rows = Grid.rows + 1
         
         Call FGrSetRowStyle(Grid, Row, "B")
         Grid.TextMatrix(Row, C_FMT) = "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         
         Grid.TextMatrix(Row, C_DESC) = "Total"
         Grid.TextMatrix(Row, C_VALINIT) = Format(SubTotal(C_VALINIT), NUMFMT)
         
      End If
   End If
   
   GridTot.TextMatrix(0, C_DESC) = "TOTAL"
   GridTot.TextMatrix(0, C_VALINIT) = Format(Total(C_VALINIT), NUMFMT)
   
   Call FGrVRows(Grid)
   
   Call CalcDep
   
   Grid.Redraw = True

End Function

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   
   If ValidaCred(True) = False Then
      Exit Sub
   End If
      
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
   
   If ValidaCred(True) = False Then
      Exit Sub
   End If
   
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
   
   If Not lViewRes Then
   
      NumWi = 1150
      
      ColWi(C_FECHA) = 750 - 50
      If ColWi(C_FECHAUTIL) > 0 Then
         ColWi(C_FECHAUTIL) = 750 - 60
      End If
      ColWi(C_NDOC) = 0    'ColWi(C_NDOC) - 700
      ColWi(C_CANTIDAD) = 450 - 100
      ColWi(C_DESC) = 2500 - 300
      ColWi(C_VAL31DICANOANT) = NumWi - 30
      ColWi(C_FACTACT) = ColWi(C_FACTACT) - 400
      ColWi(C_VALREAJUSTADO) = NumWi - 60
      If ColWi(C_CREDART33) > 0 Then
         ColWi(C_CREDART33) = NumWi - 200 - 200
      End If
      ColWi(C_VALREAJCRED) = NumWi - 60
      ColWi(C_DEPACUMHIST) = 0
      ColWi(C_DEPACUMACT) = NumWi
      ColWi(C_VALDEPRECIAR) = NumWi
      ColWi(C_FECHAVENTA) = 0
      ColWi(C_VU_TOTAL) = NumWi - 300 - 100
      ColWi(C_VU_YADEP) = NumWi - 200 - 100
      ColWi(C_VU_DISPONRESID) = 0
      ColWi(C_VU_ADEPRECIAR) = NumWi - 200 - 100
      ColWi(C_VU_RESIDUAL) = NumWi - 200 - 100
      ColWi(C_DEPMENSUAL) = 0
      ColWi(C_DEPPERIODO) = NumWi - 100
      ColWi(C_DEPACUMULADAANO) = 0
      ColWi(C_VALLIBRO) = NumWi
   End If
   
   For i = 0 To Grid.Cols - 1
      If ColWi(i) > 0 Then
         ColWi(i) = ColWi(i) * 0.92
      End If
   Next i
      
   gPrtReportes.GrFontSize = 7
   gPrtReportes.GrFontName = "Arial"
               
   Total(C_DESC) = "Total"
   For i = C_VALINIT To Grid.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.NTotLines = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SaveVal

End Sub

Private Sub Grid_Click()

   If Fr_Opciones.visible Then
      Fr_Opciones.visible = False
   End If
   
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

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
   
   If Col = C_NDOC Then
      If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
         Set Frm = New FrmDoc
         Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
         Set Frm = Nothing
      ElseIf Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
         Set Frm = New FrmComprobante
         Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
         Set Frm = Nothing
      End If
   Else
      Call PostClick(Bt_VerActivoFijo)
   End If
   
End Sub
Private Sub Grid_SelChange()
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)

End Sub
Private Sub Bt_VerComp_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmComprobante
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
      Set Frm = Nothing
   End If

End Sub
Private Sub Bt_VerDoc_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmDoc
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      Set Frm = New FrmDoc
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
      Set Frm = Nothing
   End If

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

Private Sub CalcDep()
   Dim Valor As Double
   Dim i As Integer, j As Integer
   Dim DtCompra As Long, DtUtil As Long
   Dim Año As Integer
   Dim SubTotal(NCOLS) As Double
   Dim Total(NCOLS) As Double
   Dim IdActFijo As Long
   Dim NoCalcDep As Boolean
   Dim TotCred33 As Double
   Dim FechaHasta As Long
   Dim Factor As Double
   Dim ValorReajustado As Double
   Dim FactorCompuesto As Double, FactorPeriodo As Double
   Dim Valor31DicAnoAnt As Double
   Dim NoCalcCCMM As Boolean
   Dim DepAcumHist As Double
   Dim DepMensual As Double
   Dim DepPeriodo As Double
     
   
   IdActFijo = 0
   FechaHasta = GetTxDate(Tx_Fecha)
   
   lMsgIPC = False
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      DepAcumHist = vFmt(Grid.TextMatrix(i, C_DEPACUMHIST))
               
      If vFmt(Grid.TextMatrix(i, C_IDACTFIJO)) > 0 Then
      
         IdActFijo = vFmt(Grid.TextMatrix(i, C_IDACTFIJO))
                  
         DtCompra = vFmt(Grid.TextMatrix(i, C_LNGFECHA))
         DtUtil = vFmt(Grid.TextMatrix(i, C_LNGFECHAUTIL))
         
         'no se calcula depreciación si el activo fijo es NO Depreciable o está Totalmente Depreciado
         NoCalcDep = Val(Grid.TextMatrix(i, C_NODEP)) <> 0 Or Val(Grid.TextMatrix(i, C_TOTDEP)) <> 0
         
         'se calcula depreciación sólo si no se vende o se vende después de la "Fecha Hasta" del reporte
         NoCalcDep = NoCalcDep Or Not (vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) = 0 Or vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) > lFecha)
         
         If NoCalcDep Then
            Grid.TextMatrix(i, C_TIPODEP) = ""
         End If

         
'        'no se calcula depreciación si la fecha de compra y de utilización es el 31 dic del año
'        '(esto se controla en la ventana de ingreso de activo fijo, para que el usuario tenga la opción de cambiarlo)
'         NoCalcDep = NoCalcDep Or DtCompra = DateSerial(gEmpresa.Ano, 12, 31)
            
         'calculamos el factor de acuerdo al mes de compra y reajustamos el valor
         
         'No se calcula CCMM si el activo fijo está totalmente depreciado o
         'se vende en el año y el reporte es al 31 de dic del año
         
         NoCalcCCMM = False
         If Val(Grid.TextMatrix(i, C_TOTDEP)) <> 0 Or (vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) <> 0 And FechaHasta = DateSerial(gEmpresa.Ano, 12, 31)) Then
            NoCalcCCMM = True
         End If
         
         
         ValorReajustado = CalcValReajustado(vFmt(Grid.TextMatrix(i, C_VALINIT)), DtCompra, FactorCompuesto, FactorPeriodo, Valor31DicAnoAnt)
            
         'If Val(Grid.TextMatrix(i, C_TOTDEP)) <> 0 Or (vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) <> 0 And FechaHasta = DateSerial(gEmpresa.Ano, 12, 31)) Then
         If NoCalcCCMM Then
                  
            If Grid.TextMatrix(i, C_VU_DISPONRESID) = 0 And vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) = 0 Then   'el activo fijo fue totalmente depreciado el año anterior y quedó con valor 1
               Grid.TextMatrix(i, C_FACTACT) = Format(0, DBLFMT3)
               Grid.TextMatrix(i, C_VAL31DICANOANT) = 0
               Grid.TextMatrix(i, C_VALREAJUSTADO) = 0
               Grid.TextMatrix(i, C_DEPACUMHIST) = 0
               Grid.TextMatrix(i, C_VU_TOTAL) = 0
               Grid.TextMatrix(i, C_VU_YADEP) = 0
               Grid.TextMatrix(i, C_FECHAUTIL) = ""
               NoCalcDep = True
               
            Else
               Grid.TextMatrix(i, C_FACTACT) = Format(1, DBLFMT3)
'               Grid.TextMatrix(i, C_VAL31DICANOANT) = Grid.TextMatrix(i, C_VALINIT)         'Solicitado por Joshua Nicolás Catrín el 26 de febrero 2018
               Grid.TextMatrix(i, C_VAL31DICANOANT) = Format(Valor31DicAnoAnt, NUMFMT)
               Grid.TextMatrix(i, C_VALREAJUSTADO) = Grid.TextMatrix(i, C_VAL31DICANOANT)
            End If
           
         Else
            'ValorReajustado = CalcValReajustado(vFmt(Grid.TextMatrix(i, C_VALINIT)), DtCompra, FactorCompuesto, FactorPeriodo, Valor31DicAnoAnt)
            
            Grid.TextMatrix(i, C_FACTACT) = Format(FactorPeriodo, DBLFMT3)
            'Grid.TextMatrix(i, C_FACTACT) = Format(FactorCompuesto, DBLFMT3)
            
            If Grid.TextMatrix(i, C_VAL31DICANOANT) <> "" Then      'viene del año anterior, usamos ese
               Grid.TextMatrix(i, C_VALREAJUSTADO) = Format(vFmt(Grid.TextMatrix(i, C_VAL31DICANOANT)) * vFmt(Grid.TextMatrix(i, C_FACTACT)), NUMFMT)
               
            Else    'se usa el calculado, que no funciona cuando el primer año se aplicó 33 bis
               Grid.TextMatrix(i, C_VAL31DICANOANT) = Format(Valor31DicAnoAnt, NUMFMT)
               Grid.TextMatrix(i, C_VALREAJUSTADO) = Format(ValorReajustado, NUMFMT)
               
            End If
            
         End If
                           
         'aplicamos crédito 33 BIS si corresponde (sólo el año de compra)
         If NoCalcDep Or vFmt(Tx_MaxCred33) = 0 Or Year(DtCompra) <> gEmpresa.Ano Then
            Grid.TextMatrix(i, C_CREDART33) = 0
            
         ElseIf vFmt(Grid.TextMatrix(i, C_VALCRED33)) < 0 Then
            If vFmt(Grid.TextMatrix(i, C_LNGFECHAUTIL)) < gFechaInicioDepInstantanea Then
               Grid.TextMatrix(i, C_CREDART33) = Format(vFmt(Grid.TextMatrix(i, C_TIENECREDART33)) * vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) * gCredArt33, NUMFMT)
            
            ElseIf vFmt(Grid.TextMatrix(i, C_LNGFECHAUTIL)) >= gFechaInicioDepInstantanea And vFmt(Grid.TextMatrix(i, C_LNGFECHAUTIL)) < DateAdd("yyyy", 1, gFechaInicioDepInstantanea) Then
               Grid.TextMatrix(i, C_CREDART33) = Format(vFmt(Grid.TextMatrix(i, C_TIENECREDART33)) * vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) * gCredArt33_2014, NUMFMT)
            
            Else
               Grid.TextMatrix(i, C_CREDART33) = Format(vFmt(Grid.TextMatrix(i, C_TIENECREDART33)) * vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) * gCredArt33_2015, NUMFMT)
            
            End If
            
         Else
            Grid.TextMatrix(i, C_CREDART33) = Format(vFmt(Grid.TextMatrix(i, C_TIENECREDART33)) * vFmt(Grid.TextMatrix(i, C_VALCRED33)), NUMFMT)
         End If
         
         TotCred33 = TotCred33 + vFmt(Grid.TextMatrix(i, C_CREDART33))
         
         'If Val(Grid.TextMatrix(i, C_TOTDEP)) <> 0 Or (vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) <> 0 And FechaHasta = DateSerial(gEmpresa.Ano, 12, 31)) Then
         If NoCalcCCMM Then
            'Grid.TextMatrix(i, C_VALREAJCRED) = Grid.TextMatrix(i, C_VALINIT)
            Grid.TextMatrix(i, C_VALREAJCRED) = Grid.TextMatrix(i, C_VAL31DICANOANT)
         Else
            Grid.TextMatrix(i, C_VALREAJCRED) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) - vFmt(Grid.TextMatrix(i, C_CREDART33)), NUMFMT)
         End If
         
         'si hay dep. histórica, se actualiza con Factor de enero año en curso, para que se actualice con los 12 meses
         'Grid.TextMatrix(i, C_DEPACUMACT) = Format(vFmt(Grid.TextMatrix(i, C_DEPACUMHIST)) * gIndices(1).FactorCM, NUMFMT)
         Grid.TextMatrix(i, C_DEPACUMACT) = Format(vFmt(Grid.TextMatrix(i, C_DEPACUMHIST)) * vFmt(Grid.TextMatrix(i, C_FACTACT)), NUMFMT)
         
         'Valor Libro actualizado a depreciar
         If NoCalcDep Then
            Grid.TextMatrix(i, C_VALDEPRECIAR) = 0
         ElseIf Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_INST Then
            Grid.TextMatrix(i, C_VALDEPRECIAR) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJCRED)) / 2, NUMFMT)
         Else
            Grid.TextMatrix(i, C_VALDEPRECIAR) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJCRED)) - vFmt(Grid.TextMatrix(i, C_DEPACUMACT)), NUMFMT)
         End If
                     
         '**** ADO 2747741 Tema 2
         If Grid.TextMatrix(i, C_VU_ADEPRECIAR) <> "0" Then
            'depreciación mensual
            If vFmt(Grid.TextMatrix(i, C_VU_DISPONRESID)) <> 0 Then
               DepMensual = vFmt(Grid.TextMatrix(i, C_VALDEPRECIAR)) / vFmt(Grid.TextMatrix(i, C_VU_DISPONRESID))
            Else
               DepMensual = 0
            End If
         

            'depreciación mensual
            Grid.TextMatrix(i, C_DEPMENSUAL) = Format(DepMensual, NUMFMT)
            
            If Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_ARAUCANIA Or Val(Grid.TextMatrix(i, C_DEPLEY21256)) <> 0 Then
               Grid.TextMatrix(i, C_DEPMENSUAL) = ""
            End If
         Else
            Grid.TextMatrix(i, C_DEPMENSUAL) = 0
            DepMensual = 0
         End If
            'depreciación periodo = mensual * Meses a depreciar
            Grid.TextMatrix(i, C_DEPPERIODO) = Format(DepMensual * vFmt(Grid.TextMatrix(i, C_VU_ADEPRECIAR)), NUMFMT)
            
            
            'depreciación acumulada a año en curso (lo normal)
            Grid.TextMatrix(i, C_DEPACUMULADAANO) = Format(vFmt(Grid.TextMatrix(i, C_DEPPERIODO)) + vFmt(Grid.TextMatrix(i, C_DEPACUMACT)), NUMFMT)

         'parche por Reporte 41 Victor Morales (25 ene 2010)
         'si termina la vida útil del bien dentro del periodo que se está calculando, el valor libro debe ajustarse a $1, ajustando el valor de la depreciación del periodo
         '==> Grid.TextMatrix(i, C_VALLIBRO) = 1
         
         If vFmt(Grid.TextMatrix(i, C_VU_RESIDUAL)) = 0 And vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) = 0 And vFmt(Grid.TextMatrix(i, C_DEPPERIODO)) <> 0 Then
            'depreciación acumulada a año en curso
            Grid.TextMatrix(i, C_DEPACUMULADAANO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJCRED)) - 1, NUMFMT)
            Grid.TextMatrix(i, C_DEPPERIODO) = Format(vFmt(Grid.TextMatrix(i, C_DEPACUMULADAANO)) - vFmt(Grid.TextMatrix(i, C_DEPACUMACT)), NUMFMT)
         End If
         
         'Ley de la Araucanía
         If Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_ARAUCANIA Or Val(Grid.TextMatrix(i, C_DEPLEY21256)) <> 0 Then
            Grid.TextMatrix(i, C_DEPACUMULADAANO) = Grid.TextMatrix(i, C_VALDEPRECIAR)
            Grid.TextMatrix(i, C_DEPPERIODO) = Grid.TextMatrix(i, C_VALDEPRECIAR)
            
         ElseIf Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_INST Then
            DepPeriodo = vFmt(Grid.TextMatrix(i, C_DEPPERIODO))
            
            Grid.TextMatrix(i, C_DEPPERIODO) = Grid.TextMatrix(i, C_DEPPERIODO) & " + " & Grid.TextMatrix(i, C_VALDEPRECIAR)
            Grid.TextMatrix(i, C_DEPACUMULADAANO) = Format(DepPeriodo + vFmt(Grid.TextMatrix(i, C_VALDEPRECIAR)), NUMFMT)
          
         End If
         
         'Valor Libro año en curso
         If NoCalcDep Then
            Grid.TextMatrix(i, C_VALLIBRO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJCRED)) - vFmt(Grid.TextMatrix(i, C_DEPACUMULADAANO)), NUMFMT)
            
            If Val(Grid.TextMatrix(i, C_TOTDEP)) <> 0 Then
               Grid.TextMatrix(i, C_VALLIBRO) = 1
            End If
         
            Grid.TextMatrix(i, C_DEPACUMULADAANO) = Format(DepAcumHist, NUMFMT)        'Victor Morales 28 ago 2019
            
            
         ElseIf vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) = 0 Or vFmt(Grid.TextMatrix(i, C_LNGFECHAVENTA)) > lFecha Then
            
            If vFmt(Grid.TextMatrix(i, C_VU_DISPONRESID)) = 0 Then   'el activo fijo fue totalmente depreciado el año anterior y quedó con valor 1
               Grid.TextMatrix(i, C_VALLIBRO) = 1
            ElseIf Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_ARAUCANIA Or Val(Grid.TextMatrix(i, C_DEPLEY21256)) <> 0 Then
               Grid.TextMatrix(i, C_VALLIBRO) = 1
            Else
               Grid.TextMatrix(i, C_VALLIBRO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJCRED)) - vFmt(Grid.TextMatrix(i, C_DEPACUMULADAANO)), NUMFMT)
            End If
            
         Else
            Grid.TextMatrix(i, C_VALLIBRO) = 0   'Grid.TextMatrix(i, C_VALINIT)   (reporte de Victor Morales del 13 de Oct. 2008)
         
         End If
         
         'ponemos valores columnas resumen
      
         If Val(Grid.TextMatrix(i, C_TIENECREDART33ANOINIT)) <> 0 Then   'cambio solicitado por Victor el 14 ago 2012
            Grid.TextMatrix(i, C_RESCCMMACTFIJO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) - vFmt(Grid.TextMatrix(i, C_VALINIT)) * (1 - gCredArt33), NUMFMT)

         Else
            Grid.TextMatrix(i, C_RESCCMMACTFIJO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) - vFmt(Grid.TextMatrix(i, C_VALINIT)), NUMFMT)
         End If
         
         If vFmt(Grid.TextMatrix(i, C_RESCCMMACTFIJO)) < 0 Then
            Grid.TextMatrix(i, C_RESCCMMACTFIJO) = 0
         End If
         If vFmt(Grid.TextMatrix(i, C_RESCCMMPERIODO)) < 0 Then
            Grid.TextMatrix(i, C_RESCCMMPERIODO) = 0
         End If
         Grid.TextMatrix(i, C_RESCCMMPERIODO) = Format(vFmt(Grid.TextMatrix(i, C_VALREAJUSTADO)) - vFmt(Grid.TextMatrix(i, C_VAL31DICANOANT)), NUMFMT)
         Grid.TextMatrix(i, C_RESCCMMDEPACUM) = Format(vFmt(Grid.TextMatrix(i, C_DEPACUMACT)) - vFmt(Grid.TextMatrix(i, C_DEPACUMHIST)), NUMFMT)
         Grid.TextMatrix(i, C_RESDEPEJERCICIO) = Grid.TextMatrix(i, C_DEPPERIODO)
         Grid.TextMatrix(i, C_RESCREDART33) = Grid.TextMatrix(i, C_CREDART33)
                 
         'sumamos subtotales
         
         For j = C_VAL31DICANOANT To NCOLS
            If j <> C_FACTACT And j <> C_FECHAUTIL Then
               
               If Val(Grid.TextMatrix(i, C_DEPLEY21210)) = DEP_LEY21210_INST And j = C_DEPPERIODO Then
                  SubTotal(j) = SubTotal(j) + vFmt(Grid.TextMatrix(i, C_DEPACUMULADAANO))
               Else
                  SubTotal(j) = SubTotal(j) + vFmt(Grid.TextMatrix(i, j))
               End If
            End If
         Next j
            
      Else
      
         If IdActFijo <> 0 And InStr(LCase(Grid.TextMatrix(i, C_DESC)), "total") Then
         
            For j = 0 To NCOLS
            
               If j = C_VAL31DICANOANT Or (j >= C_VALREAJUSTADO And j <= C_VALDEPRECIAR) Or (j >= C_DEPMENSUAL And j <= C_VALLIBRO) Or (j >= C_RESCCMMACTFIJO And j <= C_RESCREDART33) Then
                  If j <> C_FECHAUTIL Then
                     Grid.TextMatrix(i, j) = Format(SubTotal(j), NUMFMT)
                     Total(j) = Total(j) + SubTotal(j)
                  End If
               End If
            
            Next j
            
         End If
      
         For j = 0 To NCOLS
            SubTotal(j) = 0
         Next j
         
      End If
      
         
   Next i
   
   'ponemos el total final
   For j = 0 To NCOLS
   
      If j = C_VAL31DICANOANT Or (j >= C_VALREAJUSTADO And j <= C_VALDEPRECIAR) Or (j >= C_DEPMENSUAL And j <= C_VALLIBRO) Or (j >= C_RESCCMMACTFIJO And j <= C_RESCREDART33) Then
         If j <> C_FECHAUTIL Then
            GridTot.TextMatrix(0, j) = Format(Total(j), NUMFMT)
         End If
      End If
      
   Next j
   
   Call ValidaCred(False)
   
End Sub

Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Frm.TxSelDate(Tx_Fecha) = vbOK Then
      lFecha = GetTxDate(Tx_Fecha)
      
      If lFecha < DateSerial(gEmpresa.Ano, 1, 1) Or lFecha > DateSerial(gEmpresa.Ano, 12, 31) Then
         MsgBox1 "Fecha inválida.", vbExclamation
         Call SetTxDate(Tx_Fecha, DateSerial(gEmpresa.Ano, 12, 31))
         lFecha = DateSerial(gEmpresa.Ano, 12, 31)
         Exit Sub
      End If

'      If lFecha < DateSerial(gEmpresa.Ano, 12, 31) Then
'         MsgBox1 "Asegúrese de asignar correctamente los meses a depreciar en el año actual, en la ventana de mantención de cada Activo Fijo.", vbInformation
'      End If
      Me.MousePointer = vbHourglass
      Call LoadAll
      Me.MousePointer = vbDefault
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Tx_Fecha_GotFocus()
   Call DtGotFocus(Tx_Fecha)
End Sub
Private Sub Tx_Fecha_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_Fecha) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_Fecha)
   
   Call DtLostFocus(Tx_Fecha)
   
   lFecha = GetTxDate(Tx_Fecha)
   
   If lFecha < DateSerial(gEmpresa.Ano, 1, 1) Or lFecha > DateSerial(gEmpresa.Ano, 12, 31) Then
      MsgBox1 "Fecha inválida.", vbExclamation
      lFecha = DateSerial(gEmpresa.Ano, 12, 31)
      Call SetTxDate(Tx_Fecha, lFecha)
      Exit Sub
   End If
   
   If lFecha < DateSerial(gEmpresa.Ano, 12, 31) Then
      MsgBox1 "Asegúrese de asignar correctamente los meses a depreciar en el año actual, en la ventana de mantención de cada Activo Fijo.", vbInformation
   End If
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
End Sub

Private Sub Tx_Fecha_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_MaxCred33_Change()
   lModTope = True
End Sub

Private Sub Tx_MaxCred33_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_MaxCred33_LostFocus()
   Dim Q1 As String
   
   If lModTope Then
      Tx_MaxCred33 = Format(vFmt(Tx_MaxCred33), NUMFMT)
      
      Q1 = "UPDATE ParamEmpresa SET Valor = " & vFmt(Tx_MaxCred33) & " WHERE Tipo='MAXCRED33'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      gMaxCred33 = vFmt(Tx_MaxCred33)
      
      Me.MousePointer = vbHourglass
      Call LoadAll
      Me.MousePointer = vbDefault
      lModTope = False
      
   End If
End Sub
Private Function ValidaCred(ByVal BtCancel As Boolean) As Boolean
   
   If vFmt(GridTot.TextMatrix(0, C_CREDART33)) > vFmt(Tx_MaxCred33) And vFmt(Tx_MaxCred33) > 0 Then
      
      If MsgBox1("El Crédito Artículo 33 bis excede el máximo de $ " & Tx_MaxCred33 & ".-" & vbCrLf & vbCrLf & "Ingrese al detalle de cada activo fijo y modifique el valor del Crédito Art. 33 bis.", vbInformation + IIf(BtCancel, vbOKCancel, 0)) = vbCancel Then
         ValidaCred = False
         Exit Function
      End If
      
   End If

   ValidaCred = True
End Function
Private Function GetIPC(ByVal Fecha) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   Dim AnoMes As Long
   
   AnoMes = DateSerial(Year(Fecha), month(Fecha), 1)
   
   If gEmpresa.Ano = 2019 Then         'Diconsinuidad del INE (Victor Morales, 20 ago 2019)
      If AnoMes = DateSerial(2018, 11, 1) Then       'Nov 2018
         GetIPC = 100.74
         Exit Function
      ElseIf AnoMes = DateSerial(2018, 12, 1) Then       'Dic 2018
         GetIPC = 100.64
         Exit Function
      End If
   End If
   
   Q1 = "SELECT pIPC FROM IPC WHERE AnoMes = " & AnoMes
   Set Rs = OpenRs(DbMain, Q1)
   
   GetIPC = 0
   If Not Rs.EOF Then
      GetIPC = vFld(Rs("pIPC"))
   Else
      If Not lMsgIPC Then
         MsgBox1 "No se encontró el valor en puntos del IPC de " & Format(AnoMes, "mmm yyyy") & ".", vbExclamation
         lMsgIPC = True
      End If
   End If
   
   Call CloseRs(Rs)
   
End Function

Private Sub SaveVal()
   Dim Q1 As String
   Dim i As Integer
   
   Me.MousePointer = vbHourglass
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_IDACTFIJO) <> "" Then
         Q1 = "UPDATE MovActivoFijo SET"
         Q1 = Q1 & "  ValReajustadoNeto = " & vFmt(Grid.TextMatrix(i, C_VALREAJCRED))
         Q1 = Q1 & ", DepAcumFinal = " & vFmt(Grid.TextMatrix(i, C_DEPACUMULADAANO))
         Q1 = Q1 & ", VidaUtilResidual = " & vFmt(Grid.TextMatrix(i, C_VU_RESIDUAL))
         Q1 = Q1 & ", ValorLibro = " & vFmt(Grid.TextMatrix(i, C_VALLIBRO))
         Q1 = Q1 & " WHERE IdActFijo = " & Grid.TextMatrix(i, C_IDACTFIJO)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next i
   
   Q1 = "UPDATE EmpresasAno SET"
   Q1 = Q1 & "  CredArt33bis = " & vFmt(GridTot.TextMatrix(0, C_CREDART33))
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
          
   Me.MousePointer = vbHourglass
End Sub

Private Function CalcValReajustado(ByVal Valor As Double, ByVal DtCompra As Double, FactorCompuesto As Double, FactorPeriodo As Double, Valor31DicAnoAnt As Double) As Double
   Dim FechaHasta As Long
   Dim ValorReajustado As Double
   Dim Dt As Long, Dt_1 As Long
   Dim IPC_Dt_1 As Double, IPC_Nov As Double
   Dim DtNov As Long
   Dim Factor As Double
   Dim StrFactor As String
   
   
   FechaHasta = GetTxDate(Tx_Fecha)
      
   'lo compraron este año
   If gEmpresa.Ano = Year(DtCompra) Then
      If gFactorActAnual(month(DtCompra), month(FechaHasta)).bFact Then
         Factor = gFactorActAnual(month(DtCompra), month(FechaHasta)).Fact
      
      ElseIf gIndices(month(DtCompra) - 1).PuntosIPC <> 0 Then
         If month(FechaHasta) - 1 = month(DtCompra) Then
            Factor = gIndices(month(FechaHasta) - 1).FactorCM
         Else
            Factor = gIndices(month(FechaHasta) - 1).PuntosIPC / gIndices(month(DtCompra) - 1).PuntosIPC
         End If
         
         If gEmpresa.Ano = 2010 And FechaHasta = DateSerial(2010, 12, 31) Then     'discontinuidad 2010
            Factor = GetFactorCM(DateSerial(Year(DtCompra), month(DtCompra), 1))
         End If
         
      Else
         Factor = 0
      End If
      
      If Factor < 1 And Factor <> 0 Then
         Factor = 1
      End If
      
      StrFactor = Format(Factor, DBLFMT3)
'      Factor = Round(Factor, 3)
      ValorReajustado = Valor * vFmt(StrFactor)
      
      FactorCompuesto = Factor
      FactorPeriodo = Factor
      Valor31DicAnoAnt = Valor
         
   Else                          'es de años anteriores, iteramos actualizaciones año a año
   
      ValorReajustado = Valor
      Dt = DtCompra
      
      Do While (Year(Dt) < gEmpresa.Ano)
         
         'restamos un mes a la fecha de compra
         Dt_1 = DateAdd("m", -1, Dt)
         
         'obtenemos el IPC de este mes
         IPC_Dt_1 = GetIPC(Dt_1)
         
         'obtenemos IPC de noviembre del año, para actualizar a diciembre del año
         DtNov = DateSerial(Year(Dt), 11, 1)
         IPC_Nov = GetIPC(DtNov)
        
         If DtNov = Dt_1 Then
            If DateSerial(Year(Dt) + 1, 12, 31) < FechaHasta Then   'si es la misma fecha pero no FechaHasta
               DtNov = DateSerial(Year(Dt) + 1, 11, 1) 'pasamos a nov año siguiente
               IPC_Nov = GetIPC(DtNov)
            Else
               Exit Do    'pasamos a este año
            End If
         End If
            
         
         If IPC_Dt_1 = 0 Then
            If Not lMsgIPCCompra Then
               MsgBox1 "El punto de IPC del mes anterior a la fecha de compra no ha sido ingresado al sistema. " & vbCrLf & vbCrLf & "Se sugiere actualizar el valor de compra del bien a una fecha más cercana, verificando que esté ingresado en el sistema el punto de IPC correspondiente.", vbExclamation
            End If
            Factor = 0
            lMsgIPCCompra = True
         Else
            'Grid.TextMatrix(i, C_FACTACT) = Format(gIndices(Month(FechaHasta) - 1).PuntosIPC / IPC_Dt_1, DBLFMT3)
            
            'obtenemos el factor para actualizar a diciembre del año
            Factor = IPC_Nov / IPC_Dt_1
            
            '631703 ffv
            If Factor < 1 Then
            Factor = Factor * 1.32719715
            End If
            '631703 ffv
            
            If Year(Dt) = 2010 Then                'se asigna en duro por discontinuidad generada por el INE
               Factor = 1.025
            End If
            
         End If
         
          '3239389
         If Year(Dt) = 2018 And month(Dt) = 12 Then                'se asigna en duro por discontinuidad generada por el INE
               Factor = 1.028
         End If
         '3239389
         
         
         If Factor < 1 And Factor <> 0 Then
            Factor = 1
         End If
         
        
         
         StrFactor = Format(Factor, DBLFMT3)
         ValorReajustado = ValorReajustado * vFmt(StrFactor)
                  
         'pasamos a noviembre del año
         Dt = DateSerial(Year(DtNov), 12, 1)    'el activo queda actualizado a diciembre del año
            
      Loop
      
      'calculamos factor de este periodo
      If IPC_Nov <> 0 Then
         Factor = GetIPC(DateAdd("m", -1, FechaHasta)) / IPC_Nov
         'Factor = gIndices(Month(FechaHasta) - 1).PuntosIPC / gIndices(Month(DtOld) - 1).PuntosIPC
         
         '631703
        If month(FechaHasta) = 1 And Year(FechaHasta) = 2024 Then
           If Factor < 1 Then
                Factor = 1
           End If
        End If
         '631703
         
         If Year(FechaHasta) = 2010 And month(FechaHasta) = 12 Then       'se asigna en duro por discontinuidad generada por el INE
            Factor = 1.025
         End If
    
        '3239389
         If Year(Dt) = 2018 And month(Dt) = 12 Then                'se asigna en duro por discontinuidad generada por el INE
               Factor = 1.028
         End If
         '3239389
      
      Else
         Factor = 0
      End If
       
           '631703 ffv
            If Factor < 1 Then
            Factor = Factor * 1.32719715
            End If
            '631703 ffv
      
      If Factor < 1 And Factor <> 0 Then
         Factor = 1
      End If
      
      
      'si se ingresó este año pero se compró el 31 dic año anterior, es como si lo hubiera comprado el 1 de este año
      If DtCompra = DateSerial(gEmpresa.Ano - 1, 12, 31) Then
         ValorReajustado = Valor   'no se aplica CCMM
      End If

      Valor31DicAnoAnt = ValorReajustado
      
      'aplicamos factor de este periodo
      StrFactor = Format(Factor, DBLFMT3)
      ValorReajustado = ValorReajustado * vFmt(StrFactor)
      
      FactorPeriodo = Factor

      If Valor <> 0 Then
         Factor = vFmt(Format(ValorReajustado, NUMFMT)) / Valor
      Else
         Factor = 0
      End If
      
      FactorCompuesto = Factor
      
   End If
   
   CalcValReajustado = ValorReajustado
      
End Function
