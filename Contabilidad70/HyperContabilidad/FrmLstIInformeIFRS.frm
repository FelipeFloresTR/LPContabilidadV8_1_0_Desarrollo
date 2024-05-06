VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstInformeIFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listar Estado de Resultado Financiero"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FrmLstIInformeIFRS.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Filtro 
      Height          =   1095
      Left            =   60
      TabIndex        =   19
      Top             =   720
      Width           =   11835
      Begin VB.CheckBox Ch_SaldosVig 
         Caption         =   "Saldos Vigentes"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3435
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3435
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   10500
         Picture         =   "FrmLstIInformeIFRS.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1740
         Picture         =   "FrmLstIInformeIFRS.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3840
         Picture         =   "FrmLstIInformeIFRS.frx":0866
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   230
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   24
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   23
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2220
         TabIndex        =   21
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6795
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   11986
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   11835
      Begin VB.CommandButton Bt_Email 
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
         Left            =   3360
         Picture         =   "FrmLstIInformeIFRS.frx":0B70
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Enviar por Correo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox Ch_VerCodCta 
         Caption         =   "Ver Códificación"
         Height          =   315
         Left            =   6000
         TabIndex        =   22
         Top             =   180
         Width           =   1755
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
         Left            =   2880
         Picture         =   "FrmLstIInformeIFRS.frx":0FF3
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   2040
         Picture         =   "FrmLstIInformeIFRS.frx":141C
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2460
         Picture         =   "FrmLstIInformeIFRS.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   1500
         Picture         =   "FrmLstIInformeIFRS.frx":1B1B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Cancel          =   -1  'True
         Caption         =   "Seleccionar"
         Height          =   315
         Left            =   9240
         TabIndex        =   16
         Top             =   180
         Width           =   1155
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
         Left            =   960
         Picture         =   "FrmLstIInformeIFRS.frx":1BBF
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_View 
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
         Left            =   120
         Picture         =   "FrmLstIInformeIFRS.frx":2004
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Vista previa de la impresión"
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
         Left            =   540
         Picture         =   "FrmLstIInformeIFRS.frx":24AB
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10500
         TabIndex        =   17
         Top             =   180
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmLstInformeIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_DESC = 1
Const C_IDCUENTA = 2
Const C_NIVEL = 3
Const C_SALDO = 4
Const C_FMT = 5

Const NCOLS = C_FMT

Dim lRc As Integer
Dim lOper As Integer
Dim lCodIFRS As String
Dim lDescIFRS As String
Dim lTipoInforme As Integer

Dim lMes As Integer

Public Function FSelect(ByVal TipoInforme As Integer, CodIFRS As String, DescIFRS As String) As Integer
   
   lOper = O_SELECT
   lTipoInforme = TipoInforme
   
   Me.Show vbModal
   
   CodIFRS = lCodIFRS
   DescIFRS = lDescIFRS
   
   FSelect = lRc
   
End Function
Public Sub FView(ByVal TipoInforme As Integer, Optional ByVal Mes As Integer = 0)
   
   lOper = O_VIEW
   lTipoInforme = TipoInforme
   
   'Me.Show vbModal
   Me.Show
      
End Sub


Private Sub SetUpGrid()
   Dim i As Integer
   
  Grid.AllowUserResizing = flexResizeColumns
  
  Grid.Cols = NCOLS + 1
  Call FGrSetup(Grid)
   
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   If lOper = O_SELECT Then
      Grid.ColWidth(C_SALDO) = 0
   Else
      Grid.ColWidth(C_SALDO) = 1500
   End If
   
   Grid.ColWidth(C_DESC) = 8700 + 1500 - Grid.ColWidth(C_SALDO)
   
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
         
   Grid.TextMatrix(0, C_CODIGO) = "Código"
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   If lOper = O_SELECT Then
      Grid.TextMatrix(0, C_SALDO) = ""
   Else
      Grid.TextMatrix(0, C_SALDO) = "Saldo"
   End If
   
End Sub


Private Sub Bt_Buscar_Click()

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Titulo As String
   
   Titulo = gInformeIFRS(lTipoInforme)
'   If InStr(Titulo, "Resultado") Then
'      Titulo = gInformeIFRS(lTipoInforme) & " por Función"
'   End If
   
   Titulo = Titulo & vbTab & "Periodo: " & Tx_Desde & " - " & Tx_Hasta

   'Call FGr2Clip(Grid, Titulo)
   '2861570
    Call FGr2Clip_membr(Grid, Titulo)
   '2861570
End Sub

Private Sub Bt_Email_Click()
Dim Frm As FrmEmailAccount

  Set Frm = Nothing
  Set Frm = New FrmEmailAccount
  
 Dim vAjunto As String
  vAjunto = Export_SendEmail(Grid, Nothing, Nothing, Nothing, "EstadoFinanciero" & "_" & Tx_Desde & "_" & Tx_Hasta, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta, C_CODIGO)
   
 If Frm.FEdit(vAjunto) Then
 Frm.Show
 End If
End Sub

Private Sub Bt_OK_Click()

   If Val(Grid.TextMatrix(Grid.Row, C_NIVEL)) = IFRS_MAXNIVEL And Grid.TextMatrix(Grid.Row, C_CODIGO) <> "" Then
      lCodIFRS = VFmtCodigoIFRS(Grid.TextMatrix(Grid.Row, C_CODIGO))
      lDescIFRS = VFmtCodigoIFRS(Grid.TextMatrix(Grid.Row, C_DESC))
      lRc = vbOK
      Unload Me
   Else
      MsgBox1 "Seleccione un registro de último nivel.", vbExclamation
   End If
   
End Sub

Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
      
   Call SetUpPrtGrid
   
'   PrtOrient = Printer.Orientation
'   Printer.Orientation = cdlLandscape
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
'   Printer.Orientation = PrtOrient
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub bt_View_Click()
   Dim Frm As FrmPrintPreview
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Cb_Informe_Click()

   Me.MousePointer = vbHourglass
   
   Call LoadAll
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Ch_VerCodCta_Click()

   If Ch_VerCodCta <> 0 Then
      Grid.ColWidth(C_CODIGO) = 1200
      Grid.ColWidth(C_DESC) = 8700 + 1500 - Grid.ColWidth(C_SALDO)
   Else
      Grid.ColWidth(C_CODIGO) = 0
      Grid.ColWidth(C_DESC) = 8700 + 1500 - Grid.ColWidth(C_SALDO) + 1200
   End If
   
   
End Sub

Private Sub Form_Load()
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Frm As FrmMsgConBreak
   Dim Msg As String
   Dim CurPlan As String
   Dim i As Integer
   Dim MesActual As Integer
   
   If lOper = O_SELECT Then
      Bt_Cancel.Caption = "Cancelar"
      
      Fr_Filtro.visible = False
      Bt_Sum.visible = False
      Bt_Calc.visible = False
      Bt_ConvMoneda.visible = False
      Bt_Calendar.visible = False
      Ch_VerCodCta.visible = False
      
   Else
   
      Bt_OK.visible = False
      
      MesActual = GetMesActual()
   
      If lMes = 0 Then
         If MesActual > 0 Then
            lMes = MesActual
         Else
            lMes = GetUltimoMesConComps()
         End If
      End If
      
      ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
      
      Call FirstLastMonthDay(ActDate, D1, D2)
      Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
      Call SetTxDate(Tx_Hasta, D2)
      
      Call FillCbAreaNeg(Cb_AreaNeg, False)
      Call FillCbCCosto(Cb_CCosto, False)


   End If
   
   Me.Caption = gInformeIFRS(lTipoInforme) ' & " (Formato IFRS)"
   Call SetUpGrid
   
   Ch_VerCodCta.Value = 1
   
   Call Bt_Buscar_Click
   
   If lOper = O_VIEW Then
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
         
      If Not Rs.EOF Then
         CurPlan = vFld(Rs("Valor"))
      End If
      
      Call CloseRs(Rs)
      
      If CurPlan <> "BÁSICO" And CurPlan <> "INTERMEDIO" And CurPlan <> "AVANZADO" Then
         Set Frm = New FrmMsgConBreak
            
         Msg = "Este informe sólo se mostrará para las empresas que utilicen uno de los planes de cuenta predefinidos por el sistema (Básico, Intermedio, Avanzado o IFRS)." & vbCrLf & vbCrLf & "Si no es así, será necesario realizar la configuración de IFRS en forma manual, utilizando la opción   Definiciones >> Plan de Cuentas >> Configurar códigos IFRS"
         Call Frm.FView(Msg, "NoDispMsgIFRS")
            
         Set Frm = Nothing
      End If
      
      If SaldosSinClasifIFRS Then
   
         MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Existen cuentas con saldo distinto de cero, que no tienen su correspondiente clasificación IFRS, lo cual puede provocar" & vbCrLf & "que los Estados Financieros queden descuadrados.", vbExclamation + vbOKOnly
   
      End If
   
      MsgBox1 "Este informe sólo considera comprobantes en estado APROBADO.", vbInformation + vbOKOnly
   End If
      
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(1) As String
   Dim FontTit(0) As FontDef_t
   Dim FontNom(0) As FontDef_t
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = gInformeIFRS(lTipoInforme)
'   If InStr(Titulos(0), "Resultado") Then
'      Titulos(0) = gInformeIFRS(lTipoInforme) & " por Función"
'   End If
   
   Titulos(1) = "Periodo: " & Tx_Desde & " - " & Tx_Hasta
      
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
      
   gPrtReportes.GrFontName = "Arial"
   gPrtReportes.GrFontSize = 8
   
   If CbItemData(Cb_AreaNeg) > 0 Then
      Encabezados(0) = "Área de negocio: " & Cb_AreaNeg
   End If
   
   i = 0
   If CbItemData(Cb_CCosto) > 0 Then
      If Encabezados(0) <> "" Then
         i = 1
      End If
      
      Encabezados(i) = "Centro de Costo: " & Cb_CCosto
   End If
   
   gPrtReportes.Encabezados = Encabezados
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   ColWi(C_DESC) = ColWi(C_DESC) - 100
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDCUENTA
   gPrtReportes.NTotLines = 0
   gPrtReportes.FmtCol = C_FMT
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer, i As Integer
   Dim CurNiv As Integer
   Dim CodFather As String
   Dim Padre(IFRS_MAXNIVEL) As String
   Dim CodPadre(IFRS_MAXNIVEL) As String
   Dim Nivel As Integer
   Dim k As Integer
   Dim Total(IFRS_MAXNIVEL) As Double, TotReservas As Double
   Dim Saldo As Double
   Dim StrTotal As String
   Dim NLinTot As Integer
   Dim GananciaAntesImpuestos As Double
   Dim GastoImpuestos As Double
   Dim WhFecha As String, Wh As String
   Dim TotActivos As Double
   Dim TotPasivos As Double, RowPasivos As Integer
   Dim TotPatrimonio As Double, RowPatrimonio As Integer
   Dim TotUtilidadPerdida As Double, RowUtilidadPerdida As Integer
   Dim TotGananciasPerdidas As Double, RowGananciasPerdidas As Integer, TotOriGananciasPerdidas As Double
   Dim IdAreaNeg As Long, IdCCosto As Long
   Dim PadreN3, n As Integer
   Dim RowOtrasReservas As Integer
   Dim Codigo As String
   
   'Call MarkCompCCMM
      
   'Tbl = "IFRS_PlanIFRS"
   If lOper = O_SELECT Then
      NLinTot = 1
   Else
      NLinTot = 2
   End If
   StrTotal = "Total "
   
   WhFecha = " ( Comprobante.Fecha IS NULL OR "
   WhFecha = WhFecha & " (Comprobante.Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & "))"
   
   IdAreaNeg = ItemData(Cb_AreaNeg)
   IdCCosto = ItemData(Cb_CCosto)
   
   Wh = ""
   If ItemData(Cb_AreaNeg) > 0 Then
      Wh = Wh & " AND (MovComprobante.IdComp IS NULL OR MovComprobante.IdAreaNeg = " & IdAreaNeg & ")"
   End If
   
   If ItemData(Cb_CCosto) > 0 Then
      Wh = Wh & " AND (MovComprobante.IdComp IS NULL OR MovComprobante.IdCCosto = " & IdCCosto & ")"
   End If
      
   Q1 = "SELECT IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion as Descr"
   Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion as ClasCta "
   If lOper = O_VIEW Then
      Q1 = Q1 & ", 0 as SumDebe, 0 As SumHaber  "
   End If
   Q1 = Q1 & " FROM ( IFRS_PlanIFRS "
   '3224380
   If gDbType = SQL_ACCESS Then
        Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS ) "
   Else
        Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS AND (Cuentas.IdEmpresa IS NULL OR Cuentas.IdEmpresa = " & gEmpresa.id & ") AND (Cuentas.Ano IS NULL OR Cuentas.Ano = " & gEmpresa.Ano & ")) "
   End If
   '3224380
   
   If lTipoInforme = IFRS_ESTFIN Then
      Q1 = Q1 & " WHERE IFRS_PlanIFRS.Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ")"
   Else
      Q1 = Q1 & " WHERE IFRS_PlanIFRS.Clasificacion = " & CLASCTA_RESULTADO
   End If
   
   '3224380
   If gDbType = SQL_ACCESS Then
    Q1 = Q1 & " AND (Cuentas.IdEmpresa IS NULL OR Cuentas.IdEmpresa = " & gEmpresa.id & ") AND (Cuentas.Ano IS NULL OR Cuentas.Ano = " & gEmpresa.Ano & ")"
   End If
   '3224380
   If lOper = O_VIEW Then
      Q1 = Q1 & " UNION "
         
      Q1 = Q1 & "SELECT IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion as Descr"
      'Q1 = Q1 & ", Cuentas.Clasificacion as ClasCta "
      Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion as ClasCta "
      
      If lOper = O_VIEW Then
         Q1 = Q1 & ", Sum(MovComprobante.Debe) as SumDebe, Sum(MovComprobante.Haber) As SumHaber  "
      End If
      
      Q1 = Q1 & " FROM (( IFRS_PlanIFRS "
      Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
      
      If lOper = O_VIEW Then
         Q1 = Q1 & " LEFT JOIN MovComprobante ON Cuentas.IdCuenta = MovComprobante.IdCuenta"
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
         Q1 = Q1 & " LEFT JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
         Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Else
         Q1 = Q1 & ")"
      End If
         
      If lTipoInforme = IFRS_ESTFIN Then
         Q1 = Q1 & " WHERE IFRS_PlanIFRS.Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ")"
      Else
         Q1 = Q1 & " WHERE IFRS_PlanIFRS.Clasificacion = " & CLASCTA_RESULTADO
      End If
      
      If lOper = O_VIEW Then
         Q1 = Q1 & " AND " & WhFecha & Wh
         Q1 = Q1 & " AND (Comprobante.Estado IS NULL OR Comprobante.Estado = " & EC_APROBADO & ")"
         'Q1 = Q1 & " AND (Comprobante.EsCCMM IS NULL OR Comprobante.EsCCMM = 0) "   'no se consideran los comprobantes de CCMM (corrección Monetaria)
         Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
      End If
      
      Q1 = Q1 & " AND (Cuentas.IdEmpresa IS NULL OR Cuentas.IdEmpresa = " & gEmpresa.id & ") AND (Cuentas.Ano IS NULL OR Cuentas.Ano = " & gEmpresa.Ano & ")"

      Q1 = Q1 & " GROUP BY IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion "
      Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion"
   End If
   
   Q1 = Q1 & " ORDER BY IFRS_PlanIFRS.Codigo"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   Do While Rs.EOF = False
   
      Nivel = vFld(Rs("Nivel"))
      
      If Nivel < CurNiv Then
      
         For k = CurNiv - 1 To Nivel Step -1
                     
            If StrTotal <> "" And Padre(k) <> "" Then
               
               'cuando se selecciona una cuenta, no se muestran los totales
               If lOper <> O_SELECT Then
                  
                  Grid.rows = Grid.rows + NLinTot    '1 o 2
                  
                  If k = 1 Then
                     Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (0), " ") & UCase(StrTotal & Padre(k))
                  ElseIf k = 2 Then
                     Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (1), " ") & UCase(StrTotal & Padre(k))       'estaba FCase
                  Else
                     Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (IFRS_MAXNIVEL - 2), " ") & StrTotal & FCase(Padre(k))
                  End If
               
                  Grid.TextMatrix(Row, C_NIVEL) = k
                  Grid.TextMatrix(Row, C_SALDO) = Format(Total(k), NEGNUMFMT)
                  
                  Call FGrFontBold(Grid, Row, C_DESC, True)
                  Call FGrFontBold(Grid, Row, C_SALDO, True)
                  Grid.TextMatrix(Row, C_FMT) = "B"
                     
                  If LCase(Padre(k)) = "activos" Then
                     TotActivos = Total(k)
                                                            
                  ElseIf LCase(Padre(k)) = "ganancias (pérdidas) acumuladas" Then
                     RowGananciasPerdidas = Row
                     TotGananciasPerdidas = Total(k)
                     
                  End If
                  
                  Total(k) = 0
                  Row = Row + NLinTot
                  
               Else
                  Grid.rows = Grid.rows + 1
                  
                  Row = Row + 1
                  
                  Exit For
               End If
               
               
            End If
            
         Next k
      End If
      
      If Nivel > CurNiv Then
         If Row > 1 And Nivel <= 3 Then
             'Salto una línea
            Grid.rows = Row + 1
            Grid.TextMatrix(Row, C_IDCUENTA) = "*******"
   
            Row = Row + 1
         End If
      End If
   
      CurNiv = Nivel
      
      Padre(CurNiv) = vFld(Rs("Descr"))
      CodPadre(CurNiv) = FmtCodIFRS(vFld(Rs("Codigo")))
                  
'      If vFld(Rs("Codigo")) = "3010203" Then
'         MsgBeep vbExclamation
'      End If

      If FmtCodIFRS(vFld(Rs("Codigo"))) = Grid.TextMatrix(Row - 1, C_CODIGO) Then
         Row = Row - 1
      Else
         Grid.rows = Row + 1
      End If
      
      Grid.TextMatrix(Row, C_CODIGO) = FmtCodIFRS(vFld(Rs("Codigo")))
      
      If CurNiv = 1 Then
         Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * 0, " ") & UCase(vFld(Rs("Descr")))     'cambiamos FCase por UCase 21 nov 13
         
      ElseIf CurNiv = 2 Then
         Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * 1, " ") & UCase(vFld(Rs("Descr"), True))          'estava FCase
         
      ElseIf CurNiv > 2 And CurNiv <= IFRS_MAXNIVEL - 1 Then
         Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * 1, " ") & FCase(vFld(Rs("Descr"), True))
         
         If LCase(vFld(Rs("Descr"))) = "otras reservas" Then
            RowOtrasReservas = Row
         End If
               
      Else  'CurNiv = IFRS_MAXNIVEL
         Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (CurNiv - 1) - 1, " ") & FCase(vFld(Rs("Descr")))
      
      End If
            
      Grid.TextMatrix(Row, C_IDCUENTA) = vFld(Rs("idCuenta"))
      Grid.TextMatrix(Row, C_NIVEL) = vFld(Rs("Nivel"))
      
      'cuando se selecciona una cuenta, no se muestran los totales
      If lOper <> O_SELECT Then
         If vFld(Rs("ClasCta")) = CLASCTA_ACTIVO Then
            Saldo = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
         Else
            Saldo = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
         End If
         
         If CurNiv = IFRS_MAXNIVEL And Saldo <> 0 Then
            Grid.TextMatrix(Row, C_SALDO) = Format(Saldo, NEGNUMFMT)
         End If
                     
         If LCase(vFld(Rs("Descr"))) = "utilidad o pérdida neta del periodo" Then
            RowUtilidadPerdida = Row

         End If

         For k = 1 To IFRS_MAXNIVEL - 1
            Total(k) = Total(k) + Saldo
         Next k
      
      End If
      
      If CurNiv < IFRS_MAXNIVEL Then
         Call FGrFontBold(Grid, Row, C_DESC, True)
         Grid.TextMatrix(Row, C_FMT) = "B"
      End If
         
      Row = Row + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
         
   'Ponemos los últimos totales
   
   If lOper <> O_SELECT Then
   
      For k = CurNiv - 1 To 1 Step -1
         Grid.rows = Grid.rows + 2
                  
         If k = 1 Then
            Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (0), " ") & UCase(StrTotal & Padre(k))
            Grid.TextMatrix(Row, C_SALDO) = Format(Total(1), NEGNUMFMT)
   
         ElseIf k = 2 Then
            Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (1), " ") & UCase(StrTotal & Padre(k))       'estaba FCase
            Grid.TextMatrix(Row, C_SALDO) = Format(Total(k), NEGNUMFMT)
            
         Else
            Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (IFRS_MAXNIVEL - 2), " ") & StrTotal & Padre(k)
            Grid.TextMatrix(Row, C_SALDO) = Format(Total(k), NEGNUMFMT)
         End If
         
         Grid.TextMatrix(Row, C_NIVEL) = k
         
         If Trim(LCase(Grid.TextMatrix(Row, C_DESC))) = "total patrimonio" Then
            TotPatrimonio = vFmt(Grid.TextMatrix(Row, C_SALDO))
            RowPatrimonio = Row
   
         ElseIf Trim(LCase(Grid.TextMatrix(Row, C_DESC))) = "total pasivos" Then
            TotPasivos = vFmt(Grid.TextMatrix(Row, C_SALDO))
            RowPasivos = Row
         End If
         
         Call FGrFontBold(Grid, Row, C_DESC, True)
         Call FGrFontBold(Grid, Row, C_SALDO, True)
         Grid.TextMatrix(Row, C_FMT) = "B"
         
         'cuando se selecciona una cuenta, no se muestran los totales
         If lOper = O_SELECT Then
            Grid.TextMatrix(Row, C_SALDO) = ""
         End If
         
         Row = Row + 2
      Next k
      
      'traemos los saldos de las cuentas de Resultado Integral para asignarlos a Otras Reservas
      If lTipoInforme = IFRS_ESTFIN Then
         Q1 = "SELECT IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion as Descr"
         Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion as ClasCta "
         Q1 = Q1 & ", Sum(MovComprobante.Debe) as SumDebe, Sum(MovComprobante.Haber) As SumHaber  "
         Q1 = Q1 & " FROM (( IFRS_PlanIFRS "
         Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
         Q1 = Q1 & " LEFT JOIN MovComprobante ON Cuentas.IdCuenta = MovComprobante.IdCuenta "
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
         Q1 = Q1 & " LEFT JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
         Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
         
         Q1 = Q1 & " WHERE IFRS_PlanIFRS.Nivel = " & gLastNivelIFRS & " AND Left(IFRS_PlanIFRS.Codigo,3) = '401'"
      
         Q1 = Q1 & " AND " & WhFecha & Wh
         Q1 = Q1 & " AND (Comprobante.Estado IS NULL OR Comprobante.Estado = " & EC_APROBADO & ")"
         Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
         Q1 = Q1 & " AND (Cuentas.IdEmpresa IS NULL OR Cuentas.IdEmpresa = " & gEmpresa.id & ") AND (Cuentas.Ano IS NULL OR Cuentas.Ano = " & gEmpresa.Ano & ")"
         Q1 = Q1 & " GROUP BY IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion "
         Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion"
         
         Q1 = Q1 & " ORDER BY IFRS_PlanIFRS.Codigo "
      
         Set Rs = OpenRs(DbMain, Q1)
         
         TotReservas = 0
         
         Do While Not Rs.EOF
            
            Codigo = vFld(Rs("Codigo"))
            For i = RowOtrasReservas + 1 To RowOtrasReservas + 6
               If Right(VFmtCodigoIFRS(Grid.TextMatrix(i, C_CODIGO)), 2) = Mid(Codigo, 4, 2) Then
                  Saldo = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
                  Grid.TextMatrix(i, C_SALDO) = Format(Saldo, NUMFMT)
                  TotReservas = TotReservas + Saldo
               End If
            Next i
            
            Rs.MoveNext
         Loop
         
         Call CloseRs(Rs)
         
         If RowOtrasReservas > 0 Then
            Grid.TextMatrix(RowOtrasReservas + 7, C_SALDO) = Format(TotReservas, NUMFMT)
         End If
      End If
      
      
            
      'ajustamos las Ganancias(Pérdidas) acumuladas con el Patrimonio
      If lTipoInforme = IFRS_ESTFIN Then
      
         'ajustamos Utilidad o Pérdida Neta del Período
         TotOriGananciasPerdidas = TotGananciasPerdidas
         TotUtilidadPerdida = TotActivos - TotPasivos - TotReservas
         If RowUtilidadPerdida > 0 Then
            Grid.TextMatrix(RowUtilidadPerdida, C_SALDO) = IIf(TotUtilidadPerdida <> 0, Format(TotUtilidadPerdida, NEGNUMFMT), "")  'ponemos el total
         End If

         'ajustamos las Ganancias(Pérdidas) acumuladas con el Patrimonio
         TotGananciasPerdidas = TotGananciasPerdidas + TotUtilidadPerdida
         If RowGananciasPerdidas > 0 Then
            Grid.TextMatrix(RowGananciasPerdidas, C_SALDO) = IIf(TotGananciasPerdidas <> 0, Format(TotGananciasPerdidas, NEGNUMFMT), "")
         End If

         'ajustamos el Patrimonio
         TotPatrimonio = TotPatrimonio + TotUtilidadPerdida + TotReservas
         If RowPatrimonio > 0 Then
            Grid.TextMatrix(RowPatrimonio, C_SALDO) = IIf(TotPatrimonio <> 0, Format(TotPatrimonio, NEGNUMFMT), "")
         End If

         'ajustamos el los Pasivos
         TotPasivos = TotActivos
         If TotPasivos > 0 Then
            Grid.TextMatrix(RowPasivos, C_SALDO) = IIf(TotPasivos <> 0, Format(TotPasivos, NEGNUMFMT), "")
         End If
   
      End If
   End If
   
   'rellenamos la columna C_IDCUENTA con algo para que imprima la línea
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_IDCUENTA) = "" Then
         Grid.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   'ocultamos líneas con saldo en cero si corresponde
   
   If Ch_SaldosVig <> 0 Then
   
      'primero el último nivel
      For i = Grid.FixedRows + 1 To Grid.rows - 1
      
         If vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivelIFRS Then
            Grid.RowHeight(i) = 0
         End If
      Next i
      
      'luego el penúltimo nivel
      For i = Grid.FixedRows + 1 To Grid.rows - 1
      
      
         If Grid.RowHeight(i) > 0 Then
            If vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivelIFRS - 1 And Grid.TextMatrix(i, C_CODIGO) <> "" Then
               PadreN3 = i
               n = 0
            ElseIf vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivelIFRS Then
               n = n + 1
            End If
            
            'si es total de nivel 3 y no hay líneas de detalle con saldo en este rubro, lo coultamos completo
            If vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivelIFRS - 1 And Grid.TextMatrix(i, C_CODIGO) = "" And n = 0 Then
               Grid.RowHeight(i) = 0
               Grid.RowHeight(PadreN3) = 0
               Grid.RowHeight(i + 1) = 0
            End If
         End If
      Next i
   End If

   Call FGrVRows(Grid, 2)
   Grid.Redraw = True
   
End Sub
Private Sub Form_Resize()
    Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

'   d = Me.Width - 2 * (Grid.Left + W.xFrame)
'   If d > 1000 Then
'      Grid.Width = d
'   End If

 
   d = Me.Height - Grid.Top - W.YCaption * 2 + 80 - 30
   If d > 1000 Then
      Grid.Height = d
      If Fr_Filtro.visible = False Then
         Grid.Top = Fr_Filtro.Top
         Grid.Height = Grid.Height + Fr_Filtro.Height
      End If
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
      
   Call FGrVRows(Grid)
   
End Sub

Private Sub Grid_DblClick()

   If Bt_OK.visible Then
      Call Bt_OK_Click
   End If
   
End Sub
Private Sub tx_Desde_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Desde_GotFocus()
   Call DtGotFocus(Tx_Desde)
End Sub

Private Sub Tx_Desde_LostFocus()
   
   If Trim$(Tx_Desde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Desde)
   
End Sub

Private Sub Tx_Desde_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub tx_Hasta_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Hasta_GotFocus()
   Call DtGotFocus(Tx_Hasta)
   
End Sub

Private Sub Tx_Hasta_LostFocus()
   
   If Trim$(Tx_Hasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Hasta)
      
End Sub

Private Sub Tx_Hasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_Desde)
   Else
      Call Frm.TxSelDate(Tx_Hasta)
   End If
   
   Set Frm = Nothing
   
End Sub
Private Sub EnableFrm(bool As Boolean)
   Bt_Buscar.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid, -1, -1, C_SALDO, C_SALDO)
   
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

