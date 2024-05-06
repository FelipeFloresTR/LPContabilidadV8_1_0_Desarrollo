VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalEjecIFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Situación Financiera Ejecutivo"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FrmBalEjecIFRS.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Filtro 
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   720
      Width           =   11835
      Begin VB.CheckBox Ch_SaldosVig 
         Caption         =   "Saldos Vigentes"
         Height          =   195
         Left            =   4620
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   10500
         Picture         =   "FrmBalEjecIFRS.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1740
         Picture         =   "FrmBalEjecIFRS.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3840
         Picture         =   "FrmBalEjecIFRS.frx":0866
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
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2220
         TabIndex        =   17
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7455
      Left            =   60
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   13150
      _Version        =   393216
      FixedRows       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   11835
      Begin VB.CheckBox Ch_VerCodCta 
         Caption         =   "Ver Códificación"
         Height          =   315
         Left            =   6000
         TabIndex        =   18
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
         Picture         =   "FrmBalEjecIFRS.frx":0B70
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
         Left            =   2040
         Picture         =   "FrmBalEjecIFRS.frx":0F99
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
         Left            =   2460
         Picture         =   "FrmBalEjecIFRS.frx":1337
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
         Left            =   1500
         Picture         =   "FrmBalEjecIFRS.frx":1698
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Left            =   960
         Picture         =   "FrmBalEjecIFRS.frx":173C
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "FrmBalEjecIFRS.frx":1B81
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "FrmBalEjecIFRS.frx":2028
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10500
         TabIndex        =   13
         Top             =   180
         Width           =   1155
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridV 
      Height          =   7515
      Left            =   60
      TabIndex        =   19
      Top             =   1500
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   13256
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "FrmBalEjecIFRS"
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
Const C_INTER = 5
Const C_CODIGO_P = 6
Const C_DESC_P = 7
Const C_IDCUENTA_P = 8
Const C_NIVEL_P = 9
Const C_SALDO_P = 10
Const C_FMT = 11

Const NCOLS = C_FMT

Dim lRc As Integer
Dim lOper As Integer
Dim lCodIFRS As String
Dim lDescIFRS As String
Dim lWCodCta As Integer
Dim lWVal As Integer
Dim lWDesc As Integer

Dim lMes As Integer

Public Sub FView(Optional ByVal Mes As Integer = 0)

   lMes = Mes
   lOper = O_VIEW
   
   Me.Show vbModal
      
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.AllowUserResizing = flexResizeColumns
   
   Grid.Cols = NCOLS + 1
   GridV.Cols = NCOLS + 1
   Call FGrSetup(GridV)
  
   Grid.FixedRows = 0
   GridV.FixedRows = 0
   
   lWCodCta = 1000
   lWVal = G_DVALWIDTH + 200
   lWDesc = 4800
   
   If Ch_VerCodCta <> 0 Then
      GridV.ColWidth(C_CODIGO) = lWCodCta
      GridV.ColWidth(C_CODIGO_P) = lWCodCta
   Else
      GridV.ColWidth(C_CODIGO) = 0
      GridV.ColWidth(C_CODIGO_P) = 0
   End If
   
   Grid.ColWidth(C_DESC) = lWDesc
   GridV.ColWidth(C_DESC) = lWDesc
   GridV.ColWidth(C_SALDO) = lWVal
   GridV.ColWidth(C_NIVEL) = 0
   GridV.ColWidth(C_IDCUENTA) = 0
   
   GridV.ColWidth(C_INTER) = 300
   
   GridV.ColWidth(C_DESC_P) = lWDesc
   GridV.ColWidth(C_SALDO_P) = lWVal
   GridV.ColWidth(C_NIVEL_P) = 0
   GridV.ColWidth(C_IDCUENTA_P) = 0
   
   GridV.ColWidth(C_FMT) = 0
         
   GridV.ColAlignment(C_DESC) = flexAlignLeftCenter
   GridV.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   GridV.ColAlignment(C_SALDO) = flexAlignRightCenter
   GridV.ColAlignment(C_DESC_P) = flexAlignLeftCenter
   GridV.ColAlignment(C_CODIGO_P) = flexAlignLeftCenter
   GridV.ColAlignment(C_SALDO_P) = flexAlignRightCenter
         
'   Grid.TextMatrix(0, C_CODIGO) = "Código"
'   Grid.TextMatrix(0, C_DESC) = "Descripción"
'   Grid.TextMatrix(0, C_SALDO) = "Saldo"
   
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

   'Call FGr2Clip(GridV, "")
   '2861570
   Call FGr2Clip_membr(GridV, "")
   '2861570
End Sub

Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
      
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub bt_View_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape

   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   Printer.Orientation = PrtOrient

   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Ch_SaldosVig_Click()
   Call EnableFrm(True)

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
   
   Me.Caption = gInformeIFRS(IFRS_BALEJEC)
   
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
   
   Call SetUpGrid
   
   Ch_VerCodCta.Value = 1
   
   Call Bt_Buscar_Click
   
   If lOper = O_VIEW Then
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS' "
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
   
         MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Existen cuentas con saldo distinto de cero, que no tienen su correspondiente clasificación IFRS.", vbExclamation + vbOKOnly
   
      End If

      MsgBox1 "Este informe sólo considera comprobantes en estado APROBADO.", vbInformation + vbOKOnly
   End If
      
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(0) As FontDef_t
   Dim FontNom(0) As FontDef_t
   
   Set gPrtReportes.Grid = GridV
   
   Titulos(0) = Me.Caption
   
   Titulos(1) = "Periodo: " & Tx_Desde & " - " & Tx_Hasta
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
      
   gPrtReportes.GrFontName = "Arial"
   gPrtReportes.GrFontSize = 8
   gPrtReportes.Encabezados = Encabezados
   
   For i = 0 To GridV.Cols - 1
      ColWi(i) = GridV.ColWidth(i)
   Next i
   
   'ColWi(C_DESC) = ColWi(C_DESC) - 100
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDCUENTA
   gPrtReportes.NTotLines = 0
   gPrtReportes.FmtCol = C_FMT
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim row As Integer, i As Integer
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
   Dim WhFecha As String
   Dim TotActivos As Double
   Dim TotPasivos As Double, RowPasivos As Integer
   Dim TotPatrimonio As Double, RowPatrimonio As Integer
   Dim TotUtilidadPerdida As Double, RowUtilidadPerdida As Integer
   Dim TotGananciasPerdidas As Double, RowGananciasPerdidas As Integer, TotOriGananciasPerdidas As Double
   Dim PadreN3 As Integer, n As Integer
   Dim RowOtrasReservas As Integer
   Dim Codigo As String

   StrTotal = "Total "
   NLinTot = 2
   
   WhFecha = " (Comprobante.Fecha IS NULL OR "
   WhFecha = WhFecha & " (Comprobante.Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & "))"
      
   Q1 = "SELECT IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion as Descr"
   Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion as ClasCta "
   Q1 = Q1 & ", Sum(MovComprobante.Debe) as SumDebe, Sum(MovComprobante.Haber) As SumHaber  "
   
   Q1 = Q1 & " FROM (( IFRS_PlanIFRS "
   Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
   Q1 = Q1 & " LEFT JOIN MovComprobante ON Cuentas.IdCuenta = MovComprobante.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   
   
   Q1 = Q1 & " WHERE " & WhFecha
   Q1 = Q1 & " AND (Comprobante.Estado IS NULL OR Comprobante.Estado = " & EC_APROBADO & ")"
   Q1 = Q1 & " AND IFRS_PlanIFRS.Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND (Comprobante.IdEmpresa IS NULL OR Comprobante.IdEmpresa = " & gEmpresa.id & ") AND (Comprobante.Ano IS NULL OR Comprobante.Ano = " & gEmpresa.Ano & ")"
   Q1 = Q1 & " AND (Cuentas.IdEmpresa IS NULL OR Cuentas.IdEmpresa = " & gEmpresa.id & ") AND (Cuentas.Ano IS NULL OR Cuentas.Ano = " & gEmpresa.Ano & ")"

   Q1 = Q1 & " GROUP BY IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion "
   Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion "
   Q1 = Q1 & " ORDER BY IFRS_PlanIFRS.Codigo"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   GridV.Redraw = False
   row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   Do While Rs.EOF = False
   
      Nivel = vFld(Rs("Nivel"))
      
      If Nivel < CurNiv Then
      
         For k = CurNiv - 1 To Nivel Step -1
         
'            If k = 1 Then
'               Total(k) = 0
'               Exit For
'            End If
            
            If StrTotal & Padre(k) <> "" Then
               
               Grid.rows = Grid.rows + NLinTot    '1 o 2
                              
               If k = 1 Then
                  Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (0), " ") & UCase(StrTotal & Padre(k))
               ElseIf k = 2 Then
                  Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (1), " ") & UCase(StrTotal & Padre(k))
               Else
                  Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (IFRS_MAXNIVEL - 2), " ") & StrTotal & Padre(k)
               End If
               
               Grid.TextMatrix(row, C_NIVEL) = k
               Grid.TextMatrix(row, C_SALDO) = Format(Total(k), NEGNUMFMT)
               
               Call FGrFontBold(Grid, row, C_DESC, True)
               Call FGrFontBold(Grid, row, C_SALDO, True)
               Grid.TextMatrix(row, C_FMT) = "B"
               
               If InStr(LCase(Padre(k)), "antes de impuestos") Then
                  GananciaAntesImpuestos = Total(k)
               
               ElseIf InStr(LCase(Padre(k)), "operaciones continuadas") Then
                  Grid.TextMatrix(row, C_SALDO) = Format(GananciaAntesImpuestos + GastoImpuestos, NEGNUMFMT)
            
               ElseIf LCase(Padre(k)) = "activos" Then
                  TotActivos = Total(k)
                  
               ElseIf LCase(Padre(k)) = "ganancias (pérdidas) acumuladas" Then
                  RowGananciasPerdidas = row
                  TotGananciasPerdidas = Total(k)
               End If
               
               Total(k) = 0
               row = row + NLinTot
            End If
            
         Next k
      End If
   
       If Nivel > CurNiv Then
         If row > 0 And Nivel <= 3 Then
             'Salto una línea
            Grid.rows = row + 1
            Grid.TextMatrix(row, C_IDCUENTA) = "*******"
   
            row = row + 1
         End If
      End If
  
      CurNiv = Nivel
      
      Padre(CurNiv) = vFld(Rs("Descr"))
      CodPadre(CurNiv) = FmtCodIFRS(vFld(Rs("Codigo")))
            
      
      Grid.rows = row + 1
      Grid.TextMatrix(row, C_CODIGO) = FmtCodIFRS(vFld(Rs("Codigo")))
         
      If CurNiv = 1 Then
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * 0, " ") & UCase(vFld(Rs("Descr")))     'cambiamos FCase por UCase 21 nov 13
         
      ElseIf CurNiv = 2 Then
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * 1, " ") & UCase(vFld(Rs("Descr"), True))          'estava FCase
         
      ElseIf CurNiv > 2 And CurNiv <= IFRS_MAXNIVEL - 1 Then
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * 1, " ") & FCase(vFld(Rs("Descr"), True))
         
         If LCase(vFld(Rs("Descr"))) = "otras reservas" Then
            RowOtrasReservas = row
         End If
      
      Else  'CurNiv = IFRS_MAXNIVEL
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (CurNiv - 1) - 1, " ") & FCase(vFld(Rs("Descr")))
      
      End If
               
      Grid.TextMatrix(row, C_IDCUENTA) = vFld(Rs("idCuenta"))
      Grid.TextMatrix(row, C_NIVEL) = vFld(Rs("Nivel"))
      
      'cuando se selecciona una cuenta, no se muestran los totales
      If vFld(Rs("ClasCta")) = CLASCTA_ACTIVO Then
         Saldo = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
      Else
         Saldo = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
      End If
      
      If CurNiv = IFRS_MAXNIVEL And Saldo <> 0 Then
         Grid.TextMatrix(row, C_SALDO) = Format(Saldo, NEGNUMFMT)
      End If
      
      If LCase(vFld(Rs("Descr"))) = "utilidad o pérdida neta del periodo" Then
         RowUtilidadPerdida = row

      End If
                     
      For k = 1 To IFRS_MAXNIVEL - 1
         Total(k) = Total(k) + Saldo
      Next k
            
      If CurNiv < IFRS_MAXNIVEL Then
         Call FGrFontBold(Grid, row, C_DESC, True)
         Grid.TextMatrix(row, C_FMT) = "B"
      End If
         
      row = row + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
      
   'Ponemos los últimos totales
   
   For k = CurNiv - 1 To 1 Step -1
      Grid.rows = Grid.rows + 2
            
      If k = 1 Then
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (0), " ") & UCase(StrTotal & Padre(k))
         Grid.TextMatrix(row, C_SALDO) = Format(Total(1), NEGNUMFMT)
      
      ElseIf k = 2 Then
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (1), " ") & UCase(StrTotal & Padre(k))       'estaba FCase
         Grid.TextMatrix(row, C_SALDO) = Format(Total(k), NEGNUMFMT)
      Else
         Grid.TextMatrix(row, C_DESC) = String(REP_INDENT * (IFRS_MAXNIVEL - 2), " ") & StrTotal & Padre(k)
         Grid.TextMatrix(row, C_SALDO) = Format(Total(k), NEGNUMFMT)
      End If
      
      Grid.TextMatrix(row, C_NIVEL) = k
      
      If Trim(LCase(Grid.TextMatrix(row, C_DESC))) = "total patrimonio" Then
         TotPatrimonio = vFmt(Grid.TextMatrix(row, C_SALDO))
         RowPatrimonio = row
         
      ElseIf Trim(LCase(Grid.TextMatrix(row, C_DESC))) = "total pasivos" Then
         TotPasivos = vFmt(Grid.TextMatrix(row, C_SALDO))
         RowPasivos = row
      End If
      
      Call FGrFontBold(Grid, row, C_DESC, True)
      Call FGrFontBold(Grid, row, C_SALDO, True)
      Grid.TextMatrix(row, C_FMT) = "B"
      
      Grid.TextMatrix(row, C_SALDO) = ""
      
      row = row + 2
   Next k
   
   'traemos los saldos de las cuentas de Resultado Integral para asignarlos a Otras Reservas
   
   Q1 = "SELECT IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion as Descr"
   Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion as ClasCta "
   Q1 = Q1 & ", Sum(MovComprobante.Debe) as SumDebe, Sum(MovComprobante.Haber) As SumHaber  "
   
   Q1 = Q1 & " FROM (( IFRS_PlanIFRS "
   Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
   Q1 = Q1 & " LEFT JOIN MovComprobante ON Cuentas.IdCuenta = MovComprobante.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   
   Q1 = Q1 & " WHERE IFRS_PlanIFRS.Nivel = " & gLastNivelIFRS & " AND Left(IFRS_PlanIFRS.Codigo,3) = '401'"
   Q1 = Q1 & " AND " & WhFecha
   Q1 = Q1 & " AND (Comprobante.Estado IS NULL OR Comprobante.Estado = " & EC_APROBADO & ")"
   Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND (Comprobante.IdEmpresa IS NULL OR Comprobante.IdEmpresa = " & gEmpresa.id & ") AND (Comprobante.Ano IS NULL OR Comprobante.Ano = " & gEmpresa.Ano & ")"
   
   Q1 = Q1 & " GROUP BY IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion "
   Q1 = Q1 & ", IFRS_PlanIFRS.Clasificacion "
   
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
   
   Grid.TextMatrix(RowOtrasReservas + 7, C_SALDO) = Format(TotReservas, NUMFMT)
   

   
   'ajustamos las Ganancias(Pérdidas) acumuladas con el Patrimonio
   
   'ajustamos Utilidad o Pérdida Neta del Período
   TotOriGananciasPerdidas = TotGananciasPerdidas
   TotUtilidadPerdida = TotActivos - TotPasivos - TotReservas
   Grid.TextMatrix(RowUtilidadPerdida, C_SALDO) = Format(TotUtilidadPerdida, NEGNUMFMT)    'ponemos el total

   'ajustamos las Ganancias(Pérdidas) acumuladas con el Patrimonio
   TotGananciasPerdidas = TotGananciasPerdidas + TotUtilidadPerdida
   Grid.TextMatrix(RowGananciasPerdidas, C_SALDO) = Format(TotGananciasPerdidas, NEGNUMFMT)

   'ajustamos el Patrimonio
   TotPatrimonio = TotPatrimonio + TotUtilidadPerdida + TotReservas
   Grid.TextMatrix(RowPatrimonio, C_SALDO) = Format(TotPatrimonio, NEGNUMFMT)

   'ajustamos el los Pasivos
   TotPasivos = TotActivos
   Grid.TextMatrix(RowPasivos, C_SALDO) = Format(TotPasivos, NEGNUMFMT)

   
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
      
         If vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivel Then
            Grid.RowHeight(i) = 0
         End If
      Next i
      
      'luego el penúltimo nivel
      For i = Grid.FixedRows + 1 To Grid.rows - 1
      
      
         If Grid.RowHeight(i) > 0 Then
            If vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivel - 1 And Grid.TextMatrix(i, C_CODIGO) <> "" Then
               PadreN3 = i
               n = 0
            ElseIf vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivel Then
               n = n + 1
            End If
            
            'si es total de nivel 3 y no hay líneas de detalle con saldo en este rubro, lo coultamos completo
            If vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And vFmt(Grid.TextMatrix(i, C_NIVEL)) = gLastNivel - 1 And Grid.TextMatrix(i, C_CODIGO) = "" And n = 0 Then
               Grid.RowHeight(i) = 0
               Grid.RowHeight(PadreN3) = 0
               Grid.RowHeight(i + 1) = 0
            End If
         End If
      Next i
   End If

   Call FGrVRows(Grid, 2)

   If Ch_SaldosVig <> 0 Then
      Call SetParallelGrid_SinValCero
   Else
      Call SetParallelGrid_ConValCero
   End If
   
   GridV.Redraw = True
'   Grid.Visible = True
'   GridV.Visible = False
   Call EnableFrm(False)
   
End Sub

Private Sub SetParallelGrid_ConValCero()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim InitPasivo As Integer, InitPatrimonio As Integer
   Dim r As Integer
   Dim RowActCorrientes As Integer, RowActNoCorrientes As Integer
   Dim RowPasCorrientes As Integer, RowPasNoCorrientes As Integer
   Dim RowEndCorrientes As Integer
   Dim RowTotPatrimonio As Integer
   Dim RowActTotal As Integer
   Dim RowPasTotal As Integer, RowPatPasTotal As Integer
   Dim RowEndTotal As Integer
   Dim Descrip As String
   
   For i = 0 To Grid.rows - 1
            
      'marcamos inicio de pasivos
      If Left(Grid.TextMatrix(i, C_CODIGO), 1) = "2" And InitPasivo = 0 Then
         InitPasivo = i
      End If
      
      Descrip = LCase(Trim(Grid.TextMatrix(i, C_DESC)))
            
      'marcamos linea "Total Activos Corrientes"
      If Descrip = LCase("Total Activos Corrientes") Then
         If RowActCorrientes = 0 Then
            RowActCorrientes = i
         End If
      End If
      
      'marcamos linea "Total Activos No Corrientes"
      If Descrip = LCase("Total Activos no Corrientes") Then
         If RowActNoCorrientes = 0 Then
            RowActNoCorrientes = i
         End If
      End If
   
      'marcamos linea "Total Pasivos Corrientes"
      If Descrip = LCase("Total Pasivos Corrientes") Then
         If RowPasCorrientes = 0 Then
            RowPasCorrientes = i
         End If
      End If
      
      'marcamos linea "Total Pasivos No Corrientes"
      If Descrip = LCase("Total Pasivos No Corrientes") Then
         If RowPasNoCorrientes = 0 Then
            RowPasNoCorrientes = i
         End If
      End If
   
      'marcamos línea de "Total Activos"
      If Descrip = "total activos" Then
         If RowActTotal = 0 Then
            RowActTotal = i
         End If
         
      'marcamos línea de "Total Pasivos"
      ElseIf Descrip = "total pasivos" Then
         If RowPasTotal = 0 Then
            RowPasTotal = i
         End If
      
      'marcamos linea "Patrimonio"
      ElseIf Descrip = "patrimonio" Then
         If InitPatrimonio = 0 Then
            InitPatrimonio = i
         End If
      
      'marcamos linea "Total Patrimonio"
      ElseIf Descrip = "total patrimonio" Then
         If RowTotPatrimonio = 0 Then
            RowTotPatrimonio = i
         End If
         
      End If
            
   Next i
   
   'máximo de líneas de Corrientes
   If RowActCorrientes > RowPasCorrientes - InitPasivo Then
      RowEndCorrientes = RowActCorrientes
   Else
      RowEndCorrientes = RowPasCorrientes - InitPasivo
   End If
   
   'máximo de líneas de total (activo o pasivo)
   If RowActTotal - RowActCorrientes > RowPatPasTotal - RowPasCorrientes Then
      RowEndTotal = RowEndCorrientes + RowActTotal - RowActCorrientes
   Else
      RowEndTotal = RowEndCorrientes + RowPatPasTotal - RowPasCorrientes - 2
   End If

'   RowEndTotal = RowEndTotal + 3   'esto se hace para poder correr el patrimonio más abajo de la linea del total activos no correintes
            
   GridV.rows = 0
   
   'copiamos los activos Corrientes
   For i = 0 To RowActCorrientes - 1
      
      GridV.rows = GridV.rows + 1
      
      For j = 0 To C_SALDO
         GridV.TextMatrix(i, j) = Grid.TextMatrix(i, j)

      Next j
      
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
      
'      Call FGrSetRowStyle(GridV, i, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_SALDO)
      GridV.TextMatrix(i, C_FMT) = "FCELL"
         
   Next i
      
   'rellenamos para llegar al total Corriente
   For i = RowActCorrientes To RowEndCorrientes
      GridV.rows = GridV.rows + 1
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   'copiamos el total activos Corrientes
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowEndCorrientes, j) = Grid.TextMatrix(RowActCorrientes, j)
   Next j
   
   'ahora el resto de los activos
   
   k = RowEndCorrientes + 1
   For i = RowActCorrientes + 1 To RowActTotal - 1
      
      GridV.rows = GridV.rows + 1
      
      For j = 0 To C_SALDO
         GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
      Next j
      
      If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
         GridV.TextMatrix(k, C_IDCUENTA) = "*"
      End If
      
'      Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_SALDO)
      GridV.TextMatrix(k, C_FMT) = "FCELL"
      
      k = k + 1
      
   Next i
      
   'rellenamos para llegar a la línea de total activos
   For i = RowActTotal To RowEndTotal
      GridV.rows = GridV.rows + 1
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   'copiamos el total activo
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowEndTotal, j) = Grid.TextMatrix(RowActTotal, j)
   Next j
      
   
   'y luego los pasivos Corrientes
   r = 0
   For i = InitPasivo To RowPasCorrientes - 1
      
      If r >= GridV.rows Then
         GridV.rows = GridV.rows + 1
      End If
      
      For j = 0 To C_SALDO
         GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
      Next j
      
'      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
      GridV.TextMatrix(r, C_FMT) = "FCELL"
      
      r = r + 1
   Next i
   
   'saltamos hasta llegar al total Corrientes

   'copiamos el total pasivos Corrientes
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowEndCorrientes, j + C_CODIGO_P) = Grid.TextMatrix(RowPasCorrientes, j)
   Next j
   
   GridV.TextMatrix(RowEndCorrientes, C_FMT) = "FCELL"

   'ahora el resto de los pasivos hasta antes del total
   
   r = RowEndCorrientes + 1
   For i = RowPasCorrientes + 1 To InitPatrimonio - 1
      
      If r >= GridV.rows Then
         GridV.rows = GridV.rows + 1
      End If
      
      For j = 0 To C_SALDO
         GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
      Next j
      
'      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
      GridV.TextMatrix(r, C_FMT) = "FCELL"
     
      r = r + 1
      
   Next i

   
   'agregamos el patrimonio
'   r = r + 3     'esto se hace para poder correr el patrimonio más abajo de la linea del total activos no correintes
   
   For i = InitPatrimonio To RowTotPatrimonio

      If r >= GridV.rows Then
         GridV.rows = GridV.rows + 1
      End If

      For j = 0 To C_SALDO
         GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
      Next j

'      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
      GridV.TextMatrix(r, C_FMT) = "FCELL"

      r = r + 1
   Next i
   
   
   
   'saltamos hasta llegar a la línea de total pasivo
   
   'copiamos el total pasivo final
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowEndTotal, j + C_CODIGO_P) = Grid.TextMatrix(RowPasTotal, j)
   Next j
   
   'cambiamos el total final de pasivos
   GridV.TextMatrix(RowEndTotal, C_DESC_P) = "TOTAL PASIVOS"


   GridV.TextMatrix(RowEndTotal, C_FMT) = "FCELL"
   
   Call FGrVRows(GridV, 1)
      
   'eliminamos el código de la ahora 1ª línea de activo y de la ahora 1ª línea de pasivo
   GridV.TextMatrix(0, C_CODIGO) = ""
   GridV.TextMatrix(0, C_CODIGO_P) = ""
   
   'ponemos en mayúscula las líneas que corresponde
   GridV.TextMatrix(0, C_DESC) = UCase(GridV.TextMatrix(0, C_DESC))
   GridV.TextMatrix(0, C_DESC_P) = UCase(GridV.TextMatrix(0, C_DESC_P))
   GridV.TextMatrix(RowEndTotal, C_DESC) = UCase(GridV.TextMatrix(RowEndTotal, C_DESC))
   GridV.TextMatrix(RowEndTotal, C_DESC_P) = UCase(GridV.TextMatrix(RowEndTotal, C_DESC_P))
   
   'ponemos en bold las líneas que corresponde
   For i = 0 To GridV.rows - 1
      If GridV.TextMatrix(i, C_CODIGO) = "" And GridV.TextMatrix(i, C_DESC) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC, C_SALDO)
      End If
      If GridV.TextMatrix(i, C_CODIGO_P) = "" And GridV.TextMatrix(i, C_DESC_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC_P, C_SALDO_P)
      End If
      If Right(VFmtCodigoIFRS(GridV.TextMatrix(i, C_CODIGO)), 2) = "00" And GridV.TextMatrix(i, C_DESC) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC, C_SALDO)
      End If
      If Right(VFmtCodigoIFRS(GridV.TextMatrix(i, C_CODIGO_P)), 2) = "00" And GridV.TextMatrix(i, C_DESC_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC_P, C_SALDO_P)
      End If
      
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = GridV.TextMatrix(i, C_IDCUENTA_P)
      End If

   Next i
   


End Sub

Private Sub SetParallelGrid_SinValCero()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim InitPasivo As Integer, InitPatrimonio As Integer
   Dim r As Integer
   Dim RowActCorriente As Integer, RowActNoCorriente As Integer
   Dim RowPasCorriente As Integer, RowPasNoCorriente As Integer
   Dim RowEndCorriente As Integer, NEndCorriente As Integer
   Dim RowTotPatrimonio As Integer
   Dim NActCorriente As Integer, NPasCorriente As Integer
   Dim RowActTotal As Integer
   Dim RowPasTotal As Integer, RowPatPasTotal As Integer
   Dim RowEndTotal As Integer
   Dim Descrip As String
   Dim RowTotalActivo As Integer, RowTotalPasivo As Integer, RowTotalFinal As Integer
   
   For i = 0 To Grid.rows - 1
            
      'marcamos inicio de pasivos
      If Left(Grid.TextMatrix(i, C_CODIGO), 1) = "2" And InitPasivo = 0 Then
         InitPasivo = i
      End If
      
      
      'contamos los Activos Corriente > 0
      If RowActCorriente = 0 Then
         If Grid.RowHeight(i) > 0 Then
            NActCorriente = NActCorriente + 1
         End If
      End If
      
      'contamos los Pasivos Corriente > 0
      If InitPasivo > 0 And RowPasCorriente = 0 Then
         If Grid.RowHeight(i) > 0 Then
            NPasCorriente = NPasCorriente + 1
         End If
      End If

      Descrip = LCase(Trim(Grid.TextMatrix(i, C_DESC)))
            
      'marcamos linea "Total Activos Corrientes"
      If Descrip = LCase("Total Activos Corrientes") Then
         If RowActCorriente = 0 Then
            RowActCorriente = i
         End If
      End If
      
      'marcamos linea "Total Activos No Corrientes"
      If Descrip = LCase("Total Activos no Corrientes") Then
         If RowActNoCorriente = 0 Then
            RowActNoCorriente = i
         End If
      End If
   
      'marcamos linea "Total Pasivos Corrientes"
      If Descrip = LCase("Total Pasivos Corrientes") Then
         If RowPasCorriente = 0 Then
            RowPasCorriente = i
         End If
      End If
      
      'marcamos linea "Total Pasivos No Corrientes"
      If Descrip = LCase("Total Pasivos No Corrientes") Then
         If RowPasNoCorriente = 0 Then
            RowPasNoCorriente = i
         End If
      End If
   
      'marcamos línea de "Total Activos"
      If Descrip = "total activos" Then
         If RowActTotal = 0 Then
            RowActTotal = i
         End If
         
      'marcamos línea de "Total Pasivos"
      ElseIf Descrip = "total pasivos" Then
         If RowPasTotal = 0 Then
            RowPasTotal = i
         End If
      
      'marcamos linea "Patrimonio"
      ElseIf Descrip = "patrimonio" Then
         If InitPatrimonio = 0 Then
            InitPatrimonio = i
         End If
      
      'marcamos linea "Total Patrimonio"
      ElseIf Descrip = "total patrimonio" Then
         If RowTotPatrimonio = 0 Then
            RowTotPatrimonio = i
         End If
         
      End If
            
   Next i
   
   'máximo de líneas de Corriente
   If RowActCorriente > RowPasCorriente - InitPasivo Then
      RowEndCorriente = RowActCorriente
   Else
      RowEndCorriente = RowPasCorriente - InitPasivo
   End If
   
   If NActCorriente > NPasCorriente Then
      NEndCorriente = NActCorriente
   Else
      NEndCorriente = NPasCorriente
   End If
            
   GridV.rows = 0
   k = 0
   
   'copiamos los activos Corriente
   For i = 0 To RowActCorriente - 1
   
      If Grid.RowHeight(i) > 0 Then
      
         GridV.rows = GridV.rows + 1
         
         For j = 0 To C_SALDO
            GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
   
         Next j
         
         If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
            GridV.TextMatrix(k, C_IDCUENTA) = "*"
         End If
         
'         Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_SALDO)
         GridV.TextMatrix(k, C_FMT) = "FCELL"
      
         k = k + 1
      End If
      
   Next i
      
   GridV.rows = GridV.rows + 1
   
   'rellenamos para llegar al total Corriente
   For i = k To NEndCorriente
      GridV.rows = GridV.rows + 1
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   
   'copiamos el total activos Corrientes
   For j = 0 To C_SALDO
      GridV.TextMatrix(NEndCorriente, j) = Grid.TextMatrix(RowActCorriente, j)
      'GridV.TextMatrix(k, j) = Grid.TextMatrix(RowActCorriente, j)
   Next j
   
   'ahora el resto de los activos
   
   k = NEndCorriente + 1
   For i = RowActCorriente + 1 To RowActTotal - 1
      
      If Grid.RowHeight(i) > 0 Then
         GridV.rows = GridV.rows + 1
         
         For j = 0 To C_SALDO
            GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
         Next j
         
         If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
            GridV.TextMatrix(k, C_IDCUENTA) = "*"
         End If
         
   '      Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_SALDO)
         GridV.TextMatrix(k, C_FMT) = "FCELL"
         
         k = k + 1
      End If
      
   Next i
         
   RowTotalActivo = k
   
   'y luego los pasivos Corriente
   r = 0
   For i = InitPasivo To RowPasCorriente - 1
   
      If Grid.RowHeight(i) > 0 Then
      
         If r >= GridV.rows Then
            GridV.rows = GridV.rows + 1
         End If
         
         For j = 0 To C_SALDO
            GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
         Next j
         
   '      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
         GridV.TextMatrix(r, C_FMT) = "FCELL"
         
         r = r + 1
      End If
   Next i
   
   'saltamos hasta llegar al total Corriente

   'copiamos el total pasivos Corriente
   For j = 0 To C_SALDO
      GridV.TextMatrix(NEndCorriente, j + C_CODIGO_P) = Grid.TextMatrix(RowPasCorriente, j)
   Next j
   
   GridV.TextMatrix(NEndCorriente, C_FMT) = "FCELL"

   'ahora el resto de los pasivos hasta antes del total
   
   r = NEndCorriente + 1
   For i = RowPasCorriente + 1 To InitPatrimonio - 1
      
      If Grid.RowHeight(i) > 0 Then
         If r >= GridV.rows Then
            GridV.rows = GridV.rows + 1
         End If
         
         For j = 0 To C_SALDO
            GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
         Next j
         
   '      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
         GridV.TextMatrix(r, C_FMT) = "FCELL"
        
         r = r + 1
      End If
      
   Next i

   
   'agregamos el patrimonio
'   r = r + 3     'esto se hace para poder correr el patrimonio más abajo de la linea del total activos no corrientes
   
   For i = InitPatrimonio To RowTotPatrimonio

      If Grid.RowHeight(i) > 0 Then
      
         If r >= GridV.rows Then
            GridV.rows = GridV.rows + 1
         End If
   
         For j = 0 To C_SALDO
            GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
         Next j
   
   '      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_SALDO)
         GridV.TextMatrix(r, C_FMT) = "FCELL"
   
         r = r + 1
      End If
   Next i
   
   
   
   'saltamos hasta llegar a la línea de total pasivo
   
   RowTotalPasivo = r
   
   If RowTotalActivo > RowTotalPasivo Then
      RowTotalFinal = RowTotalActivo
   Else
      RowTotalFinal = RowTotalPasivo
   End If
   
   GridV.rows = GridV.rows + 1

   'copiamos el total activo
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowTotalFinal, j) = Grid.TextMatrix(RowActTotal, j)
   Next j
      
   'copiamos el total pasivo final
   For j = 0 To C_SALDO
      GridV.TextMatrix(RowTotalFinal, j + C_CODIGO_P) = Grid.TextMatrix(RowPasTotal, j)
   Next j

   GridV.TextMatrix(RowTotalFinal, C_FMT) = "FCELL"
   
   Call FGrVRows(GridV, 1)
      
   'eliminamos el código de la ahora 1ª línea de activo y de la ahora 1ª línea de pasivo
   GridV.TextMatrix(0, C_CODIGO) = ""
   GridV.TextMatrix(0, C_CODIGO_P) = ""
   
   'ponemos en mayúscula las líneas que corresponde
   GridV.TextMatrix(0, C_DESC) = UCase(GridV.TextMatrix(0, C_DESC))
   GridV.TextMatrix(0, C_DESC_P) = UCase(GridV.TextMatrix(0, C_DESC_P))
   GridV.TextMatrix(RowTotalFinal, C_DESC) = UCase(GridV.TextMatrix(RowTotalFinal, C_DESC))
   GridV.TextMatrix(RowTotalFinal, C_DESC_P) = UCase(GridV.TextMatrix(RowTotalFinal, C_DESC_P))
   
   'ponemos en bold las líneas que corresponde
   For i = 0 To GridV.rows - 1
      If GridV.TextMatrix(i, C_CODIGO) = "" And GridV.TextMatrix(i, C_DESC) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC, C_SALDO)
      End If
      If GridV.TextMatrix(i, C_CODIGO_P) = "" And GridV.TextMatrix(i, C_DESC_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC_P, C_SALDO_P)
      End If
      If Right(VFmtCodigoIFRS(GridV.TextMatrix(i, C_CODIGO)), 2) = "00" And GridV.TextMatrix(i, C_DESC) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC, C_SALDO)
      End If
      If Right(VFmtCodigoIFRS(GridV.TextMatrix(i, C_CODIGO_P)), 2) = "00" And GridV.TextMatrix(i, C_DESC_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_DESC_P, C_SALDO_P)
      End If
      
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = GridV.TextMatrix(i, C_IDCUENTA_P)
      End If
      
   Next i
   


End Sub

Private Sub Form_Resize()
    Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - 2 * (GridV.Left + W.xFrame)
   If d > 1000 Then
      GridV.Width = d
   End If

 
   d = Me.Height - GridV.Top - W.YCaption * 2 + 80 - 30
   If d > 1000 Then
      GridV.Height = d
      If Fr_Filtro.visible = False Then
         GridV.Top = Fr_Filtro.Top
         GridV.Height = Grid.Height + Fr_Filtro.Height
      End If
   Else
      Me.Height = GridV.Top + 1000 + W.YCaption * 2
   End If
      
   Call FGrVRows(GridV)
   
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
   
   Call EnableFrm(True)

   
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

Private Sub Ch_VerCodCta_Click()
   Dim i As Integer
      
   If Ch_VerCodCta <> 0 Then
      GridV.ColWidth(C_CODIGO) = lWCodCta
      GridV.ColWidth(C_CODIGO_P) = lWCodCta
      
      GridV.ColWidth(C_DESC) = lWDesc
      GridV.ColWidth(C_SALDO) = lWVal
      GridV.ColWidth(C_DESC_P) = lWDesc
      GridV.ColWidth(C_SALDO_P) = lWVal
      
   Else
      GridV.ColWidth(C_CODIGO) = 0
      GridV.ColWidth(C_CODIGO_P) = 0
      
      GridV.ColWidth(C_DESC) = lWDesc + 900
      GridV.ColWidth(C_SALDO) = lWVal + 400
      GridV.ColWidth(C_DESC_P) = lWDesc + 900
      GridV.ColWidth(C_SALDO_P) = lWVal + 400
   End If
   
End Sub

