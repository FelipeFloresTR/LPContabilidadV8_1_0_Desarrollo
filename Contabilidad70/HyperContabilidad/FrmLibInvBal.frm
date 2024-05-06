VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLibInvBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Inventario y Balance"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "FrmLibInvBal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5835
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10292
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   9855
      Begin VB.CheckBox Ch_VerCodCuenta 
         Caption         =   "Ver Código Cuenta"
         Height          =   195
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   1815
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
         Picture         =   "FrmLibInvBal.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "FrmLibInvBal.frx":0435
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "FrmLibInvBal.frx":07D3
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "FrmLibInvBal.frx":0B34
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Picture         =   "FrmLibInvBal.frx":0BD8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8640
         TabIndex        =   18
         Top             =   180
         Width           =   1095
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
         Left            =   120
         Picture         =   "FrmLibInvBal.frx":1092
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   960
         Picture         =   "FrmLibInvBal.frx":1539
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   60
      TabIndex        =   19
      Top             =   660
      Width           =   9855
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3480
         Picture         =   "FrmLibInvBal.frx":197E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1680
         Picture         =   "FrmLibInvBal.frx":1C88
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   230
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   660
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   2460
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   3435
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   3435
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   8640
         Picture         =   "FrmLibInvBal.frx":1F92
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   3
         Left            =   3900
         TabIndex        =   25
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   2
         Left            =   3900
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7560
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Desde:"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   23
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
      Height          =   195
      Index           =   7
      Left            =   2040
      TabIndex        =   22
      Top             =   1080
      Width           =   465
   End
End
Attribute VB_Name = "FrmLibInvBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_CUENTA = 1
Const C_MES = 2
Const C_DEBE = 3
Const C_HABER = 4
Const C_SALDO = 5
Const C_IDCUENTA = 6
Const C_OBLIGATORIA = 7
Const C_FMT = 8

Const NCOLS = C_FMT

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean
Dim lWCodCta As Integer

Dim lMes As Integer

Private Sub SetUpGrid()
   Dim Col As Integer
   
   lWCodCta = 1400
   
   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
   Else
      Grid.ColWidth(C_CODIGO) = 0
   End If
   
   Grid.ColWidth(C_CUENTA) = 3170
   Grid.ColWidth(C_MES) = 1000
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_SALDO) = 1300
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_MES) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   Grid.TextMatrix(0, C_CODIGO) = "Código"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_MES) = "Mes"
   Grid.TextMatrix(0, C_DEBE) = "Débitos"
   Grid.TextMatrix(0, C_HABER) = "Créditos"
   Grid.TextMatrix(0, C_SALDO) = "Saldo"
   Grid.TextMatrix(0, C_FMT) = "       .FMT"
   
   Call FGrVRows(Grid)
   
   
End Sub

Private Sub Bt_Buscar_Click()
   Dim F1 As Long
   Dim F2 As Long
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
      
   If F1 > F2 Then
      MsgBeep vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   Me.MousePointer = vbHourglass
   
   Call LoadAll
   
   Me.MousePointer = vbDefault
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call LP_FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview

   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtLibros.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call SetPrtNotas(False)  'dejamos nota Art. 100 como para balances
   
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
   
   lPapelFoliado = False
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(LIBOF_INVBAL, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(LIBOF_INVBAL) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
      
      lPapelFoliado = True
   
   End If
   
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Call SetUpPrtGrid
      nFolio = gPrtLibros.PrtFlexGrid(Printer)
      
      If lPapelFoliado And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(LIBOF_INVBAL, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, nFolio)
      
      lInfoPreliminar = False
      Printer.Orientation = OldOrientacion
      
   End If
   
   Call SetPrtNotas(False)  'dejamos nota Art. 100 como para balances
   Call ResetPrtBas(gPrtLibros)
      
End Sub


Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoAjuste_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_VerCodCuenta_Click()

   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
   Else
      Grid.ColWidth(C_CODIGO) = 0
   End If
   
   GridTot.ColWidth(C_CODIGO) = Grid.ColWidth(C_CODIGO)
   

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   
   ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
   
   Call FirstLastMonthDay(ActDate, D1, D2)
   Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
   Call SetTxDate(Tx_Hasta, D2)
   
   Call FillCbAreaNeg(Cb_AreaNeg, False)
   Call FillCbCCosto(Cb_CCosto, False)
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)


   
   lOrientacion = ORIENT_VER
   
   Ch_VerCodCuenta = 1
   
   Call SetUpGrid
   Call LoadAll
   Call SetupPriv
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim FontTit(1) As FontDef_t
   Dim Encabezados(3) As String
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   Dim F1 As Long, F2 As Long
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
      
   Titulos(0) = Caption
   FontTit(0).FontBold = True
   
   If lInfoPreliminar Then
      Titulos(1) = INFO_PRELIMINAR
      FontTit(1).FontBold = True
   End If
   
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
   If GetTxDate(Tx_Desde) <> DateSerial(gEmpresa.Ano, 1, 1) Then
      Encabezados(0) = Format(GetTxDate(Tx_Desde), DATEFMT) & " a "
   Else
      Encabezados(0) = "Al "
   End If
   Encabezados(0) = Encabezados(0) & Format(GetTxDate(Tx_Hasta), DATEFMT)
   
   i = 1
   'PS 26/10/2005 para distinguir q filtro se hizo
   If Cb_AreaNeg.ListIndex > 0 Then
      Encabezados(i) = "Area de Negocio   : " & Cb_AreaNeg
      i = i + 1
   End If
   
   If Cb_CCosto.ListIndex > 0 Then
      Encabezados(i) = "Centro de Gestión : " & Cb_CCosto
   End If
   '****
   
   gPrtLibros.Encabezados = Encabezados
   
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
     
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_OBLIGATORIA
   gPrtLibros.NTotLines = 1
   
   Call SetPrtNotas(True)  'dejamos nota Art. 100 como para libros
     
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
  
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim F1 As Long, F2 As Long
   Dim Total(C_SALDO) As Double
   Dim IdCuenta As Long
   Dim WhEstado As String
   Dim Wh As String
   Dim TotDebe As Double
   Dim TotHaber As Double
   
   'Limpio Grilla TOTALES
   GridTot.Clear
   
   Grid.Redraw = False
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)

   If Ch_LibOficial <> 0 Then
      WhEstado = " AND Comprobante.Estado=" & EC_APROBADO
      MsgBox1 "Dado que es Libro Oficial, sólo se seleccionarán los comprobantes APROBADOS.", vbInformation + vbOKOnly
   Else
      WhEstado = " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   End If
   
   If ItemData(Cb_AreaNeg) > 0 Then
      Wh = Wh & " AND MovComprobante.IdAreaNeg = " & ItemData(Cb_AreaNeg)
   End If
   
   If ItemData(Cb_CCosto) > 0 Then
      Wh = Wh & " AND MovComprobante.IdCCosto = " & ItemData(Cb_CCosto)
   End If
   
   If ItemData(Cb_TipoAjuste) > 0 Then
      If ItemData(Cb_TipoAjuste) = TAJUSTE_FINANCIERO Then
         Wh = Wh & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
      Else
         Wh = Wh & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
      End If
   End If


   Q1 = "SELECT Codigo, Descripcion, MovComprobante.idCuenta, Cuentas.Debe as CtaDebe "
   Q1 = Q1 & ", Cuentas.Haber as CtaHaber, Sum(MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber "
   Q1 = Q1 & ", " & SqlMonthLng("Comprobante.Fecha") & " as Mes"
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp = Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.idCuenta = Cuentas.idCuenta  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   Q1 = Q1 & " WHERE (Fecha BETWEEN  " & F1 & " AND " & F2 & ")"
   Q1 = Q1 & " AND Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & "," & CLASCTA_ORDEN & ") "
   Q1 = Q1 & WhEstado & Wh
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Codigo, Descripcion, MovComprobante.idCuenta, Cuentas.Debe, Cuentas.Haber "
   Q1 = Q1 & " , " & SqlMonthLng("Fecha")
   Q1 = Q1 & " ORDER BY Codigo, " & SqlMonthLng("Fecha")
   
   Set Rs = OpenRs(DbMain, Q1)

   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   IdCuenta = 0
   
   Do While Rs.EOF = False

      Grid.rows = Grid.rows + 1
      
      If IdCuenta <> vFld(Rs("IdCuenta")) Then
      
         If IdCuenta <> 0 Then   'ponemos total cuenta anterior
            Call FGrSetRowStyle(Grid, Row, "B")
            Grid.TextMatrix(Row, C_FMT) = "B"
            Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
            
            Grid.TextMatrix(Row, C_MES) = "Total"
            Grid.TextMatrix(Row, C_DEBE) = Format(TotDebe, NUMFMT)
            Grid.TextMatrix(Row, C_HABER) = Format(TotHaber, NUMFMT)
            Grid.TextMatrix(Row, C_SALDO) = Format(TotDebe - TotHaber, NEGNUMFMT)
            
            TotDebe = 0
            TotHaber = 0
            
            Grid.rows = Grid.rows + 1
            Row = Row + 1
         End If
         
         Grid.TextMatrix(Row, C_CODIGO) = FmtCodCuenta(vFld(Rs("Codigo")))
         Grid.TextMatrix(Row, C_CUENTA) = FCase(vFld(Rs("Descripcion"), True))
         Grid.TextMatrix(Row, C_IDCUENTA) = vFld(Rs("IdCuenta"))
         Call FGrSetRowStyle(Grid, Row, "B")
         Grid.TextMatrix(Row, C_FMT) = "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         
         IdCuenta = vFld(Rs("IdCuenta"))
         Grid.rows = Grid.rows + 1
         Row = Row + 1
      End If
      
      Grid.TextMatrix(Row, C_MES) = Format(DateSerial(gEmpresa.Ano, vFld(Rs("Mes")), 1), "mmm yyyy")
      Grid.TextMatrix(Row, C_DEBE) = Format(vFld(Rs("Debe")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_HABER) = Format(vFld(Rs("Haber")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_SALDO) = Format(vFld(Rs("Debe")) - vFld(Rs("Haber")), NEGBL_NUMFMT)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      
      TotDebe = TotDebe + vFld(Rs("Debe"))
      TotHaber = TotHaber + vFld(Rs("Haber"))
      
      Total(C_DEBE) = Total(C_DEBE) + vFld(Rs("Debe"))
      Total(C_HABER) = Total(C_HABER) + vFld(Rs("Haber"))
      Total(C_SALDO) = Total(C_SALDO) + vFld(Rs("Debe")) - vFld(Rs("Haber"))
      
      Row = Row + 1
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   'ponemos el total de la última cuenta
   
   If IdCuenta <> 0 Then
      
      Grid.rows = Grid.rows + 1
      
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      
      Grid.TextMatrix(Row, C_MES) = "Total"
      Grid.TextMatrix(Row, C_DEBE) = Format(TotDebe, NUMFMT)
      Grid.TextMatrix(Row, C_HABER) = Format(TotHaber, NUMFMT)
      Grid.TextMatrix(Row, C_SALDO) = Format(TotDebe - TotHaber, NEGNUMFMT)
      
   End If

   GridTot.TextMatrix(0, C_CUENTA) = "TOTAL"
   GridTot.TextMatrix(0, C_DEBE) = Format(Total(C_DEBE), NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(Total(C_HABER), NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(Total(C_SALDO), NUMFMT)
   
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = C_CUENTA
   Grid.ColSel = Grid.Col
   
   Grid.Redraw = True
   
   Call EnableFrm(False)
   
End Sub
Private Sub Form_Resize()
   Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Height - Grid.Top - W.YCaption * 2 - GridTot.Height + 100
   If d > 1000 Then
      Grid.Height = d
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   GridTot.Top = Grid.Top + Grid.Height + 30
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid)

End Sub
Private Sub EnableFrm(bool As Boolean)
   Bt_Buscar.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBE, C_HABER)
   
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

Public Function FView(ByVal Mes As Integer)
   Dim MesActual As Integer
   
   MesActual = GetMesActual()

   lMes = Mes
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConComps()
      End If
   End If
      
   Me.Show vbModeless
   
End Function
Private Sub Ch_LibOficial_Click()
   If Ch_LibOficial <> 0 Then
      Cb_AreaNeg.ListIndex = 0
      Cb_AreaNeg.Enabled = False
      Cb_CCosto.ListIndex = 0
      Cb_CCosto.Enabled = False
   Else
      Cb_AreaNeg.Enabled = True
      Cb_CCosto.Enabled = True
   End If
   
   Call EnableFrm(True)

End Sub

Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub

