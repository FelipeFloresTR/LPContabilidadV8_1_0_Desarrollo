VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLibMayor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Mayor"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "FrmLibMayor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Ch_VerNumDoc 
      Caption         =   "Ver N° Documento"
      Height          =   255
      Left            =   9540
      TabIndex        =   32
      Top             =   7800
      Width           =   1635
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   7740
      Width           =   9375
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11535
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
         Left            =   120
         Picture         =   "FrmLibMayor.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   540
         Picture         =   "FrmLibMayor.frx":042F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Detalle documento seleccionado"
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
         Left            =   2520
         Picture         =   "FrmLibMayor.frx":08A3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Left            =   3480
         Picture         =   "FrmLibMayor.frx":0947
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   3060
         Picture         =   "FrmLibMayor.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Convertir moneda"
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
         Left            =   3900
         Picture         =   "FrmLibMayor.frx":1046
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Calendario"
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
         Left            =   1980
         Picture         =   "FrmLibMayor.frx":146F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Copiar Excel"
         Top             =   180
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
         Left            =   1080
         Picture         =   "FrmLibMayor.frx":18B4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10320
         TabIndex        =   22
         Top             =   180
         Width           =   1095
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
         Left            =   1500
         Picture         =   "FrmLibMayor.frx":1D5B
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      TabIndex        =   23
      Top             =   660
      Width           =   11535
      Begin VB.CommandButton Bt_SelCuenta 
         Height          =   315
         Left            =   5160
         Picture         =   "FrmLibMayor.frx":2215
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Buscar cuenta"
         Top             =   540
         Width           =   315
      End
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   900
         Width           =   1635
      End
      Begin VB.CheckBox Ch_SaldosCtasSinMov 
         Caption         =   "Ver saldos de cuentas sin movimiento en periodo seleccionado"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   960
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   3315
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   3315
      End
      Begin VB.ComboBox Cb_Cuentas 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   4395
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3600
         Picture         =   "FrmLibMayor.frx":260A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1800
         Picture         =   "FrmLibMayor.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   230
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   10320
         Picture         =   "FrmLibMayor.frx":2C1E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   5700
         TabIndex        =   33
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   0
         Left            =   5700
         TabIndex        =   31
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   5700
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2100
         TabIndex        =   24
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   30
      TabIndex        =   0
      Top             =   1980
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483643
      BackColorBkg    =   16777215
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7380
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmLibMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_FECHA = 0
Const C_NUMERO = 1
Const C_NDOC = 2
Const C_GLOSA = 3
Const C_DEBITOS = 4
Const C_CREDITOS = 5
Const C_SALDO = 6
Const C_TSALDO = 7
Const C_IDCOMP = 8
Const C_IDDOC = 9
Const C_OBLIGATORIA = 10
Const C_FMT = 11

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lMes As Integer
Dim lDesde As Long
Dim lHasta As Long
Dim lIdCuenta As Long

Dim lIdCCosto As Long
Dim lIdAreaNeg As Long

Dim lTotDebe As Double
Dim lTotHaber As Double

Dim lInLoad As Boolean

Dim lTipoAjuste As Integer

Private Sub Bt_Buscar_Click()
   Dim F1 As Long, F2 As Long
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
   
   If F1 > F2 Then
      MsgBeep vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   Call LoadAll
   
   MousePointer = vbDefault
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call LP_FGr2Clip(Grid, "Cuentas: " & Cb_Cuentas & vbTab & " Fecha Inicio: " & Tx_Desde & vbTab & " Fecha Término: " & Tx_Hasta)
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
   
   If Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(LIBOF_MAYOR, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(LIBOF_MAYOR) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
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
         Call AppendLogImpreso(LIBOF_MAYOR, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, nFolio)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
   End If
      
   Call SetPrtNotas(False)  'dejamos nota Art. 100 como para balances
   Call ResetPrtBas(gPrtLibros)
      
End Sub

Private Sub Bt_SelCuenta_Click()
   Dim FrmPlan As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Nombre As String
   Dim Descrip As String
   
   Set FrmPlan = New FrmPlanCuentas
   If FrmPlan.FSelect(IdCuenta, Codigo, Nombre, Descrip, False) = vbOK Then
      Call CbSelItem(Cb_Cuentas, IdCuenta)
   End If
   
   Set FrmPlan = Nothing
   
End Sub

Private Sub Cb_Cuentas_Click()
   Call EnableFrm(True)
End Sub

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

Private Sub Ch_SaldosCtasSinMov_Click()
   Call EnableFrm(True)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   Dim D1 As Long, D2 As Long
   Dim Q1 As String
   Dim ActDate As Long
   
   lInLoad = True
   
   If lMes > 0 Then
      ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
      
      Call FirstLastMonthDay(ActDate, D1, D2)
      Call SetTxDate(Tx_Desde, D1)
      Call SetTxDate(Tx_Hasta, D2)
   
   Else
      Call SetTxDate(Tx_Desde, lDesde)
      Call SetTxDate(Tx_Hasta, lHasta)
   End If
   
   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))
   
   lOrientacion = ORIENT_VER
   
   Call FillCbCuentas(Cb_Cuentas)
   If lIdCuenta > 0 Then
      Call CbSelItem(Cb_Cuentas, lIdCuenta)
   End If

   Call FillCbAreaNeg(Cb_AreaNeg, False)
   If lIdAreaNeg > 0 Then
      Call CbSelItem(Cb_AreaNeg, lIdAreaNeg)
   End If
   
   Call FillCbCCosto(Cb_CCosto, False)
   If lIdCCosto > 0 Then
      Call CbSelItem(Cb_CCosto, lIdCCosto)
   End If
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   If lTipoAjuste > 0 Then
      Call CbSelItem(Cb_TipoAjuste, lTipoAjuste)
   Else
      Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)
   End If

   
   Ch_VerNumDoc = 1
   
   Call SetUpGrid
   Call LoadAll
   
   Call SetupPriv
   
   lInLoad = False
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - Tx_CurrCell.Height - 500
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 230
   GridTot.Width = Grid.Width - 230
   Ch_VerNumDoc.Left = GridTot.Left + GridTot.Width - Ch_VerNumDoc.Width
   Tx_CurrCell.Top = GridTot.Top + GridTot.Height + 60
   Tx_CurrCell.Width = Ch_VerNumDoc.Left - 200
   Ch_VerNumDoc.Top = Tx_CurrCell.Top
   
   Call FGrVRows(Grid)
   
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
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.ColWidth(C_FECHA) = FW_FECHA - 200
   Grid.ColWidth(C_NUMERO) = 1400
   Grid.ColWidth(C_NDOC) = 2200
   Grid.ColWidth(C_GLOSA) = 2500 + 200
   Grid.ColWidth(C_DEBITOS) = 1300
   Grid.ColWidth(C_CREDITOS) = 1300
   Grid.ColWidth(C_SALDO) = 1300
   Grid.ColWidth(C_TSALDO) = 0   '400
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_IDDOC) = 0
   
   Grid.ColWidth(C_OBLIGATORIA) = 0
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_FECHA) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMERO) = flexAlignLeftCenter
   Grid.ColAlignment(C_NDOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBITOS) = flexAlignRightCenter
   Grid.ColAlignment(C_CREDITOS) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   
   Call FGrTotales(Grid, GridTot)
   
    
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(C_FMT) As Integer
   Dim Total(C_FMT) As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim FontNom(0) As FontDef_t
   Dim Nombres(5) As String
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = "LIBRO MAYOR"
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
   gPrtLibros.Encabezados = Encabezados
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   ColWi(C_NUMERO) = Grid.ColWidth(C_NUMERO) - 300
   ColWi(C_FECHA) = Grid.ColWidth(C_FECHA) - 100
   
   If ColWi(C_NDOC) = 0 Then
      ColWi(C_GLOSA) = Grid.ColWidth(C_GLOSA) - 300
      ColWi(C_DEBITOS) = Grid.ColWidth(C_DEBITOS) + 150
      ColWi(C_CREDITOS) = Grid.ColWidth(C_CREDITOS) + 150
      ColWi(C_SALDO) = Grid.ColWidth(C_SALDO) + 150
   Else
      ColWi(C_NDOC) = Grid.ColWidth(C_NDOC) - 100
      ColWi(C_GLOSA) = Grid.ColWidth(C_GLOSA) - 100
      ColWi(C_DEBITOS) = Grid.ColWidth(C_DEBITOS) + 150
      ColWi(C_CREDITOS) = Grid.ColWidth(C_CREDITOS) + 150
      ColWi(C_SALDO) = Grid.ColWidth(C_SALDO) + 150
   End If
   
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
   Dim RsSaldos As Recordset
   Dim Row As Integer
   Dim SumDebitos As Double
   Dim SumCreditos As Double
   Dim Credito As Double, Debito As Double
   Dim IdCuenta As String
   Dim Wh As String
   Dim CodCuenta As String
   Dim Idx As Integer
   Dim i As Integer
   Dim WhEstado As String
   Dim Mes As Integer
   Dim SumDebMes As Double
   Dim SumCredMes As Double
   Dim SoloUnMes As Boolean
   Dim SumSaldosDebe As Double
   Dim SumSaldosHaber As Double
   
   CodCuenta = ""
   
   lTotDebe = 0
   lTotHaber = 0
   IdCuenta = "0"
   
   Grid.Redraw = False
   
   If ItemData(Cb_Cuentas) <> -1 Then
   
      Idx = InStr(Cb_Cuentas, " ")
      
      If Idx > 0 Then
         
         CodCuenta = Left(Cb_Cuentas, Idx - 1)
         Wh = " AND " & GenWhereCuentas(CodCuenta)
         
      End If
   
   End If
   
   If Ch_LibOficial <> 0 Then
      WhEstado = " AND Comprobante.Estado=" & EC_APROBADO
      MsgBox1 "Dado que es Libro Oficial, sólo se seleccionarán los comprobantes APROBADOS.", vbInformation + vbOKOnly
   Else
      WhEstado = " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   End If
   
   If month(GetTxDate(Tx_Desde)) = month(GetTxDate(Tx_Hasta)) Then
      SoloUnMes = True
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


   
   'primero obtenemos los saldos anteriores a la fecha solicitada en el informe
   Q1 = "SELECT MovComprobante.IdCuenta, Cuentas.Codigo, Descripcion, Sum(MovComprobante.Debe) As SumDebe, Sum(MovComprobante.Haber) as SumHaber "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Fecha < " & GetTxDate(Tx_Desde)
   Q1 = Q1 & WhEstado & Wh
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY MovComprobante.idCuenta, Cuentas.Codigo, Descripcion "
   Q1 = Q1 & " ORDER BY Cuentas.Codigo "
   
   Set RsSaldos = OpenRs(DbMain, Q1)

   'ahora el detalle del periodo
   
   'primero los no resumidos
   Q1 = "SELECT Cuentas.Codigo, Descripcion, MovComprobante.Debe as Debe, MovComprobante.Haber as Haber "
   Q1 = Q1 & ",MovComprobante.Glosa, Comprobante.idComp, Comprobante.Correlativo, ImpResumido, TipoAjuste "
   Q1 = Q1 & ",Cuentas.idCuenta, Fecha, Tipo, idMov "
   Q1 = Q1 & ",Documento.TipoLib, Documento.IdDoc, Documento.TipoDoc, Documento.NumDoc, Entidades.RUT, Entidades.Nombre, Entidades.NotValidRut "
   'damos orden a comprobantes por tipo: Apertura, Ingreso, Egreso, Traspaso
   Q1 = Q1 & ", iif(Tipo = " & TC_APERTURA & ", 1, iif(Tipo = " & TC_INGRESO & ", 2, iif(Tipo = " & TC_EGRESO & ", 3, 4))) as OrdenTipo "
   Q1 = Q1 & ", Comprobante.Glosa as GlosaComp "
   Q1 = Q1 & " FROM (((( MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta  "
   ' 2825535 se comenta la siguiente linea y se agrega la segunda linea
   'Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa )"
   ' FIN 2825535
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True) & " )"
   Q1 = Q1 & " WHERE (Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
   Q1 = Q1 & WhEstado & Wh
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
  
   If gFunciones.ComprobanteResumido Then
      '3132792
        Q1 = Q1 & " AND (ImpResumido = 0 or  ImpResumido is null) "
       'Q1 = Q1 & " AND ImpResumido = 0  "
      '3132792
   
      Q1 = Q1 & " UNION "
   
      'y ahora los resumidos
      Q1 = Q1 & "SELECT Cuentas.Codigo, Descripcion, Sum(MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber "
      Q1 = Q1 & ", ' ' as Glosa, Comprobante.idComp, Comprobante.Correlativo, ImpResumido, TipoAjuste "
      Q1 = Q1 & ",Cuentas.idCuenta, Fecha, Tipo, 0 as idMov "
      Q1 = Q1 & ",0 as TipoLib, 0 as IdDoc, 0 as TipoDoc, ' ' as NumDoc, ' ' as RUT, ' ' as Nombre, 0 as NotValidRut "
      'damos orden a comprobantes por tipo: Apertura, Ingreso, Egreso, Traspaso
      Q1 = Q1 & ", iif(Tipo = " & TC_APERTURA & ", 1, iif(Tipo = " & TC_INGRESO & ", 2, iif(Tipo = " & TC_EGRESO & ", 3, 4))) as OrdenTipo "
      Q1 = Q1 & ", Comprobante.Glosa as GlosaComp "
      
      Q1 = Q1 & " FROM ((((MovComprobante "
      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta  "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      
      Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True) & " )"
      Q1 = Q1 & " WHERE (Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
      Q1 = Q1 & WhEstado & Wh
      Q1 = Q1 & " AND ImpResumido <> 0 "
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
  
      Q1 = Q1 & " GROUP BY Cuentas.Codigo, Descripcion"
      Q1 = Q1 & ", Comprobante.idComp, Comprobante.Correlativo, ImpResumido, TipoAjuste "
      Q1 = Q1 & ", Cuentas.idCuenta, Fecha, Tipo"
      Q1 = Q1 & ", iif(Tipo = " & TC_APERTURA & ", 1, iif(Tipo = " & TC_INGRESO & ", 2, iif(Tipo = " & TC_EGRESO & ", 3, 4))) "
      Q1 = Q1 & ", Comprobante.Glosa "
   
   End If
   
   Q1 = Q1 & " ORDER BY Cuentas.Codigo, Fecha "
   
   'damos orden a comprobantes por tipo: Apertura, Ingreso, Egreso, Traspaso
   If gFunciones.ComprobanteResumido Then
      Q1 = Q1 & ", OrdenTipo "
   Else
      Q1 = Q1 & ", iif(Tipo = " & TC_APERTURA & ", 1, iif(Tipo = " & TC_INGRESO & ", 2, iif(Tipo = " & TC_EGRESO & ", 3, 4)))"
   End If
   Q1 = Q1 & ", Comprobante.IdComp, IdMov"


   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Clear
   
   Row = 1
   Grid.rows = 1
   
   'para que formatee al imprimir (al final (después de FGrVRows) le ponemos RowHeight=0 para que no se vea
   Grid.TextMatrix(0, C_FMT) = ".FMT"
   
   Mes = 0
   SumDebMes = 0
   SumCredMes = 0
   
   Do While Rs.EOF = False
   
      If FGrChkMaxSize(Grid) = True Then
         MsgBox1 "Se mostrarán " & Grid.rows & " registros.", vbInformation + vbOKOnly
         Exit Do
      End If
      
      Grid.rows = Grid.rows + 1
      
       
      If IdCuenta <> vFld(Rs("Codigo")) Then  'cambio de cuenta
      
         'If Not SoloUnMes And Mes <> 0 Then
         If Mes <> 0 Then
            Call SubTotMes(Row, SumDebMes, SumCredMes, Mes, False, vFmt(Grid.TextMatrix(Row - 1, C_SALDO)))
            
            Mes = month(vFld(Rs("Fecha")))
            SumDebMes = 0
            SumCredMes = 0
         End If
         
         If IdCuenta <> "0" Then
            Call Totales(Row, SumDebitos, SumCreditos, False)
         End If
          
         If Ch_SaldosCtasSinMov <> 0 Then
            'insertamos saldos de  cuentas que no tienen movimientos este mes
            Call InsertSaldosCtasSinMov(vFld(Rs("Codigo")), Row, RsSaldos, SumSaldosDebe, SumSaldosHaber)
             
            lTotDebe = lTotDebe + SumSaldosDebe
            lTotHaber = lTotHaber + SumSaldosHaber
         End If
         
         Call Encabezado(Format(vFld(Rs("Codigo")), gFmtCodigoCta), FCase(vFld(Rs("Descripcion"), True)), Row)
         Call FixedRows(Row)
         'IdCuenta = vFld(Rs("IdCuenta"))
         IdCuenta = vFld(Rs("Codigo"))
         
         If InsertSaldoAnterior(vFld(Rs("Codigo")), Row, RsSaldos) = True Then
            SumDebitos = vFld(RsSaldos("SumDebe"))
            SumCreditos = vFld(RsSaldos("SumHaber"))
            If Ch_SaldosCtasSinMov <> 0 Then
               RsSaldos.MoveNext
            End If
         End If
         
         Mes = month(vFld(Rs("Fecha")))
         
         SumDebMes = 0
         SumCredMes = 0
        
      End If
      
      'If Not SoloUnMes And Mes <> Month(vFld(Rs("Fecha"))) Then   'cambio de mes
      If Mes <> month(vFld(Rs("Fecha"))) Then   'cambio de mes
      
         Call SubTotMes(Row, SumDebMes, SumCredMes, Mes, False, vFmt(Grid.TextMatrix(Row - 1, C_SALDO)))
         
         Mes = month(vFld(Rs("Fecha")))
         SumDebMes = 0
         SumCredMes = 0

      End If
      
      Grid.TextMatrix(Row, C_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
      Grid.TextMatrix(Row, C_NUMERO) = Left(gTipoComp(vFld(Rs("Tipo"))), 1) & " " & vFld(Rs("Correlativo")) & IIf(vFld(Rs("TipoAjuste")) = TAJUSTE_TRIBUTARIO, "-T", "")
      Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("IdComp"))
      
      If vFld(Rs("IdDoc")) <> 0 Then
         Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("IdDoc"))
         Grid.TextMatrix(Row, C_NDOC) = "[" & GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc")) & "]   " & IIf(vFld(Rs("Rut")) <> "", FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False), "") & " " & vFld(Rs("Nombre"), True)
      ElseIf vFld(Rs("ImpResumido")) And gFunciones.ComprobanteResumido Then
         Grid.TextMatrix(Row, C_NDOC) = "[res.]"
      End If
      
      If vFld(Rs("Glosa")) <> "" Then
         Grid.TextMatrix(Row, C_GLOSA) = vFld(Rs("Glosa"))
      Else
         Grid.TextMatrix(Row, C_GLOSA) = vFld(Rs("GlosaComp"))
      End If
      
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
      
      Debito = vFld(Rs("Debe"))
      Credito = vFld(Rs("Haber"))
      
      If Debito > 0 Then
      
         Grid.TextMatrix(Row, C_DEBITOS) = Format(Debito, BL_NUMFMT)
         'Grid.TextMatrix(Row, C_SALDO) = Format(Debito, BL_NUMFMT)
         'Grid.TextMatrix(Row, C_TSALDO) = "DB"
         
         SumDebitos = SumDebitos + Debito
         SumDebMes = SumDebMes + Debito
      End If
      
      If Credito > 0 Then       'FCA 17/04/2009 se cambia Else por If dado que cuando el comprobante es resumido, podemos tener valores en el debe y en el haber
      
         Grid.TextMatrix(Row, C_CREDITOS) = Format(Credito, BL_NUMFMT)
         'Grid.TextMatrix(Row, C_SALDO) = Format(Credito, BL_NUMFMT)
         'Grid.TextMatrix(Row, C_TSALDO) = "CR"
         
         SumCreditos = SumCreditos + Credito
         SumCredMes = SumCredMes + Credito
         
      End If
      
      Grid.TextMatrix(Row, C_SALDO) = Format(Debito - Credito + vFmt(Grid.TextMatrix(Row - 1, C_SALDO)), NEGNUMFMT)
            
      Row = Row + 1
      Grid.rows = Row + 1
      
      Rs.MoveNext
   Loop
   
   'ponemos el total del último mes de la última cuenta
   'If Not SoloUnMes And SumDebMes > 0 Or SumCredMes > 0 Then
   If SumDebMes > 0 Or SumCredMes > 0 Then
      Call SubTotMes(Row, SumDebMes, SumCredMes, Mes, False, vFmt(Grid.TextMatrix(Row - 1, C_SALDO)))
   End If
   
   'ponemos el total de la última cuenta
   If SumDebitos > 0 Or SumCreditos > 0 Then
      Call Totales(Row, SumDebitos, SumCreditos, True)
   End If
   
   If Ch_SaldosCtasSinMov <> 0 Then
      'insertamos saldos de  cuentas que no tienen movimientos este mes
      If Row > 1 Then
         Row = Row + 2
      End If
      Grid.rows = Row + 1
      Call InsertSaldosCtasSinMov("", Row, RsSaldos, SumSaldosDebe, SumSaldosHaber)
      
      '2857956
       lTotDebe = lTotDebe + SumSaldosDebe
       lTotHaber = lTotHaber + SumSaldosHaber
      'fin 2857956
   End If

   Call CloseRs(Rs)
   Call CloseRs(RsSaldos)
   
   'totales finales
   GridTot.TextMatrix(0, C_GLOSA) = "TOTAL"
   GridTot.TextMatrix(0, C_DEBITOS) = Format(lTotDebe, BL_NUMFMT)
   GridTot.TextMatrix(0, C_CREDITOS) = Format(lTotHaber, BL_NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(lTotDebe - lTotHaber, NEGNUMFMT)
   
   'para que formatee al imprimir (por si se borró)
   Grid.TextMatrix(0, C_FMT) = "      .FMT"
   
   If Grid.rows <= 1 Then
      Grid.rows = 2
   End If
   
   Call FGrVRows(Grid, 2)
   Grid.RowHeight(0) = 0  'Row con el formateo
   Grid.rows = Grid.rows + 1
  
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = 0
   Grid.ColSel = 0
   
   Call EnableFrm(False)
   
   Grid.Redraw = True

End Sub
Private Sub Encabezado(CodCuenta As String, Cuenta As String, Row As Integer)
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
      
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "Align", flexAlignLeftCenter)
      
   Grid.TextMatrix(Row, C_FECHA) = "Cuenta: "
   Grid.TextMatrix(Row, C_NUMERO) = CodCuenta
   If Grid.ColWidth(C_NDOC) > 0 Then
      Grid.TextMatrix(Row, C_NDOC) = Cuenta
   Else
      Grid.TextMatrix(Row, C_GLOSA) = Cuenta
   End If
   
   
   Row = Row + 1
   Grid.rows = Row + 1
   
End Sub
Private Sub FixedRows(Row As Integer)

   Grid.TextMatrix(Row, C_FMT) = "LB"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
   
   Grid.TextMatrix(Row, C_FECHA) = "Fecha"
   Grid.TextMatrix(Row, C_NUMERO) = "N° Comp."
   If Grid.ColWidth(C_NDOC) > 0 Then
      Grid.TextMatrix(Row, C_NDOC) = "Nº Doc."
   End If
   Grid.TextMatrix(Row, C_GLOSA) = "Glosa"
   Grid.TextMatrix(Row, C_DEBITOS) = "Débitos"
   Grid.TextMatrix(Row, C_CREDITOS) = "Créditos"
   Grid.TextMatrix(Row, C_SALDO) = "Saldo"
   
   Call FGrSetRowStyle(Grid, Row, "BC", vbButtonFace)
   Call FGrSetRowStyle(Grid, Row, "Align", flexAlignCenterCenter)
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   
End Sub
Private Function InsertSaldoAnterior(ByVal CodCuenta As String, Row As Integer, RsSaldos As Recordset) As Boolean
      
   InsertSaldoAnterior = False
      
   Do While RsSaldos.EOF = False
      
      If CodCuenta = vFld(RsSaldos("Codigo")) Then
         Grid.TextMatrix(Row, C_GLOSA) = "Saldo Anterior"
         Grid.TextMatrix(Row, C_DEBITOS) = Format(vFld(RsSaldos("SumDebe")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_CREDITOS) = Format(vFld(RsSaldos("SumHaber")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_SALDO) = Format(vFld(RsSaldos("SumDebe")) - vFld(RsSaldos("SumHaber")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_FMT) = Grid.TextMatrix(Row, C_FMT) & "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
         Call FGrSetRowStyle(Grid, Row, "B")
         
         Row = Row + 1
         Grid.rows = Row + 1
         
         InsertSaldoAnterior = True
         
         Exit Function
      
      ElseIf CodCuenta < vFld(RsSaldos("Codigo")) Then   'no está la cuenta
         Grid.TextMatrix(Row, C_GLOSA) = "Saldo Anterior"
         Grid.TextMatrix(Row, C_DEBITOS) = 0
         Grid.TextMatrix(Row, C_CREDITOS) = 0
         Grid.TextMatrix(Row, C_SALDO) = 0
         Grid.TextMatrix(Row, C_FMT) = Grid.TextMatrix(Row, C_FMT) & "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
         Call FGrSetRowStyle(Grid, Row, "B")
         
         Row = Row + 1
         Grid.rows = Row + 1
                  
         Exit Function
      
      Else
         RsSaldos.MoveNext
         
      End If
      
   Loop
    
End Function
Private Function InsertSaldosCtasSinMov(ByVal CodCuentaHasta As String, Row As Integer, RsSaldos As Recordset, SumSaldosDebe As Double, SumSaldosHaber As Double) As Boolean
      
   SumSaldosDebe = 0
   SumSaldosHaber = 0
   
   InsertSaldosCtasSinMov = False
      
   Do While RsSaldos.EOF = False
      
      If vFld(RsSaldos("Codigo")) < CodCuentaHasta Or CodCuentaHasta = "" Then
         Call Encabezado(Format(vFld(RsSaldos("Codigo")), gFmtCodigoCta), FCase(vFld(RsSaldos("Descripcion"), True)), Row)
         Call FixedRows(Row)

         Grid.TextMatrix(Row, C_GLOSA) = "Saldo Anterior"
         Grid.TextMatrix(Row, C_DEBITOS) = Format(vFld(RsSaldos("SumDebe")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_CREDITOS) = Format(vFld(RsSaldos("SumHaber")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_SALDO) = Format(vFld(RsSaldos("SumDebe")) - vFld(RsSaldos("SumHaber")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_FMT) = Grid.TextMatrix(Row, C_FMT) & "B"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
         Call FGrSetRowStyle(Grid, Row, "B")
         
         SumSaldosDebe = SumSaldosDebe + vFld(RsSaldos("SumDebe"))
         SumSaldosHaber = SumSaldosHaber + vFld(RsSaldos("SumHaber"))
         
         
         Row = Row + 1
         Grid.rows = Row + 1
         Row = Row + 1
         Grid.rows = Row + 1
         
         InsertSaldosCtasSinMov = True
               
         RsSaldos.MoveNext
         
      Else
         Exit Function
         
      End If
      
   Loop
    
End Function

Private Sub Totales(Row As Integer, Debe As Double, Haber As Double, LastRs As Boolean)
   Dim Saldo As Double
   
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   
   Grid.Col = C_GLOSA
   Grid.Row = Row
   'Grid.CellAlignment = flexAlignRightCenter
   Grid.TextMatrix(Row, C_GLOSA) = "Total Cuenta"
   
   Grid.TextMatrix(Row, C_DEBITOS) = Format(Debe, BL_NUMFMT)
   Grid.TextMatrix(Row, C_CREDITOS) = Format(Haber, BL_NUMFMT)
   
   Saldo = Abs(Debe - Haber)
   Grid.TextMatrix(Row, C_SALDO) = Format(Debe - Haber, NEGNUMFMT)
   
   'If Debe > Haber Then
   '   Grid.TextMatrix(Row, C_TSALDO) = "DB"
   'ElseIf Debe < Haber Then
   '   Grid.TextMatrix(Row, C_TSALDO) = "CR"
   'End If
   
   lTotDebe = lTotDebe + Debe
   lTotHaber = lTotHaber + Haber
   
   Debe = 0
   Haber = 0
   
   If LastRs = False Then
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
      Grid.TextMatrix(Row, C_FMT) = "L"
      
      Row = Row + 1
      Grid.rows = Row + 1
   End If
   
   
End Sub
Private Sub SubTotMes(Row As Integer, ByVal Debe As Double, ByVal Haber As Double, ByVal Mes As Integer, LastRs As Boolean, ByVal SaldoPrev As Double)
   Dim Saldo As Double
   
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
   Call FGrSetRowStyle(Grid, Row, "B")
   
   Grid.Col = C_GLOSA
   Grid.Row = Row
   'Grid.CellAlignment = flexAlignRightCenter
   'Grid.TextMatrix(Row, C_GLOSA) = "SubTotal " & Left(gNomMes(Mes), 3)
   Grid.TextMatrix(Row, C_GLOSA) = "SubTotal " & gNomMes(Mes)
   
   Grid.TextMatrix(Row, C_DEBITOS) = Format(Debe, BL_NUMFMT)
   Grid.TextMatrix(Row, C_CREDITOS) = Format(Haber, BL_NUMFMT)
   
   'Saldo = Debe - Haber
   Saldo = SaldoPrev
   Grid.TextMatrix(Row, C_SALDO) = Format(Saldo, NEGNUMFMT)
         
   If LastRs = False Then
      Row = Row + 1
      Grid.rows = Row + 1
   End If
   
   
End Sub
Private Sub EnableFrm(bool As Boolean)
   Bt_Buscar.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
   
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub

Public Function FView(ByVal Mes As Integer)
   Dim MesActual As Integer

   lMes = Mes
   
   MesActual = GetMesActual()
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConComps()
      End If
   End If
               
   Me.Show vbModeless
   
End Function
Public Function FViewChain(ByVal desde As Long, ByVal Hasta As Long, Optional ByVal IdCuenta As Long = 0, Optional ByVal TipoAjuste As Integer = TAJUSTE_FINANCIERO, Optional ByVal IdCCosto As Long = 0, Optional ByVal IdAreaNeg As Long = 0)
   Dim MesActual As Integer

   lDesde = desde
   lHasta = Hasta
   
   If lDesde = 0 Then
      lDesde = DateSerial(gEmpresa.Ano, 1, 1)
   End If
   
   If lHasta = 0 Then
      lHasta = DateSerial(gEmpresa.Ano, 31, 12)
   End If
   
   lIdCuenta = IdCuenta
   
   lTipoAjuste = TipoAjuste
   
   lIdCCosto = IdCCosto
   lIdAreaNeg = IdAreaNeg
   
   Me.Show vbModal
   
End Function
Private Sub Grid_SelChange()
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBITOS, C_CREDITOS)
   
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
Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As Form
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Col = C_NUMERO Then
      If Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
'         Set Frm = New FrmComprobante
'         Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
'         Set Frm = Nothing
         
         Set Frm = New FrmLibDiario
         Call Frm.FViewChain(Val(Grid.TextMatrix(Row, C_IDCOMP)), CbItemData(Cb_TipoAjuste))
         Set Frm = Nothing
         
      End If
   ElseIf Col = C_NDOC Then
      If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
         Set Frm = New FrmDoc
         Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
         Set Frm = Nothing
      End If
   End If
      
End Sub

Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub

Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub

Private Sub Ch_VerNumDoc_Click()

   If Ch_VerNumDoc <> 0 Then
      Grid.ColWidth(C_NDOC) = 2200
      If Grid.ColWidth(C_DEBITOS) > 1200 Then
         Grid.ColWidth(C_DEBITOS) = Grid.ColWidth(C_DEBITOS) - 200
         Grid.ColWidth(C_CREDITOS) = Grid.ColWidth(C_CREDITOS) - 200
         Grid.ColWidth(C_SALDO) = Grid.ColWidth(C_SALDO) - 200
         Grid.ColWidth(C_GLOSA) = Grid.ColWidth(C_GLOSA) - 1630
      End If
   Else
      Grid.ColWidth(C_NDOC) = 0
      Grid.ColWidth(C_DEBITOS) = Grid.ColWidth(C_DEBITOS) + 200
      Grid.ColWidth(C_CREDITOS) = Grid.ColWidth(C_CREDITOS) + 200
      Grid.ColWidth(C_SALDO) = Grid.ColWidth(C_SALDO) + 200
      Grid.ColWidth(C_GLOSA) = Grid.ColWidth(C_GLOSA) + 1630
      
   End If
   
   If Not lInLoad Then
      Call LoadAll
   End If
   
End Sub

Private Sub Cb_TipoAjuste_Click()
   Call EnableFrm(True)

End Sub

