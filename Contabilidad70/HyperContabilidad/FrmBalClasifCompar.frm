VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalClasifCompar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Clasificado Comparativo"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   Icon            =   "FrmBalClasifCompar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5955
      Left            =   60
      TabIndex        =   11
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10504
      _Version        =   393216
      Rows            =   25
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      AllowUserResizing=   2
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Top             =   660
      Width           =   11955
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   2955
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   2955
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2100
         Picture         =   "FrmBalClasifCompar.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   4200
         Picture         =   "FrmBalClasifCompar.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   230
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   3180
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   10620
         Picture         =   "FrmBalClasifCompar.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1155
      End
      Begin VB.CheckBox Ch_VerSubTot 
         Caption         =   "Ver Sub-totales"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   2
         Left            =   6240
         TabIndex        =   29
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel Cuentas:"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   26
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2460
         TabIndex        =   25
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11955
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
         Left            =   3840
         Picture         =   "FrmBalClasifCompar.frx":0A5E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Enviar por Correo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerLibMayor 
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
         Picture         =   "FrmBalClasifCompar.frx":0EE1
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ir a Libro Mayor para cuenta seleccionada"
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox Ch_VerCodCuenta 
         Caption         =   "Ver Código Cuenta"
         Height          =   195
         Left            =   6060
         TabIndex        =   20
         Top             =   240
         Width           =   1815
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
         Picture         =   "FrmBalClasifCompar.frx":124F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   3000
         Picture         =   "FrmBalClasifCompar.frx":12F3
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   2580
         Picture         =   "FrmBalClasifCompar.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   3420
         Picture         =   "FrmBalClasifCompar.frx":19F2
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   1500
         Picture         =   "FrmBalClasifCompar.frx":1E1B
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   660
         Picture         =   "FrmBalClasifCompar.frx":2260
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10620
         TabIndex        =   21
         Top             =   180
         Width           =   1155
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
         Left            =   1080
         Picture         =   "FrmBalClasifCompar.frx":2707
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7680
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmBalClasifCompar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_CUENTA = 1
Const C_VALORANT = 2
Const C_VALOR = 3
Const C_DIF = 4
Const C_NIVEL = 5
Const C_IDCUENTA = 6
Const C_CLASCTA = 7
Const C_DEBITOSANT = 8
Const C_CREDITOSANT = 9
Const C_DEBITOS = 10
Const C_CREDITOS = 11
Const C_FMT = 12

Const NCOLS = C_FMT

Const TOT_CUENTA = -1

Const COLWI_CUENTAS = 5200

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lMes As Integer
Dim lCaption As String

Dim lClasCta As String
Dim lMensual As Boolean
Dim lComparativo As Boolean
Dim lBalClasif As Boolean
Dim lResEje(MAX_NIVELES) As Cuenta_t
Dim lWCodCta As Integer
Dim lWVal As Integer

Dim lLibOf As Integer

Private Sub Bt_Buscar_Click()
   Dim F1 As Long, F2 As Long
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
      
   If F1 > F2 Then
      MsgBox1 "Fecha de inicio es posterior a la fecha de término del reporte.", vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   If Year(F1) <> gEmpresa.Ano Then
      MsgBox1 "La fecha de inicio no corresponde al periodo actual.", vbExclamation
      Exit Sub
   End If
   If Year(F2) <> gEmpresa.Ano Then
      MsgBox1 "La fecha de término no corresponde al periodo actual.", vbExclamation
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
   
   'Call LP_FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   '2861570
   Call LP_FGr2Clip_Membr(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   '2861570
   
End Sub

Private Sub Bt_Email_Click()
Dim Frm As FrmEmailAccount

  Set Frm = Nothing
  Set Frm = New FrmEmailAccount
  
 Dim vAjunto As String
  vAjunto = Export_SendEmail(Grid, GridTot, Nothing, Nothing, lCaption & "_" & Tx_Desde & "_" & Tx_Hasta, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta, C_CODIGO)
   
 If Frm.FEdit(vAjunto) Then
 Frm.Show
 End If
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   gPrtLibros.CallEndDoc = False
   
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   
   '2861570
   'Call PrtPieBalance(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
   Call PrtPieBalanceFirma(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, 0)
   
   '2861570
   
   gPrtLibros.CallEndDoc = True
         
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
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
   Dim Pag As Integer
   
   lPapelFoliado = False
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(lLibOf, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(lLibOf) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
         
      lPapelFoliado = True
   End If
      
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      gPrtLibros.CallEndDoc = False
      
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
      
      
       '2861570
        'Call PrtPieBalance(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
        Call PrtPieBalanceFirma(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, 1)
        '2861570
      
      gPrtLibros.CallEndDoc = True
      
      Me.MousePointer = vbDefault
      
      If lPapelFoliado And Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(lLibOf, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, Pag)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
   End If
   
   Call ResetPrtBas(gPrtLibros)
  
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   Dim Mes As Integer
   
   Grid.Cols = NCOLS + 1
   
   lWCodCta = 1200
   lWVal = 1700
   
   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
   Else
      Grid.ColWidth(C_CODIGO) = 0
   End If
   
   Grid.ColWidth(C_CUENTA) = COLWI_CUENTAS
   Grid.ColWidth(C_VALORANT) = lWVal
   Grid.ColWidth(C_VALOR) = lWVal
   Grid.ColWidth(C_DIF) = lWVal
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_CREDITOS) = 0
   Grid.ColWidth(C_DEBITOS) = 0
   Grid.ColWidth(C_CREDITOSANT) = 0
   Grid.ColWidth(C_DEBITOSANT) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CLASCTA) = 0
   
   Grid.TextMatrix(1, C_CODIGO) = "Cód. Cuenta"
   Grid.TextMatrix(1, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_VALORANT) = "Saldo"
   Grid.TextMatrix(1, C_VALORANT) = "Periodo " & gEmpresa.Ano - 1
   Grid.TextMatrix(0, C_VALOR) = "Saldo"
   Grid.TextMatrix(1, C_VALOR) = "Periodo " & gEmpresa.Ano
   Grid.TextMatrix(0, C_DIF) = "Diferencia"
   Grid.TextMatrix(1, C_DIF) = gEmpresa.Ano & " - " & gEmpresa.Ano - 1

   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALORANT) = flexAlignRightCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_DIF) = flexAlignRightCenter

   If lBalClasif Then
      GridTot.visible = False
   End If
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
         
   Grid.TextMatrix(0, C_FMT) = "             .FMT"
   
   Call FGrVRows(Grid, 1)
   
End Sub

Private Sub Bt_VerLibMayor_Click()
   Dim Frm As FrmLibMayor
   Dim IdCuenta As Long
   
   IdCuenta = vFmt(Grid.TextMatrix(Grid.row, C_IDCUENTA))
   
   If IdCuenta > 0 Then
   
      Set Frm = New FrmLibMayor
      Call Frm.FViewChain(GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta), IdCuenta, CbItemData(Cb_TipoAjuste))
      Set Frm = Nothing
   
   End If

End Sub

Private Sub Cb_Nivel_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoAjuste_Click()
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

Private Sub Ch_VerCodCuenta_Click()
   Dim i As Integer
   
   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
      Grid.ColWidth(C_CUENTA) = COLWI_CUENTAS
      Grid.ColWidth(C_VALORANT) = lWVal
      Grid.ColWidth(C_VALOR) = lWVal
      Grid.ColWidth(C_DIF) = lWVal
   Else
      Grid.ColWidth(C_CODIGO) = 0
      Grid.ColWidth(C_CUENTA) = COLWI_CUENTAS + 600
      Grid.ColWidth(C_VALORANT) = lWVal + 200
      Grid.ColWidth(C_VALOR) = lWVal + 200
      Grid.ColWidth(C_DIF) = lWVal + 200
   End If
   
   For i = 0 To Grid.Cols - 1
      GridTot.ColWidth(i) = Grid.ColWidth(i)
   Next i
   
End Sub


Private Sub Ch_VerSubTot_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   Dim MesActual As Integer
      
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
   
   Me.Caption = lCaption
   
   Call SetUpGrid
   
   If gNiveles.nNiveles >= 3 Then
      Call FillNivel(Cb_Nivel, 3)
   Else
      Call FillNivel(Cb_Nivel, gNiveles.nNiveles)
   End If
   Cb_Nivel.RemoveItem (0)   'eliminamos el nivel 1 porque no tiene sentido
   
   Call FillCbAreaNeg(Cb_AreaNeg, False)
   Call FillCbCCosto(Cb_CCosto, False)
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)
   
   lOrientacion = ORIENT_VER
   
   Ch_VerCodCuenta = 1
   
   Ch_VerSubTot.visible = False
   
   If lBalClasif = False Then
      Ch_LibOficial.visible = False
      Ch_VerSubTot.visible = True
      Ch_VerSubTot = 1
   End If
   
   Call ReadResEje
   Call LoadAll
   Call SetupPriv
   
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi() As Integer
   Dim Total() As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = Me.Caption & " " & gEmpresa.Ano - 1 & " - " & gEmpresa.Ano
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
   If Cb_AreaNeg.ListIndex > 0 Then
      Encabezados(i) = "Area de Negocio   : " & Cb_AreaNeg
      i = i + 1
   End If
   
   If Cb_CCosto.ListIndex > 0 Then
      Encabezados(i) = "Centro de Gestión : " & Cb_CCosto
   End If
   
   gPrtLibros.Encabezados = Encabezados
   
   FontEnc(0).FontBold = True
   FontEnc(0).FontName = "Arial"
   FontEnc(0).FontSize = 10
   Call gPrtLibros.FntEncabezados(FontEnc())
   
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
    
   ReDim ColWi(Grid.Cols - 1)
   ReDim Total(Grid.Cols - 1)
    
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) > 0 Then
         ColWi(i) = Grid.ColWidth(i) - 100
      End If
   Next i
 
   For i = 0 To Grid.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_IDCUENTA
   If GridTot.visible = True Then
      gPrtLibros.NTotLines = 1
   Else
      gPrtLibros.NTotLines = 0
   End If
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String, QAnt As String
   Dim Nivel As Integer
   Dim Rs As Recordset
   Dim Total(MAX_NIVELES) As RepNiv_t
   Dim TotalAnt(MAX_NIVELES) As RepNiv_t
   Dim CurNiv As Integer
   Dim CurCta As String
   Dim i As Integer, j As Integer, k As Integer
   Dim row As Integer
   Dim Diff As Double, DiffAnt As Double
   Dim CodPadre As String, NomPadre As String
   Dim WhFecha As String
   Dim TotalFinal As Double, TotalFinalAnt As Double
   Dim Wh As String, WhFechaAnt
   Dim FDesde As Long, FDesdeAnt As Long
   Dim FHasta As Long, FHastaAnt As Long
   Dim LinPatrimonio As Integer
   Dim LinResEjercicio(MAX_NIVELES) As Integer
   Dim ClasifPadre As Integer
   Dim TotClasif(MAX_CLASCTA) As Double
   Dim TotClasifAnt(MAX_CLASCTA) As Double
   Dim LinTotClasif(MAX_CLASCTA) As Integer
   Dim ResEjercicio As Double
   Dim ResEjercicioAnt As Double
   Dim RowVisible As Boolean
   Dim Col As Integer
   Dim TmpQry As String, TmpQryAnt As String
   Dim Rc As Long
   Dim FileAnoAnt As String
   
   If Not gEmpresa.TieneAnoAnt Then
      Exit Sub
   End If
   
   TmpQry = DbGenTmpName2(SQL_ACCESS, "bc_")           'Forzamos Access para que no le ponga # en el caso de SQL, ya que no se permite para vistas
   TmpQryAnt = DbGenTmpName2(SQL_ACCESS, "bcant_")
   
   Nivel = Val(Cb_Nivel)
      
   'año actual
   FDesde = GetTxDate(Tx_Desde)
   FHasta = GetTxDate(Tx_Hasta)
   WhFecha = "(Comprobante.Fecha BETWEEN " & FDesde & " AND " & FHasta & ")"
   
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


   Q1 = GenQueryPorNiveles(Nivel, WhFecha & Wh, Ch_LibOficial <> 0, lClasCta, False, "", "", 0, False)

   'año anterior
   If Not HayAnoAnterior() Then
      MsgBox1 "Esta empresa no tiene año anterior ingresado en el sistema. No es posible presentar el informe comparativo completo.", vbInformation
      Exit Sub
   End If
               
   FDesdeAnt = DateSerial(gEmpresa.Ano - 1, Month(GetTxDate(Tx_Desde)), Day(GetTxDate(Tx_Desde)))
   FHastaAnt = DateSerial(gEmpresa.Ano - 1, Month(GetTxDate(Tx_Hasta)), Day(GetTxDate(Tx_Hasta)))
   WhFechaAnt = "(Comprobante.Fecha BETWEEN " & FDesdeAnt & " AND " & FHastaAnt & ")"
   
   QAnt = GenQueryPorNiveles(Nivel, WhFechaAnt & Wh, Ch_LibOficial <> 0, lClasCta, False, "", "", gEmpresa.Ano - 1, False)
   
#If DATACON = 1 Then       'Access
   If gEmprSeparadas Then
      FileAnoAnt = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
      If Not ExistFile(FileAnoAnt) Then
         MsgBox1 "No se encuentra el archivo de base de datos del año anterior. No es posible presentar el informe comparativo completo.", vbInformation
         Exit Sub
      End If
      
      'cerramos el año actual y abrimos el año anterior
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
      Call LinkMdbAdm
      
      'corrige base del año anterior, por si las moscas
      Call CorrigeBase
   
      'cerramos el año anterior  y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
      
      'linkeamos las tablas de Comprobante y MovComprobante del año anterior
      Call LinkMdbTable(DbMain, FileAnoAnt, "Comprobante", "ComprobanteAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, FileAnoAnt, "MovComprobante", "MovComprobanteAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, FileAnoAnt, "Cuentas", "CuentasAnt", , , gEmpresa.ConnStr)
      
      QAnt = ReplaceStr(QAnt, "Comprobante", "ComprobanteAnt")
      QAnt = ReplaceStr(QAnt, "Cuentas", "CuentasAnt")
   End If
   
#End If

   Grid.Redraw = False

   Rc = CreateQry(DbMain, TmpQry, Q1)
   Rc = CreateQry(DbMain, TmpQryAnt, QAnt)
   
   Q1 = "SELECT " & TmpQry & ".IdQ, " & TmpQry & ".idCuenta, " & TmpQry & ".Codigo, " & TmpQry & ".Nivel, "
   Q1 = Q1 & TmpQry & ".Descripcion, " & TmpQry & ".Debe, " & TmpQry & ".Haber, " & TmpQry & ".Clasificacion, "
   Q1 = Q1 & TmpQryAnt & ".Debe as DebeAnt, " & TmpQryAnt & ".Haber as HaberAnt, " & TmpQryAnt & ".Codigo as CodigoAnt"
   Q1 = Q1 & " FROM " & TmpQry & " LEFT JOIN " & TmpQryAnt & " ON " & TmpQry & ".IdQ = " & TmpQryAnt & ".IdQ AND " & TmpQry & ".Codigo = " & TmpQryAnt & ".Codigo "
   
   Q1 = Q1 & " UNION "

   Q1 = Q1 & "SELECT " & TmpQryAnt & ".IdQ, " & TmpQryAnt & ".idCuenta, " & TmpQryAnt & ".Codigo, " & TmpQryAnt & ".Nivel, "
   Q1 = Q1 & TmpQryAnt & ".Descripcion, " & TmpQry & ".Debe, " & TmpQry & ".Haber, " & TmpQryAnt & ".Clasificacion, "
   Q1 = Q1 & TmpQryAnt & ".Debe as DebeAnt, " & TmpQryAnt & ".Haber as HaberAnt, " & TmpQry & ".Codigo as CodigoActual "
   Q1 = Q1 & " FROM " & TmpQryAnt & " LEFT JOIN " & TmpQry & " ON " & TmpQry & ".IdQ = " & TmpQryAnt & ".IdQ AND " & TmpQry & ".Codigo = " & TmpQryAnt & ".Codigo "
   Q1 = Q1 & " WHERE " & TmpQry & ".Codigo IS NULL "
   
   Q1 = Q1 & " ORDER BY " & TmpQry & ".Codigo, " & TmpQry & ".IdQ"
      
   Set Rs = OpenRs(DbMain, Q1)
   
#If DATACON = 1 Then       'Access
   If gEmprSeparadas Then
      Call UnLinkTable(DbMain, "ComprobanteAnt")
      Call UnLinkTable(DbMain, "MovComprobanteAnt")
   End If
#End If

   For j = 0 To MAX_NIVELES
      Total(j).Debe = 0
      Total(j).Haber = 0
      Total(j).Linea = 0
      
      TotalAnt(j).Debe = 0
      TotalAnt(j).Haber = 0
      TotalAnt(j).Linea = 0
   Next j
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows - 1
   
   CurNiv = 0
   CurCta = ""
   CodPadre = ""
   NomPadre = ""
   
   Do While Rs.EOF = False
   
      'Obtengo el Padre de la cuenta y pongo el total del padre al cambiar a otro
      
      If vFld(Rs("Nivel")) = 1 Then
      
         If CodPadre <> "" Then 'había uno antes
            
            'agregamos la cuenta de resultado del ejercicio
            If lBalClasif And ClasifPadre = CLASCTA_PASIVO Then
               Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
            End If
            
            i = i + 1
            Grid.rows = i + 1
            Call FGrSetRowStyle(Grid, i, "B")
            
            Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)
            Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
            Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
            Grid.TextMatrix(i, C_FMT) = "B"
            
            Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, NEGNUMFMT)
            Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, NEGNUMFMT)
            
            Grid.TextMatrix(i, C_DEBITOSANT) = Format(TotalAnt(1).Debe, NEGNUMFMT)
            Grid.TextMatrix(i, C_CREDITOSANT) = Format(TotalAnt(1).Haber, NEGNUMFMT)
               
            If ClasifPadre = CLASCTA_RESULTADO Then     'si es Resultado, pueden haber 2 cuentas de primer nivel de este tipo, por lo que el total no se pone a este nivel sino abajo en la grilla de totales
               If Ch_VerSubTot = 0 Then
                  Grid.RowHeight(i) = 0   'ocultamos el total
               End If
            End If
            
            If ClasifPadre = CLASCTA_ACTIVO Then
               TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
               TotClasifAnt(ClasifPadre) = TotalAnt(1).Debe - TotalAnt(1).Haber
            ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
               TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
               TotClasifAnt(ClasifPadre) = TotalAnt(1).Haber - TotalAnt(1).Debe
            End If
            
            LinTotClasif(ClasifPadre) = i
                        
            'Salto una línea
            i = i + 1
            Grid.rows = i + 1
            Grid.TextMatrix(i, C_IDCUENTA) = "*"
            
         End If
         
         CodPadre = vFld(Rs("Codigo"))
         NomPadre = FCase(vFld(Rs("Descripcion"), True))
         ClasifPadre = vFld(Rs("Clasificacion"))

      End If
            
      If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
         'asignamos los totales hacia arriba
         For j = CurNiv - 1 To vFld(Rs("Nivel")) Step -1
            If j > 1 Then
               Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
               Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
               
               Grid.TextMatrix(TotalAnt(j).Linea, C_DEBITOSANT) = Format(TotalAnt(j).Debe, BL_NUMFMT)
               Grid.TextMatrix(TotalAnt(j).Linea, C_CREDITOSANT) = Format(TotalAnt(j).Haber, BL_NUMFMT)
               
            End If
            
            Total(j).Debe = 0
            Total(j).Haber = 0
            Total(j).Linea = 0
            
            TotalAnt(j).Debe = 0
            TotalAnt(j).Haber = 0
            TotalAnt(j).Linea = 0
                        
         Next j
      End If
         

      If CurCta <> vFld(Rs("Codigo")) Then
      
         If CurCta <> "" And CurNiv > 1 Then
            'ponemos totales de cuenta actual
            Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
            
            Grid.TextMatrix(TotalAnt(CurNiv).Linea, C_DEBITOSANT) = Format(TotalAnt(CurNiv).Debe, BL_NUMFMT)
            Grid.TextMatrix(TotalAnt(CurNiv).Linea, C_CREDITOSANT) = Format(TotalAnt(CurNiv).Haber, BL_NUMFMT)
            
         End If
      
         If lBalClasif Then
            
            If LinPatrimonio = 0 Then
               If vFld(Rs("IdCuenta")) = gCtasBas.IdCtaPatrimonio Then
                  LinPatrimonio = i + 1 'se agrega después
               End If
            ElseIf LinResEjercicio(0) = 0 Then
               If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
                  Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
               End If
            End If
            
         End If
                    
         'actualizamos el nivel
         CurNiv = vFld(Rs("Nivel"))
         
         'agregamos la nueva cuenta
         i = i + 1
         Grid.rows = i + 1
         CurCta = vFld(Rs("Codigo"))
  
         Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("idCuenta"))
         Grid.TextMatrix(i, C_CLASCTA) = vFld(Rs("Clasificacion"))
         Grid.TextMatrix(i, C_NIVEL) = CurNiv
         
         
         If CurNiv <> 1 Then 'Es un hijo
            Call FGrSetRowStyle(Grid, i, "FC", gColores(CurNiv))
            Grid.TextMatrix(i, C_FMT) = "C" & gColores(CurNiv)
            Grid.TextMatrix(i, C_CODIGO) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 2), " ") & FCase(vFld(Rs("Descripcion")))
                        
         Else 'Es el Padre  Nivel 1
            Call FGrSetRowStyle(Grid, i, "B")
            Grid.TextMatrix(i, C_CUENTA) = UCase(vFld(Rs("Descripcion")))
            Grid.TextMatrix(i, C_FMT) = "B"
            
         End If
                  
         Total(CurNiv).Debe = 0
         Total(CurNiv).Haber = 0
         Total(CurNiv).Linea = i
                  
         TotalAnt(CurNiv).Debe = 0
         TotalAnt(CurNiv).Haber = 0
         TotalAnt(CurNiv).Linea = i
                  
      End If
   
      'sumamos los totales al nivel actual y a los niveles anteriores
      For j = CurNiv To 1 Step -1
         Total(j).Debe = Total(j).Debe + vFld(Rs("Debe"))
         Total(j).Haber = Total(j).Haber + vFld(Rs("Haber"))
         
         TotalAnt(j).Debe = TotalAnt(j).Debe + vFld(Rs("DebeAnt"))
         TotalAnt(j).Haber = TotalAnt(j).Haber + vFld(Rs("HaberAnt"))
      Next j
            
      Rs.MoveNext
      
   Loop
      
   'ponemos el total de la última línea
   If CurCta <> "" Then
      'ponemos totales de cuenta actual
      Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
      Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
      
      Grid.TextMatrix(TotalAnt(CurNiv).Linea, C_DEBITOSANT) = Format(TotalAnt(CurNiv).Debe, BL_NUMFMT)
      Grid.TextMatrix(TotalAnt(CurNiv).Linea, C_CREDITOSANT) = Format(TotalAnt(CurNiv).Haber, BL_NUMFMT)
            
      'asignamos los totales hacia arriba
      For j = CurNiv - 1 To 2 Step -1
         If j > 1 Then
            Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
            
            Grid.TextMatrix(TotalAnt(j).Linea, C_DEBITOSANT) = Format(TotalAnt(j).Debe, BL_NUMFMT)
            Grid.TextMatrix(TotalAnt(j).Linea, C_CREDITOSANT) = Format(TotalAnt(j).Haber, BL_NUMFMT)
         End If
      Next j
      
      If CodPadre <> "" Then 'había uno antes
         
         'agregamos la cuenta de resultado del ejercicio
         If lBalClasif And ClasifPadre = CLASCTA_PASIVO Then
            Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
         End If
         
         i = i + 1
         Grid.rows = i + 1
         Call FGrSetRowStyle(Grid, i, "B")
         
         Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)
         Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
         Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
         Grid.TextMatrix(i, C_FMT) = "B"
         
         Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, BL_NUMFMT)
         Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, BL_NUMFMT)
         
         Grid.TextMatrix(i, C_DEBITOSANT) = Format(TotalAnt(1).Debe, BL_NUMFMT)
         Grid.TextMatrix(i, C_CREDITOSANT) = Format(TotalAnt(1).Haber, BL_NUMFMT)
         
         If ClasifPadre = CLASCTA_RESULTADO Then   'si es Resultado, pueden haber 2 cuentas de primer nivel de este tipo, por lo que el total no se pone a este nivel sino abajo en la grilla de totales
            If Ch_VerSubTot = 0 Then
               Grid.RowHeight(i) = 0   'ocultamos el total
            End If
         End If
         
         If ClasifPadre = CLASCTA_ACTIVO Then
            TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
            TotClasifAnt(ClasifPadre) = TotalAnt(1).Debe - TotalAnt(1).Haber
         ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
            TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
            TotClasifAnt(ClasifPadre) = TotalAnt(1).Haber - TotalAnt(1).Debe
         End If
         
         LinTotClasif(ClasifPadre) = i
                  
      End If
      
   End If
   
   Call CloseRs(Rs)
      
   TotalFinal = 0
   
   'calculamos la columna Valor como la diferencia de Créditos y Débitos y ocultamos filas con valor 0
   For row = Grid.FixedRows To Grid.rows - 1
   
      If Trim(Grid.TextMatrix(row, C_CODIGO)) <> "" And vFmt(Grid.TextMatrix(row, C_IDCUENTA)) <> TOT_CUENTA Then
         
         If Val(Grid.TextMatrix(row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(row, C_DEBITOS)) - vFmt(Grid.TextMatrix(row, C_CREDITOS))
            DiffAnt = vFmt(Grid.TextMatrix(row, C_DEBITOSANT)) - vFmt(Grid.TextMatrix(row, C_CREDITOSANT))
         Else
            Diff = vFmt(Grid.TextMatrix(row, C_CREDITOS)) - vFmt(Grid.TextMatrix(row, C_DEBITOS))
            DiffAnt = vFmt(Grid.TextMatrix(row, C_CREDITOSANT)) - vFmt(Grid.TextMatrix(row, C_DEBITOSANT))
         End If
         
         Grid.TextMatrix(row, C_VALOR) = Format(Diff, NEGNUMFMT)
         Grid.TextMatrix(row, C_VALORANT) = Format(DiffAnt, NEGNUMFMT)
                                    
         If Diff = 0 And DiffAnt = 0 Then
            Grid.RowHeight(row) = 0
         End If
         
         
      ElseIf vFmt(Grid.TextMatrix(row, C_IDCUENTA)) = TOT_CUENTA Then
         
         Call FGrSetRowStyle(Grid, row, "B")
         
         If Val(Grid.TextMatrix(row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(row, C_DEBITOS)) - vFmt(Grid.TextMatrix(row, C_CREDITOS))
            DiffAnt = vFmt(Grid.TextMatrix(row, C_DEBITOSANT)) - vFmt(Grid.TextMatrix(row, C_CREDITOSANT))
         Else
            Diff = vFmt(Grid.TextMatrix(row, C_CREDITOS)) - vFmt(Grid.TextMatrix(row, C_DEBITOS))
            DiffAnt = vFmt(Grid.TextMatrix(row, C_CREDITOSANT)) - vFmt(Grid.TextMatrix(row, C_DEBITOSANT))
         End If
         
         Grid.TextMatrix(row, C_VALOR) = Format(Diff, NEGNUMFMT)
         Grid.TextMatrix(row, C_VALORANT) = Format(DiffAnt, NEGNUMFMT)
         
         TotalFinalAnt = TotalFinalAnt + DiffAnt
         
      End If
      
   Next row
   
   'ponemos el resultado del ejercicio si correponde
   If lBalClasif Then
   
      ResEjercicio = TotClasif(CLASCTA_ACTIVO) - TotClasif(CLASCTA_PASIVO)
      ResEjercicioAnt = TotClasifAnt(CLASCTA_ACTIVO) - TotClasifAnt(CLASCTA_PASIVO)
      
      If ResEjercicio <> 0 Or ResEjercicioAnt <> 0 Then
      
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            
            Grid.TextMatrix(LinResEjercicio(k), C_VALOR) = Format(ResEjercicio, NEGNUMFMT)
            Grid.TextMatrix(LinResEjercicio(k), C_VALORANT) = Format(ResEjercicioAnt, NEGNUMFMT)
            
           'If (vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALOR)) <> 0 Or vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALORANT))) And Nivel >= Val(Grid.TextMatrix(LinResEjercicio(k), C_NIVEL)) Then
            '2817660
            If (vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALOR)) <> 0 Or vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALORANT)) <> 0) And Nivel >= Val(Grid.TextMatrix(LinResEjercicio(k), C_NIVEL)) Then
               Grid.RowHeight(LinResEjercicio(k)) = Grid.RowHeight(0)
            End If
            
         Next k
         
         If LinPatrimonio > 0 Then
            
            Grid.TextMatrix(LinPatrimonio, C_VALOR) = Format(vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) + ResEjercicio, NEGNUMFMT)
            Grid.TextMatrix(LinPatrimonio, C_VALORANT) = Format(vFmt(Grid.TextMatrix(LinPatrimonio, C_VALORANT)) + ResEjercicioAnt, NEGNUMFMT)
            
            If (vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) <> 0 Or vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) <> 0) And Nivel >= Val(Grid.TextMatrix(LinPatrimonio, C_NIVEL)) Then
               Grid.RowHeight(LinPatrimonio) = Grid.RowHeight(0)
            End If
            
         End If
         
         If LinTotClasif(CLASCTA_PASIVO) > 0 Then
            Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR) = Format(vFmt(Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR)) + ResEjercicio, NEGNUMFMT)
            Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALORANT) = Format(vFmt(Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALORANT)) + ResEjercicioAnt, NEGNUMFMT)
         End If
      
         TotalFinal = TotalFinal + ResEjercicio
         TotalFinalAnt = TotalFinalAnt + ResEjercicioAnt
         
      Else
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            Grid.RowHeight(LinResEjercicio(k)) = 0
         Next k
         
      End If
   End If

   'calculamos la diferencia entre los dos periodos
   For row = Grid.FixedRows To Grid.rows - 1
      If Trim(Grid.TextMatrix(row, C_VALOR)) <> "" Or Trim(Grid.TextMatrix(row, C_VALORANT)) <> "" Then
         Grid.TextMatrix(row, C_DIF) = Format(vFmt(Grid.TextMatrix(row, C_VALOR)) - vFmt(Grid.TextMatrix(row, C_VALORANT)), NEGNUMFMT)
      End If
   Next row

   'aquí no se pone el nombre de la cuenta de resultado que gana, como en el Estado de Resultado Clasificado o Mensual porque un periodo puede ganar una (ej. ganancia) y el otro periodo ganar otra (ej. pérdida)
   GridTot.TextMatrix(0, C_CUENTA) = "TOTAL"
   GridTot.TextMatrix(0, C_VALOR) = Format(TotalFinal, NEGNUMFMT)
   GridTot.TextMatrix(0, C_VALORANT) = Format(TotalFinalAnt, NEGNUMFMT)
      
   Call FGrVRows(Grid, 1)
   
   'Grid.Rows = Grid.Rows + 25
      
   Grid.TopRow = Grid.FixedRows
   Grid.row = Grid.FixedRows
   Grid.Col = C_CODIGO
   Grid.RowSel = Grid.row
   Grid.ColSel = Grid.Col
   
   Grid.Redraw = True
   Call EnableFrm(False)
  
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 500
   If lBalClasif Then
      Grid.Height = Grid.Height + 300
   End If
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 200
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid, 1)

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

Private Sub Grid_DblClick()
   Call Bt_VerLibMayor_Click
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

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
Public Function FViewBalClasif(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lClasCta = CLASCTA_ACTIVO & "," & CLASCTA_PASIVO
   lCaption = "Balance Clasificado Comparativo"
   lLibOf = LIBOF_CLASIFICADO
   
   lBalClasif = True
   
   Me.Show vbModeless
End Function
Public Function FViewEstResultClasif(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lClasCta = CLASCTA_RESULTADO
   lCaption = "Estado de Resultado Clasificado Comparativo"
   lLibOf = LIBOF_ESTRESCLASIF
   
   Me.Show vbModeless
End Function

Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub

Private Sub AddResEjercicio(i As Integer, LinPatrimonio As Integer, LinResEjercicio() As Integer)
   Dim Nivel As Integer
   Dim NivRes As Integer, NivAux As Integer
   Dim NivPat As Integer
   Dim SpIndent As String
   Dim j As Integer, k As Integer
   Dim IdPadre As Long, IdPadreAux As Long
   Dim DescRes As String
   Dim CodRes As String
      
   If LinPatrimonio = 0 Then
      MsgBox1 "No se encontró la cuenta de Patrimonio, de acuerdo a la ""Definición de Cuentas Básicas"" de la ""Configuración Inicial"" de la empresa.", vbExclamation
      Exit Sub
   End If
   
   If LinResEjercicio(0) > 0 Then   'ya se agregó el resultado
      Exit Sub
   End If
   
   Nivel = Val(Cb_Nivel)
   NivPat = Val(Grid.TextMatrix(LinPatrimonio, C_NIVEL))
   
   If lResEje(gLastNivel).Codigo = "" Then
      MsgBox1 "No se encontró la cuenta de Resultado Ejercicio, de acuerdo a la ""Definición de Cuentas Básicas"" de la ""Configuración Inicial"" de la empresa.", vbExclamation
      Exit Sub
   End If
     
   k = 0
   For j = NivPat + 1 To Nivel
      
      i = i + 1
      Grid.rows = i + 1
      
      SpIndent = String(REP_INDENT * (j - 2), " ")
      Call FGrSetRowStyle(Grid, i, "FC", gColores(j))
      Grid.TextMatrix(i, C_FMT) = "C" & gColores(j)
      Grid.TextMatrix(i, C_CODIGO) = FmtCodCuenta(lResEje(j).Codigo)
      Grid.TextMatrix(i, C_CUENTA) = SpIndent & lResEje(j).Descripcion
      Grid.TextMatrix(i, C_NIVEL) = lResEje(j).Nivel
      Grid.TextMatrix(i, C_IDCUENTA) = "*"
      
      LinResEjercicio(k) = i
      k = k + 1
   Next j
   
End Sub
Private Sub ReadResEje()
   Dim i As Integer
   
   i = gLastNivel
   lResEje(gLastNivel).id = gCtasBas.IdCtaResEje
   
   Do While lResEje(i).id > 0
      lResEje(i).Codigo = GetDatosCuenta(lResEje(i).id, lResEje(i).Descripcion, lResEje(i).IdPadre, lResEje(i).Nivel)
      i = i - 1
      lResEje(i).id = lResEje(i + 1).IdPadre
   Loop

End Sub
Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub

