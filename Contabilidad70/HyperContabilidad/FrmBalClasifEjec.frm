VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalClasifEjec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BalanceClasificado Ejecutivo"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "FrmBalClasifEjec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6255
      Left            =   60
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   25
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   60
      TabIndex        =   12
      Top             =   660
      Width           =   12075
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1275
      End
      Begin VB.CheckBox Ch_SaldosVig 
         Caption         =   "Saldos Vigentes"
         Height          =   195
         Left            =   4440
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   3315
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   3315
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2100
         Picture         =   "FrmBalClasifEjec.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   4020
         Picture         =   "FrmBalClasifEjec.frx":0316
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
         Left            =   3000
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   10740
         Picture         =   "FrmBalClasifEjec.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1155
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
         Left            =   6120
         TabIndex        =   28
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel cuentas"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   26
         Top             =   600
         Width           =   975
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
      Left            =   60
      TabIndex        =   23
      Top             =   -60
      Width           =   12075
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
         Picture         =   "FrmBalClasifEjec.frx":0A5E
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
         Picture         =   "FrmBalClasifEjec.frx":0EE1
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir a Libro Mayor para cuenta seleccionada"
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox Ch_VerCodCuenta 
         Caption         =   "Ver Código Cuenta"
         Height          =   195
         Left            =   5760
         TabIndex        =   21
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
         Picture         =   "FrmBalClasifEjec.frx":124F
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "FrmBalClasifEjec.frx":12F3
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "FrmBalClasifEjec.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "FrmBalClasifEjec.frx":19F2
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "FrmBalClasifEjec.frx":1E1B
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "FrmBalClasifEjec.frx":2260
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10800
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
         Left            =   1080
         Picture         =   "FrmBalClasifEjec.frx":2707
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridV 
      Height          =   6255
      Left            =   120
      TabIndex        =   29
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   25
      Cols            =   9
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmBalClasifEjec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_CUENTA = 1
Const C_VALOR = 2
Const C_NIVEL = 3
Const C_IDCUENTA = 4
Const C_CLASCTA = 5
Const C_DEBITOS = 6
Const C_CREDITOS = 7
Const C_INTER = 8
Const C_CODIGO_P = 9
Const C_CUENTA_P = 10
Const C_VALOR_P = 11
Const C_NIVEL_P = 12
Const C_IDCUENTA_P = 13
Const C_CLASCTA_P = 14
Const C_DEBITOS_P = 15
Const C_CREDITOS_P = 16
Const C_FMT = 17

Const NCOLS = C_FMT

'Const C_INI_MES = C_FMT + 1
'Const C_SALDOFIN = C_INI_MES + 12
'Const C_INI_COMP = C_SALDOFIN + 1

Const TOT_CUENTA = -1

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
   
   'Call FGr2Clip(GridV, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   '2861570
   Call LP_FGr2Clip_Membr(GridV, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   '2861570
End Sub

Private Sub Bt_Email_Click()
Dim Frm As FrmEmailAccount

  Set Frm = Nothing
  Set Frm = New FrmEmailAccount
  
  Dim vAjunto As String
  vAjunto = Export_SendEmail(Grid, Nothing, Nothing, Nothing, lCaption & "_" & Tx_Desde & "_" & Tx_Hasta, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta, C_CODIGO)
   
 If Frm.FEdit(vAjunto) Then
 Frm.Show
 End If
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
    Dim PrtOrient As Integer
   Dim Pag As Integer
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   If lMensual And Not lComparativo Then
      If MsgBox1("Es muy probable que al imprimir este informe no quepan todas las columnas." & vbLf & vbLf & "Se sugiere copiarlo a Excel e imprimirlo desde ahí.", vbInformation + vbOKCancel) = vbCancel Then
         Exit Sub
      End If
   End If
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape
   
   
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
   Printer.Orientation = PrtOrient
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
   
   If lMensual And Not lComparativo Then
      If MsgBox1("Es muy probable que al imprimir este informe no quepan todas las columnas." & vbLf & vbLf & "Se sugiere copiarlo a Excel e imprimirlo desde ahí.", vbInformation + vbOKCancel) = vbCancel Then
         Exit Sub
      End If
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
      Printer.Orientation = lOrientacion
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
   GridV.Cols = NCOLS + 1
   Grid.FixedRows = 0
   GridV.FixedRows = 0
   
   lWCodCta = 1200
   lWVal = G_DVALWIDTH + 200
   
   If Ch_VerCodCuenta <> 0 Then
      GridV.ColWidth(C_CODIGO) = lWCodCta
      GridV.ColWidth(C_CODIGO_P) = lWCodCta
   Else
      GridV.ColWidth(C_CODIGO) = 0
      GridV.ColWidth(C_CODIGO_P) = 0
   End If
   
   GridV.ColWidth(C_CUENTA) = 5500
   GridV.ColWidth(C_VALOR) = lWVal
   GridV.ColWidth(C_NIVEL) = 0
   GridV.ColWidth(C_CREDITOS) = 0
   GridV.ColWidth(C_DEBITOS) = 0
   GridV.ColWidth(C_IDCUENTA) = 0
   GridV.ColWidth(C_CLASCTA) = 0
   
   GridV.ColWidth(C_INTER) = 300
   
   GridV.ColWidth(C_CUENTA_P) = 5500
   GridV.ColWidth(C_VALOR_P) = lWVal
   GridV.ColWidth(C_NIVEL_P) = 0
   GridV.ColWidth(C_CREDITOS_P) = 0
   GridV.ColWidth(C_DEBITOS_P) = 0
   GridV.ColWidth(C_IDCUENTA_P) = 0
   GridV.ColWidth(C_CLASCTA_P) = 0
   
   GridV.ColWidth(C_FMT) = 0
   
'   GridV.TextMatrix(0, C_CODIGO) = "Cód. Cuenta"
'   GridV.TextMatrix(0, C_CUENTA) = "Cuenta"
'   GridV.TextMatrix(0, C_VALOR) = "Saldo"
'
'   GridV.TextMatrix(0, C_CODIGO_P) = "Cód. Cuenta"
'   GridV.TextMatrix(0, C_CUENTA_P) = "Cuenta"
'   GridV.TextMatrix(0, C_VALOR_P) = "Saldo"

   GridV.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   GridV.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   GridV.ColAlignment(C_VALOR) = flexAlignRightCenter
   GridV.ColAlignment(C_CODIGO_P) = flexAlignLeftCenter
   GridV.ColAlignment(C_CUENTA_P) = flexAlignLeftCenter
   GridV.ColAlignment(C_VALOR_P) = flexAlignRightCenter
   
   Call FGrSetup(GridV)
           
   Call FGrVRows(GridV, 1)
      
End Sub

Private Sub Bt_VerLibMayor_Click()
   Dim Frm As FrmLibMayor
   Dim IdCuenta As Long
   
   If GridV.Col >= C_CODIGO_P Then
      IdCuenta = vFmt(GridV.TextMatrix(GridV.Row, C_IDCUENTA_P))
   Else
      IdCuenta = vFmt(GridV.TextMatrix(GridV.Row, C_IDCUENTA))
   End If
   
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

Private Sub Ch_SaldosVig_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_VerCodCuenta_Click()
   Dim i As Integer
      
   If Ch_VerCodCuenta <> 0 Then
      GridV.ColWidth(C_CODIGO) = lWCodCta
      GridV.ColWidth(C_CODIGO_P) = lWCodCta
      
      GridV.ColWidth(C_CUENTA) = 5500
      GridV.ColWidth(C_VALOR) = lWVal
      GridV.ColWidth(C_CUENTA_P) = 5500
      GridV.ColWidth(C_VALOR_P) = lWVal
      
   Else
      GridV.ColWidth(C_CODIGO) = 0
      GridV.ColWidth(C_CODIGO_P) = 0
      
      GridV.ColWidth(C_CUENTA) = 5500 + 900
      GridV.ColWidth(C_VALOR) = lWVal + 400
      GridV.ColWidth(C_CUENTA_P) = 5500 + 900
      GridV.ColWidth(C_VALOR_P) = lWVal + 400
   End If
   
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

   lOrientacion = ORIENT_HOR
   
   Ch_VerCodCuenta = 1
   Ch_SaldosVig = 1
   
   Call ReadResEje
   Call LoadAll
   Call SetupPriv
   
   If lBalClasif = False Then
      Ch_LibOficial.visible = False
   End If
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi() As Integer
   Dim Total() As String
   Dim Titulos(2) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   
   Set gPrtLibros.Grid = GridV
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = "Balance General Clasificado"
   Titulos(1) = "Desde el " & Format(GetTxDate(Tx_Desde), "dd \d\e mmmm") & " hasta el " & Format(GetTxDate(Tx_Hasta), "dd \d\e mmmm \d\e yyyy")
   
   FontTit(0).FontBold = True
   
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
   
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
'   If GetTxDate(Tx_Desde) <> DateSerial(gEmpresa.Ano, 1, 1) Then
'      Encabezados(0) = Format(GetTxDate(Tx_Desde), DATEFMT) & " a "
'   Else
'      Encabezados(0) = "Al "
'   End If
'   Encabezados(0) = Encabezados(0) & Format(GetTxDate(Tx_Hasta), DATEFMT)
   
   i = 0
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
    
   ReDim ColWi(GridV.Cols - 1)
   ReDim Total(GridV.Cols - 1)
    
   For i = 0 To GridV.Cols - 1
      If GridV.ColWidth(i) > 0 Then
         If i = C_CUENTA Or i = C_CUENTA_P Then
            ColWi(i) = GridV.ColWidth(i) - 850
         Else
            ColWi(i) = GridV.ColWidth(i) - 100
         End If
      End If
   Next i
   
   ColWi(C_INTER) = 60
    
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_IDCUENTA
   gPrtLibros.NTotLines = 0
   gPrtLibros.FmtCol = C_FMT
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Nivel As Integer
   Dim Rs As Recordset
   Dim Total(MAX_NIVELES) As RepNiv_t
   Dim CurNiv As Integer
   Dim CurCta As String
   Dim i As Integer, j As Integer, k As Integer
   Dim Row As Integer
   Dim Diff As Double
   Dim FirstDiaMes As Long, LastDiaMes As Long
   Dim CodPadre As String, NomPadre As String
   Dim CodPadre2 As String, NomPadre2 As String
   Dim WhFecha As String
   Dim TotalFinal As Double
   Dim Wh As String
   Dim Mes As Integer
   Dim FDesde As Long
   Dim FHasta As Long
   Dim TotFinHasta As Double
   Dim TotFinMenos1 As Double
   Dim LinPatrimonio As Integer, LinTotalPatrimonio As Integer
   Dim LinResEjercicio(MAX_NIVELES) As Integer
   Dim ClasifPadre As Integer
   Dim ClasifPadre2 As Integer
   Dim TotClasif(MAX_CLASCTA) As Double
   Dim LinTotClasif(MAX_CLASCTA) As Integer
   Dim ResEjercicio As Double
   Dim RowVisible As Boolean
   Dim Col As Integer
   Dim Repetido As Boolean
   
   Nivel = Val(Cb_Nivel)
   
   GridV.Redraw = False
   
   FDesde = GetTxDate(Tx_Desde)
   FHasta = GetTxDate(Tx_Hasta)
   WhFecha = "Comprobante.Fecha BETWEEN " & FDesde & " AND " & FHasta
   
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


   Q1 = GenQueryPorNiveles(Nivel, WhFecha & Wh, Ch_LibOficial <> 0, lClasCta, lMensual)

   Set Rs = OpenRs(DbMain, Q1)
   
   For j = 0 To MAX_NIVELES
      Total(j).Debe = 0
      Total(j).Haber = 0
      Total(j).Linea = 0
   Next j
   
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows - 1
   
   CurNiv = 0
   CurCta = ""
   CodPadre = ""
   NomPadre = ""
   CodPadre2 = ""
   NomPadre2 = ""
   
   Do While Rs.EOF = False
   
      'Obtengo el Padre de la cuenta y pongo el total del padre al cambiar a otro
      If vFld(Rs("Codigo")) = 1110000 Then
       i = i
       Rs.MoveNext
      End If
      
      
      If vFld(Rs("Nivel")) = 1 Or vFld(Rs("Nivel")) = 2 Then
      
         If CodPadre2 <> "" Then 'había uno antes
            If Not Repetido Then
               i = i + 2
               Grid.rows = i + 2
            Else
               i = i - 1
            End If
            Call FGrSetRowStyle(Grid, i, "B")
            
            Grid.TextMatrix(i, C_CUENTA) = FCase("TOTAL " & UCase(NomPadre2))
            Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
            Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre2
            Grid.TextMatrix(i, C_FMT) = "B"
            
            If LinPatrimonio > 0 Then
               If UCase(NomPadre2) = UCase(Grid.TextMatrix(LinPatrimonio, C_CUENTA)) Then
                  'estamos poniendo TOTAL PATRIMONIO
                  LinTotalPatrimonio = i
               End If
            End If
            
            Grid.TextMatrix(i, C_DEBITOS) = Format(Total(2).Debe, NEGNUMFMT)
            Grid.TextMatrix(i, C_CREDITOS) = Format(Total(2).Haber, NEGNUMFMT)
                                              
            'Salto una línea
            i = i + 1
            Grid.rows = i + 1
            Grid.TextMatrix(i, C_IDCUENTA) = "*"
            
         End If
         
         If vFld(Rs("Nivel")) = 2 Then
         
            If CodPadre2 = vFld(Rs("Codigo")) Then
               Repetido = True
            Else
               Repetido = False
               CodPadre2 = vFld(Rs("Codigo"))
               NomPadre2 = FCase(vFld(Rs("Descripcion"), True))
               ClasifPadre2 = vFld(Rs("Clasificacion"))
            End If
         
         ElseIf vFld(Rs("Nivel")) = 1 Then
      
            If CodPadre <> "" Then 'había uno antes
                              
               i = i + 1
               Grid.rows = i + 1
               Call FGrSetRowStyle(Grid, i, "B")
               
               Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)
               Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
               Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
               Grid.TextMatrix(i, C_FMT) = "B"
               
               Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, NEGNUMFMT)
               Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, NEGNUMFMT)
               
               If TotClasif(ClasifPadre) <> 0 Then
                  MsgBox1 "ATENCIÓN:" & vbCrLf & vbCrLf & "Revise su plan de cuentas. Hay dos o más cuentas de primer nivel clasificadas como " & UCase(gClasCta(ClasifPadre)) & "." & vbCrLf & "Esto puede generar errores en el Balance.", vbInformation
               End If
               
               If ClasifPadre = CLASCTA_ACTIVO Then
                  TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
               ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                  TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
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
            CodPadre2 = ""
            NomPadre2 = ""
            ClasifPadre2 = 0
        End If
         
      End If
            
      If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
         'asignamos los totales hacia arriba
         For j = CurNiv - 1 To vFld(Rs("Nivel")) Step -1
            If j > 1 Then
               Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
               Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
               
            End If
            Total(j).Debe = 0
            Total(j).Haber = 0
            Total(j).Linea = 0
                        
         Next j
      End If
         
      'cambia la cuenta
      If CurCta <> vFld(Rs("Codigo")) Then
      
         If CurCta <> "" And CurNiv > 1 Then
            'ponemos totales de cuenta actual
            Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
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
         
         'agregamos la nueva cuenta, sólo si es nivel > 1
         If CurNiv > 1 Then 'Es un hijo
            
            i = i + 1
            Grid.rows = i + 1
            CurCta = vFld(Rs("Codigo"))
            
            Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("idCuenta"))
            Grid.TextMatrix(i, C_CLASCTA) = vFld(Rs("Clasificacion"))
            Grid.TextMatrix(i, C_NIVEL) = CurNiv
         End If
         
         'nivel 2 destaca y los siguientes siguen igual
         If CurNiv > 2 Then
            Call FGrSetRowStyle(Grid, i, "FC", gColores(CurNiv))
            Grid.TextMatrix(i, C_FMT) = "C" & gColores(CurNiv)
            Grid.TextMatrix(i, C_CODIGO) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 2), " ") & FCase(vFld(Rs("Descripcion")))
'            If Trim(FCase(vFld(Rs("Descripcion")))) = "Capital Pagado" Then
'               MsgBeep vbExclamation
'            End If
                        
         ElseIf CurNiv = 2 Then
            Call FGrSetRowStyle(Grid, i, "B")
            Grid.TextMatrix(i, C_CUENTA) = FCase(vFld(Rs("Descripcion")))
            Grid.TextMatrix(i, C_FMT) = "B"
            
         End If
                  
         Total(CurNiv).Debe = 0
         Total(CurNiv).Haber = 0
         Total(CurNiv).Linea = i
                  
      End If
   
      'sumamos los totales al nivel actual y a los niveles anteriores
      For j = CurNiv To 1 Step -1
         Total(j).Debe = Total(j).Debe + vFld(Rs("Debe"))
         Total(j).Haber = Total(j).Haber + vFld(Rs("Haber"))
      Next j
                     
      Rs.MoveNext
      
   Loop
      
   'ponemos el total de la última línea
   If CurCta <> "" Then
      'ponemos totales de cuenta actual
      Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
      Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
      
      'asignamos los totales hacia arriba
      For j = CurNiv - 1 To 2 Step -1
         If j > 1 Then
            Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
         End If
      Next j
      
      If CodPadre2 <> "" Then 'había uno antes
                  
          'agregamos la cuenta de resultado del ejercicio
         If lBalClasif And ClasifPadre = CLASCTA_PASIVO Then
            Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)    'OJO
         End If
                 
         If LinTotalPatrimonio = 0 Then
                 
            i = i + 1
            Grid.rows = i + 1
            Grid.TextMatrix(i, C_IDCUENTA) = "*"
            i = i + 1
            Grid.rows = i + 1
            Call FGrSetRowStyle(Grid, i, "B")
            
            Grid.TextMatrix(i, C_CUENTA) = FCase("TOTAL " & UCase(NomPadre2))     'Aquí se pone TOTAL PATRIMONIO
            Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
            Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre2
            Grid.TextMatrix(i, C_FMT) = "B"
            
            Grid.TextMatrix(i, C_DEBITOS) = Format(Total(2).Debe, BL_NUMFMT)
            Grid.TextMatrix(i, C_CREDITOS) = Format(Total(2).Haber, BL_NUMFMT)
            
            LinTotalPatrimonio = i
            
         Else
            Grid.TextMatrix(LinTotalPatrimonio, C_DEBITOS) = Format(Total(2).Debe, BL_NUMFMT)
            Grid.TextMatrix(LinTotalPatrimonio, C_CREDITOS) = Format(Total(2).Haber, BL_NUMFMT)
          
         End If
         
         'Salto una línea
         i = i + 1
         Grid.rows = i + 1
         Grid.TextMatrix(i, C_IDCUENTA) = "*"
         
      End If
      
      If CodPadre <> "" Then 'había uno antes
         
         
         i = i + 1
         Grid.rows = i + 1
         Call FGrSetRowStyle(Grid, i, "B")
         
         Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)               'Aquí se pone TOTAL PASIVO
         Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
         Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
         Grid.TextMatrix(i, C_FMT) = "B"
         
         Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, BL_NUMFMT)
         Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, BL_NUMFMT)
         
         If TotClasif(ClasifPadre) <> 0 Then
            MsgBox1 "ATENCIÓN:" & vbCrLf & vbCrLf & "Revise su plan de cuentas. Hay dos o más cuentas de primer nivel clasificadas como " & UCase(gClasCta(ClasifPadre)) & "." & vbCrLf & "Esto puede generar errores en el Balance.", vbExclamation
         End If
         
         If ClasifPadre = CLASCTA_ACTIVO Then
            TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
         ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
            TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
         End If
         
         LinTotClasif(ClasifPadre) = i
                  
      End If
      
   End If
   
   Call CloseRs(Rs)
      
   TotalFinal = 0
   
   'calculamos la columna Valor como la diferencia de Créditos y Débitos y ocultamos filas con valor 0
   For Row = Grid.FixedRows To Grid.rows - 1
   
      If Trim(Grid.TextMatrix(Row, C_CODIGO)) <> "" And vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) <> TOT_CUENTA Then
         
         If Val(Grid.TextMatrix(Row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(Row, C_DEBITOS)) - vFmt(Grid.TextMatrix(Row, C_CREDITOS))
         Else
            Diff = vFmt(Grid.TextMatrix(Row, C_CREDITOS)) - vFmt(Grid.TextMatrix(Row, C_DEBITOS))
         End If
         
         Grid.TextMatrix(Row, C_VALOR) = Format(Diff, NEGNUMFMT)
                  
         If Diff = 0 Then
            Grid.RowHeight(Row) = 0
         End If
                  
      ElseIf vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) = TOT_CUENTA Then
         
         Call FGrSetRowStyle(Grid, Row, "B")
         
         If Val(Grid.TextMatrix(Row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(Row, C_DEBITOS)) - vFmt(Grid.TextMatrix(Row, C_CREDITOS))
         Else
            Diff = vFmt(Grid.TextMatrix(Row, C_CREDITOS)) - vFmt(Grid.TextMatrix(Row, C_DEBITOS))
         End If
         
         Grid.TextMatrix(Row, C_VALOR) = Format(Diff, NEGNUMFMT)
         TotalFinal = TotalFinal + Diff
         
      End If
      
   Next Row
   
   'ponemos el resultado del ejercicio si correponde
   If lBalClasif Then

      ResEjercicio = TotClasif(CLASCTA_ACTIVO) - TotClasif(CLASCTA_PASIVO)

      If ResEjercicio <> 0 Then
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            Grid.TextMatrix(LinResEjercicio(k), C_VALOR) = Format(ResEjercicio, NEGNUMFMT)
            If vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALOR)) <> 0 And Nivel >= Val(Grid.TextMatrix(LinResEjercicio(k), C_NIVEL)) Then
               Grid.RowHeight(LinResEjercicio(k)) = Grid.RowHeight(k)
            End If
         Next k
'         If LinPatrimonio > 0 Then          'FCA 13 jun 2014 no puede estar en este reporte pero si en el clasificado normal
'            Grid.TextMatrix(LinPatrimonio, C_VALOR) = Format(vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) + ResEjercicio, NEGNUMFMT)
'            If vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) <> 0 And Nivel >= Val(Grid.TextMatrix(LinPatrimonio, C_NIVEL)) Then
'               Grid.RowHeight(LinPatrimonio) = Grid.RowHeight(0)
'            End If
'         End If


         If LinTotalPatrimonio > 0 Then
            Grid.TextMatrix(LinTotalPatrimonio, C_VALOR) = Format(vFmt(Grid.TextMatrix(LinTotalPatrimonio, C_VALOR)) + ResEjercicio, NEGNUMFMT)
            If vFmt(Grid.TextMatrix(LinTotalPatrimonio, C_VALOR)) <> 0 And Nivel >= Val(Grid.TextMatrix(LinTotalPatrimonio, C_NIVEL)) Then
               Grid.RowHeight(LinTotalPatrimonio) = Grid.RowHeight(0)
            End If
         End If

         If LinTotClasif(CLASCTA_PASIVO) > 0 Then
            Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR) = Format(vFmt(Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR)) + ResEjercicio, NEGNUMFMT)
         End If

         TotalFinal = TotalFinal + ResEjercicio
      Else
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            Grid.RowHeight(LinResEjercicio(k)) = 0
         Next k
      End If
   End If

   Call FGrVRows(Grid, 1)
   
   'borramos los títulos de las columnas con ancho 0
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) = 0 Then
         Grid.TextMatrix(0, i) = ""
      End If
   Next i
   
   'Grid.Rows = Grid.Rows + 25
      
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.Col = C_CODIGO
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col

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
Private Sub SetParallelGrid_SinValCero()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim InitPasivo As Integer
   Dim r As Integer
   Dim RowActCirculante As Integer, NActCirculante As Integer
   Dim RowPasCirculante As Integer, NPasCirculante As Integer
   Dim RowEndCirculante As Integer, NEndCirculante As Integer
   Dim RowActTotal As Integer
   Dim RowPasTotal As Integer
   Dim RowEndTotal As Integer
   Dim RowTotalFinal As Integer
   Dim RowTotalPasivo As Integer
   Dim RowTotalActivo As Integer
   
   For i = 0 To Grid.rows - 1
      
      'marcamos inicio de pasivos
      If Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_PASIVO Then
         If InitPasivo = 0 Then
            InitPasivo = i
         End If
      End If
      
      'contamos los Activos Circulantes > 0
      If RowActCirculante = 0 Then
         If Grid.RowHeight(i) > 0 Then
            NActCirculante = NActCirculante + 1
         End If
      End If
      
      'contamos los Pasivos Circulantes > 0
      If InitPasivo > 0 And RowPasCirculante = 0 Then
         If Grid.RowHeight(i) > 0 Then
            NPasCirculante = NPasCirculante + 1
         End If
      End If
      
      'marcamos linea "Total Activo Circulante"
      If InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "total") <> 0 And (InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "circulante") <> 0 Or InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "corriente") <> 0) Then
         If Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_ACTIVO Then
            If RowActCirculante = 0 Then
               RowActCirculante = i
            End If
         'marcamos linea "Total Pasivo Circulante"
         ElseIf Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_PASIVO Then
            If RowPasCirculante = 0 Then
               RowPasCirculante = i
            End If
         End If
   
      'marcamos línea de "Total Activos"
      ElseIf LCase(Grid.TextMatrix(i, C_CUENTA)) = "total activos" Then
         If RowActTotal = 0 Then
            RowActTotal = i
         End If
      'marcamos línea de "Total Pasivos"
      ElseIf LCase(Grid.TextMatrix(i, C_CUENTA)) = "total pasivos" Then
         If RowPasTotal = 0 Then
            RowPasTotal = i
         End If
      End If
   
   Next i
   
   'máximo de líneas de circulante
   If RowActCirculante > RowPasCirculante - InitPasivo Then
      RowEndCirculante = RowActCirculante
   Else
      RowEndCirculante = RowPasCirculante - InitPasivo
   End If
   
   If NActCirculante > NPasCirculante Then
      NEndCirculante = NActCirculante
   Else
      NEndCirculante = NPasCirculante
   End If
   
   NEndCirculante = NEndCirculante - 1
   
   'máximo de líneas de total (activo o pasivo)
   If RowActTotal - RowActCirculante > RowPasTotal - RowPasCirculante Then
      RowEndTotal = RowEndCirculante + RowActTotal - RowActCirculante
   Else
      RowEndTotal = RowEndCirculante + RowPasTotal - RowPasCirculante
   End If


            
   GridV.rows = 0
   k = 0
   
   'copiamos los activos circulantes
   For i = 0 To RowActCirculante - 1
        
      If Grid.RowHeight(i) > 0 Then
         GridV.rows = GridV.rows + 1
         
         For j = 0 To C_CREDITOS
            GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
         Next j
         
         If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
            GridV.TextMatrix(k, C_IDCUENTA) = "*"
         End If
                  
         Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_CREDITOS)
         GridV.TextMatrix(k, C_FMT) = "FCELL"
         
         k = k + 1
      End If
         
   Next i
   
   GridV.rows = GridV.rows + 1
      
   For i = k To NEndCirculante
      GridV.rows = GridV.rows + 1
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   
   'copiamos el total activo circulante
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(NEndCirculante, j) = Grid.TextMatrix(RowActCirculante, j)
   Next j
   
   'ahora el resto de los activos
   
   k = NEndCirculante + 1
   For i = RowActCirculante + 1 To RowActTotal - 1
      
      If Grid.RowHeight(i) > 0 Then
      
         GridV.rows = GridV.rows + 1
         
         For j = 0 To C_CREDITOS
            GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
         Next j
         
         If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
            GridV.TextMatrix(k, C_IDCUENTA) = "*"
         End If
         
         Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_CREDITOS)
         GridV.TextMatrix(k, C_FMT) = "FCELL"
         
         k = k + 1
      End If
      
   Next i
   
   'rellenamos para llegar a la línea de total activos
'   For i = RowActTotal To RowEndTotal
'      GridV.Rows = GridV.Rows + 1
'      k = k + 1
'   Next i
   
   RowTotalActivo = k
           
   'y luego los pasivos circulantes
   r = 0
   For i = InitPasivo To RowPasCirculante - 1
      
      If Grid.RowHeight(i) > 0 Then
         If r >= GridV.rows Then
            GridV.rows = GridV.rows + 1
         End If
         
         For j = 0 To C_CREDITOS
            GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
         Next j
         
         Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_CREDITOS)
         GridV.TextMatrix(r, C_FMT) = "FCELL"
         
         r = r + 1
      End If
      
   Next i
   
   'saltamos hasta llegar al total circulante

   'copiamos el total pasivo circulante
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(NEndCirculante, j + C_CODIGO_P) = Grid.TextMatrix(RowPasCirculante, j)
   Next j
   
   GridV.TextMatrix(NEndCirculante, C_FMT) = "FCELL"

   'ahora el resto de los pasivos hasta antes del total
   
   r = NEndCirculante + 1
   For i = RowPasCirculante + 1 To RowPasTotal - 1
      
      If Grid.RowHeight(i) > 0 Then
      
         If r >= GridV.rows Then
            GridV.rows = GridV.rows + 1
         End If
         
         For j = 0 To C_CREDITOS
            GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
         Next j
         
         Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_CREDITOS)
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
   
   'copiamos el TOTAL ACTIVO
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowTotalFinal, j) = Grid.TextMatrix(RowActTotal, j)
   Next j
   
   'copiamos el total pasivo final
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowTotalFinal, j + C_CODIGO_P) = Grid.TextMatrix(RowPasTotal, j)
   Next j

   GridV.TextMatrix(RowTotalFinal, C_FMT) = "FCELL"
   
   Call FGrVRows(GridV, 1)
      
   'ponemos en bold las líneas que corresponde
   For i = 0 To GridV.rows - 1
      If GridV.TextMatrix(i, C_CODIGO) = "" And GridV.TextMatrix(i, C_CUENTA) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_CUENTA, C_VALOR)
         GridV.TextMatrix(i, C_CUENTA) = UCase(GridV.TextMatrix(i, C_CUENTA))
      End If
      If GridV.TextMatrix(i, C_CODIGO_P) = "" And GridV.TextMatrix(i, C_CUENTA_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_CUENTA_P, C_VALOR_P)
         GridV.TextMatrix(i, C_CUENTA_P) = UCase(GridV.TextMatrix(i, C_CUENTA_P))
      End If
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = GridV.TextMatrix(i, C_IDCUENTA_P)
      End If
   Next i

End Sub
Private Sub SetParallelGrid_ConValCero()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim InitPasivo As Integer
   Dim r As Integer
   Dim RowActCirculante As Integer, NActCirculante As Integer
   Dim RowPasCirculante As Integer, NPasCirculante As Integer
   Dim RowEndCirculante As Integer
   Dim RowActTotal As Integer
   Dim RowPasTotal As Integer
   Dim RowEndTotal As Integer
   
   For i = 0 To Grid.rows - 1
      
      'marcamos inicio de pasivos
      If Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_PASIVO Then
         If InitPasivo = 0 Then
            InitPasivo = i
         End If
      End If
            
      'marcamos linea "Total Activo Circulante"
      If InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "total") <> 0 And (InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "circulante") <> 0 Or InStr(LCase(Grid.TextMatrix(i, C_CUENTA)), "corriente") <> 0) Then
         If Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_ACTIVO Then
            If RowActCirculante = 0 Then
               RowActCirculante = i
            End If
         'marcamos linea "Total Pasivo Circulante"
         ElseIf Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_PASIVO Then
            If RowPasCirculante = 0 Then
               RowPasCirculante = i
            End If
         End If
   
      'marcamos línea de "Total Activos"
      ElseIf LCase(Grid.TextMatrix(i, C_CUENTA)) = "total activos" Then
         If RowActTotal = 0 Then
            RowActTotal = i
         End If
      'marcamos línea de "Total Pasivos"
      ElseIf LCase(Grid.TextMatrix(i, C_CUENTA)) = "total pasivos" Then
         If RowPasTotal = 0 Then
            RowPasTotal = i
         End If
      End If
   
   Next i
   
   'máximo de líneas de circulante
   If RowActCirculante > RowPasCirculante - InitPasivo Then
      RowEndCirculante = RowActCirculante
   Else
      RowEndCirculante = RowPasCirculante - InitPasivo
   End If
   
   'máximo de líneas de total (activo o pasivo)
   If RowActTotal - RowActCirculante > RowPasTotal - RowPasCirculante Then
      RowEndTotal = RowEndCirculante + RowActTotal - RowActCirculante
   Else
      RowEndTotal = RowEndCirculante + RowPasTotal - RowPasCirculante
   End If


            
   GridV.rows = 0
   
   'copiamos los activos circulantes
   For i = 0 To RowActCirculante - 1
      
      GridV.rows = GridV.rows + 1
      
      For j = 0 To C_CREDITOS
         GridV.TextMatrix(i, j) = Grid.TextMatrix(i, j)

      Next j
      
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
      
      Call FGrSetRowStyle(GridV, i, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_CREDITOS)
      GridV.TextMatrix(i, C_FMT) = "FCELL"
         
   Next i
      
   'rellenamos para llegar al total circulante
   For i = RowActCirculante To RowEndCirculante
      GridV.rows = GridV.rows + 1
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = "*"
      End If
   Next i
   
   'copiamos el total activo circulante
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowEndCirculante, j) = Grid.TextMatrix(RowActCirculante, j)
   Next j
   
   'ahora el resto de los activos
   
   k = RowEndCirculante + 1
   For i = RowActCirculante + 1 To RowActTotal - 1
      
      GridV.rows = GridV.rows + 1
      
      For j = 0 To C_CREDITOS
         GridV.TextMatrix(k, j) = Grid.TextMatrix(i, j)
      Next j
      
      If GridV.TextMatrix(k, C_IDCUENTA) = "" Then
         GridV.TextMatrix(k, C_IDCUENTA) = "*"
      End If
      
      Call FGrSetRowStyle(GridV, k, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), 0, C_CREDITOS)
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
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowEndTotal, j) = Grid.TextMatrix(RowActTotal, j)
   Next j
      
   
   'y luego los pasivos circulantes
   r = 0
   For i = InitPasivo To RowPasCirculante - 1
      
      If r >= GridV.rows Then
         GridV.rows = GridV.rows + 1
      End If
      
      For j = 0 To C_CREDITOS
         GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
      Next j
      
      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_CREDITOS)
      GridV.TextMatrix(r, C_FMT) = "FCELL"
      
      r = r + 1
   Next i
   
   'saltamos hasta llegar al total circulante

   'copiamos el total pasivo circulante
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowEndCirculante, j + C_CODIGO_P) = Grid.TextMatrix(RowPasCirculante, j)
   Next j
   
   GridV.TextMatrix(RowEndCirculante, C_FMT) = "FCELL"

   'ahora el resto de los pasivos hasta antes del total
   
   r = RowEndCirculante + 1
   For i = RowPasCirculante + 1 To RowPasTotal - 1
      
      If r >= GridV.rows Then
         GridV.rows = GridV.rows + 1
      End If
      
      For j = 0 To C_CREDITOS
         GridV.TextMatrix(r, j + C_CODIGO_P) = Grid.TextMatrix(i, j)
      Next j
      
      Call FGrSetRowStyle(GridV, r, "FC", gColores(Val(Grid.TextMatrix(i, C_NIVEL))), C_CODIGO_P, C_CODIGO_P + C_CREDITOS)
      GridV.TextMatrix(r, C_FMT) = "FCELL"
     
      r = r + 1
      
   Next i
   
   'saltamos hasta llegar a la línea de total pasivo
   
   'copiamos el total pasivo final
   For j = 0 To C_CREDITOS
      GridV.TextMatrix(RowEndTotal, j + C_CODIGO_P) = Grid.TextMatrix(RowPasTotal, j)
   Next j

   GridV.TextMatrix(RowEndTotal, C_FMT) = "FCELL"
   
   Call FGrVRows(GridV, 1)
   
   'ponemos en bold las líneas que corresponde y llenamos C_IDCUENTA si corresponde
   For i = 0 To GridV.rows - 1
      If GridV.TextMatrix(i, C_CODIGO) = "" And GridV.TextMatrix(i, C_CUENTA) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_CUENTA, C_VALOR)
         GridV.TextMatrix(i, C_CUENTA) = UCase(GridV.TextMatrix(i, C_CUENTA))
      End If
      If GridV.TextMatrix(i, C_CODIGO_P) = "" And GridV.TextMatrix(i, C_CUENTA_P) <> "" Then
         Call FGrSetRowStyle(GridV, i, "B", 0, C_CUENTA_P, C_VALOR_P)
         GridV.TextMatrix(i, C_CUENTA_P) = UCase(GridV.TextMatrix(i, C_CUENTA_P))
      End If
      
      If GridV.TextMatrix(i, C_IDCUENTA) = "" Then
         GridV.TextMatrix(i, C_IDCUENTA) = GridV.TextMatrix(i, C_IDCUENTA_P)
      End If
   Next i


End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   GridV.Height = Me.Height - GridV.Top - 500
   GridV.Width = Me.Width - 330
   
   Call FGrVRows(GridV, 1)
   
   GridV.rows = GridV.rows + 2

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

Private Sub GridV_DblClick()
   Call Bt_VerLibMayor_Click

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
   lCaption = "Balance Clasificado Ejecutivo"
   lLibOf = LIBOF_CLASIFICADO
   
   lBalClasif = True
   
   Me.Show vbModeless
End Function

Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub

Private Sub AddResEjercicio(i As Integer, ByVal LinPatrimonio As Integer, LinResEjercicio() As Integer)
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
'      Call FGrSetRowStyle(Grid, i, "FC", gColores(j))
'      Grid.TextMatrix(i, C_FMT) = "C" & gColores(j)
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

Private Sub CopyCellStyle(GridFrom As Control, GridTo As Control, ByVal RowFrom As Integer, ByVal ColFrom As Integer, Optional ByVal RowTo As Integer = -1, Optional ByVal ColTo As Integer = -1)
   Dim GrFrom As Control, GrTo As Control
   Dim oRowFrom As Integer, oColFrom As Integer
   Dim oRowTo As Integer, oColTo As Integer
   Dim CurrFillStyle As Integer
   
   If TypeName(GridFrom) = "MSFlexGrid" Then
      Set GrFrom = GridFrom
   Else
      Set GrFrom = GridFrom.FlxGrid
   End If

   If TypeName(GridTo) = "MSFlexGrid" Then
      Set GrTo = GridTo
   Else
      Set GrTo = GridTo.FlxGrid
   End If

   oRowFrom = GrFrom.Row
   oColFrom = GrFrom.Col

   oRowTo = GrTo.Row
   oColTo = GrTo.Col

   GrFrom.Row = RowFrom
   GrFrom.RowSel = RowFrom
   GrFrom.Col = ColFrom
   GrFrom.ColSel = ColFrom
   
   If RowTo < 0 Then
      RowTo = RowFrom
   End If
   
   If ColTo < 0 Then
      ColTo = ColFrom
   End If
   
   GrTo.Row = RowTo
   GrTo.RowSel = RowTo
   GrTo.Col = ColTo
   GrTo.ColSel = ColTo
   
   CurrFillStyle = GrTo.FillStyle
   
   GrTo.FillStyle = flexFillRepeat

   
   GrTo.CellFontBold = GrFrom.CellFontBold
   GrTo.CellFontItalic = GrFrom.CellFontItalic
   GrTo.CellFontUnderline = GrFrom.CellFontUnderline
   GrTo.CellForeColor = GrFrom.CellForeColor
   GrTo.CellBackColor = GrFrom.CellBackColor
   GrTo.CellTextStyle = GrFrom.CellTextStyle
   GrTo.CellAlignment = GrFrom.CellAlignment
   
   GrFrom.Row = oRowFrom
   GrFrom.RowSel = oRowFrom
   
   GrFrom.Col = oColFrom
   GrFrom.ColSel = oColFrom

   GrFrom.Row = oRowFrom
   GrFrom.RowSel = oRowFrom
   
   GrTo.Col = oColTo
   GrTo.ColSel = oColTo
   
   GrTo.FillStyle = CurrFillStyle

End Sub
