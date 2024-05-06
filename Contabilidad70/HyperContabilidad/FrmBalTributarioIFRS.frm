VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalTributarioIFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance General 8 Columnas Formato IFRS"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   Icon            =   "FrmBalTributarioIFRS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   60
      TabIndex        =   19
      Top             =   600
      Width           =   11115
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   3315
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   3315
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3840
         Picture         =   "FrmBalTributarioIFRS.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1740
         Picture         =   "FrmBalTributarioIFRS.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   9840
         Picture         =   "FrmBalTributarioIFRS.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   2
         Left            =   4740
         TabIndex        =   27
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   4740
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2220
         TabIndex        =   24
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel cuentas:"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   20
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   11115
      Begin VB.CheckBox Ch_VerCodCta 
         Caption         =   "Ver Código Cuenta"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1695
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
         Picture         =   "FrmBalTributarioIFRS.frx":0A5E
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   2460
         Picture         =   "FrmBalTributarioIFRS.frx":0B02
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   2040
         Picture         =   "FrmBalTributarioIFRS.frx":0E63
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2880
         Picture         =   "FrmBalTributarioIFRS.frx":1201
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   960
         Picture         =   "FrmBalTributarioIFRS.frx":162A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   120
         Picture         =   "FrmBalTributarioIFRS.frx":1A6F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9840
         TabIndex        =   17
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
         Left            =   540
         Picture         =   "FrmBalTributarioIFRS.frx":1F16
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   60
      TabIndex        =   8
      Top             =   1680
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   4
      Cols            =   13
      FixedRows       =   2
      FixedCols       =   2
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7080
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7380
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7680
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmBalTributarioIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_CUENTA = 1
Const C_DEBITOS = 2
Const C_CREDITOS = 3
Const C_DEUDOR = 4
Const C_ACREEDOR = 5
Const C_INVACTIVO = 6
Const C_INVPASIVO = 7
Const C_PERDIDA = 8
Const C_GANANCIA = 9
Const C_IDCUENTA = 10
Const C_CLASIF = 11
Const C_NIVEL = 12

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lMes As Integer

Dim lWCodCta As Integer

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
   Dim Clip As String
   Dim Membrete As String
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   'Call FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   Clip = LP_FGr2String(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta, False, C_CODIGO)
   
   If Clip <> "" Then
    '2861570 tema 1
    If MsgBox1("¿Desea agregar datos básicos de la empresa (rut, nombre, dirección giro, rep. Legal)?.", vbInformation + vbYesNo) = vbYes Then
      Membrete = "Razón Social " & vbTab & gEmpresa.RazonSocial & vbCrLf
      Membrete = Membrete & " Rut " & vbTab & gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut) & vbCrLf
      Membrete = Membrete & " Dirección " & vbTab & gEmpresa.Direccion & ", " & IIf(gEmpresa.Ciudad <> "", FCase(gEmpresa.Ciudad), FCase(gEmpresa.Comuna)) & vbCrLf
      Membrete = Membrete & " Giro " & vbTab & gEmpresa.Giro & vbCrLf
      Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal1 & vbCrLf
      If gEmpresa.RutRepLegal1 <> "" Then
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & "-" & DV_Rut(gEmpresa.RutRepLegal1) & vbCrLf
      Else
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & vbCrLf
      End If
      
     If gEmpresa.RepConjunta Then
        Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal2 & vbCrLf
        If gEmpresa.RutRepLegal2 <> "" Then
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & "-" & DV_Rut(gEmpresa.RutRepLegal2) & vbCrLf & vbCrLf
        Else
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & vbCrLf & vbCrLf
        End If
      End If

      Clip = Membrete & Clip
      
      End If
      'fin 2861570 tema 1
      Clip = Clip & FGr2String(GridTot(0))
      Clip = Clip & FGr2String(GridTot(1))
      Clip = Clip & FGr2String(GridTot(2))
      
      Clipboard.Clear
      Clipboard.SetText Clip
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
   
   gPrtLibros.TotFntBold = True
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
      
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      gPrtLibros.CallEndDoc = False
      
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
      Call PrtPieBalance(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
      
      gPrtLibros.CallEndDoc = True
      
      Me.MousePointer = vbDefault
            
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, Pag)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
      gPrtLibros.TotFntBold = True

   End If
   
   Call ResetPrtBas(gPrtLibros)
  
End Sub

Private Sub Cb_Nivel_Click()
   Call EnableFrm(True)
End Sub

Private Sub Ch_VerCodCta_Click()
   Dim i As Integer, j As Integer

   If Ch_VerCodCta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
      For i = C_DEBITOS To Grid.Cols - 1
         Grid.ColWidth(i) = IIf(Grid.ColWidth(i) > 0, G_VALWIDTH, 0)
      Next i
   Else
      Grid.ColWidth(C_CODIGO) = 0
      For i = C_DEBITOS To Grid.Cols - 1
         Grid.ColWidth(i) = IIf(Grid.ColWidth(i) > 0, G_DVALWIDTH, 0)
      Next i
   End If
   
   For i = 0 To 2
      For j = 0 To Grid.Cols - 1
         GridTot(i).ColWidth(j) = Grid.ColWidth(j)
      Next j
   Next i
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   Dim MesActual As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim CurPlan As String
   Dim Frm As FrmMsgConBreak
   Dim Msg As String
   
   Me.Caption = gInformeIFRS(IFRS_BAL8COL)
   
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
   
   Call FillNivel(Cb_Nivel, 2)
   
   Call FillCbAreaNeg(Cb_AreaNeg, False)
   Call FillCbCCosto(Cb_CCosto, False)
   
   lOrientacion = ORIENT_HOR
   
   Ch_VerCodCta = 1
   
   Call SetUpGrid
   Call LoadAll
      
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
   
      MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Existen cuentas con saldo distinto de cero, que no tienen su correspondiente clasificación IFRS.", vbExclamation + vbOKOnly
   
   End If

   MsgBox1 "Este informe sólo considera comprobantes en estado APROBADO.", vbInformation + vbOKOnly

End Sub

Private Sub Form_Resize()
   Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - 2 * (Grid.Left + W.xFrame)
   If d > 1000 Then
      Grid.Width = d + 30
   End If
 
   d = Me.Height - Grid.Top - W.YCaption * 2 - GridTot(0).Height * 3 + 100
   If d > 1000 Then
      Grid.Height = d
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   GridTot(0).Top = Grid.Top + Grid.Height
   GridTot(0).Width = Grid.Width - 300
   
   GridTot(1).Top = GridTot(0).Top + GridTot(0).Height
   GridTot(1).Width = Grid.Width - 300
   
   GridTot(2).Top = GridTot(1).Top + GridTot(1).Height
   GridTot(2).Width = Grid.Width - 300
   
   Call FGrVRows(Grid)
End Sub
Private Sub SetUpGrid()
   Dim Col As Integer
   Dim i As Integer
   
   lWCodCta = 1200
   
   If Ch_VerCodCta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
   Else
      Grid.ColWidth(C_CODIGO) = 0
   End If
   
   Grid.ColWidth(C_CUENTA) = 3000
   Grid.ColWidth(C_DEBITOS) = G_VALWIDTH
   Grid.ColWidth(C_CREDITOS) = G_VALWIDTH
   Grid.ColWidth(C_DEUDOR) = G_VALWIDTH
   Grid.ColWidth(C_ACREEDOR) = G_VALWIDTH
   Grid.ColWidth(C_INVACTIVO) = G_VALWIDTH
   Grid.ColWidth(C_INVPASIVO) = G_VALWIDTH
   Grid.ColWidth(C_PERDIDA) = G_VALWIDTH
   Grid.ColWidth(C_GANANCIA) = G_VALWIDTH
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_CLASIF) = 0
   
   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBITOS) = flexAlignRightCenter
   Grid.ColAlignment(C_CREDITOS) = flexAlignRightCenter
   Grid.ColAlignment(C_DEUDOR) = flexAlignRightCenter
   Grid.ColAlignment(C_ACREEDOR) = flexAlignRightCenter
   Grid.ColAlignment(C_INVACTIVO) = flexAlignRightCenter
   Grid.ColAlignment(C_INVPASIVO) = flexAlignRightCenter
   Grid.ColAlignment(C_PERDIDA) = flexAlignRightCenter
   Grid.ColAlignment(C_GANANCIA) = flexAlignRightCenter
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot(0))
   Call FGrTotales(Grid, GridTot(1))
   Call FGrTotales(Grid, GridTot(2))
   
   Grid.TextMatrix(0, C_DEUDOR) = "Saldo"
   Grid.TextMatrix(0, C_ACREEDOR) = "Saldo"
   Grid.TextMatrix(0, C_INVACTIVO) = "Inventario"
   Grid.TextMatrix(0, C_INVPASIVO) = "Inventario"
   Grid.TextMatrix(0, C_PERDIDA) = "Resultado"
   Grid.TextMatrix(0, C_GANANCIA) = "Resultado"
   
   Grid.TextMatrix(1, C_CODIGO) = "Código"
   Grid.TextMatrix(1, C_CUENTA) = "Cuentas"
   Grid.TextMatrix(1, C_DEBITOS) = "Debitos"
   Grid.TextMatrix(1, C_CREDITOS) = "Créditos"
   Grid.TextMatrix(1, C_DEUDOR) = "Deudor"
   Grid.TextMatrix(1, C_ACREEDOR) = "Acreedor"
   Grid.TextMatrix(1, C_INVACTIVO) = "Activo"
   Grid.TextMatrix(1, C_INVPASIVO) = "Pasivo"
   Grid.TextMatrix(1, C_PERDIDA) = "Pérdida"
   Grid.TextMatrix(1, C_GANANCIA) = "Ganancia"
      
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Nivel As Integer
   Dim Rs As Recordset
   Dim Total(MAX_NIVELES) As RepNiv_t
   Dim CurNiv As Integer
   Dim CurCta As String
   Dim i As Integer, j As Integer
   Dim row As Integer
   Dim Diff As Double
   Dim SumTotal(C_GANANCIA) As Double
   Dim FirstDiaMes As Long, LastDiaMes As Long, F1 As Long
   Dim UbiCol As Integer
   Dim WhFecha As String
   Dim Wh As String
   
   Grid.Redraw = False
   
   Nivel = Val(Cb_Nivel)
   
   WhFecha = "(Comprobante.Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
   
   If ItemData(Cb_AreaNeg) > 0 Then
      Wh = Wh & " AND MovComprobante.IdAreaNeg = " & ItemData(Cb_AreaNeg)
   End If
   
   If ItemData(Cb_CCosto) > 0 Then
      Wh = Wh & " AND MovComprobante.IdCCosto = " & ItemData(Cb_CCosto)
   End If
   
   Wh = Wh & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"

   Q1 = GenQueryIFRSporNiveles(Nivel, WhFecha & Wh, True, 0)

   Set Rs = OpenRs(DbMain, Q1)
   
   For j = 0 To MAX_NIVELES
      Total(j).Debe = 0
      Total(j).Haber = 0
      Total(j).Linea = 0
   Next j
   
   i = Grid.FixedRows - 1
   Grid.rows = Grid.FixedRows
   
   CurNiv = 0
   CurCta = ""
   
   Do While Rs.EOF = False
   
      If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
         For j = CurNiv - 1 To vFld(Rs("Nivel")) Step -1
            
            If Total(j).Linea > 0 Then
            
               Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
               Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
               
               Total(j).Debe = 0
               Total(j).Haber = 0
               Total(j).Linea = 0
               
            End If
            
         Next j
      End If
   
      If CurCta <> vFld(Rs("Codigo")) Then
      
         If CurCta <> "" Then
            'ponemos totales de cuenta actual
            Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
                        
         End If
      
         'actualizamos el nivel
         CurNiv = vFld(Rs("Nivel"))
         
         'agregamos la nueva cuenta
         i = i + 1
         Grid.rows = i + 1
         CurCta = vFld(Rs("Codigo"))
  
         'Call FGrSetRowStyle(Grid, i, "FC", gColores(CurNiv))
         
         Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("idCuenta"))
         Grid.TextMatrix(i, C_NIVEL) = CurNiv
         Grid.TextMatrix(i, C_CODIGO) = Format(vFld(Rs("Codigo")), gFmtCodigoIFRS)
         'Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 1), " ") & vFld(Rs("Descripcion"))
         Grid.TextMatrix(i, C_CUENTA) = FCase(vFld(Rs("Descripcion"), True))
         'Grid.TextMatrix(i, C_CLASIF) = vFld(Rs("Clasificacion"))
         Grid.TextMatrix(i, C_CLASIF) = Left(vFld(Rs("Codigo")), 1)
         
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
      For j = CurNiv - 1 To 1 Step -1
         If Total(j).Linea > 0 Then
            Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
         End If
      Next j
      
   End If
   
   Call CloseRs(Rs)
   
   For row = Grid.FixedRows To Grid.rows - 1
      If Trim(Grid.TextMatrix(row, C_CODIGO)) = "" Then
         Exit For
      End If
      
      Diff = vFmt(Grid.TextMatrix(row, C_DEBITOS)) - vFmt(Grid.TextMatrix(row, C_CREDITOS))
      If Diff > 0 Then
         Grid.TextMatrix(row, C_DEUDOR) = Format(Diff, BL_NUMFMT)
      Else
         Grid.TextMatrix(row, C_ACREEDOR) = Format(Abs(Diff), BL_NUMFMT)
      End If
      
      Select Case Val(Grid.TextMatrix(row, C_CLASIF))
         
         Case CLASCTA_ACTIVO, CLASCTA_PASIVO, CLASCTA_ORDEN
         
            If vFmt(Grid.TextMatrix(row, C_DEUDOR)) > 0 Then
               Grid.TextMatrix(row, C_INVACTIVO) = Grid.TextMatrix(row, C_DEUDOR)
            Else
               Grid.TextMatrix(row, C_INVPASIVO) = Grid.TextMatrix(row, C_ACREEDOR)
            End If
            
         Case CLASCTA_RESULTADO
      
            If vFmt(Grid.TextMatrix(row, C_DEUDOR)) > 0 Then
               Grid.TextMatrix(row, C_PERDIDA) = Grid.TextMatrix(row, C_DEUDOR)
            Else
               Grid.TextMatrix(row, C_GANANCIA) = Grid.TextMatrix(row, C_ACREEDOR)
            End If
     
      End Select
      
      'sólo mostramos cuentas del nivel seleccionado y con movimiento
      If Val(Grid.TextMatrix(row, C_NIVEL)) <> Val(Cb_Nivel) Or (vFmt(Grid.TextMatrix(row, C_DEBITOS)) = 0 And vFmt(Grid.TextMatrix(row, C_CREDITOS)) = 0) Then
         Grid.RowHeight(row) = 0
      End If
               
      'If Val(Grid.TextMatrix(Row, C_NIVEL)) = 1 Then   'sumamos totales finales
      If Grid.RowHeight(row) <> 0 Then
         'Suma de Totales
         SumTotal(C_DEBITOS) = SumTotal(C_DEBITOS) + vFmt(Grid.TextMatrix(row, C_DEBITOS))
         SumTotal(C_CREDITOS) = SumTotal(C_CREDITOS) + vFmt(Grid.TextMatrix(row, C_CREDITOS))
         SumTotal(C_DEUDOR) = SumTotal(C_DEUDOR) + vFmt(Grid.TextMatrix(row, C_DEUDOR))
         SumTotal(C_ACREEDOR) = SumTotal(C_ACREEDOR) + vFmt(Grid.TextMatrix(row, C_ACREEDOR))
         SumTotal(C_INVACTIVO) = SumTotal(C_INVACTIVO) + vFmt(Grid.TextMatrix(row, C_INVACTIVO))
         SumTotal(C_INVPASIVO) = SumTotal(C_INVPASIVO) + vFmt(Grid.TextMatrix(row, C_INVPASIVO))
         SumTotal(C_PERDIDA) = SumTotal(C_PERDIDA) + vFmt(Grid.TextMatrix(row, C_PERDIDA))
         SumTotal(C_GANANCIA) = SumTotal(C_GANANCIA) + vFmt(Grid.TextMatrix(row, C_GANANCIA))
      
      End If
      
   Next row
   
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   Grid.row = Grid.FixedRows
   Grid.Col = C_CODIGO
   Grid.RowSel = Grid.row
   Grid.ColSel = Grid.Col
   
   Grid.Redraw = True
   
   'Pongo totales finales
   
   For i = C_DEBITOS To C_GANANCIA
      GridTot(0).TextMatrix(0, i) = ""
      GridTot(1).TextMatrix(0, i) = ""
      GridTot(2).TextMatrix(0, i) = ""
   Next i

   
   'SUBTOTAL
   GridTot(0).TextMatrix(0, C_CUENTA) = "Sub Total"
   For i = C_DEBITOS To C_GANANCIA
      GridTot(0).TextMatrix(0, i) = Format(SumTotal(i), NUMFMT)
      GridTot(1).TextMatrix(0, i) = ""
      GridTot(2).TextMatrix(0, i) = ""
   Next i
   
   'UTILIDAD O PERDIDA
   
   'debitos-creditos
   Diff = vFmt(GridTot(0).TextMatrix(0, C_DEBITOS)) - vFmt(GridTot(0).TextMatrix(0, C_CREDITOS))
   If Diff < 0 Then
      GridTot(1).TextMatrix(0, C_DEBITOS) = Format(Abs(Diff), NUMFMT)
   Else
      GridTot(1).TextMatrix(0, C_CREDITOS) = Format(Abs(Diff), NUMFMT)
   End If
   
   'deudor-acreedor
   Diff = vFmt(GridTot(0).TextMatrix(0, C_DEUDOR)) - vFmt(GridTot(0).TextMatrix(0, C_ACREEDOR))
   If Diff < 0 Then
      GridTot(1).TextMatrix(0, C_DEUDOR) = Format(Abs(Diff), NUMFMT)
   Else
      GridTot(1).TextMatrix(0, C_ACREEDOR) = Format(Abs(Diff), NUMFMT)
   End If
   
   'InvActivo-InvPasivo
   Diff = vFmt(GridTot(0).TextMatrix(0, C_INVACTIVO)) - vFmt(GridTot(0).TextMatrix(0, C_INVPASIVO))
   If Diff < 0 Then
      GridTot(1).TextMatrix(0, C_INVACTIVO) = Format(Abs(Diff), NUMFMT)
   Else
      GridTot(1).TextMatrix(0, C_INVPASIVO) = Format(Abs(Diff), NUMFMT)
   End If
   
   'Pérdida-Ganancia
   Diff = vFmt(GridTot(0).TextMatrix(0, C_PERDIDA)) - vFmt(GridTot(0).TextMatrix(0, C_GANANCIA))
   If Diff < 0 Then
      GridTot(1).TextMatrix(0, C_PERDIDA) = Format(Abs(Diff), NUMFMT)
      GridTot(1).TextMatrix(0, C_CUENTA) = "Utilidad"
   Else
      GridTot(1).TextMatrix(0, C_GANANCIA) = Format(Abs(Diff), NUMFMT)
      If Diff = 0 Then
         GridTot(1).TextMatrix(0, C_CUENTA) = "Utilidad"
      Else
         GridTot(1).TextMatrix(0, C_CUENTA) = "Pérdida"
      End If
   End If
   
   'TOTALES FINALES
   GridTot(2).TextMatrix(0, C_CUENTA) = "TOTALES"
   For i = C_DEBITOS To C_GANANCIA
      GridTot(2).TextMatrix(0, i) = Format(vFmt(GridTot(0).TextMatrix(0, i)) + vFmt(GridTot(1).TextMatrix(0, i)), NUMFMT)
   Next i
   
   Call EnableFrm(False)
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(C_NIVEL) As Integer
   Dim Total(C_NIVEL * 4) As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   Dim AddColWi As Integer
   Dim FontGrid(1) As FontDef_t
   Dim PorcAjuste As Single
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = Me.Caption & " " & gEmpresa.Ano
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
   gPrtLibros.Encabezados = Encabezados
   FontEnc(0).FontBold = True
   FontEnc(0).FontName = "Arial"
   FontEnc(0).FontSize = 10
   Call gPrtLibros.FntEncabezados(FontEnc())
    
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
   
   'PS
   If Printer.Orientation = vbPRORPortrait Then
      'VERTICAL
      gPrtLibros.GrFontSize = 7
      gPrtLibros.GrFontName = "Arial"
      gPrtLibros.TotFntBold = False
      PorcAjuste = 0.75
      If Ch_VerCodCta <> 0 Then
         AddColWi = 100 'OK
      Else
         AddColWi = 60  'OK
      End If
   Else
      'HORIZONTAL
      PorcAjuste = 1
      If Ch_VerCodCta <> 0 Then
         AddColWi = 100 + 50
      Else
         AddColWi = 60 + 50
      End If
   End If
   '***
   

   For i = 0 To Grid.Cols - 1
   
      If Grid.ColWidth(i) > 0 Then
         ColWi(i) = Grid.ColWidth(i) * PorcAjuste
         If i >= C_DEBITOS Then
            ColWi(i) = ColWi(i) + AddColWi
         End If
      End If
      
   Next i
   
   ColWi(C_CUENTA) = ColWi(C_CUENTA) - 100
   
   j = 0
   For i = 0 To Grid.Cols - 1
      Total(j) = GridTot(0).TextMatrix(0, i)
      j = j + 1
   Next i
   
   For i = 0 To Grid.Cols - 1
      Total(j) = GridTot(1).TextMatrix(0, i)
      j = j + 1
   Next i
   
    For i = 0 To Grid.Cols - 1
      Total(j) = GridTot(2).TextMatrix(0, i)
      j = j + 1
   Next i
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_CODIGO
   gPrtLibros.NTotLines = 3
   
   'gPrtLibros.TotFntBold = False
   
End Sub

Private Sub Grid_DblClick()
   Dim IdCuenta As Long
   Dim Frm As FrmLibMayor

   If gPlanCuentas = "IFRS" Then
      
      IdCuenta = vFmt(Grid.TextMatrix(Grid.row, C_IDCUENTA))
      
      If IdCuenta > 0 Then
      
         Set Frm = New FrmLibMayor
         Call Frm.FViewChain(GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta), IdCuenta, TAJUSTE_FINANCIERO)
         Set Frm = Nothing
      
      End If
   End If
      
End Sub
Private Sub Bt_VerLibMayor_Click()
   Dim Frm As FrmLibMayor
   Dim IdCuenta As Long
   
   IdCuenta = vFmt(Grid.TextMatrix(Grid.row, C_IDCUENTA))
   
   If IdCuenta > 0 Then
   
      Set Frm = New FrmLibMayor
      Call Frm.FViewChain(GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta), IdCuenta, TAJUSTE_AMBOS)
      Set Frm = Nothing
   
   End If

End Sub

Private Sub Grid_Scroll()
   GridTot(0).LeftCol = Grid.LeftCol
   GridTot(1).LeftCol = Grid.LeftCol
   GridTot(2).LeftCol = Grid.LeftCol
   
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
Public Function FView(Optional ByVal Mes As Integer = 0, Optional ByVal ShowModal As Boolean = False)
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
         
   If ShowModal Then
      Me.Show vbModal
   Else
      Me.Show vbModeless
   End If
End Function

Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub
