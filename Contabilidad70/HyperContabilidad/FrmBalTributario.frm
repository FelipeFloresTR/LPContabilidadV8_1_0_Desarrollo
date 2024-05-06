VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBalTributario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance General 8 Columnas"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FrmBalTributario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   60
      TabIndex        =   22
      Top             =   600
      Width           =   11835
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1275
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   3315
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   3315
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   4620
         TabIndex        =   4
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   4140
         Picture         =   "FrmBalTributario.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2100
         Picture         =   "FrmBalTributario.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   10560
         Picture         =   "FrmBalTributario.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   2
         Left            =   5940
         TabIndex        =   30
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   5940
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2460
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel cuentas:"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   23
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   21
      Top             =   0
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
         Left            =   3840
         Picture         =   "FrmBalTributario.frx":0A5E
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "FrmBalTributario.frx":0EE1
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ir a Libro Mayor para cuenta seleccionada"
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox Ch_VerCodCta 
         Caption         =   "Ver Código Cuenta"
         Height          =   255
         Left            =   5880
         TabIndex        =   19
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
         Left            =   2040
         Picture         =   "FrmBalTributario.frx":124F
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "FrmBalTributario.frx":12F3
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "FrmBalTributario.frx":1654
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "FrmBalTributario.frx":19F2
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "FrmBalTributario.frx":1E1B
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   600
         Picture         =   "FrmBalTributario.frx":2260
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10560
         TabIndex        =   20
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
         Left            =   1020
         Picture         =   "FrmBalTributario.frx":2707
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   60
      TabIndex        =   10
      Top             =   1680
      Width           =   11835
      _ExtentX        =   20876
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
      Left            =   0
      TabIndex        =   24
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
      TabIndex        =   25
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
      TabIndex        =   26
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
Attribute VB_Name = "FrmBalTributario"
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

'3054289 se cambio long por double
Public Patrimonio As Double
'3054289

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
      'Clip = Clip & Membrete
      
      Clipboard.Clear
      Clipboard.SetText Clip
   End If
   
End Sub

Private Sub Bt_Email_Click()
Dim Frm As FrmEmailAccount

  Set Frm = Nothing
  Set Frm = New FrmEmailAccount
   
 Dim vAjunto As String
  vAjunto = Export_SendEmail(Grid, GridTot(0), GridTot(1), GridTot(2), "BalTributario_" & Tx_Desde & "_" & Tx_Hasta, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta, C_CODIGO)
   
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
   Call PrtPieBalanceFirma(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, 0) ' 0 = PREVIEW
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
   
   If Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(LIBOF_TRIBUTARIO, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(LIBOF_TRIBUTARIO) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
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
      Call PrtPieBalanceFirma(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight, 1) ' 1 = BT_ PRINT
      '2861570
      gPrtLibros.CallEndDoc = True
      
      Me.MousePointer = vbDefault
      
      If lPapelFoliado And Ch_LibOficial.visible = True And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(LIBOF_TRIBUTARIO, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, Pag)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
      gPrtLibros.TotFntBold = True

   End If
   
   Call ResetPrtBas(gPrtLibros)
  
End Sub

Private Sub Bt_VerLibMayor_Click()
   Dim Frm As FrmLibMayor
   Dim IdCuenta As Long
   
   IdCuenta = vFmt(Grid.TextMatrix(Grid.Row, C_IDCUENTA))
   
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
   
   If CbItemData(Cb_TipoAjuste) = TAJUSTE_TRIBUTARIO Then
      If Not gEmpresa.TieneAnoAnt And Not CompAperturaTribTieneMovs Then
         MsgBox1 "ATENCIÓN: Recuerde llenar el comprobante de Apertura Tributario, ingresando los movimientos correspondientes en la ventana de edición de este comprobante.", vbInformation
      End If
   End If

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
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)

   
   lOrientacion = ORIENT_HOR
   
   Ch_VerCodCta = 1
   
   Call SetUpGrid
   Call LoadAll
   
   Call SetupPriv
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
   Dim Row As Integer
   Dim Diff As Double
   Dim SumTotal(C_GANANCIA) As Double
   Dim FirstDiaMes As Long, LastDiaMes As Long, F1 As Long
   Dim UbiCol As Integer
   Dim WhFecha As String
   Dim Wh As String
   
   Patrimonio = 0
   Grid.Redraw = False
   
   Nivel = Val(Cb_Nivel)
   
   WhFecha = "(Comprobante.Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
   
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
   
   Q1 = GenQueryPorNiveles(Nivel, WhFecha & Wh, Ch_LibOficial <> 0)

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
         Grid.TextMatrix(i, C_CODIGO) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
         'Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 1), " ") & vFld(Rs("Descripcion"))
         Grid.TextMatrix(i, C_CUENTA) = FCase(vFld(Rs("Descripcion"), True))
         Grid.TextMatrix(i, C_CLASIF) = vFld(Rs("Clasificacion"))
         
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
   
   For Row = Grid.FixedRows To Grid.rows - 1
      If Trim(Grid.TextMatrix(Row, C_CODIGO)) = "" Then
         Exit For
      End If
      
      Diff = vFmt(Grid.TextMatrix(Row, C_DEBITOS)) - vFmt(Grid.TextMatrix(Row, C_CREDITOS))
      If Diff > 0 Then
         Grid.TextMatrix(Row, C_DEUDOR) = Format(Diff, BL_NUMFMT)
      Else
         Grid.TextMatrix(Row, C_ACREEDOR) = Format(Abs(Diff), BL_NUMFMT)
      End If
      
      If Replace(Trim(Grid.TextMatrix(Row, C_CODIGO)), "-", "") = "2030000" Or Replace(Trim(Grid.TextMatrix(Row, C_CODIGO)), "-", "") = "2300000" Then
         Patrimonio = IIf(Grid.TextMatrix(Row, C_ACREEDOR) = "", "0", Grid.TextMatrix(Row, C_ACREEDOR))
      End If

      
      Select Case Val(Grid.TextMatrix(Row, C_CLASIF))
         
         Case CLASCTA_ACTIVO, CLASCTA_PASIVO, CLASCTA_ORDEN
         
            If vFmt(Grid.TextMatrix(Row, C_DEUDOR)) > 0 Then
               Grid.TextMatrix(Row, C_INVACTIVO) = Grid.TextMatrix(Row, C_DEUDOR)
            Else
               Grid.TextMatrix(Row, C_INVPASIVO) = Grid.TextMatrix(Row, C_ACREEDOR)
            End If
            
         Case CLASCTA_RESULTADO
      
            If vFmt(Grid.TextMatrix(Row, C_DEUDOR)) > 0 Then
               Grid.TextMatrix(Row, C_PERDIDA) = Grid.TextMatrix(Row, C_DEUDOR)
            Else
               Grid.TextMatrix(Row, C_GANANCIA) = Grid.TextMatrix(Row, C_ACREEDOR)
            End If
     
      End Select
      
      'sólo mostramos cuentas del nivel seleccionado y con movimiento
      If Val(Grid.TextMatrix(Row, C_NIVEL)) <> Val(Cb_Nivel) Or (vFmt(Grid.TextMatrix(Row, C_DEBITOS)) = 0 And vFmt(Grid.TextMatrix(Row, C_CREDITOS)) = 0) Then
         Grid.RowHeight(Row) = 0
      End If
               
      'If Val(Grid.TextMatrix(Row, C_NIVEL)) = 1 Then   'sumamos totales finales
      If Grid.RowHeight(Row) <> 0 Then
         'Suma de Totales
         SumTotal(C_DEBITOS) = SumTotal(C_DEBITOS) + vFmt(Grid.TextMatrix(Row, C_DEBITOS))
         SumTotal(C_CREDITOS) = SumTotal(C_CREDITOS) + vFmt(Grid.TextMatrix(Row, C_CREDITOS))
         SumTotal(C_DEUDOR) = SumTotal(C_DEUDOR) + vFmt(Grid.TextMatrix(Row, C_DEUDOR))
         SumTotal(C_ACREEDOR) = SumTotal(C_ACREEDOR) + vFmt(Grid.TextMatrix(Row, C_ACREEDOR))
         SumTotal(C_INVACTIVO) = SumTotal(C_INVACTIVO) + vFmt(Grid.TextMatrix(Row, C_INVACTIVO))
         SumTotal(C_INVPASIVO) = SumTotal(C_INVPASIVO) + vFmt(Grid.TextMatrix(Row, C_INVPASIVO))
         SumTotal(C_PERDIDA) = SumTotal(C_PERDIDA) + vFmt(Grid.TextMatrix(Row, C_PERDIDA))
         SumTotal(C_GANANCIA) = SumTotal(C_GANANCIA) + vFmt(Grid.TextMatrix(Row, C_GANANCIA))
      
      End If
      
   Next Row
   
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.Col = C_CODIGO
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col
      
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
   
   Grid.Redraw = True
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
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
 
End Sub

Private Sub Grid_DblClick()
   Call Bt_VerLibMayor_Click
   
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
Public Function FView(ByVal Mes As Integer, Optional ByVal ShowModal As Boolean = False)
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
Public Function GetPatrimonio()
   Call SetUpGrid
   Call LoadAll
End Function


Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
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
Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub


