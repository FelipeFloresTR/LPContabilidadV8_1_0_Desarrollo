VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLibDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Diario"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "FrmLibDiario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Ch_VerNumDoc 
      Caption         =   "Ver N° Documento"
      Height          =   255
      Left            =   9600
      TabIndex        =   28
      Top             =   7740
      Width           =   1635
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7740
      Width           =   9435
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5595
      Left            =   30
      TabIndex        =   9
      Top             =   1740
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9869
      _Version        =   393216
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483643
      GridColorFixed  =   12632256
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   720
      Width           =   11415
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1635
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   3435
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   3435
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1740
         Picture         =   "FrmLibDiario.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3780
         Picture         =   "FrmLibDiario.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   10200
         Picture         =   "FrmLibDiario.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   1800
         TabIndex        =   29
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   4860
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   0
         Left            =   4860
         TabIndex        =   26
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   60
      Width           =   11415
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
         Picture         =   "FrmLibDiario.frx":0A5E
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   120
         Picture         =   "FrmLibDiario.frx":0ED2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   2460
         Picture         =   "FrmLibDiario.frx":12F5
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
         Left            =   3420
         Picture         =   "FrmLibDiario.frx":1399
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
         Left            =   3000
         Picture         =   "FrmLibDiario.frx":16FA
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
         Left            =   3840
         Picture         =   "FrmLibDiario.frx":1A98
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
         Left            =   1920
         Picture         =   "FrmLibDiario.frx":1EC1
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
         Left            =   1080
         Picture         =   "FrmLibDiario.frx":2306
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10200
         TabIndex        =   19
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
         Picture         =   "FrmLibDiario.frx":27AD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7380
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmLibDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NCUENTA = 0
Const C_CUENTA = 1
Const C_DOC = 2
Const C_DEBE = 3
Const C_HABER = 4
Const C_GLOSA = 5
Const C_IDCOMP = 6
Const C_IDDOC = 7
Const C_OBLIGATORIA = 8
Const C_FMT = 9

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lMes As Integer
Dim lidComp As Long

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
   
   Call LP_FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & vbTab & " Fecha Término: " & Tx_Hasta)
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
   Dim Q1 As String
   Dim nFolio As Integer
   
   lPapelFoliado = False
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(LIBOF_DIARIO, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(LIBOF_DIARIO) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
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
         Call AppendLogImpreso(LIBOF_DIARIO, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, nFolio)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
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

Private Sub Ch_VerNumDoc_Click()

   If Ch_VerNumDoc <> 0 Then
      Grid.ColWidth(C_DOC) = 3000
      If Grid.ColWidth(C_DEBE) > 1200 Then
         Grid.ColWidth(C_DEBE) = Grid.ColWidth(C_DEBE) - 300
         Grid.ColWidth(C_HABER) = Grid.ColWidth(C_HABER) - 300
      End If
   Else
      Grid.ColWidth(C_DOC) = 0
      Grid.ColWidth(C_DEBE) = Grid.ColWidth(C_DEBE) + 300
      Grid.ColWidth(C_HABER) = Grid.ColWidth(C_HABER) + 300
      
   End If
   
   If Not lInLoad Then
      Call LoadAll
   End If
   
End Sub

Private Sub Form_Load()
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   
   lInLoad = True
   
   ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
   
   Call FirstLastMonthDay(ActDate, D1, D2)
   Call SetTxDate(Tx_Desde, D1)
   Call SetTxDate(Tx_Hasta, D2)
   
   lOrientacion = ORIENT_VER
   
   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))
         
   'Call FillCbCuentas(Cb_Cuentas)
   Call FillCbAreaNeg(Cb_AreaNeg, False)
   Call FillCbCCosto(Cb_CCosto, False)
   
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
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.ColWidth(C_NCUENTA) = 1250
   Grid.ColWidth(C_CUENTA) = 2000
   Grid.ColWidth(C_DOC) = 3000
   Grid.ColWidth(C_DEBE) = 1200
   Grid.ColWidth(C_HABER) = 1200
   Grid.ColWidth(C_GLOSA) = 6200
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_IDDOC) = 0
      
   Grid.ColAlignment(C_NCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   
   Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim IdComp As Long
   Dim EstadoComp As Integer
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim Wh As String
   Dim Idx As Integer
   Dim CodCuenta As String
   Dim WhEstado As String
   Dim TxtRes As String
   Dim StrAnulado As String
   Dim D1 As Long, D2 As Long
   
   lTotDebe = 0
   lTotHaber = 0
   
   Grid.Redraw = False
   
   If gCompAnuladoLibDiario Then
      StrAnulado = "," & EC_ANULADO
   End If
      
   If Ch_LibOficial <> 0 Then
      WhEstado = " AND Comprobante.Estado IN(" & EC_APROBADO & StrAnulado & ")"
      MsgBox1 "Dado que es Libro Oficial, sólo se seleccionarán los comprobantes APROBADOS.", vbInformation + vbOKOnly
   Else
      WhEstado = " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & StrAnulado & ")"
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

   
   'primero los no resumidos
   Q1 = "SELECT Cuentas.Codigo, Descripcion, MovComprobante.Debe as Debe, MovComprobante.Haber As Haber"
   Q1 = Q1 & ",MovComprobante.Glosa as GlosaMov, Comprobante.idComp, Fecha, Comprobante.Tipo, iif(Comprobante.Tipo = " & TC_APERTURA & ",0,Comprobante.Tipo) as TipoAp "
   Q1 = Q1 & ",Comprobante.Correlativo, Comprobante.Glosa as GlosaComp, Orden as SOrden, ImpResumido, TipoAjuste "
   Q1 = Q1 & ",Documento.TipoLib, Documento.IdDoc, Documento.TipoDoc, Documento.NumDoc, Entidades.RUT, Entidades.Nombre, Entidades.NotValidRut "
   Q1 = Q1 & ", Comprobante.Estado as EstadoComp, MovComprobante.IdMov "
   Q1 = Q1 & " FROM (((( MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
   
   If lidComp > 0 Then
      Q1 = Q1 & " WHERE Comprobante.IdComp = " & lidComp
   Else
      Q1 = Q1 & " WHERE (Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
   End If
   
   Q1 = Q1 & WhEstado & Wh
   
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   If gFunciones.ComprobanteResumido Then
      Q1 = Q1 & " AND ImpResumido = 0 "
      
      Q1 = Q1 & " UNION "
      
      'y ahora los resumidos
      Q1 = Q1 & "SELECT Cuentas.Codigo, Descripcion, Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) as Haber "
      Q1 = Q1 & ",' ' as GlosaMov, Comprobante.idComp, Fecha, Comprobante.Tipo, iif(Comprobante.Tipo = " & TC_APERTURA & ",0,Comprobante.Tipo) as TipoAp "
      Q1 = Q1 & ",Comprobante.Correlativo, Comprobante.Glosa as GlosaComp, 0, ImpResumido, TipoAjuste "
      Q1 = Q1 & ",0, 0, 0, ' ', ' ', ' ', 0 "
      Q1 = Q1 & ", Comprobante.Estado as EstadoComp, 0"
      Q1 = Q1 & " FROM ((((MovComprobante "
      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp = Comprobante.idComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
      Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
      
      If lidComp > 0 Then
         Q1 = Q1 & " WHERE Comprobante.IdComp = " & lidComp
      Else
         Q1 = Q1 & " WHERE (Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
      End If
      Q1 = Q1 & WhEstado & Wh
      Q1 = Q1 & " AND ImpResumido <> 0 "
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      
      Q1 = Q1 & " GROUP BY Cuentas.Codigo, Descripcion "
      Q1 = Q1 & ", Comprobante.idComp, Fecha, Tipo, iif(Comprobante.Tipo = " & TC_APERTURA & ",0,Comprobante.Tipo) "
      Q1 = Q1 & ", Comprobante.Correlativo, Comprobante.Glosa, ImpResumido, TipoAjuste, Comprobante.Estado "
   End If
   
   'Q1 = Q1 & " ORDER BY Fecha, Comprobante.Tipo, Comprobante.Correlativo, Orden"
   Q1 = Q1 & " ORDER BY Fecha, Comprobante.Correlativo, TipoAp, TipoAjuste, SOrden"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Clear
   Grid.TextMatrix(0, C_FMT) = ".FMT"
   
   Row = 1
   Grid.rows = 1
   EstadoComp = 0
   
   'para que formatee al imprimir (al final (después de FGrVRows) le ponemos RowHeight=0 para que no se vea
   Grid.TextMatrix(0, C_FMT) = "      .FMT"
  
   Do While Rs.EOF = False
      
      If FGrChkMaxSize(Grid) = True Then
         MsgBox1 "Se mostrarán " & Grid.rows & " registros.", vbInformation + vbOKOnly
         Exit Do
      End If
      
      Grid.rows = Grid.rows + 1

      If IdComp <> vFld(Rs("idComp")) Then
         If Row > 1 Then
           Call Totales(Row, SumDebe, SumHaber, False, EstadoComp)
           
         End If
         
         Call Encabezado(vFld(Rs("Correlativo")), vFld(Rs("Tipo")), vFld(Rs("Fecha")), Row, vFld(Rs("GlosaComp"), True), vFld(Rs("IdComp")), vFld(Rs("ImpResumido")), vFld(Rs("EstadoComp")), vFld(Rs("TipoAjuste")))
         
         If lMes = 0 Then   'viene de FViewChain
            
            Call FirstLastMonthDay(vFld(Rs("Fecha")), D1, D2)
            Call SetTxDate(Tx_Desde, D1)
            Call SetTxDate(Tx_Hasta, D2)
            
         End If
         
         Call FixedRows(Row)
         IdComp = vFld(Rs("idComp"))
         EstadoComp = vFld(Rs("EstadoComp"))
         If vFld(Rs("ImpResumido")) And gFunciones.ComprobanteResumido Then
           TxtRes = ""    ' "[res.]"
         End If
          
      End If
      Grid.TextMatrix(Row, C_NCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta) ' vFld(Rs("Codigo"))
      Grid.TextMatrix(Row, C_CUENTA) = FCase(vFld(Rs("Descripcion"), True))
      
      If vFld(Rs("IdDoc")) <> 0 Then
         Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("IdDoc"))
         Grid.TextMatrix(Row, C_DOC) = "[" & GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc")) & "]   " & IIf(vFld(Rs("Rut")) <> "", FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False), "") & " " & vFld(Rs("Nombre"), True)
      ElseIf vFld(Rs("ImpResumido")) Then
         Grid.TextMatrix(Row, C_DOC) = TxtRes
         TxtRes = ""
      End If
      
      Grid.TextMatrix(Row, C_DEBE) = Format(vFld(Rs("Debe")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_HABER) = Format(vFld(Rs("Haber")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_GLOSA) = Left(vFld(Rs("GlosaMov"), True), 30)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "        O"
         
      SumDebe = vFld(Rs("Debe")) + SumDebe
      SumHaber = vFld(Rs("Haber")) + SumHaber
      
      Row = Row + 1
      Grid.rows = Row + 1
      
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   'ponemos el total del último comprobante
   If SumDebe > 0 Or SumHaber > 0 Then
      Call Totales(Row, SumDebe, SumHaber, True, EstadoComp)
   End If
   
   'totales finales
   GridTot.TextMatrix(0, C_DOC) = "TOTAL"
   GridTot.TextMatrix(0, C_DEBE) = Format(lTotDebe, BL_NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(lTotHaber, BL_NUMFMT)

   
   If Grid.rows <= 1 Then
      Grid.rows = 2
   End If
   Call FGrVRows(Grid)
   
   Grid.RowHeight(0) = 0  'Row con el formateo
   Grid.rows = Grid.rows + 1
   
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = 0
   Grid.ColSel = 0

   Grid.Redraw = True
   
   Call EnableFrm(False)

End Sub
Private Sub FixedRows(Row As Integer)

   Grid.TextMatrix(Row, C_FMT) = "LB"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "          O"
      
   Grid.TextMatrix(Row, C_NCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(Row, C_CUENTA) = "Cuenta Contable"
   Grid.TextMatrix(Row, C_DOC) = "Documento"
   Grid.TextMatrix(Row, C_DEBE) = "Debe"
   Grid.TextMatrix(Row, C_HABER) = "Haber"
   Grid.TextMatrix(Row, C_GLOSA) = "Glosa"
   
   Call FGrSetRowStyle(Grid, Row, "BC", vbButtonFace)
   Call FGrSetRowStyle(Grid, Row, "Align", flexAlignCenterCenter)
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   
End Sub

Private Sub Encabezado(ByVal Comprobante As Long, ByVal Tipo As Byte, ByVal Fecha As Long, Row As Integer, ByVal Glosa As String, ByVal IdComp As Long, ByVal ImpResumido As Integer, ByVal EstadoComp As Integer, ByVal TipoAjuste As Integer)
   Dim Res As String
   Dim Estado As String
   
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "         O"
      
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "Align", flexAlignLeftCenter)
      
   Grid.TextMatrix(Row, C_NCUENTA) = "Comprobante: "
   If ImpResumido <> 0 Then
      Res = ""    '"[res.]"
   End If
   
   If EstadoComp = EC_ANULADO Then
      Estado = " (Anulado)"
   End If
   
   Grid.TextMatrix(Row, C_CUENTA) = UCase(Left(gTipoComp(Tipo), 1)) & " " & Comprobante & IIf(TipoAjuste = TAJUSTE_TRIBUTARIO, "-T", "") & Estado
   Grid.TextMatrix(Row, C_IDCOMP) = IdComp
   If Grid.ColWidth(C_DOC) <> 0 Then
      Grid.TextMatrix(Row, C_DOC) = "Fecha: " & Format(Fecha, EDATEFMT)
   Else
      Grid.TextMatrix(Row, C_DEBE) = Format(Fecha, EDATEFMT)
   End If
   
   Grid.Row = Row
   Grid.Col = C_HABER
   Grid.CellAlignment = flexAlignRightCenter
   Grid.TextMatrix(Row, C_HABER) = "Glosa:"
   
   Grid.TextMatrix(Row, C_GLOSA) = Glosa
   
   Row = Row + 1
   Grid.rows = Row + 1
 
End Sub

Private Sub Totales(Row As Integer, Debe As Double, Haber As Double, ByVal LastRs As Boolean, ByVal EstadoComp As Integer)
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "            O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   
   Grid.Col = C_DOC
   Grid.Row = Row
   Grid.CellAlignment = flexAlignRightCenter
   If Grid.ColWidth(C_DOC) <> 0 Then
      Grid.TextMatrix(Row, C_DOC) = "Totales"
   Else
      Grid.TextMatrix(Row, C_CUENTA) = "Totales"
   End If
   
   
   Grid.TextMatrix(Row, C_DEBE) = Format(Debe, BL_NUMFMT)
   Grid.TextMatrix(Row, C_HABER) = Format(Haber, BL_NUMFMT)
   
   
   If EstadoComp <> EC_ANULADO Then
      lTotDebe = lTotDebe + Debe
      lTotHaber = lTotHaber + Haber
   End If
   
   Debe = 0
   Haber = 0
   
   If LastRs = False Then
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "            O"
      Grid.TextMatrix(Row, C_FMT) = "L"
      
      Row = Row + 1
      Grid.rows = Row + 1
   End If
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If

   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(C_FMT) As Integer
   Dim Total(C_FMT) As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = "LIBRO DIARIO"
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
   
   If Grid.ColWidth(C_DOC) <> 0 Then
      ColWi(C_GLOSA) = 3200
      ColWi(C_DOC) = ColWi(C_DOC) - 700
   Else
      ColWi(C_GLOSA) = ColWi(C_GLOSA) - 1200
   End If
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_OBLIGATORIA
   gPrtLibros.NTotLines = 1
   
    If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
  
   Call SetPrtNotas(True)  'vemos si hay que poner nota Art. 100
   
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
Private Sub EnableFrm(ByVal bool As Boolean)
   Bt_Buscar.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
   
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
Public Function FViewChain(ByVal IdComp As Long, Optional ByVal TipoAjuste As Integer = 0)
   Dim MesActual As Integer
   
   lidComp = IdComp
   lTipoAjuste = TipoAjuste
   
   lMes = 0
   
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
   
   If Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
      Set Frm = Nothing
   ElseIf Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      Set Frm = New FrmDoc
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
      Set Frm = Nothing
   End If
      
End Sub
Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub

