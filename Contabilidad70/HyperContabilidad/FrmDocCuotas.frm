VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmDocCuotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Documento a Crédito"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_TotCuotas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   8460
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   8895
      Begin VB.TextBox Tx_Estado 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton Bt_FInicio 
         Height          =   315
         Left            =   5040
         Picture         =   "FrmDocCuotas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   780
         Width           =   230
      End
      Begin VB.TextBox Tx_FInicio 
         Height          =   315
         Left            =   3900
         TabIndex        =   0
         Top             =   780
         Width           =   1155
      End
      Begin VB.CommandButton Bt_GenCuotas 
         Caption         =   "Generar Cuotas"
         Height          =   855
         Left            =   7440
         Picture         =   "FrmDocCuotas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox Tx_NumCuotas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6540
         TabIndex        =   2
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox Tx_Total 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox Tx_FEmision 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_Doc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado Doc.:"
         Height          =   195
         Left            =   5460
         TabIndex        =   30
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Primera cuota:"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   28
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "N° de Cuotas:"
         Height          =   315
         Index           =   2
         Left            =   5460
         TabIndex        =   25
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha emisión:"
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Documento:"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton Bt_Del 
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
         Picture         =   "FrmDocCuotas.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar cuota seleccionada"
         Top             =   120
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
         Left            =   1440
         Picture         =   "FrmDocCuotas.frx":0BB0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerLibCaja 
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
         Left            =   480
         Picture         =   "FrmDocCuotas.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ver Libro de Caja con el Ingreso Percibido"
         Top             =   120
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
         Left            =   60
         Picture         =   "FrmDocCuotas.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   7740
         TabIndex        =   16
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6600
         TabIndex        =   15
         Top             =   180
         Width           =   1035
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
         Left            =   1980
         Picture         =   "FrmDocCuotas.frx":13E5
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   120
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
         Left            =   2400
         Picture         =   "FrmDocCuotas.frx":188C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   120
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
         Left            =   2820
         Picture         =   "FrmDocCuotas.frx":1D46
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Copiar Excel"
         Top             =   120
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
         Left            =   3780
         Picture         =   "FrmDocCuotas.frx":218B
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Calculadora"
         Top             =   120
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
         Left            =   3360
         Picture         =   "FrmDocCuotas.frx":24EC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Convertir moneda"
         Top             =   120
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
         Left            =   4200
         Picture         =   "FrmDocCuotas.frx":288A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Calendario"
         Top             =   120
         Width           =   375
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   6315
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   11139
      Cols            =   2
      Rows            =   4
      FixedCols       =   1
      FixedRows       =   2
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Label Lb_TotCuotas 
      Caption         =   "Total"
      Height          =   195
      Left            =   1440
      TabIndex        =   27
      Top             =   8520
      Width           =   495
   End
End
Attribute VB_Name = "FrmDocCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOCCUOTA = 0
Const C_NUMCUOTA = 1
Const C_LNGFECHAEXIGPAGO = 2
Const C_FECHAEXIGPAGO = 3
Const C_MONTOCUOTA = 4
Const C_FECHAINGPERCIBIDO = 5
Const C_IDCOMPPAGO = 6
Const C_COMPPAGO = 7
Const C_IDLIBCAJA = 8
Const C_LIBCAJA = 9
Const C_UPDATE = 10

Const NCOLS = C_UPDATE

Dim lIdDoc As Long
Dim lOper As Integer
Dim lRc As Integer
Dim lFEmision As Long
Dim lTotal As Double
Dim lTipoLib As Integer
Dim lTipoDoc As Integer
Dim lGenCuotas As Boolean
Dim lFVenc As Long
Dim lNumCuotas As Integer


Public Function FEdit(ByVal IdDoc As Long, FVenc As Long, NumCuotas As Integer) As Integer
   
   lOper = O_EDIT
   
   lIdDoc = IdDoc
   Me.Show vbModal
   
   FVenc = lFVenc
   NumCuotas = lNumCuotas
   FEdit = lRc
   
End Function

Public Sub FView(ByVal IdDoc As Long)
   
   lOper = O_VIEW
   
   lIdDoc = IdDoc
   Me.Show vbModal
   
End Sub

Private Function SetUpGrid()
   
   Grid.Cols = NCOLS + 1
   Grid.rows = 4
   Grid.FixedRows = 2
   Grid.FixedCols = C_NUMCUOTA + 1
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDDOCCUOTA) = 0
   Grid.ColWidth(C_NUMCUOTA) = 1000
   Grid.ColWidth(C_LNGFECHAEXIGPAGO) = 0
   Grid.ColWidth(C_FECHAEXIGPAGO) = 1500
   Grid.ColWidth(C_MONTOCUOTA) = 1200
   Grid.ColWidth(C_FECHAINGPERCIBIDO) = 1300
   Grid.ColWidth(C_IDCOMPPAGO) = 0
   Grid.ColWidth(C_COMPPAGO) = 1800
   Grid.ColWidth(C_IDLIBCAJA) = 0
   Grid.ColWidth(C_LIBCAJA) = 1600
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_NUMCUOTA) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHAEXIGPAGO) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOCUOTA) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHAINGPERCIBIDO) = flexAlignRightCenter
   Grid.ColAlignment(C_COMPPAGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_LIBCAJA) = flexAlignLeftCenter
   
   Grid.TextMatrix(1, C_NUMCUOTA) = "N° Cuota"
   Grid.TextMatrix(0, C_FECHAEXIGPAGO) = "Fecha"
   Grid.TextMatrix(1, C_FECHAEXIGPAGO) = "Exigibilidad Pago"
   Grid.TextMatrix(0, C_MONTOCUOTA) = "Monto Cuota"
   Grid.TextMatrix(1, C_MONTOCUOTA) = "$"
   Grid.TextMatrix(0, C_FECHAINGPERCIBIDO) = "Fecha"
   Grid.TextMatrix(1, C_FECHAINGPERCIBIDO) = "Ingreso Percibido"
   Grid.TextMatrix(0, C_COMPPAGO) = "Comprobante"
   Grid.TextMatrix(1, C_COMPPAGO) = "de Pago"
   Grid.TextMatrix(0, C_LIBCAJA) = "Libro"
   Grid.TextMatrix(1, C_LIBCAJA) = "de Caja"
   
   Call FGrLocateCntrl(Grid, Tx_TotCuotas, C_MONTOCUOTA)
   Lb_TotCuotas.Left = Tx_TotCuotas.Left - Lb_TotCuotas.Width - 30
   
   Call FGrVRows(Grid, 1)
   
   
End Function

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
   
End Sub

Private Sub Bt_Del_Click()
   Dim NumCuota As Long
   Dim Row As Integer
   
   Row = Grid.Row
   
   NumCuota = Val(Grid.TextMatrix(Row, C_NUMCUOTA))
   If NumCuota = 0 Then
      Exit Sub
   End If
   
   'verificamos que sea la última cuota
   If Val(Grid.TextMatrix(Row + 1, C_NUMCUOTA)) <> 0 Then
      MsgBox1 "Debe eliminar la última cuota primero.", vbExclamation
      Exit Sub
   End If
      
   If Trim(Grid.TextMatrix(Row + 1, C_FECHAINGPERCIBIDO)) <> "" Then
      MsgBox1 "No puede eliminar esta cuota, ya fue pagada.", vbExclamation
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row + 1, C_IDCOMPPAGO)) <> 0 Then
      MsgBox1 "No puede eliminar esta cuota, ya fue contabilizado el pago.", vbExclamation
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row + 1, C_IDLIBCAJA)) <> 0 Then
      MsgBox1 "No puede eliminar esta cuota, ya fue incorporada al libro de caja.", vbExclamation
      Exit Sub
   End If
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDDOCCUOTA, C_UPDATE)
   
   Grid.TextMatrix(Row, C_NUMCUOTA) = "0"          'se asigna "0" no "" para diferenciar con registro en blanco
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot

End Sub

Private Sub Bt_GenCuotas_Click()
   Dim Total As Double
   Dim NCuotas As Integer
   Dim MontoCuota As Double, Resto As Integer
   Dim i As Integer, Row As Integer
   Dim FInicio As Long, Dt As Long
   
   Total = vFmt(Tx_Total)
   NCuotas = vFmt(Tx_NumCuotas)
   FInicio = GetTxDate(Tx_FInicio)
   
   If Total = 0 Then
      MsgBox1 "Total es cero", vbExclamation
      Exit Sub
   End If
   
   If NCuotas = 0 Then
      If MsgBox1("Número de cuotas es cero." & vbCrLf & vbCrLf & "¿Desea eliminar todas las cuotas?", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
   End If
   
   If GetTxDate(Tx_FInicio) < lFEmision Then
      MsgBox1 "Fecha primera cuota anterior a fecha de emisión del documento", vbExclamation
      Exit Sub
   End If
   
   If GetTxDate(Tx_FInicio) > CLng(DateAdd("d", 30, lFEmision)) Then
      If MsgBox1("La fecha primera cuota es a más de 30 días de la fecha de emisión del documento." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
      End If
   End If
   
   If NCuotas = 0 Then
      Grid.rows = Grid.FixedRows
      Call FGrVRows(Grid)
      
   Else
      MontoCuota = Int(Total / NCuotas)
      Resto = Total Mod NCuotas
      
      Row = Grid.FixedRows
      
      Grid.FlxGrid.Redraw = False
      Grid.rows = Grid.FixedRows
      
      For i = 1 To NCuotas
         Grid.rows = Grid.rows + 1
         
         Grid.TextMatrix(Row, C_NUMCUOTA) = i
         If Day(FInicio) <= 28 Then
            Dt = DateAdd("m", 1 * (i - 1), FInicio)
            Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = CLng(DateSerial(Year(Dt), Month(Dt), Day(FInicio)))
         Else
            Dt = DateAdd("d", 30 * (i - 1), FInicio)
            Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = Dt
         End If
            
         Grid.TextMatrix(Row, C_FECHAEXIGPAGO) = Format(Dt, EDATEFMT)
         If i = 1 Then
            Grid.TextMatrix(Row, C_MONTOCUOTA) = Format(MontoCuota + Resto, NUMFMT)
         Else
            Grid.TextMatrix(Row, C_MONTOCUOTA) = Format(MontoCuota, NUMFMT)
         End If
         
         Call FGrModRow(Grid, Row, FGR_U, C_IDDOCCUOTA, C_UPDATE)
         
         Row = Row + 1
         
      Next i
   
   End If
   
   Call CalcTot
   
   Call FGrVRows(Grid, 1)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_NUMCUOTA
   Grid.FlxGrid.Redraw = True
   lGenCuotas = True

End Sub

Private Sub bt_OK_Click()
   
   If Valida() Then
      Call SaveGrid
      lRc = vbOK
      lFVenc = GetTxDate(Tx_FInicio)
      lNumCuotas = vFmt(Tx_NumCuotas)
      Unload Me
   End If
   
End Sub

Private Sub Bt_VerLibCaja_Click()
   Dim IdLibCaja As Long, Frm As FrmLibCaja
   Dim Q1 As String, Rs As Recordset
   Dim Mes As Integer
   
   IdLibCaja = Val(Grid.TextMatrix(Grid.Row, C_IDLIBCAJA))
   
   If IdLibCaja <= 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT FechaIngresoLibro FROM LibroCaja WHERE IdLibroCaja = " & IdLibCaja
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Mes = Month(vFld(Rs("FechaIngresoLibro")))
   End If
   
   Call CloseRs(Rs)
   
   If Mes <> 0 Then
      Set Frm = New FrmLibCaja
      Frm.FView (Mes)
      Set Frm = Nothing
   End If

End Sub

Private Sub Form_Load()

   Call SetUpGrid
   
   If lOper = O_VIEW Then
      Bt_OK.visible = False
      Bt_Cancel.Caption = "Cerrar"
      Grid.Locked = True
      Bt_GenCuotas.visible = False
      
   End If
   
   lFVenc = 0
   lNumCuotas = 0
   
   Call LoadAll
   
   Call SetupPriv
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim NCuotas As Integer
   Dim FPrimeraCuota As Long
   Dim FVenc As Long
   Dim HayCuotasPagadas As Boolean
   Dim EstadoDoc As Integer
   Dim SaldoDoc As Double
   
   If lIdDoc = 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT TipoLib, TipoDoc, NumDoc, FEmisionOri, FVenc, Total, Estado, SaldoDoc FROM Documento WHERE IdDoc = " & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      lTipoLib = vFld(Rs("TipoLib"))
      lTipoDoc = vFld(Rs("TipoDoc"))
      lFEmision = vFld(Rs("FEmisionOri"))
      FVenc = vFld(Rs("FVenc"))
      lTotal = vFld(Rs("Total"))
      Tx_Doc = gTipoDoc(GetTipoDoc(lTipoLib, lTipoDoc)).Diminutivo & "-" & vFld(Rs("NumDoc"))
      Tx_FEmision = Format(lFEmision, DATEFMT)
      Tx_Total = Format(lTotal, NUMFMT)
      EstadoDoc = vFld(Rs("Estado"))
      Tx_Estado = gEstadoDoc(EstadoDoc)
      SaldoDoc = vFld(Rs("SaldoDoc"))
   End If
   Call CloseRs(Rs)
   
   
   Q1 = "SELECT IdDocCuota, NumCuota, DocCuotas.FechaExigPago, MontoCuota, FechaIngPercibido, IdCompPago, IdLibCaja "
   Q1 = Q1 & ", LibroCaja.FechaIngresoLibro As FechaLibCaja, LibroCaja.TipoOper, Comprobante.Tipo As CompTipo, Comprobante.Fecha As CompFecha, Comprobante.Correlativo as CompCorr "
   Q1 = Q1 & " FROM (DocCuotas LEFT JOIN Comprobante ON DocCuotas.IdCompPago = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "DocCuotas") & " )"
   Q1 = Q1 & " LEFT JOIN LibroCaja ON DocCuotas.IdLibCaja = LibroCaja.IdLibroCaja "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "DocCuotas")
   Q1 = Q1 & " WHERE DocCuotas.IdDoc = " & lIdDoc
   Q1 = Q1 & " AND DocCuotas.IdEmpresa = " & gEmpresa.id & " AND DocCuotas.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY NumCuota "
      
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   NCuotas = 0
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDOCCUOTA) = vFld(Rs("IdDocCuota"))
      Grid.TextMatrix(i, C_NUMCUOTA) = vFld(Rs("NumCuota"))
      Grid.TextMatrix(i, C_LNGFECHAEXIGPAGO) = vFld(Rs("FechaExigPago"))
      Grid.TextMatrix(i, C_FECHAEXIGPAGO) = Format(vFld(Rs("FechaExigPago")), EDATEFMT)
      If FPrimeraCuota = 0 Then
         FPrimeraCuota = vFld(Rs("FechaExigPago"))
      End If
      
      Grid.TextMatrix(i, C_MONTOCUOTA) = Format(vFld(Rs("MontoCuota")), NUMFMT)
      
      If vFld(Rs("FechaIngPercibido")) > 0 Then
         Grid.TextMatrix(i, C_FECHAINGPERCIBIDO) = Format(vFld(Rs("FechaIngPercibido")), EDATEFMT)
      End If
      
      Grid.TextMatrix(i, C_IDCOMPPAGO) = vFld(Rs("IdCompPago"))
      If vFld(Rs("IdCompPago")) > 0 Then
         Grid.TextMatrix(i, C_COMPPAGO) = UCase(Left(gTipoComp(vFld(Rs("CompTipo"))), 1)) & "-" & vFld(Rs("CompCorr")) & "   " & Format(vFld(Rs("CompFecha")), EDATEFMT)
      End If
      
      Grid.TextMatrix(i, C_IDLIBCAJA) = vFld(Rs("IdLibCaja"))
      If vFld(Rs("IdLibCaja")) > 0 Then
'         Grid.TextMatrix(i, C_LIBCAJA) = gTipoOperCaja(vFld(Rs("TipoOper"))) & "s " & Format(vFld(Rs("CompFecha")), "mm/yyyy")
         Grid.TextMatrix(i, C_LIBCAJA) = "Mes " & Format(vFld(Rs("CompFecha")), "mm/yyyy")
      End If
      
      If vFld(Rs("FechaIngPercibido")) > 0 Or vFld(Rs("IdCompPago")) Or vFld(Rs("IdLibCaja")) Then
         HayCuotasPagadas = True
      End If
      
      Rs.MoveNext
      NCuotas = NCuotas + 1
      i = i + 1
   
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_NUMCUOTA
   Grid.FlxGrid.Redraw = True
      
   If NCuotas = 0 Then
      Call SetTxDate(Tx_FInicio, FVenc)
      Tx_NumCuotas = ""
   Else
      Tx_NumCuotas = NCuotas
      Call SetTxDate(Tx_FInicio, FPrimeraCuota)
   End If
   
   If HayCuotasPagadas Or (EstadoDoc = ED_PAGADO And NCuotas = 0 And SaldoDoc = 0) Then
      Bt_GenCuotas.Enabled = False
      Call SetTxRO(Tx_NumCuotas, True)
      Call SetTxRO(Tx_FInicio, True)
      Bt_FInicio.Enabled = False
   End If
   
   Call CalcTot
   
   If EstadoDoc + ED_PAGADO And NCuotas = 0 Then
      MsgBox1 "No es posible definir cuotas para este documento porque está Pagado y el saldo es cero.", vbInformation
   End If
   
End Sub
Private Sub SaveGrid()
   Dim i As Integer
   Dim NCuotas As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim id As Long
   Dim Fecha As Long
   
   If lGenCuotas Then
'      Q1 = "DELETE * FROM DocCuotas WHERE IdDoc = " & lIdDoc
'      Call ExecSQL(DbMain, Q1)
      Q1 = " WHERE IdDoc = " & lIdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "DocCuotas", Q1)
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_NUMCUOTA) = "" Then     'ya terminó la lista de cuotas
         Exit For
      End If
      
      NCuotas = NCuotas + 1
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
      
         Q1 = "INSERT INTO DocCuotas (IdDoc, NumCuota, FechaExigPago, MontoCuota, FechaIngPercibido, IdCompPago, IdLibCaja, Estado, IdEmpresa, Ano )         "
         Q1 = Q1 & "VALUES(" & lIdDoc & "," & vFmt(Grid.TextMatrix(i, C_NUMCUOTA))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_LNGFECHAEXIGPAGO))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_MONTOCUOTA))
         Q1 = Q1 & ", 0, 0, 0, " & ED_PENDIENTE & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
         
         Call ExecSQL(DbMain, Q1)
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM DocCuotas WHERE IdDocCuota = " & Val(Grid.TextMatrix(i, C_IDDOCCUOTA))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdDocCuota = " & Val(Grid.TextMatrix(i, C_IDDOCCUOTA))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "DocCuotas", Q1)
         
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then 'Update
         Q1 = "UPDATE DocCuotas SET "
         Q1 = Q1 & "  FechaExigPago = " & vFmt(Grid.TextMatrix(i, C_LNGFECHAEXIGPAGO))
         Q1 = Q1 & ", MontoCuota = " & vFmt(Grid.TextMatrix(i, C_MONTOCUOTA))
         
         Q1 = Q1 & " WHERE IdDocCuota = " & Val(Grid.TextMatrix(i, C_IDDOCCUOTA))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
                  
      End If
      
   Next i

   'actualizamos la fecha de vencimiento del documento con la fecha de exig. de pago de la primera cuota
   If NCuotas > 0 Then
   
      Fecha = vFmt(Grid.TextMatrix(Grid.FixedRows, C_LNGFECHAEXIGPAGO))
      
      If Fecha > 0 Then
         Q1 = "UPDATE Documento SET FVenc = " & Fecha & ", NumCuotas = " & NCuotas
         Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      
      End If
      
   Else   'sólo actualizamos las cuotas en 0
      Q1 = "UPDATE Documento SET NumCuotas = 0"
      Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
   End If
   

End Sub
Private Function Valida() As Boolean
   Dim i As Integer
   Dim Total As Double
   Dim NCuotas As Integer
   
   Valida = False
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_NUMCUOTA) = "" Then   'los registros borrados tienen el NUmCuota en 0
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 And Val(Grid.TextMatrix(i, C_NUMCUOTA)) > 0 Then
         NCuotas = NCuotas + 1
      End If
       
   Next i
   
   If NCuotas <> vFmt(Tx_NumCuotas) Then
      MsgBox1 "La cantidad de cuotas no calza con el detalle de cuotas.", vbExclamation
      Exit Function
   End If
      
   
   Call CalcTot
      
   If vFmt(Tx_TotCuotas) <> vFmt(Tx_Total) And Val(Tx_NumCuotas) > 0 Then
      MsgBox1 "Suma total de cuotas no calza con el total del documento.", vbExclamation
      Exit Function
   End If
   
   Valida = True
   
End Function

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Fecha As Long

   Action = vbOK
   
   If Col = C_FECHAEXIGPAGO Then
      Fecha = GetDate(Value, "dmy")
      If Row > Grid.FixedRows Then
         If Fecha < GetDate(Grid.TextMatrix(Row - 1, C_FECHAEXIGPAGO), "dmy") And Grid.RowHeight(Row - 1) > 0 Then
            MsgBox1 "Fecha de exigibiliad de pago anterior a cuota previa.", vbExclamation + vbOKOnly
            Action = vbCancel
            Exit Sub
         End If
      End If
      If Val(Grid.TextMatrix(Row + 1, C_NUMCUOTA)) > 0 Then
         If Fecha > GetDate(Grid.TextMatrix(Row + 1, C_FECHAEXIGPAGO), "dmy") Then
            MsgBox1 "Fecha de exigibiliad de pago posterior a cuota siguiente.", vbExclamation + vbOKOnly
            Action = vbCancel
            Exit Sub
         End If
      End If
      
      Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO) = Fecha
      
      Value = Format(Grid.TextMatrix(Row, C_LNGFECHAEXIGPAGO), EDATEFMT)

   ElseIf Col = C_MONTOCUOTA Then
      Value = Format(vFmt(Value), NUMFMT)
      Grid.TextMatrix(Row, C_MONTOCUOTA) = Value
      Call CalcTot
      
   End If
         
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDDOCCUOTA, C_UPDATE)
   End If
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   Dim i As Integer

   If Val(Grid.TextMatrix(Row, C_IDCOMPPAGO)) > 0 Or Val(Grid.TextMatrix(Row, C_IDLIBCAJA)) > 0 Then
      Exit Sub
   End If
   
   If Grid.Col <> C_FECHAEXIGPAGO And Grid.Col <> C_MONTOCUOTA Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_NUMCUOTA)) = 0 Then
      For i = Row - 1 To Grid.FixedRows - 1 Step -1
         If Val(Grid.TextMatrix(i, C_NUMCUOTA)) > 0 Then
            Exit For
         End If
      Next i
         
      Grid.TextMatrix(Row, C_NUMCUOTA) = Val(Grid.TextMatrix(i, C_NUMCUOTA)) + 1
      Grid.rows = Grid.rows + 1
   End If

   EdType = FEG_Edit
   
End Sub


Private Sub Grid_DblClick()
   Dim Row As Integer

   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Col = C_COMPPAGO Then
      Call Bt_VerComp_Click
   ElseIf Grid.Col = C_LIBCAJA Then
      Call Bt_VerLibCaja_Click
   End If
   
      
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   If Grid.Col = C_FECHAEXIGPAGO Then
      Call KeyDate(KeyAscii)
   ElseIf Grid.Col = C_MONTOCUOTA Then
      Call KeyNumPos(KeyAscii)
   End If

End Sub

Private Sub Tx_FInicio_GotFocus()
   Call DtGotFocus(Tx_FInicio)
   
End Sub

Private Sub Tx_FInicio_LostFocus()
   
   If Trim$(Tx_FInicio) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FInicio)
      
End Sub

Private Sub Tx_FInicio_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_FInicio_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FInicio)
   Set Frm = Nothing
   
End Sub

Private Sub Tx_NumCuotas_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
   
End Sub

Private Sub CalcTot()
   Dim Total As Double
   Dim i As Integer
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_NUMCUOTA) = "" Then     'ya terminó la lista de cuotas, los registros borrados tienen valor 0
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 And Val(Grid.TextMatrix(i, C_NUMCUOTA)) > 0 Then
         Total = Total + vFmt(Grid.TextMatrix(i, C_MONTOCUOTA))
      End If
      
   Next i

   Tx_TotCuotas = Format(Total, NUMFMT)
   
End Sub

Private Function SetupPriv()
   Dim Enab As Boolean
     
   Enab = True
   If lOper = O_EDIT Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         Enab = False
      End If
   Else   'lOper = O_VIEW
      Enab = False

   End If
      
   If Not Enab Then
      Call EnableForm(Me, False)
      Bt_VerComp.Enabled = True
      Bt_VerLibCaja.Enabled = True
      Bt_Sum.Enabled = True
   End If

End Function

Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   Dim Encabezados(0) As String
   
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   
   gPrtReportes.Encabezados = Encabezados
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Encabezados(0) As String

   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   gPrtReportes.Encabezados = Encabezados
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(1) As String
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
   
   Encabezados(0) = "Documento: " & Tx_Doc
   Encabezados(1) = "Fecha Emisión: " & Tx_FEmision
   gPrtReportes.Encabezados = Encabezados
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   Total(C_FECHAEXIGPAGO) = "TOTAL"
   Total(C_MONTOCUOTA) = Tx_Total
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_NUMCUOTA
   gPrtReportes.NTotLines = 1
   
   
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
   Dim IdComp As Long, Frm As FrmComprobante
   
   IdComp = Val(Grid.TextMatrix(Grid.Row, C_IDCOMPPAGO))
   
   If IdComp <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmComprobante
   Call Frm.FView(IdComp, False)
   Set Frm = Nothing

End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_CopyExcel_Click()
   Dim Tit As String, Clip As String
   
   Tit = Me.Caption & vbCr & "Documento:" & vbTab & Tx_Doc & vbCr & "Fecha emisión:" & vbTab & Tx_FEmision & vbCr & "Total:" & vbTab & Tx_Total & vbCr & "Num. Cuotas:" & vbTab & Tx_NumCuotas & vbCr
   
'   Call FGr2Clip(Grid, Tit)
   
   Clip = FGr2String(Grid, Tit, False, C_NUMCUOTA)
   Clip = Clip & vbTab & "Total" & vbTab & Tx_TotCuotas
   
   Clipboard.Clear
   Clipboard.SetText Clip
   
End Sub
