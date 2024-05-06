VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResCartolas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartolas Bancarias"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "FrmResCartolas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_SaldoIni 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   7200
      Width           =   1515
   End
   Begin VB.TextBox Tx_TCargos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   7200
      Width           =   1515
   End
   Begin VB.TextBox Tx_TAbonos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   7200
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   19
      Top             =   600
      Width           =   9795
      Begin VB.ComboBox Cb_Concil 
         Height          =   315
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Tx_Desde 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Tx_Hasta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Cartola 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
      Begin VB.ComboBox Cb_CartBanco 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   6480
         TabIndex        =   20
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartola:"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   9795
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8340
         TabIndex        =   11
         Top             =   180
         Width           =   1275
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
         Picture         =   "FrmResCartolas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "FrmResCartolas.frx":00B0
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "FrmResCartolas.frx":0411
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "FrmResCartolas.frx":07AF
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "FrmResCartolas.frx":0BD8
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "FrmResCartolas.frx":101D
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "FrmResCartolas.frx":14C4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.TextBox Tx_Saldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7200
      Width           =   1515
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   60
      TabIndex        =   3
      Top             =   1380
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   30
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      Caption         =   "Nota: en verde se muestran los movimientos conciliados de este año y en azul los movimientos conciliados del año anterior"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   180
      TabIndex        =   26
      Top             =   7740
      Width           =   8775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   25
      Top             =   7260
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Saldo:"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   23
      Top             =   7260
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Saldo Contable:"
      Height          =   195
      Index           =   0
      Left            =   6780
      TabIndex        =   16
      Top             =   6540
      Width           =   1125
   End
End
Attribute VB_Name = "FrmResCartolas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_FECHA = 0
Private Const C_COMPR = 1
Private Const C_GLOSA = 2
Private Const C_DOC = 3
Private Const C_DEBE = 4
Private Const C_HABER = 5

Const NCOLS = C_HABER

Private lidBanco As Long
Private lidCartola As Long
Private lConcil As Integer

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Cb_CartBanco_Click()
   Dim Q1 As String
   
   Cb_Cartola.Clear
   
   Q1 = "SELECT " & SqlConcat(gDbType, "Ano", "' - '", "Cartola") & " as Cart, idCartola"
   Q1 = Q1 & " FROM Cartola"
   Q1 = Q1 & " WHERE idCuentaBco=" & ItemData(Cb_CartBanco)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Ano Desc, cartola Desc"
   Call FillCombo(Cb_Cartola, DbMain, Q1, lidCartola)

End Sub

Private Sub FillCartola()
   Dim Q1 As String, Rs As Recordset, r As Integer
   Dim TCargos As Double, TAbonos As Double, Saldo As Double, SaldoIni As Double

   If Cb_Cartola.ListCount <= 0 Then
      Exit Sub
   End If

   Tx_Desde = ""
   Tx_Hasta = ""
   Tx_Saldo = ""

   MousePointer = vbHourglass
   DoEvents

   Q1 = "SELECT FDesde, FHasta, SaldoIni FROM Cartola WHERE idCartola=" & ItemData(Cb_Cartola)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      SaldoIni = vFld(Rs("SaldoIni"))
      Tx_SaldoIni = Format(SaldoIni, NUMFMT)
      Tx_Desde = FmtFecha(vFld(Rs("FDesde")))
      Tx_Hasta = FmtFecha(vFld(Rs("FHasta")))
   Else
      SaldoIni = 0
   End If
   Call CloseRs(Rs)
   
   
   Q1 = "SELECT DetCartola.Fecha, Comprobante.Tipo, Comprobante.Correlativo, MovComprobante.Glosa, DetCartola.NumDoc, DetCartola.Detalle, DetCartola.Cargo, DetCartola.Abono, MovComprobante.Debe, MovComprobante.Haber, Cartola.FDesde, Cartola.FHasta, DetCartola.IdMov As DetCartIdMov"
   Q1 = Q1 & " FROM ((Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cartola", "DetCartola") & " )"
   Q1 = Q1 & " LEFT JOIN MovComprobante ON DetCartola.IdCartola = MovComprobante.IdCartola AND DetCartola.IdMov = MovComprobante.IdMov "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "DetCartola") & " )"
   Q1 = Q1 & " LEFT JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cartola.idCartola=" & ItemData(Cb_Cartola)
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   
   If ItemData(Cb_Concil) >= 0 Then
      Q1 = Q1 & " AND"
      If ItemData(Cb_Concil) Then
         Q1 = Q1 & " (DetCartola.IdMov < 0 OR (DetCartola.IdMov > 0 AND NOT MovComprobante.IdCartola IS NULL))"
      Else
         Q1 = Q1 & " DetCartola.IdMov >= 0 AND MovComprobante.IdCartola IS NULL"
      End If
   End If
   
   Q1 = Q1 & " ORDER BY DetCartola.Fecha, DetCartola.IdDetCartola"
   
   Grid.Redraw = False
   
   Set Rs = OpenRs(DbMain, Q1)
   r = 0
   Saldo = SaldoIni
   TCargos = 0
   TAbonos = 0
   Grid.rows = 1
   Do Until Rs.EOF
      r = r + 1
      Grid.rows = r + 1
      
      Grid.TextMatrix(r, C_FECHA) = FmtFecha(vFld(Rs("Fecha")))
      
      If IsNull(Rs("Tipo")) = False And IsNull(Rs("Correlativo")) = False Then
         Grid.TextMatrix(r, C_COMPR) = Left(gTipoComp(vFld(Rs("Tipo"))), 1) & " " & vFld(Rs("Correlativo"))
         Grid.TextMatrix(r, C_GLOSA) = vFld(Rs("Glosa"), True)
'         Grid.TextMatrix(r, C_DEBE) = Format(vFld(Rs("Debe")), NEGBL_NUMFMT)
'         Grid.TextMatrix(r, C_HABER) = Format(vFld(Rs("Haber")), NEGBL_NUMFMT)
         Grid.TextMatrix(r, C_DEBE) = Format(vFld(Rs("Haber")), NEGBL_NUMFMT)   'Abono  FCA 8 sep 2014 Se cambió orden para que calcen abonos y cargos
         Grid.TextMatrix(r, C_HABER) = Format(vFld(Rs("Debe")), NEGBL_NUMFMT)   'Cargo
         
         'If ItemData(Cb_Concil) = -1 Then    'todos
            Call FGrSetRowStyle(Grid, r, "FC", COLOR_VERDEOSCURO)
         'End If
         
         'TCargos = TCargos + vFld(Rs("Debe"))
         'TAbonos = TAbonos + vFld(Rs("Haber"))
         'FCA 12 Sep 2006
         TAbonos = TAbonos + vFld(Rs("Debe"))
         TCargos = TCargos + vFld(Rs("Haber"))
         
         Saldo = Saldo + vFld(Rs("Debe")) - vFld(Rs("Haber"))
      Else
         Grid.TextMatrix(r, C_GLOSA) = vFld(Rs("Detalle"), True)
         Grid.TextMatrix(r, C_DEBE) = Format(vFld(Rs("Cargo")), NEGBL_NUMFMT)
         Grid.TextMatrix(r, C_HABER) = Format(vFld(Rs("Abono")), NEGBL_NUMFMT)
         
         Saldo = Saldo + vFld(Rs("Abono")) - vFld(Rs("Cargo"))
         
         TCargos = TCargos + vFld(Rs("Cargo"))
         TAbonos = TAbonos + vFld(Rs("Abono"))
         
         If vFld(Rs("DetCartIdMov")) < 0 Then
            Call FGrSetRowStyle(Grid, r, "FC", COLOR_AZULOSCURO)
         End If
      End If
      
      Grid.TextMatrix(r, C_DOC) = vFld(Rs("NumDoc"))
   
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Debug.Print "Saldo:" & TAbonos - TCargos
   
   Tx_TCargos = Format(TCargos, NEGNUMFMT)
   Tx_TAbonos = Format(TAbonos, NEGNUMFMT)
   Tx_Saldo = Format(Saldo, NEGNUMFMT)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   Grid.Redraw = True

   MousePointer = vbDefault

End Sub

Private Sub Cb_Cartola_Click()
   
   MousePointer = vbHourglass
   DoEvents
   
   Call FillCartola
   
   MousePointer = vbDefault
   
End Sub


Private Sub Cb_Concil_Click()
   Call Cb_Cartola_Click
End Sub

Private Sub Form_Load()
   Dim Q1 As String

   If lidBanco = 0 Then
      lidBanco = -1
      lidCartola = -1
      lConcil = -1
   End If

   Call AddItem(Cb_Concil, "(todos)", -1)
   Call AddItem(Cb_Concil, "Conciliados", 1)
   Call AddItem(Cb_Concil, "No conciliados", 0)
   Call SelItem(Cb_Concil, lConcil)

   Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(Cb_CartBanco, DbMain, Q1, lidBanco)

   Call SetupForm

End Sub
Private Sub SetupForm()

   Call FGrSetup(Grid)

   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_COMPR) = "Compr."
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_DOC) = "Documento"
   Grid.TextMatrix(0, C_DEBE) = "Cargos"
   Grid.TextMatrix(0, C_HABER) = "Abonos"

   Grid.ColWidth(C_FECHA) = 1100
   Grid.ColWidth(C_COMPR) = 900
   Grid.ColWidth(C_GLOSA) = 4000
   Grid.ColWidth(C_DOC) = 1000
   Grid.ColWidth(C_DEBE) = W_MONTO
   Grid.ColWidth(C_HABER) = W_MONTO

   Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
   Grid.ColAlignment(C_COMPR) = flexAlignLeftCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter

   Call FGrLocateCntrl(Grid, Tx_TCargos, C_DEBE)
   Call FGrLocateCntrl(Grid, Tx_TAbonos, C_HABER)
   'Call FGrLocateCntrl(Grid, Tx_Saldo, C_HABER)

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call PostClick(Bt_CopyExcel)
   End If
   
End Sub
' Concil: -1 (todos), 1 conciliados, 0 no conciliados
Public Sub ShowCartola(ByVal idBanco As Long, ByVal idCartola As Long, ByVal Concil As Integer)

   lidBanco = idBanco
   lidCartola = idCartola
   lConcil = Concil

   Me.Show vbModal

End Sub
Private Sub Bt_Preview_Click()

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
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption & " - " & Cb_CartBanco & " - " & Cb_Cartola & " - " & Cb_Concil)
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
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(1) As String
   Dim Encabezados(1) As String
   Dim FontTit(0) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   Dim Totales(NCOLS) As String
   
   Set gPrtReportes.Grid = Grid
   
   Printer.Orientation = ORIENT_VER
   
   Titulos(0) = Caption
   
   If ItemData(Cb_Concil) = 1 Then
      Titulos(1) = "Movimientos Conciliados"
   ElseIf ItemData(Cb_Concil) = 0 Then
      Titulos(1) = "Movimientos No Conciliados"
   Else ' todos
      Titulos(1) = "Movimientos Conciliados y No Conciliados"
   End If
   
   Encabezados(0) = "Banco: " & vbTab & Cb_CartBanco
   Encabezados(1) = "Cartola: " & vbTab & Cb_Cartola
   
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   gPrtReportes.Encabezados = Encabezados
   FontEnc(0).FontBold = True
   FontEnc(0).FontName = "Arial"
   FontEnc(0).FontSize = 10
   Call gPrtReportes.FntEncabezados(FontEnc())
    
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Totales(i) = ""
   Next i
      
   Totales(C_HABER) = Tx_Saldo
      
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_FECHA
   gPrtReportes.Total = Totales
   gPrtReportes.NTotLines = 1
   
   
End Sub

