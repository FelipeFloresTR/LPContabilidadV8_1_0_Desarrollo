VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResInfConcil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Conciliación Bancaria"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "FrmResInfConcil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   13
      Top             =   660
      Width           =   10215
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   180
         Width           =   1695
      End
      Begin VB.ComboBox Cb_CartBanco 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8940
         TabIndex        =   10
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
         Picture         =   "FrmResInfConcil.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
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
         Picture         =   "FrmResInfConcil.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "FrmResInfConcil.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Copiar Excel"
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
         Picture         =   "FrmResInfConcil.frx":0DB2
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "FrmResInfConcil.frx":11DB
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "FrmResInfConcil.frx":1579
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "FrmResInfConcil.frx":18DA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.CommandButton Bt_Det 
      Caption         =   "Detalle Cartola..."
      Height          =   855
      Left            =   8940
      Picture         =   "FrmResInfConcil.frx":197E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   30
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "(*) El saldo conciliado no corresponde al saldo de la cartola."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   5055
   End
End
Attribute VB_Name = "FrmResInfConcil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_ANO = 0
Private Const C_CART = 1
Private Const C_FDESDE = 2
Private Const C_FHASTA = 3
Private Const C_INICIAL = 4
Private Const C_CARGOS = 5
Private Const C_ABONOS = 6
Private Const C_SALDO = 7
Private Const C_IDCART = 8
Const NCOLS = C_IDCART

Const E_CONCILIADO = 1
Const E_NOCONCILIADO = 2

Private Sub FillGrid()
   Dim Q1 As String, Rs As Recordset, r As Integer, Estado As Integer
   Dim SaldCon As Double

   Estado = ItemData(Cb_Estado)

   Q1 = "SELECT Cartola.Cartola, Cartola.idCartola, Cartola.TotAbono, Cartola.TotCargo, Cartola.FDesde, Cartola.FHasta, Cartola.SaldoIni, Sum(DetCartola.Abono) AS Abonos, Sum(DetCartola.Cargo) AS Cargos"
   Q1 = Q1 & " FROM Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cartola", "DetCartola")
   Q1 = Q1 & " WHERE Cartola.IdCuentaBco=" & ItemData(Cb_CartBanco)
   If Estado = E_CONCILIADO Then
      Q1 = Q1 & " AND DetCartola.idMov <> 0"
   ElseIf Estado = E_NOCONCILIADO Then
      Q1 = Q1 & " AND (DetCartola.idMov IS NULL OR DetCartola.idMov = 0)"
   End If
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cartola.Cartola, Cartola.idCartola, Cartola.TotAbono, Cartola.TotCargo, Cartola.FDesde, Cartola.FHasta, Cartola.SaldoIni"
   Q1 = Q1 & " ORDER BY Cartola.Cartola"

   Set Rs = OpenRs(DbMain, Q1)
   
   r = 0
   Grid.rows = 1
   Do Until Rs.EOF
      r = r + 1
      Grid.rows = r + 1
   
      SaldCon = vFld(Rs("SaldoIni")) + vFld(Rs("Abonos")) - vFld(Rs("Cargos"))
      
      Grid.TextMatrix(r, C_ANO) = gEmpresa.Ano
      Grid.TextMatrix(r, C_FDESDE) = FmtFecha(vFld(Rs("FDesde")))
      Grid.TextMatrix(r, C_FHASTA) = FmtFecha(vFld(Rs("FHasta")))
      Grid.TextMatrix(r, C_INICIAL) = Format(vFld(Rs("SaldoIni")), NEGBL_NUMFMT)
      Grid.TextMatrix(r, C_ABONOS) = Format(vFld(Rs("Abonos")), NEGBL_NUMFMT)
      Grid.TextMatrix(r, C_CARGOS) = Format(vFld(Rs("Cargos")), NEGBL_NUMFMT)
      Grid.TextMatrix(r, C_SALDO) = Format(SaldCon, NEGBL_NUMFMT)
      Grid.TextMatrix(r, C_IDCART) = vFld(Rs("idCartola"))
   
      If vFld(Rs("TotAbono")) - vFld(Rs("TotCargo")) <> SaldCon Then
         Q1 = " (*)"
      Else
         Q1 = ""
      End If
      
      Grid.TextMatrix(r, C_CART) = vFld(Rs("Cartola")) & Q1
   
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
      
   Call FGrVRows(Grid)

End Sub

Private Sub Bt_Buscar_Click()
   MousePointer = vbHourglass
   DoEvents
   
   Call FillGrid

   MousePointer = vbDefault

End Sub

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Det_Click()
   Dim Frm As FrmResCartolas
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_IDCART) = "" Then
      Exit Sub
   End If

   Set Frm = New FrmResCartolas
   Call Frm.ShowCartola(ItemData(Cb_CartBanco), Val(Grid.TextMatrix(Row, C_IDCART)), ItemData(Cb_Estado))
   Set Frm = Nothing

End Sub




Private Sub Cb_CartBanco_Click()
   Call Bt_Buscar_Click
End Sub

Private Sub Cb_Estado_Click()
   Call Bt_Buscar_Click

End Sub

Private Sub Form_Load()
   Dim Q1 As String

   Call AddItem(Cb_Estado, "(todos)", 0)
   Call AddItem(Cb_Estado, "Conciliados", 1, True)    '1
   Call AddItem(Cb_Estado, "No conciliados", 0)       '0

   Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(Cb_CartBanco, DbMain, Q1, -1)

   Call SetupForm
   
   Call Bt_Buscar_Click

End Sub

Private Sub SetupForm()

   Call FGrSetup(Grid)

   Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_CART) = "Cartola"
   Grid.TextMatrix(0, C_FDESDE) = "Desde"
   Grid.TextMatrix(0, C_FHASTA) = "Hasta"
   Grid.TextMatrix(0, C_INICIAL) = "Inicial"
   Grid.TextMatrix(0, C_CARGOS) = "Cargos"
   Grid.TextMatrix(0, C_ABONOS) = "Abonos"
   Grid.TextMatrix(0, C_SALDO) = "Saldo"

   Grid.ColWidth(C_ANO) = 600
   Grid.ColWidth(C_CART) = 900
   Grid.ColWidth(C_FDESDE) = 1000
   Grid.ColWidth(C_FHASTA) = 1000
   Grid.ColWidth(C_INICIAL) = W_MONTO
   Grid.ColWidth(C_ABONOS) = W_MONTO
   Grid.ColWidth(C_CARGOS) = W_MONTO
   Grid.ColWidth(C_SALDO) = W_MONTO
   Grid.ColWidth(C_IDCART) = 0

   Grid.ColAlignment(C_ANO) = flexAlignCenterCenter
   Grid.ColAlignment(C_CART) = flexAlignCenterCenter
   Grid.ColAlignment(C_FDESDE) = flexAlignRightCenter
   Grid.ColAlignment(C_FHASTA) = flexAlignRightCenter
   Grid.ColAlignment(C_INICIAL) = flexAlignRightCenter
   Grid.ColAlignment(C_ABONOS) = flexAlignRightCenter
   Grid.ColAlignment(C_CARGOS) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter

End Sub

Private Sub Grid_DblClick()

   Call PostClick(Bt_Det)
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call PostClick(Bt_CopyExcel)
   End If
   
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
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Printer.Orientation = OldOrientation
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption & " - " & Cb_CartBanco)
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_ABONOS, C_CARGOS)
   
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
   Dim Encabezados(0) As String
   Dim FontTit(0) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer, e As Integer
   
   Set gPrtReportes.Grid = Grid
   
   Printer.Orientation = ORIENT_VER
   
   Titulos(0) = Caption
   e = ItemData(Cb_Estado)
   If e = 1 Then
      Titulos(1) = "Movimientos Conciliados"
   ElseIf e = 0 Then
      Titulos(1) = "Movimientos No Conciliados"
   Else
      Titulos(1) = "Movimientos"
   End If
   
   Encabezados(0) = "Banco: " & vbTab & Cb_CartBanco
   
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
   Next i
            
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDCART
   
   
End Sub


