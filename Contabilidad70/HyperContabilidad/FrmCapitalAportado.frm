VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmCapitalAportado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capital Aportado"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   9975
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
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   180
         Width           =   1035
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
         Picture         =   "FrmCapitalAportado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   9420
         TabIndex        =   9
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
         Left            =   120
         Picture         =   "FrmCapitalAportado.frx":04BA
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "FrmCapitalAportado.frx":0961
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "FrmCapitalAportado.frx":0DA6
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "FrmCapitalAportado.frx":11CF
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "FrmCapitalAportado.frx":156D
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "FrmCapitalAportado.frx":18CE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6660
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmCapitalAportado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDSOCIO = 0
Const C_RUT = 1
Const C_NOMBRE = 2
Const C_MONTOPAGADO = 3
Const C_MONTOINGRESADOUSR = 4
Const C_MONTOATRASPASAR = 5
Const C_UPDATE = 6

Const NCOLS = C_UPDATE

Dim lRc As Integer
Dim lValor As Double

Public Function FEdit(CapitalAportado As Double) As Integer

   Me.Show vbModal
   
   If lRc = vbOK Then
      CapitalAportado = lValor
   End If
   
   FEdit = lRc
   
End Function


Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   If Valida() Then
      Call SaveAll
      
      lValor = vFmt(GridTot.TextMatrix(0, C_MONTOATRASPASAR))

      lRc = vbOK
      
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   
   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double

   Q1 = "SELECT IdSocio, RUT, Nombre, MontoPagado, MontoIngresadoUsuario, MontoATraspasar "
   Q1 = Q1 & " FROM Socios "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Nombre"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDSOCIO) = vFld(Rs("IdSocio"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_MONTOPAGADO) = Format(vFld(Rs("MontoPagado")), NUMFMT)
      Grid.TextMatrix(i, C_MONTOINGRESADOUSR) = Format(vFld(Rs("MontoIngresadoUsuario")), NUMFMT)
      Grid.TextMatrix(i, C_MONTOATRASPASAR) = IIf(vFld(Rs("MontoATraspasar")) <> 0, Format(vFld(Rs("MontoATraspasar")), NUMFMT), Grid.TextMatrix(i, C_MONTOPAGADO))
      
      
      Total = Total + vFmt(Grid.TextMatrix(i, C_MONTOATRASPASAR))
      
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_RUT
   
   Grid.FlxGrid.Redraw = True
   
   GridTot.TextMatrix(0, C_MONTOATRASPASAR) = Format(Total, NUMFMT)

End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   
   For i = Grid.FixedRows To Grid.rows - 1

      If Val(Grid.TextMatrix(i, C_IDSOCIO)) = 0 Then
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
         Q1 = "UPDATE Socios SET "
         Q1 = Q1 & " MontoIngresadoUsuario = " & vFmt(Grid.TextMatrix(i, C_MONTOINGRESADOUSR))
         Q1 = Q1 & ", MontoATraspasar = " & vFmt(Grid.TextMatrix(i, C_MONTOATRASPASAR))
         Q1 = Q1 & " WHERE IdSocio = " & Grid.TextMatrix(i, C_IDSOCIO)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      End If
      
   Next i

   Q1 = "UPDATE EmpresasAno SET "
   Q1 = Q1 & " CPS_CapitalAportado = " & vFmt(GridTot.TextMatrix(0, C_MONTOATRASPASAR))
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   
End Sub

Private Function Valida() As Boolean

   Valida = True
   
End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid)

   Grid.ColWidth(C_IDSOCIO) = 0
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_NOMBRE) = 4400
   Grid.ColWidth(C_MONTOPAGADO) = 1500
   Grid.ColWidth(C_MONTOINGRESADOUSR) = 1500
   Grid.ColWidth(C_MONTOATRASPASAR) = 1500
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOPAGADO) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOINGRESADOUSR) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOATRASPASAR) = flexAlignRightCenter
   
   Grid.TextMatrix(1, C_RUT) = "RUT"
   Grid.TextMatrix(1, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_MONTOPAGADO) = "Monto"
   Grid.TextMatrix(1, C_MONTOPAGADO) = "Pagado"
   Grid.TextMatrix(0, C_MONTOINGRESADOUSR) = "Monto"
   Grid.TextMatrix(1, C_MONTOINGRESADOUSR) = "Ingresado Usuario"
   Grid.TextMatrix(0, C_MONTOATRASPASAR) = "Monto"
   Grid.TextMatrix(1, C_MONTOATRASPASAR) = "a Traspasar"
   
   Call FGrVRows(Grid)
   Call FGrTotales(Grid, GridTot)
      
   GridTot.TextMatrix(0, C_NOMBRE) = "Total"
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   Call ResetPrtBas(gPrtReportes)
   

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
   Call LP_FGr2Clip(Grid, Me.Caption)

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

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
'   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   Total(C_NOMBRE) = GridTot.TextMatrix(0, C_NOMBRE)
   Total(C_MONTOATRASPASAR) = GridTot.TextMatrix(0, C_MONTOATRASPASAR)
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDSOCIO
   gPrtReportes.NTotLines = 1

End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   Action = vbOK

   If Col = C_MONTOINGRESADOUSR Then
      Value = Format(vFmt(Value), NUMFMT)
      Grid.TextMatrix(Row, Col) = Value
      
      If vFmt(Value) <> 0 Then
         Grid.TextMatrix(Row, C_MONTOATRASPASAR) = Value
      Else
         Grid.TextMatrix(Row, C_MONTOATRASPASAR) = Grid.TextMatrix(Row, C_MONTOPAGADO)
      End If
      
      Call CalcTot
   End If
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDSOCIO, C_UPDATE)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Grid.TextMatrix(Row, C_IDSOCIO) = "" Then
      Exit Sub
   End If
      
   If Col = C_MONTOINGRESADOUSR Then
      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 12
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub CalcTot()
   Dim Total As Double
   Dim i As Integer

   For i = Grid.FixedRows To Grid.rows
      If Val(Grid.TextMatrix(i, C_IDSOCIO)) = 0 Then
         Exit For
      End If
      
      Total = Total + vFmt(Grid.TextMatrix(i, C_MONTOATRASPASAR))
   Next i
   
   GridTot.TextMatrix(0, C_MONTOATRASPASAR) = Format(Total, NUMFMT)
   
End Sub
