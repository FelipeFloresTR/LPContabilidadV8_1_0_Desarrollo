VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmDetCapPropioSimplAcum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Imponible Primera Categoría Acumuladas"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4048
      Cols            =   2
      Rows            =   3
      FixedCols       =   1
      FixedRows       =   1
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
      TabIndex        =   11
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   5340
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   6480
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":04BA
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":0961
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":0DA6
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":11CF
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":156D
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
         Picture         =   "FrmDetCapPropioSimplAcum.frx":18CE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Total:"
      Height          =   195
      Left            =   4620
      TabIndex        =   12
      Top             =   3420
      Width           =   855
   End
End
Attribute VB_Name = "FrmDetCapPropioSimplAcum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_ANO = 1
Const C_VALOR = 2
Const C_INGRESOMANUAL = 3
Const C_UPDATE = 4

Const NCOLS = C_UPDATE

Dim lRc As Integer
Dim lTipoDetCapPropioSimpl As Integer
Dim lValorAno As Double

Public Function FEdit(ByVal TipoDetCapPropioSimpl As Integer, ValorAno As Double) As Integer

   lTipoDetCapPropioSimpl = TipoDetCapPropioSimpl
   
   Me.Show vbModal
   
   ValorAno = lValorAno
   
   FEdit = lRc
   
End Function


Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   If Valida() Then
      Call SaveAll
   
      lRc = vbOK
      
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()

   Me.Caption = "Acumulado Anual " & gTipoDetCapPropioSimpl(lTipoDetCapPropioSimpl)
   Call SetUpGrid
   
   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double
   Dim Row As Integer

   Grid.rows = Grid.FixedRows
   Row = Grid.rows
   
   If gEmpresa.Ano < 2017 Then
      Exit Sub
   End If
   
   Grid.FlxGrid.Redraw = False
   
   For i = 2017 To gEmpresa.Ano
   
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(Row, C_ANO) = i
      
      Row = Row + 1
      
   Next i
   
   'llenamos los años
   Row = Row - 1
   
   Q1 = "SELECT IdCapPropioSimplAnual, AnoValor, Valor, IngresoManual "
   Q1 = Q1 & " FROM CapPropioSimplAnual "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND TipoDetCPS = " & lTipoDetCapPropioSimpl
   Q1 = Q1 & " ORDER BY AnoValor"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = Grid.FixedRows
   
   Do While Not Rs.EOF
   
      If Grid.TextMatrix(Row, C_ANO) = vFld(Rs("AnoValor")) Then
      
         Grid.TextMatrix(Row, C_VALOR) = Format(vFld(Rs("Valor")), NUMFMT)
         Grid.TextMatrix(Row, C_ID) = Format(vFld(Rs("IdCapPropioSimplAnual")), NUMFMT)
         Grid.TextMatrix(Row, C_INGRESOMANUAL) = vFld(Rs("IngresoManual"))
         
         Row = Row + 1
         
         If Row >= Grid.rows Then
            Exit Do
         End If
         
         Rs.MoveNext
      
      ElseIf Grid.TextMatrix(Row, C_ANO) < vFld(Rs("AnoValor")) Then
         Row = Row + 1
      
      Else
         Rs.MoveNext
         
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   Call CalcTot
      
   Grid.Row = Grid.FixedRows
   Grid.Col = C_ANO
   Call FGrVRows(Grid)
   
   Grid.FlxGrid.Redraw = True
   
End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   

   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_ANO) = "" Then
         Exit For
      End If
         
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
         Q1 = "INSERT INTO CapPropioSimplAnual (TipoDetCPS, IngresoManual, AnoValor, Valor, IdEmpresa ) "
         Q1 = Q1 & " VALUES ("
         Q1 = Q1 & lTipoDetCapPropioSimpl
         Q1 = Q1 & ",1"
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_ANO))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_VALOR))
         Q1 = Q1 & "," & gEmpresa.id & ") "
         
         Call ExecSQL(DbMain, Q1)
      
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
         Q1 = "UPDATE CapPropioSimplAnual SET "
         Q1 = Q1 & "  Valor = " & vFmt(Grid.TextMatrix(i, C_VALOR))
         Q1 = Q1 & ", IngresoManual = 1"
         
         Q1 = Q1 & " WHERE IdCapPropioSimplAnual =" & Grid.TextMatrix(i, C_ID)
         Q1 = Q1 & " AND TipoDetCPS = " & lTipoDetCapPropioSimpl & " AND AnoValor = " & vFmt(Grid.TextMatrix(i, C_ANO))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         
         Call ExecSQL(DbMain, Q1)
      
'      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
'
''         Q1 = "DELETE * FROM Socios "
'         Q1 = " WHERE IdSocio = " & Grid.TextMatrix(i, C_IDSOCIO)
'         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
''         Call ExecSQL(DbMain, Q1)
'
'         Call DeleteSQL(DbMain, "Socios", Q1)
      End If
      
      If Grid.TextMatrix(i, C_ANO) = gEmpresa.Ano Then
         lValorAno = vFmt(Grid.TextMatrix(i, C_VALOR))
      End If
         
   Next i

End Sub

Private Function Valida() As Boolean

   Valida = True
   
End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ANO) = 4900
   Grid.ColWidth(C_VALOR) = 1500
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_INGRESOMANUAL) = 0
   
   Grid.ColAlignment(C_ANO) = flexAlignCenterCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call FGrLocateCntrl(Grid, Tx_Total, C_VALOR)
   
      
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
                  
   Total(C_ANO) = "Total"
   Total(C_VALOR) = Tx_Total
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total

   gPrtReportes.ColObligatoria = C_ANO
   gPrtReportes.NTotLines = 1

End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   Action = vbOK

   If Col = C_VALOR Then
      If Value <> "" Then
         Value = Format(vFmt(Value), NUMFMT)
         Grid.TextMatrix(Row, Col) = Value
      End If
   End If
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
      Call CalcTot
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Grid.TextMatrix(Row, C_ANO) = "" Or Grid.TextMatrix(Row, C_INGRESOMANUAL) = "0" Then
      Exit Sub
   End If
   
   If Col = C_VALOR Then
      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 12
   End If
         
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   Call KeyNum(KeyAscii)

End Sub

Private Sub CalcTot()
   Dim i As Integer
   Dim Total As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_ANO) = "" Then
         Exit For
      End If
      
      Total = Total + vFmt(Grid.TextMatrix(i, C_VALOR))
      
   Next i
   
   Tx_Total = Format(Total, NUMFMT)
      
End Sub
