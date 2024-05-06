VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmDetBaseImponible14DFull 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Imponible Primera Categoría Reg 14 D"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_PerdidaAnoAnterior 
      Caption         =   "Obtener Perdida Año Anterior"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Tx_Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4380
      Width           =   1695
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   3375
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
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
      TabIndex        =   0
      Top             =   0
      Width           =   9075
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
         Left            =   120
         Picture         =   "FrmDetBaseImponible14DFull.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminar registro seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6600
         TabIndex        =   9
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
         Left            =   1200
         Picture         =   "FrmDetBaseImponible14DFull.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7740
         TabIndex        =   7
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
         Left            =   780
         Picture         =   "FrmDetBaseImponible14DFull.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   1620
         Picture         =   "FrmDetBaseImponible14DFull.frx":0D5D
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
         Left            =   3540
         Picture         =   "FrmDetBaseImponible14DFull.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2700
         Picture         =   "FrmDetBaseImponible14DFull.frx":15CB
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   3120
         Picture         =   "FrmDetBaseImponible14DFull.frx":1969
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   2160
         Picture         =   "FrmDetBaseImponible14DFull.frx":1CCA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Total:"
      Height          =   195
      Left            =   5700
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "FrmDetBaseImponible14DFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_CODIGO = 1
Const C_FECHA = 2
Const C_LNGFECHA = 3
Const C_DESCRIP = 4
Const C_VALOR = 5
Const C_UPDATE = 6

Const NCOLS = C_UPDATE

Dim lTipo As Integer
Dim lCodigo As Integer
Dim lDescrip As String
Dim lRc As Integer

Public Function FEdit(ByVal Tipo As Integer, ByVal Codigo As Integer, ByVal Descrip As String, Valor As Double) As Integer

   lTipo = Tipo
   lCodigo = Codigo
   lDescrip = Trim(Descrip)
   
   '2699582

        If lCodigo = 9600 And gEmpresa.Ano >= 2022 Then
        Bt_PerdidaAnoAnterior.visible = True
        
        Else
         Bt_PerdidaAnoAnterior.visible = False
        End If
   'fin 2699582
   
   Me.Show vbModal
   
   Valor = 0
   
   If lRc = vbOK Then
      Valor = vFmt(Tx_Total)
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
   
      lRc = vbOK
      
      Unload Me
   End If
   
End Sub

'2699582
Private Sub Bt_PerdidaAnoAnterior_Click()
  TraerPerdidaAnterior (lTipo)
  LoadAll
End Sub
'fin 2699582

Private Sub Form_Load()
   
   Call SetUpGrid
   
   Me.Caption = lDescrip
   Grid.TextMatrix(Grid.FixedRows, C_DESCRIP) = lDescrip
   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double
      
   Q1 = "SELECT IdBaseImponible14D, Fecha, Valor "
   Q1 = Q1 & " FROM BaseImponible14D "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND Codigo = " & lCodigo & " AND Nivel = " & BIMP14D_MAXNIV + 1
   Q1 = Q1 & " ORDER BY Fecha, IdBaseImponible14D "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.FlxGrid.Redraw = False
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_ID) = vFld(Rs("IdBaseImponible14D"))
      Grid.TextMatrix(i, C_CODIGO) = lCodigo
      Grid.TextMatrix(i, C_LNGFECHA) = vFld(Rs("Fecha"))
      Grid.TextMatrix(i, C_FECHA) = IIf(vFld(Rs("Fecha")) > 0, Format(vFld(Rs("Fecha")), EDATEFMT), "")
      Grid.TextMatrix(i, C_DESCRIP) = lDescrip
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("Valor")), NUMFMT)
      
      Total = Total + vFld(Rs("Valor"))
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.Row = Grid.FixedRows
   Grid.Col = C_FECHA
   
   Tx_Total = Format(Total, NUMFMT)
   
   Grid.FlxGrid.Redraw = True
   
End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
      
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
      
         Q1 = "INSERT INTO BaseImponible14D (IdEmpresa, Ano, Tipo, Nivel, Codigo, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & lTipo
         Q1 = Q1 & ", " & BIMP14D_MAXNIV + 1 & ", " & lCodigo & ", " & Grid.TextMatrix(i, C_LNGFECHA) & ", " & vFmt(Grid.TextMatrix(i, C_VALOR)) & ")"
   
         Call ExecSQL(DbMain, Q1)
   
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then

         Q1 = "UPDATE BaseImponible14D SET "
         Q1 = Q1 & " Fecha = " & Grid.TextMatrix(i, C_LNGFECHA) & ", Valor = " & vFmt(Grid.TextMatrix(i, C_VALOR))
         Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND IdBaseImponible14D = " & Grid.TextMatrix(i, C_ID)
         
         Call ExecSQL(DbMain, Q1)
  
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then

         Wh = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND IdBaseImponible14D = " & Grid.TextMatrix(i, C_ID)
         Call DeleteSQL(DbMain, "BaseImponible14D", Wh)
  
      End If
      
   Next i
   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un registro a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.RowHeight(Row) = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar este registro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_ID, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot
End Sub

Private Function Valida() As Boolean
   Dim i As Integer, j As Integer

   Valida = False
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_LNGFECHA)) > 0 And Year(Val(Grid.TextMatrix(i, C_LNGFECHA))) <> gEmpresa.Ano Then
         Grid.Row = i
         Grid.Col = C_FECHA
         Grid.SetFocus
         MsgBox1 "Línea " & i & ": Fecha no pertenece al año actual.", vbExclamation
         Exit Function
      End If
      
      For j = Grid.FixedRows To Grid.rows - 1
         If i <> j And Val(Grid.TextMatrix(i, C_LNGFECHA)) > 0 And Val(Grid.TextMatrix(i, C_LNGFECHA)) = Val(Grid.TextMatrix(j, C_LNGFECHA)) Then
            Grid.Row = j
            Grid.Col = C_FECHA
            Grid.SetFocus
            MsgBox1 "Línea " & j & ": Fecha repetida.", vbExclamation
            Exit Function
         End If
      Next j
   Next i
   
   Valida = True

End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_CODIGO) = 0
   Grid.ColWidth(C_LNGFECHA) = 0
   Grid.ColWidth(C_FECHA) = 1200
   Grid.ColWidth(C_DESCRIP) = 5100
   Grid.ColWidth(C_VALOR) = 1500
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(0, C_VALOR) = "Valor Nominal"
   
      
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
   Call LP_FGr2Clip(Grid, Me.Caption & vbTab & "Año " & gEmpresa.Ano)

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
   
   Titulos(0) = Left(Me.Caption, 70)
   
   gPrtReportes.Titulos = Titulos
'   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   Total(C_FECHA) = "Total"
   Total(C_VALOR) = Tx_Total
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total

   gPrtReportes.ColObligatoria = C_FECHA
   gPrtReportes.NTotLines = 1

End Sub


Private Sub CalcTot()
   Dim i As Integer
   Dim Total As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_FECHA) = "" Then
         Exit For
      End If
      
      Total = Total + vFmt(Grid.TextMatrix(i, C_VALOR))
      
   Next i
   
   Tx_Total = Format(Total, NUMFMT)
      
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Fecha As Long
   
   Action = vbOK

   If Col = C_FECHA Then
      If Value <> "" Then
         Fecha = GetDate(Value, "dmy")
         
         If Year(Fecha) <> gEmpresa.Ano Then
            MsgBox1 "Esta fecha no pertenece al año actual.", vbExclamation
            Action = vbRetry
            Exit Sub
         End If
         
         Value = Format(Fecha, EDATEFMT)
         Grid.TextMatrix(Row, Col) = Value
         Grid.TextMatrix(Row, C_LNGFECHA) = Fecha
         
         Grid.TextMatrix(Row, C_DESCRIP) = lDescrip
         
         Grid.rows = Grid.rows + 1
      Else
         Grid.TextMatrix(Row, C_LNGFECHA) = 0
         Grid.TextMatrix(Row, C_DESCRIP) = ""
      
      End If
   
   ElseIf Col = C_VALOR Then
      Value = Format(vFmt(Value), NUMFMT)
      Grid.TextMatrix(Row, Col) = Value
      Call CalcTot
   End If

   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   End If


End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Row > Grid.FixedRows And (Val(Grid.TextMatrix(Row - 1, C_LNGFECHA)) = 0 Or vFmt(Grid.TextMatrix(Row - 1, C_VALOR)) = 0) Then
      MsgBox1 "Debe completar la línea anterior.", vbExclamation
      Exit Sub
   End If
   
   If Col = C_FECHA Or Col = C_VALOR Then
      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 12
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   If Grid.Col = C_FECHA Then
      Call KeyDate(KeyAscii)
   Else
      Call KeyNumPos(KeyAscii)
   End If

End Sub
