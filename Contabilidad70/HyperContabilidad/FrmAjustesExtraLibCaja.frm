VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmAjustesExtraLibCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste Extra Libro de Caja"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12855
   StartUpPosition =   1  'CenterOwner
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   8055
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   14208
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   0
      ScrollBars      =   2
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12795
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":00A4
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":0405
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Convertir moneda"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_DetDoc 
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":07A3
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   10020
         TabIndex        =   6
         Top             =   180
         Width           =   1275
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":1031
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":1476
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   11340
         TabIndex        =   2
         Top             =   180
         Width           =   1275
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
         Picture         =   "FrmAjustesExtraLibCaja.frx":191D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAjustesExtraLibCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_TIPOAJUSTE = 1
Const C_IDITEM = 2
Const C_CONCEPTO = 3
Const C_VALOR = 4
Const C_SUBTOTAL = 5
Const C_TIPOING = 6
Const C_FMT = 7
Const C_COLOBLIGATORIA = 8
Const C_UPD = 9

Const NCOLS = C_UPD

Dim lRow_SaldoFinal As Integer
Dim lRow_Agregados As Integer
Dim lRow_Deducciones As Integer
Dim lRow_BaseImponible As Integer
Dim lRow_GastosPresuntos As Integer
Dim lRow_IngFEmision As Integer
Dim lRow_IngFExigibilidad As Integer

Dim lValUTM As Double

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub


Private Sub Bt_DetDoc_Click()
   Dim Row As Integer
   Dim Frm As Form
   
   Row = Grid.Row
   
   If Row <= lRow_SaldoFinal Then
      Me.MousePointer = vbHourglass
      Set Frm = New FrmLibCaja
      Call Frm.FView(-1)
      Set Frm = Nothing
      Me.MousePointer = vbDefault
      Exit Sub
      
   
   ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "33 bis") > 0 Then
      Me.MousePointer = vbHourglass
      Set Frm = New FrmRepActivoFijo
      Call Frm.FView
      Set Frm = Nothing
      Me.MousePointer = vbDefault
      Exit Sub
   
   End If
   
End Sub

Private Sub Bt_OK_Click()

   If Valida() Then
      SaveAll
      Unload Me
   End If
   
End Sub

Public Sub FEdit()
   Me.Show vbModal
End Sub
Private Sub Form_Load()
   lOrientacion = ORIENT_VER

   Call SetUpGrid
   Call LoadAll
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_TIPOAJUSTE) = 0
   Grid.ColWidth(C_IDITEM) = 0
   Grid.ColWidth(C_CONCEPTO) = 9000
   Grid.ColWidth(C_VALOR) = 1600
   Grid.ColWidth(C_SUBTOTAL) = 1600
   Grid.ColWidth(C_TIPOING) = 0
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
   Grid.ColWidth(C_UPD) = 0
   
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_SUBTOTAL) = flexAlignRightCenter
   
   Grid.FlxGrid.BackColor = vbButtonFace
   
   Call FGrVRows(Grid)

End Sub
Private Sub LoadAll()
   Dim i As Integer, Row As Integer
   Dim TipoIngreso As Integer
   Dim Valor As Double
   Dim FmtLine As String
   Dim id As Long
   Dim SaldoIngresos As Double, SaldoEgresos As Double
   
   lValUTM = GetValUTM()
   
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = 0
   
   Call CalcSaldosIngEgr(SaldoIngresos, SaldoEgresos)
   
   Row = 0
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Saldo Libro de Caja - Ingresos"
   Grid.TextMatrix(Row, C_VALOR) = Format(SaldoIngresos, NEGNUMFMT)
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Saldo Libro de Caja - Egresos"
   Grid.TextMatrix(Row, C_VALOR) = Format(SaldoEgresos, NEGNUMFMT)
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Saldo Final Libro de Caja"
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   lRow_SaldoFinal = Row
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Agregados"
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   lRow_Agregados = Row

   FmtLine = "L"
   For i = 1 To MAX_ITEMAJUSTESEC
      If gAjustesExtraCont(TAEC_AGREGADOS, i).Nombre = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_TIPOAJUSTE) = TAEC_AGREGADOS
      Grid.TextMatrix(Row, C_CONCEPTO) = gAjustesExtraCont(TAEC_AGREGADOS, i).Nombre
      
      TipoIngreso = gAjustesExtraCont(TAEC_AGREGADOS, i).TipoIngresoAjuste
      Grid.TextMatrix(Row, C_TIPOING) = TipoIngreso
      
      Grid.TextMatrix(Row, C_VALOR) = Format(GetStoredVal(TAEC_AGREGADOS, i, id), NEGNUMFMT)
      Grid.TextMatrix(Row, C_ID) = id
      If TipoIngreso = TIA_INGDIRECTO Then
         Call FGrBackColor(Grid, Row, C_VALOR, vbWhite)
      ElseIf TipoIngreso = TIA_CTASASOCIADAS Then
         Valor = LoadValCuentasAjustes(TAEC_AGREGADOS, gAjustesExtraCont(TAEC_AGREGADOS, i).IdItemCtasAsociadas)
         Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NEGNUMFMT)
      End If
      
      'H: Crédito art 33 Bis
      If InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "33 bis") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetVal33Bis(), NEGNUMFMT)
      
      'K: Ingresos devengados y no percibidos de Empresas Relacionadas al cierre del ejercicio
      ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "relacionadas") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetIngNoPercibRel(), NEGNUMFMT)
      
      'L: Ingresos devengados y no percibidos con  plazo mayor a  12 meses desde que  se emitió dcto.
      ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "emitió") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetIngNoPercib12MesesEmiNoRel(), NEGNUMFMT)
         lRow_IngFEmision = Row
      
      'M: Ingresos devengados y no percibidos con plazo mayor a 12 meses desde que pago es  exigible
      ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "exigible") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetIngNoPercib12MesesExigNoRel(), NEGNUMFMT)
         lRow_IngFExigibilidad = Row
      
      'Ñ: Pago de gastos adeudados antes de ingresar al régimen art 14 ter
      ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "gastos adeudados") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetIngEgrDevengados(TOPERCAJA_EGRESO), NEGNUMFMT)
     
      End If
      
      Grid.TextMatrix(Row, C_FMT) = FmtLine
      FmtLine = ""
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
      
   Next i
      
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Deducciones"
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   lRow_Deducciones = Row
   
   FmtLine = "L"
   For i = 1 To MAX_ITEMAJUSTESEC
      If gAjustesExtraCont(TAEC_DEDUCCIONES, i).Nombre = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_TIPOAJUSTE) = TAEC_DEDUCCIONES
      Grid.TextMatrix(Row, C_CONCEPTO) = gAjustesExtraCont(TAEC_DEDUCCIONES, i).Nombre
      TipoIngreso = gAjustesExtraCont(TAEC_DEDUCCIONES, i).TipoIngresoAjuste
      Grid.TextMatrix(Row, C_TIPOING) = TipoIngreso
      
      Grid.TextMatrix(Row, C_VALOR) = Format(GetStoredVal(TAEC_DEDUCCIONES, i, id), NEGNUMFMT)
      Grid.TextMatrix(Row, C_ID) = id
      
      If TipoIngreso = TIA_INGDIRECTO Then
         Call FGrBackColor(Grid, Row, C_VALOR, vbWhite)
      ElseIf TipoIngreso = TIA_CTASASOCIADAS Then
         Valor = LoadValCuentasAjustes(TAEC_DEDUCCIONES, gAjustesExtraCont(TAEC_DEDUCCIONES, i).IdItemCtasAsociadas)
         Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NEGNUMFMT)
      End If
      
      'R: Gastos presuntos equivalente al 0,5% de los ingresos brutos (Min UTM, Max 15 UTM)
      If InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "presuntos") > 0 Then
         lRow_GastosPresuntos = Row
       
      'W: Percepción de ingresos devengados antes de ingresar al regimen art 14 ter
      ElseIf InStr(LCase(Grid.TextMatrix(Row, C_CONCEPTO)), "ingresos devengados antes") > 0 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(GetIngEgrDevengados(TOPERCAJA_INGRESO), NEGNUMFMT)
      
      End If
      
      Grid.TextMatrix(Row, C_FMT) = FmtLine
      FmtLine = ""
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   Next i
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_CONCEPTO) = "Base Imponible afecta a IDPC"
   Grid.TextMatrix(Row, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, Row, "B")
   Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   lRow_BaseImponible = Row
   
   Call CalcTot
   
   Grid.TopRow = 0
   
   Grid.FlxGrid.Redraw = True

End Sub

Private Sub CalcSaldosIngEgr(SaldoIngresos As Double, SaldoEgresos As Double)
   Dim Q1 As String
   Dim Rs As Recordset

   SaldoIngresos = 0
   SaldoEgresos = 0
   
   'primero los ingresos
   Q1 = "SELECT Sum(Pagado) as TotIngreso"
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE (FechaIngresoLibro BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND LibroCaja.TipoOper = " & TOPERCAJA_INGRESO
'   Q1 = Q1 & " AND (LibroCaja.TipoLib <> " & LIB_CAJAING & " OR ( LibroCaja.TipoLib = " & LIB_CAJAING & " AND LibroCaja.TipoDoc <> " & LIBCAJA_OTROSINGINI & "))"   'Solicitado por Victor Morales el 5 de mar 2020
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
      
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      SaldoIngresos = vFld(Rs("TotIngreso"))
   End If
   
   Call CloseRs(Rs)
   
   'y ahora los egresos
   Q1 = "SELECT Sum(Pagado) as TotEgreso"
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE (FechaIngresoLibro BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND LibroCaja.TipoOper = " & TOPERCAJA_EGRESO
'   Q1 = Q1 & " AND (LibroCaja.TipoLib <> " & LIB_CAJAEGR & " OR ( LibroCaja.TipoLib = " & LIB_CAJAEGR & " AND LibroCaja.TipoDoc <> " & LIBCAJA_OTROSEGRINI & "))"     'Solicitado por Victor Morales el 5 de mar 2020
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      SaldoEgresos = vFld(Rs("TotEgreso"))
   End If
   
   Call CloseRs(Rs)
  

'      If Val(Grid.TextMatrix(Row, C_IDTIPOOPER)) = TOPERCAJA_INGRESO Then
'         If Val(Grid.TextMatrix(Row, C_ESREBAJA)) <> 0 Then
'            Grid.TextMatrix(Row, C_INGRESO) = ""
'            Grid.TextMatrix(Row, C_EGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'         Else
'            Grid.TextMatrix(Row, C_INGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'            Grid.TextMatrix(Row, C_EGRESO) = ""
'         End If
'      Else
'         If Val(Grid.TextMatrix(Row, C_ESREBAJA)) <> 0 Then
'            Grid.TextMatrix(Row, C_EGRESO) = ""
'            Grid.TextMatrix(Row, C_INGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'         Else
'            Grid.TextMatrix(Row, C_EGRESO) = Grid.TextMatrix(Row, C_PAGADO)
'            Grid.TextMatrix(Row, C_INGRESO) = ""
'         End If
'      End If

End Sub
Private Function GetStoredVal(ByVal TipoAjuste As Integer, ByVal IdItemAjuste As Integer, id As Long) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   
   id = 0
   GetStoredVal = 0
   
   If TipoAjuste = 0 Or IdItemAjuste = 0 Then
      Exit Function
   End If

   Q1 = "SELECT IdAjustesExtLibCaja, Valor FROM AjustesExtLibCaja "
   Q1 = Q1 & " WHERE TipoAjuste = " & TipoAjuste & " AND IdItemAjuste = " & IdItemAjuste
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetStoredVal = vFld(Rs("Valor"))
      id = vFld(Rs("IdAjustesExtLibCaja"))
   End If
   
   Call CloseRs(Rs)

End Function
'K: Ingresos devengados y no percibidos de Empresas Relacionadas al cierre del ejercicio
Private Function GetIngNoPercibRel() As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetIngNoPercibRel = 0
   
   Q1 = "SELECT Sum(SaldoDoc) as Saldo "
   Q1 = Q1 & " FROM Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad AND Documento.IdEmpresa = Entidades.IdEmpresa "
   
   Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS
   Q1 = Q1 & " AND (FEmision BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND NOT  Entidades.EntRelacionada IS NULL AND Entidades.EntRelacionada <> 0 "       'relacionada
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      GetIngNoPercibRel = vFld(Rs("Saldo"))
   End If
   
   Call CloseRs(Rs)
   
End Function
'L: Ingresos devengados y no percibidos con  plazo mayor a  12 meses desde que  se emitió dcto.
Private Function GetIngNoPercib12MesesEmiNoRel() As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetIngNoPercib12MesesEmiNoRel = 0
   
   Q1 = "SELECT Sum(SaldoDoc) as Saldo"
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad AND Documento.IdEmpresa = Entidades.IdEmpresa ) "
   Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS
   Q1 = Q1 & " AND (NumCuotas IS NULL OR NumCuotas = 0) "    'al contado
   Q1 = Q1 & " AND (Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0)"       'no relacionada
'   Q1 = Q1 & " AND abs(" & CLng(Now) & " - FEmisionOri ) >= 365 "
'   Q1 = Q1 & " AND DateDiff( 'm', FEmisionOri, " & CLng(Now) & " ) > 12 "
   If gDbType = SQL_SERVER Then
      Q1 = Q1 & " AND DateDiff( month, FEmisionOri, " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & " ) > 12 "   '10 ago 2018 Claudio Villegas
   Else   'access
      Q1 = Q1 & " AND DateDiff( 'm', FEmisionOri, " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & " ) > 12 "   '10 ago 2018 Claudio Villegas
   End If
'   Q1 = Q1 & " AND (FEmision BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      GetIngNoPercib12MesesEmiNoRel = vFld(Rs("Saldo"))
   End If
   
   Call CloseRs(Rs)
   
End Function
'M: Ingresos devengados y no percibidos con plazo mayor a 12 meses desde que pago es  exigible
Private Function GetIngNoPercib12MesesExigNoRel() As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetIngNoPercib12MesesExigNoRel = 0
   
   Q1 = "SELECT Sum(MontoCuota) as Saldo"
   Q1 = Q1 & " FROM (Documento "
   Q1 = Q1 & " INNER JOIN DocCuotas ON Documento.IdDoc = DocCuotas.IdDoc " & JoinEmpAno(gDbType, "DocCuotas", "Documento") & " )"                'a Crédito
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad " & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS
   Q1 = Q1 & " AND (Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0)"     'No relacionada
   Q1 = Q1 & " AND (FechaIngPercibido IS NULL or FechaIngPercibido = 0) "                 'no se ha pagado
'   Q1 = Q1 & " AND DateDiff( 'm', FEmisionOri, " & CLng(Now) & " ) > 12 "
   If gDbType = SQL_SERVER Then
      Q1 = Q1 & " AND DateDiff( month, FEmisionOri, " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & " ) > 12 "   '10 ago 2018 Claudio Villegas
   Else    'access
      Q1 = Q1 & " AND DateDiff( 'm', FEmisionOri, " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & " ) > 12 "   '10 ago 2018 Claudio Villegas
   End If
'   Q1 = Q1 & " AND (FEmision BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      GetIngNoPercib12MesesExigNoRel = vFld(Rs("Saldo"))
   End If
   
   Call CloseRs(Rs)
   
End Function
'Ñ: Pago de gastos adeudados antes de ingresar al régimen art 14 ter
'W: Percepción de ingresos devengados antes de ingresar al regimen art 14 ter
Private Function GetIngEgrDevengados(ByVal TipoOper As Integer) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetIngEgrDevengados = 0
   
   Q1 = "SELECT Sum(Pagado) as Saldo"
   Q1 = Q1 & " FROM LibroCaja "
   Q1 = Q1 & " WHERE TipoOper = " & TipoOper
   Q1 = Q1 & " AND NOT OperDevengada IS NULL AND OperDevengada <> 0"
   Q1 = Q1 & " AND (FechaIngresoLibro BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31)) & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      GetIngEgrDevengados = vFld(Rs("Saldo"))
   End If
   
   Call CloseRs(Rs)
   
End Function
Private Sub CalcTot()
   Dim Tot As Double
   Dim Row As Integer
   Dim i As Integer
   Dim Fecha As Long
   Dim MaxR As Double
   Dim MinR As Double
   
   
   'Saldo Final Libro Caja
   Row = 0
   
   Tot = vFmt(Grid.TextMatrix(lRow_SaldoFinal - 2, C_VALOR)) - vFmt(Grid.TextMatrix(lRow_SaldoFinal - 1, C_VALOR))    'Saldo Caja Ingresos - Saldo Caja Egresos
   Grid.TextMatrix(lRow_SaldoFinal, C_SUBTOTAL) = Format(Tot, NEGNUMFMT)
   
   'R: Gastos presuntos equivalente al 0,5% de los ingresos brutos (Min UTM, Max 15 UTM)
   Tot = vFmt(Grid.TextMatrix(lRow_SaldoFinal - 2, C_VALOR)) 'Saldo Caja-Ingresos
   Tot = Tot + vFmt(Grid.TextMatrix(lRow_IngFEmision, C_VALOR)) + vFmt(Grid.TextMatrix(lRow_IngFExigibilidad, C_VALOR)) 'Saldos Ing. No percibidos Emp No relacionadas
   Tot = Tot * 0.005    '5%
   
   'topamos en valor: mín: 1 UTM, máx: 15 UTM
      
   If lValUTM > 0 Then   'última UTM ingresada en el sistema que sea a lo más del 31 dic del año actual
      MaxR = vFmt(Format(15 * lValUTM, NUMFMT))   'redondeamos a 0 decimales
      MinR = vFmt(Format(lValUTM, NUMFMT))

      If Tot > MaxR Then
         Tot = MaxR
      ElseIf Tot < MinR Then
         Tot = MinR
      End If

   End If
   
   Grid.TextMatrix(lRow_GastosPresuntos, C_VALOR) = Format(Tot, NEGNUMFMT)
   
   'Agregados
   Tot = 0
   For i = lRow_Agregados + 1 To lRow_Deducciones - 1
      Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
   Next i
   Grid.TextMatrix(lRow_Agregados, C_SUBTOTAL) = Format(Tot, NEGNUMFMT)
   
   'Deducciones
   Tot = 0
   For i = lRow_Deducciones + 1 To lRow_BaseImponible - 1
      Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
   Next i
   Grid.TextMatrix(lRow_Deducciones, C_SUBTOTAL) = Format(Tot, NEGNUMFMT)
   
   Tot = vFmt(Grid.TextMatrix(lRow_SaldoFinal, C_SUBTOTAL)) + vFmt(Grid.TextMatrix(lRow_Agregados, C_SUBTOTAL)) - vFmt(Grid.TextMatrix(lRow_Deducciones, C_SUBTOTAL))
   
   'Base Imponible
   Grid.TextMatrix(lRow_BaseImponible, C_SUBTOTAL) = Format(Tot, NEGNUMFMT)
   
End Sub


Private Sub SaveAll()
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxId As Long
   
   For Row = Grid.FixedRows To Grid.rows - 1
   
'      If Val(Grid.TextMatrix(Row, C_TIPOING)) = TIA_INGDIRECTO And Grid.TextMatrix(Row, C_UPD) <> "" Then
      If Val(Grid.TextMatrix(Row, C_TIPOAJUSTE)) > 0 Then
         
         If Val(Grid.TextMatrix(Row, C_ID)) <> 0 Then
            Q1 = "UPDATE AjustesExtLibCaja SET Valor = " & vFmt(Grid.TextMatrix(Row, C_VALOR))
            Q1 = Q1 & " WHERE IdAjustesExtLibCaja = " & Grid.TextMatrix(Row, C_ID)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Else
            MaxId = 0
            Q1 = "SELECT Max(IdAjustesExtLibCaja) FROM AjustesExtLibCaja "
            Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            If Not Rs.EOF Then
               MaxId = vFld(Rs(0)) + 1
            End If
            Call CloseRs(Rs)
            
            Q1 = "INSERT INTO AjustesExtLibCaja (IdAjustesExtLibCaja, TipoAjuste, IdItemAjuste, Valor, IdEmpresa, Ano)"
            Q1 = Q1 & " VALUES(" & MaxId & ", " & Grid.TextMatrix(Row, C_TIPOAJUSTE) & "," & Grid.TextMatrix(Row, C_IDITEM) & "," & vFmt(Grid.TextMatrix(Row, C_VALOR)) & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
         End If
         
         Call ExecSQL(DbMain, Q1)
      
      End If
      
   Next Row

End Sub

Private Function Valida() As Boolean
   Valida = False
   Valida = True
End Function


Private Sub Form_Resize()
   
   Grid.Height = Me.Height - Grid.Top - 800

   If Grid.Height > 10930 Then
      Grid.Height = 10930
   End If
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim nFolio As Integer
   Dim Pag As Integer
   
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
      
'      nFolio = gPrtLibros.PrtFlexGrid(Printer)
      
'      If lPapelFoliado And Ch_LibOficial <> 0 Then
'         Call AppendLogImpreso(LIBOF_MAYOR, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
'      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, nFolio)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
   End If
      
   Call ResetPrtBas(gPrtLibros)
      
   
'   PrtOrient = Printer.Orientation
'
'   Call SetUpPrtGrid
'
'   Me.MousePointer = vbHourglass
'
'   Call gPrtReportes.PrtFlexGrid(Printer)
'
'   Me.MousePointer = vbDefault
'
'   Printer.Orientation = PrtOrient
'   gPrtReportes.FmtCol = -1
'   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   Dim Pag As Integer
      
   PrtOrient = Printer.Orientation
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   
   gPrtLibros.CallEndDoc = False
   
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   Call PrtPieBalance(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
   
   gPrtLibros.CallEndDoc = True
   Set Frm.PrtControl = Bt_Print
   
   Me.MousePointer = vbDefault
   
'   Me.MousePointer = vbHourglass
'   Call gPrtLibros.PrtFlexGrid(Frm)
'   Set Frm.PrtControl = Bt_Print
'   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
      
   Call ResetPrtBas(gPrtLibros)
   Printer.Orientation = PrtOrient
   gPrtLibros.FmtCol = -1
           
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(1) As String
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
      
   Titulos(0) = Me.Caption
            
   If lInfoPreliminar Then
      Titulos(1) = INFO_PRELIMINAR
   End If
            
   gPrtLibros.Titulos = Titulos
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i) * 0.9
   Next i
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.ColObligatoria = C_COLOBLIGATORIA
   gPrtLibros.FmtCol = C_FMT
   gPrtLibros.NTotLines = 0
      
   If lPapelFoliado <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
'   Printer.Orientation = ORIENT_VER
'   Set gPrtReportes.Grid = Grid
'
'   Titulos(0) = Me.Caption
'   gPrtReportes.Titulos = Titulos
'
'   For i = 0 To Grid.Cols - 1
'      ColWi(i) = Grid.ColWidth(i) * 0.9
'   Next i
'
'   gPrtReportes.ColWi = ColWi
'   gPrtReportes.ColObligatoria = C_COLOBLIGATORIA
'   gPrtReportes.FmtCol = C_FMT
'   gPrtReportes.NTotLines = 0
'
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = LP_FGr2String(Grid, Me.Caption)
   Clipboard.Clear
   Clipboard.SetText Clip
      
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   Value = Format(vFmt(Value), NEGNUMFMT)
   Grid.TextMatrix(Row, Col) = Value
   Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPD)
   Call CalcTot
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)

   If Val(Grid.TextMatrix(Row, C_TIPOING)) = TIA_INGDIRECTO And Col = C_VALOR Then
      EdType = FEG_Edit
   Else
      Call Bt_DetDoc_Click
      
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
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

