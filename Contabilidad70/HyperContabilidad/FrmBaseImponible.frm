VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmBaseImponible 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Imponible Primera Categoría 14 TER A)"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13035
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
         Left            =   480
         Picture         =   "FrmBaseImponible.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   11640
         TabIndex        =   9
         Top             =   180
         Width           =   1275
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
         Left            =   60
         Picture         =   "FrmBaseImponible.frx":04BA
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   900
         Picture         =   "FrmBaseImponible.frx":0961
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   2820
         Picture         =   "FrmBaseImponible.frx":0DA6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   10320
         TabIndex        =   5
         Top             =   180
         Width           =   1275
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
         Left            =   1980
         Picture         =   "FrmBaseImponible.frx":11CF
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2400
         Picture         =   "FrmBaseImponible.frx":156D
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1440
         Picture         =   "FrmBaseImponible.frx":18CE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7858
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
End
Attribute VB_Name = "FrmBaseImponible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_TIPOBASEIMP = 1
Const C_IDITEM = 2
Const C_CONCEPTO = 3
Const C_VALOR = 4
Const C_SUBTOTAL = 5
Const C_FMT = 6
Const C_COLOBLIGATORIA = 7
Const C_UPD = 8

Const NCOLS = C_UPD

Dim lRowIniIngresos As Integer
Dim lRowIniEgresos As Integer
Dim lRowBaseImponible As Integer
Dim lRowMayorValor As Integer

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Private Sub Bt_OK_Click()

   If Valida() Then
      SaveAll
      Unload Me
   End If
   
End Sub
Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Grid.rows = 18
   
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_TIPOBASEIMP) = 0
   Grid.ColWidth(C_IDITEM) = 0
   Grid.ColWidth(C_CONCEPTO) = 9900
   Grid.ColWidth(C_VALOR) = 1500
   Grid.ColWidth(C_SUBTOTAL) = 1500
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
   Grid.ColWidth(C_UPD) = 0
   
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_SUBTOTAL) = flexAlignRightCenter
   
   Grid.FlxGrid.BackColor = vbButtonFace
   
End Sub
Private Sub LoadAll()
   Dim i As Integer, Row As Integer
   Dim TipoIngreso As Integer
   Dim Valor As Double
   Dim FmtLine As String
   Dim id As Long
   Dim SaldoIngresos As Double, SaldoEgresos As Double
   Dim Q1 As String
   Dim Rs As Recordset
   
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = 0
   Row = -1
   
   FmtLine = "B"
   For i = 0 To MAX_ITEMBASEIMP
   
      If gBaseImp14Ter(BASEIMP_INGRESOS, i) = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_TIPOBASEIMP) = BASEIMP_INGRESOS
      Grid.TextMatrix(Row, C_CONCEPTO) = gBaseImp14Ter(BASEIMP_INGRESOS, i)
      
      'obtenemos el ID y el valor guardado
      Q1 = "SELECT IdBaseImponible14Ter, Valor FROM BaseImponible14Ter "
      Q1 = Q1 & " WHERE TipoBaseImp = " & Grid.TextMatrix(Row, C_TIPOBASEIMP) & " AND IdItemBaseImp = " & Grid.TextMatrix(Row, C_IDITEM)
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdBaseImponible14Ter"))
      End If
      Call CloseRs(Rs)
      
      Select Case i
                           
         Case 0      'Total de ingresos anuales percibidos en el ejercicio (y devengados en los casos que corresponda), a valor nominal
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
            lRowIniIngresos = Row
            
         Case 1      'Ingresos percibidos
            Valor = GetTotCta_CodF22_14Ter(628, "C") - GetValAjustesELC(TAEC_DEDUCCIONES, 2) - GetTotCta_CodF22_14Ter_NC(628, "D")
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 2      'Ingreso diferido imputado en el ejercicio
            Valor = GetValAjustesELC(TAEC_AGREGADOS, 16) + GetValAjustesELC(TAEC_AGREGADOS, 17) + GetValAjustesELC(TAEC_AGREGADOS, 18)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 3      'Ingresos devengados
            Valor = GetValAjustesELC(TAEC_AGREGADOS, 11) + GetValAjustesELC(TAEC_AGREGADOS, 12) + GetValAjustesELC(TAEC_AGREGADOS, 13)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 4      'Participaciones e intereses percibidos
'            Valor = GetTotCta_CodF22_14T er(629, "C")
            Valor = GetTotCta_CodF22_14Ter(629, "C") + GetValAjustesELC(TAEC_AGREGADOS, 8) + GetValAjustesELC(TAEC_AGREGADOS, 9)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 5      'Otros ingresos percibidos o devengados
'            Valor = GetTotCta_CodF22_14Ter(651, "C") + GetValAjustesELC(TAEC_AGREGADOS, 15)
            Valor = GetTotCta_CodF22_14Ter(651, "C") + GetValAjustesELC(TAEC_AGREGADOS, 15) + GetValAjustesELC(TAEC_AGREGADOS, 19)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 6      'Crédito sobre activos fijos adquiridos y pagados en el ejercicio
'            Valor = GetValAjustesELC(TAEC_AGREGADOS, 6)
            Valor = GetVal33Bis()
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
      End Select
      
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
      
   Next i
      
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   FmtLine = "B"
   For i = 0 To MAX_ITEMBASEIMP
   
      If gBaseImp14Ter(BASEIMP_EGRESOS, i) = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_TIPOBASEIMP) = BASEIMP_EGRESOS
      Grid.TextMatrix(Row, C_CONCEPTO) = gBaseImp14Ter(BASEIMP_EGRESOS, i)
            
      'obtenemos el ID y el valor guardado
      Q1 = "SELECT IdBaseImponible14Ter, Valor FROM BaseImponible14Ter "
      Q1 = Q1 & " WHERE TipoBaseImp = " & Grid.TextMatrix(Row, C_TIPOBASEIMP) & " AND IdItemBaseImp = " & Grid.TextMatrix(Row, C_IDITEM)
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdBaseImponible14Ter"))
      End If
      Call CloseRs(Rs)
      
      Select Case i
      
         Case 0      'Total de egresos anuales efectivamente pagados en el ejercicio, a valor nominal
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
            lRowIniEgresos = Row
            
         Case 1      'Costo directo de los bienes o servicios
            Valor = GetTotCta_CodF22_14Ter(630, "D") - GetValAjustesELC(TAEC_AGREGADOS, 1) - GetTotCta_CodF22_14Ter_NC(630, "C")
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 2      'Remuneraciones
            Valor = GetTotCta_CodF22_14Ter(631, "D")
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 3      'Adquisición de bienes del activo realizable y fijo
'            Valor = GetTotCta_CodF22_14Ter(632, "D")
            Valor = GetTotCta_CodF22_14Ter(632, "D") + GetValAjustesELC(TAEC_DEDUCCIONES, 13) + GetValAjustesELC(TAEC_DEDUCCIONES, 14)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 4      'Intereses pagados
            Valor = GetTotCta_CodF22_14Ter(633, "D")
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 5      'Pérdidas de ejercicios anteriores
'            Valor = GetValAjustesELC(TAEC_DEDUCCIONES, 7)
            Valor = GetValAjustesELC(TAEC_DEDUCCIONES, 7) + GetValAjustesELC(TAEC_DEDUCCIONES, 16)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
         Case 6      'Otros gastos deducidos de los ingresos
'            Valor = GetTotCta_CodF22_14Ter(635, "D") + GetValAjustesELC(TAEC_DEDUCCIONES, 17) + GetValAjustesELC(TAEC_DEDUCCIONES, 5) + GetValAjustesELC(TAEC_DEDUCCIONES, 15)
            Valor = GetTotCta_CodF22_14Ter(635, "D") + GetValAjustesELC(TAEC_DEDUCCIONES, 17) + GetValAjustesELC(TAEC_DEDUCCIONES, 5) + GetValAjustesELC(TAEC_DEDUCCIONES, 15) + GetValAjustesELC(TAEC_DEDUCCIONES, 8)
            Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            
            
      End Select
      
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
      
   Next i
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
   FmtLine = "L"
   For i = 0 To MAX_ITEMBASEIMP
      If gBaseImp14Ter(BASEIMP_TOTALES, i) = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_TIPOBASEIMP) = BASEIMP_TOTALES
      Grid.TextMatrix(Row, C_CONCEPTO) = gBaseImp14Ter(BASEIMP_TOTALES, i)
            
      'obtenemos el ID y el valor guardado
      Q1 = "SELECT IdBaseImponible14Ter, Valor FROM BaseImponible14Ter "
      Q1 = Q1 & " WHERE TipoBaseImp = " & Grid.TextMatrix(Row, C_TIPOBASEIMP) & " AND IdItemBaseImp = " & Grid.TextMatrix(Row, C_IDITEM)
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdBaseImponible14Ter"))
         Grid.TextMatrix(Row, C_SUBTOTAL) = Format(vFld(Rs("Valor")), NUMFMT)
      End If
      Call CloseRs(Rs)
      
      Grid.TextMatrix(Row, C_FMT) = FmtLine
      Grid.TextMatrix(Row, C_FMT) = "B"
      Call FGrSetRowStyle(Grid, Row, "B")
      If i = 0 Then       'Base Imponible
         Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
         lRowBaseImponible = Row
      Else
         Call FGrSetRowStyle(Grid, Row, "BC", vbWhite, C_SUBTOTAL, C_SUBTOTAL)
         lRowMayorValor = Row
      End If
     Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
      
   Next i
   
   Call CalcTot
      
   Grid.FlxGrid.Redraw = True

End Sub
Private Sub CalcTot()
   Dim Tot As Double
   Dim i As Integer
   
   Tot = 0
   For i = lRowIniIngresos + 1 To lRowIniIngresos + MAX_ITEMBASEIMP
      Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
   Next i
   
   Grid.TextMatrix(lRowIniIngresos, C_SUBTOTAL) = Format(Tot, NUMFMT)
     
   Tot = 0
   For i = lRowIniEgresos + 1 To lRowIniEgresos + MAX_ITEMBASEIMP
      Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
   Next i

   Grid.TextMatrix(lRowIniEgresos, C_SUBTOTAL) = Format(Tot, NUMFMT)
   
   Tot = vFmt(Grid.TextMatrix(lRowIniIngresos, C_SUBTOTAL)) - vFmt(Grid.TextMatrix(lRowIniEgresos, C_SUBTOTAL))
   Grid.TextMatrix(lRowBaseImponible, C_SUBTOTAL) = Format(Tot, NUMFMT)

End Sub

Private Sub SaveAll()
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxId As Long
   Dim Col As Integer
   
   For Row = Grid.FixedRows To Grid.rows - 1
            
      If Val(Grid.TextMatrix(Row, C_TIPOBASEIMP)) > 0 And Val(Grid.TextMatrix(Row, C_IDITEM)) >= 0 Then
      
         If Row = lRowIniIngresos Or Row = lRowIniEgresos Or Row = lRowBaseImponible Or Row = lRowMayorValor Then
            Col = C_SUBTOTAL
         Else
            Col = C_VALOR
         End If
         
      
         If Val(Grid.TextMatrix(Row, C_ID)) <> 0 Then
            Q1 = "UPDATE BaseImponible14Ter SET Valor = " & vFmt(Grid.TextMatrix(Row, Col))
            Q1 = Q1 & " WHERE IdBaseImponible14Ter = " & Grid.TextMatrix(Row, C_ID)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Else
            MaxId = 0
            Q1 = "SELECT Max(IdBaseImponible14Ter) FROM BaseImponible14Ter"
            Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               MaxId = vFld(Rs(0)) + 1
            End If
            Call CloseRs(Rs)
            
            Q1 = "INSERT INTO BaseImponible14Ter (IdBaseImponible14Ter, TipoBaseImp, IdItemBaseImp, Valor, IdEmpresa, Ano)"
            Q1 = Q1 & " VALUES(" & MaxId & ", " & Grid.TextMatrix(Row, C_TIPOBASEIMP) & "," & Grid.TextMatrix(Row, C_IDITEM) & "," & vFmt(Grid.TextMatrix(Row, Col)) & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
         End If
         
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next Row

End Sub
Private Function Valida() As Boolean
   Valida = False
   Valida = True
End Function

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
'
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
      
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = LP_FGr2String(Grid, Me.Caption)
   Clipboard.Clear
   Clipboard.SetText Clip
      
End Sub


Private Sub Form_Load()
   lOrientacion = ORIENT_VER

   Call SetUpGrid
   Call LoadAll

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Col = C_SUBTOTAL And Row = lRowMayorValor Then
      Value = Format(vFmt(Value), NUMFMT)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   If Col = C_SUBTOTAL And Row = lRowMayorValor Then
      EdType = FEG_Edit
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
