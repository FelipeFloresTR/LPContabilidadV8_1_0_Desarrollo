VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmBaseImponible14D 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Imponible Primera Categoría Reg 14 D"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_BaseImpAcum 
      Caption         =   "Base Imponible Acumulada..."
      Height          =   315
      Left            =   4980
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   855
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1508
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
      Width           =   7755
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   6540
         TabIndex        =   10
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
         Picture         =   "FrmBaseImponible14D.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "FrmBaseImponible14D.frx":04BA
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
         Left            =   960
         Picture         =   "FrmBaseImponible14D.frx":0961
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
         Picture         =   "FrmBaseImponible14D.frx":0DA6
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
         Left            =   2040
         Picture         =   "FrmBaseImponible14D.frx":11CF
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
         Left            =   2460
         Picture         =   "FrmBaseImponible14D.frx":156D
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
         Left            =   1500
         Picture         =   "FrmBaseImponible14D.frx":18CE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nota: Recuerde hacer doble clic en el monto para verificar la BI 14D "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   6795
   End
End
Attribute VB_Name = "FrmBaseImponible14D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_REGIMEN = 0
Const C_MONTO = 1
Const C_UPDATE = 2

Const NCOLS = C_UPDATE

Const R_14DN3 = 1
Const R_14DN8 = 2

Dim lRc As Integer
Dim lBaseImponible As Double
Dim lTipoInforme As Integer

Public Function FEdit(ByVal TipoInforme As Integer, BaseImponible As Double) As Integer

   lTipoInforme = TipoInforme
   Me.Show vbModal
   
   BaseImponible = 0
   
   If lRc = vbOK Then
      BaseImponible = lBaseImponible
   End If
   
   FEdit = lRc
   
End Function


Private Sub Bt_BaseImpAcum_Click()
   Dim Frm As FrmDetCapPropioSimplAcum
   Dim Rc As Integer
   Dim Valor As Double
   
   If Valida() Then
      Call SaveAll

      Set Frm = New FrmDetCapPropioSimplAcum
      Rc = Frm.FEdit(CPS_BASEIMPONIBLE, Valor)
      
      If Rc = vbOK Then
      
         If gEmpresa.ProPymeGeneral <> 0 Then
            Grid.TextMatrix(R_14DN3, C_MONTO) = Format(Valor, NUMFMT)
         Else
            Grid.TextMatrix(R_14DN8, C_MONTO) = Format(Valor, NUMFMT)
         End If
      
      End If
         
      Set Frm = Nothing
   End If
   
End Sub

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

   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Bt_BaseImpAcum.visible = False
   End If
   
   Call SetUpGrid
   
   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double

   Grid.TextMatrix(R_14DN3, C_REGIMEN) = "14 D N°3 Régimen Pro Pyme General"
   Grid.TextMatrix(R_14DN8, C_REGIMEN) = "14 D N°8 Régimen Pro Pyme Transparente"
      
   Q1 = "SELECT CPS_BaseImpPrimCat_14DN3, CPS_BaseImpPrimCat_14DN8 "
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.FlxGrid.Redraw = False
   
   If Not Rs.EOF Then
      
      If gEmpresa.ProPymeGeneral <> 0 Then
         Grid.TextMatrix(R_14DN3, C_MONTO) = Format(vFld(Rs("CPS_BaseImpPrimCat_14DN3")), NUMFMT)
      ElseIf gEmpresa.ProPymeTransp <> 0 Then
         Grid.TextMatrix(R_14DN8, C_MONTO) = Format(vFld(Rs("CPS_BaseImpPrimCat_14DN8")), NUMFMT)
     End If
      
   End If
   
   Call CloseRs(Rs)
   
   Grid.Row = Grid.FixedRows
   Grid.Col = C_REGIMEN
   
   Grid.FlxGrid.Redraw = True
   
End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   lBaseImponible = 0

   Q1 = "UPDATE EmpresasAno SET "
   
   If gEmpresa.ProPymeGeneral <> 0 Then
      lBaseImponible = vFmt(Grid.TextMatrix(R_14DN3, C_MONTO))
      Q1 = Q1 & "  CPS_BaseImpPrimCat_14DN3 = " & lBaseImponible
      Q1 = Q1 & ", CPS_BaseImpPrimCat_14DN8 = 0 "
       
   ElseIf gEmpresa.ProPymeTransp <> 0 Then
      lBaseImponible = vFmt(Grid.TextMatrix(R_14DN8, C_MONTO))
      Q1 = Q1 & "  CPS_BaseImpPrimCat_14DN3 = 0 "
      Q1 = Q1 & ", CPS_BaseImpPrimCat_14DN8 = " & lBaseImponible
       
   End If

   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
  
   Q1 = "SELECT IdCapPropioSimplAnual FROM CapPropioSimplAnual "
   Q1 = Q1 & " WHERE TipoDetCPS = " & CPS_BASEIMPONIBLE & " AND AnoValor = " & gEmpresa.Ano
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Q1 = "UPDATE CapPropioSimplAnual SET Valor = " & lBaseImponible & ", IngresoManual = 0 "
      Q1 = Q1 & " WHERE TipoDetCPS = " & CPS_BASEIMPONIBLE & " AND AnoValor = " & gEmpresa.Ano
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Else
      Q1 = "INSERT INTO CapPropioSimplAnual (TipoDetCPS, IngresoManual, AnoValor, Valor, IdEmpresa )"
      Q1 = Q1 & " VALUES( " & CPS_BASEIMPONIBLE
      Q1 = Q1 & ", 0"
      Q1 = Q1 & ", " & gEmpresa.Ano
      Q1 = Q1 & ", " & lBaseImponible
      Q1 = Q1 & ", " & gEmpresa.id & ") "
      
   End If
   
   Call ExecSQL(DbMain, Q1)
   
   Call CloseRs(Rs)

End Sub

Private Function Valida() As Boolean

   Valida = True
   
End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_REGIMEN) = 5160
   Grid.ColWidth(C_MONTO) = 1500
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_MONTO) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_REGIMEN) = "Régimen"
   Grid.TextMatrix(0, C_MONTO) = "Monto"
   
      
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
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   Titulos(1) = "Año " & gEmpresa.Ano
   gPrtReportes.Titulos = Titulos
'   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   'Total(C_DESC) = "Capital Pripio Tributario"
   'Total(C_TOTAL) = ""
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_REGIMEN
   gPrtReportes.NTotLines = 0

End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
'   Action = vbOK


'   If Col = C_MONTO Then
'      If Value <> "" Then
'         Value = Format(vFmt(Value), NUMFMT)
'         Grid.TextMatrix(Row, Col) = Value
'      End If
'   End If
   
'   If Action = vbOK Then
'      Call FGrModRow(Grid, Row, FGR_U, C_IDSOCIO, C_UPDATE)
'   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
'   If (gEmpresa.ProPymeGeneral <> 0 And Row = R_14DN3) Or (gEmpresa.ProPymeTransp <> 0 And Row = R_14DN8) Then
'
'      If Col = C_MONTO Then
'         EdType = FEG_Edit
'         Grid.TxBox.MaxLength = 12
'      End If
'
'   End If
   
End Sub

Private Sub Grid_DblClick()
   Dim Frm As FrmBaseImponible14DFull
   Dim Row As Integer
   Dim ValorBaseImp As Double
   
   Row = Grid.Row
   
   If (gEmpresa.ProPymeGeneral <> 0 And Row = R_14DN3) Or (gEmpresa.ProPymeTransp <> 0 And Row = R_14DN8) Then
   
      Set Frm = New FrmBaseImponible14DFull
      If Frm.FEdit(ValorBaseImp) = vbOK Then
         Grid.TextMatrix(Row, C_MONTO) = Format(ValorBaseImp, NUMFMT)
      End If
      Set Frm = Nothing
      
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   Call KeyNum(KeyAscii)

End Sub
