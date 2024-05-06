VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResDocs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Vales de Pago Electrónico"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11895
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
         Picture         =   "FrmResDocs.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   1020
         Picture         =   "FrmResDocs.frx":0465
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   600
         Picture         =   "FrmResDocs.frx":091F
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1500
         Picture         =   "FrmResDocs.frx":0DC6
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   3420
         Picture         =   "FrmResDocs.frx":120B
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   2580
         Picture         =   "FrmResDocs.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   3000
         Picture         =   "FrmResDocs.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   2040
         Picture         =   "FrmResDocs.frx":1D33
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10500
         TabIndex        =   16
         Top             =   180
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9446
      _Version        =   393216
   End
   Begin VB.Frame Fr_Periodo 
      Height          =   975
      Left            =   60
      TabIndex        =   17
      Top             =   720
      Width           =   11835
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   5640
         MaxLength       =   15
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1260
         Picture         =   "FrmResDocs.frx":1DD7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   2820
         Picture         =   "FrmResDocs.frx":1E4C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Search 
         Height          =   375
         Left            =   10440
         Picture         =   "FrmResDocs.frx":1EC1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Lb_NDoc 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento:"
         Height          =   195
         Left            =   5640
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   5
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7140
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   11
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmResDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOC = 0
Const C_NUMLIN = 1
Const C_FRECEPCION = 2
Const C_FEMISION = 3
Const C_TIPODOC = 4
Const C_DTE = 5
Const C_NUMDOC = 6
Const C_AFECTO = 7
Const C_EXENTO = 8
Const C_IVA = 9
Const C_OTROIMP = 10
Const C_IVAIRREC = 11
Const C_TOTAL = 12
Const C_RUT = 13
Const C_ENTIDAD = 14
Const C_DESCRIP = 15
Const C_ESTADO = 16

Const NCOLS = C_ESTADO

Const F_INICIO = 0
Const F_FIN = 1

Dim lSupermercado  As Boolean
Dim lVPE  As Boolean
Dim lVerOtrosImp As Boolean

Public Sub FViewSupermercado()
   
   lSupermercado = True
   lVerOtrosImp = True
   Me.Show vbModal
   
End Sub
Public Sub FViewVPE()
   
   lVPE = True
   Me.Show vbModal
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Search_Click()
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim F1 As Long
   Dim F2 As Long
   Dim Q1 As String
   Dim i As Integer

   MesActual = GetMesActual()
   If MesActual = 0 Then
      MesActual = GetUltimoMesConComps()
   End If
   
   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
   Call SetTxDate(Tx_Fecha(F_INICIO), F1)
   Call SetTxDate(Tx_Fecha(F_FIN), F2)
   
   Call BtFechaImg(Bt_Fecha(F_INICIO))
   Call BtFechaImg(Bt_Fecha(F_FIN))
   
   Cb_Estado.AddItem ""
   Cb_Estado.ItemData(Cb_Estado.NewIndex) = 0
   
   For i = 1 To MAX_ESTADODOC
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   
   If lSupermercado Then
      Me.Caption = "Resumen Supermercados y/o Comercios Similares"
   End If

   Call SetUpGrid

   Call LoadAll
   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   If lVPE Then
      Lb_NDoc = "N° Máquina"
   End If
   
   Grid.Cols = NCOLS + 1
   Grid.FixedCols = C_NUMLIN + 1
   
   Grid.ColWidth(C_IDDOC) = 0
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NUMLIN) = 400
   Grid.ColWidth(C_FRECEPCION) = 800
   Grid.ColWidth(C_FEMISION) = 800
   Grid.ColWidth(C_TIPODOC) = 900
   Grid.ColWidth(C_DTE) = 400
   Grid.ColWidth(C_NUMDOC) = 900
   
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_ENTIDAD) = 2000
   Grid.ColWidth(C_DESCRIP) = 2000
   Grid.ColWidth(C_AFECTO) = 1200
   Grid.ColWidth(C_EXENTO) = 1200
   Grid.ColWidth(C_IVA) = 1200
   Grid.ColWidth(C_OTROIMP) = 1200
   Grid.ColWidth(C_IVAIRREC) = IIf(lSupermercado, 1200, 0)
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_ESTADO) = 1050
      
   Grid.ColAlignment(C_NUMLIN) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESCRIP) = flexAlignLeftCenter
   Grid.ColAlignment(C_AFECTO) = flexAlignRightCenter
   Grid.ColAlignment(C_EXENTO) = flexAlignRightCenter
   Grid.ColAlignment(C_IVA) = flexAlignRightCenter
   Grid.ColAlignment(C_OTROIMP) = flexAlignRightCenter
   Grid.ColAlignment(C_IVAIRREC) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter

   Grid.TextMatrix(0, C_NUMLIN) = "Lín."
   Grid.TextMatrix(0, C_FRECEPCION) = "F. Rec."
   Grid.TextMatrix(0, C_FEMISION) = "F.Emisión"
   Grid.TextMatrix(0, C_TIPODOC) = "Tipo Doc."
   Grid.TextMatrix(0, C_DTE) = "DTE"
   Grid.TextMatrix(0, C_NUMDOC) = "N° Doc."
   If lVPE Then
      Grid.TextMatrix(0, C_NUMDOC) = "N° Máq."
   End If
   Grid.TextMatrix(0, C_ESTADO) = "Estado"
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_ENTIDAD) = "Razón Social"
   Grid.TextMatrix(0, C_AFECTO) = "Afecto"
   Grid.TextMatrix(0, C_EXENTO) = "Exento"
   Grid.TextMatrix(0, C_IVA) = "IVA"
   Grid.TextMatrix(0, C_OTROIMP) = "Otros Imp."
   Grid.TextMatrix(0, C_IVAIRREC) = IIf(lSupermercado, "IVA Irrec.", "")
   Grid.TextMatrix(0, C_TOTAL) = "Total"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
       
   Call FGrVRows(Grid)
   
   Call FGrTotales(Grid, GridTot)
   GridTot.ColAlignment(C_NUMDOC) = flexAlignLeftCenter
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Row As Integer
   Dim Total(NCOLS) As Double
   Dim TotSaldo As Double
   Dim NotValidRut As Boolean
   Dim TipoDoc As Integer
   Dim TipoLib As Integer
   Dim ValOtros As Double
   Dim k As Integer
   Dim EsRebaja As Boolean
   Dim IVAIrrecuperable As Double
   Dim TotOtrosImp As Double
   Dim IVAActFijo As Double
         
   Grid.Redraw = False
   
   If lVPE Then
        
      TipoLib = LIB_VENTAS
      Where = " Documento.TipoLIB = " & TipoLib
   
      TipoDoc = FindTipoDoc(LIB_VENTAS, TDOC_VALEPAGOELECTR)
      
      Where = Where & " AND Documento.TipoDoc = " & TipoDoc
   
   ElseIf lSupermercado Then
      
      TipoLib = LIB_COMPRAS
      Where = " Documento.TipoLIB = " & TipoLib
         
      Where = Where & " AND Entidades.EsSupermercado<> 0 "
   End If
   
   
   If ItemData(Cb_Estado) > 0 Then
      Where = Where & " AND Documento.Estado = " & ItemData(Cb_Estado)
   End If
   
   If Trim(Tx_NumDoc) <> "" Then
      Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
   End If

   
   If Tx_Fecha(F_INICIO) <> "" And Tx_Fecha(F_FIN) <> "" Then
      Where = Where & " AND (Documento.FEmision BETWEEN " & GetTxDate(Tx_Fecha(F_INICIO)) & " AND " & GetTxDate(Tx_Fecha(F_FIN)) & ")"
   End If

   Q1 = "SELECT IdDoc, NumDoc, TipoDoc, DTE, Documento.IdEntidad, Entidades.Rut, Entidades.Nombre, "
   Q1 = Q1 & " Entidades.NotValidRut, FEmision, FEmisionOri, Afecto, Exento, IVA, Total, Descrip, Documento.Estado "
   Q1 = Q1 & " FROM Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Entidades.IdEmpresa = Documento.IdEmpresa "
   Q1 = Q1 & " WHERE " & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY FEmision, IdDoc "

   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(TipoLib, vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_DTE) = IIf(vFld(Rs("DTE")) <> 0, "x", "")
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      
      If vFld(Rs("IdEntidad")) <> 0 Then
         Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
         Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Nombre"), True)
      End If
      
      Grid.TextMatrix(i, C_FRECEPCION) = Format(vFld(Rs("FEmision")), SDATEFMT)
      Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FEmisionOri")), SDATEFMT)
     
      EsRebaja = gTipoDoc(GetTipoDoc(TipoLib, vFld(Rs("TipoDoc")))).EsRebaja

      ValOtros = vFld(Rs("Total")) - (vFld(Rs("Exento")) + vFld(Rs("Afecto")) + vFld(Rs("IVA")))         'FCA 30 nov 2017
      If EsRebaja Then
         ValOtros = ValOtros * -1
      End If
      
      TotOtrosImp = 0
      If lVerOtrosImp <> 0 Then
         TotOtrosImp = GetDetOtroImp(TipoLib, vFld(Rs("IdDoc")), EsRebaja, IVAActFijo, IVAIrrecuperable)          'esta función, a diferencia de la que se usa en CompraVenta, retorna el total de otros impuestos (incluyendo el genérico OTROS IMPUESTOS) que no son IVAActivoFijo o IVAIrrecuperable
'         ValOtros = ValOtros - Abs(TotOtrosImp)
'         ValOtros = Abs(TotOtrosImp)           'FCA 30 nov 2017
         ValOtros = TotOtrosImp      'FCA 30 nov 2017
      End If
      
      If EsRebaja Then
         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")) * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")) * -1, NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - Abs(IVAActFijo) - Abs(IVAIrrecuperable)) * -1, NEGNUMFMT)
         If lSupermercado Then
            Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - Abs(IVAIrrecuperable)) * -1, NEGNUMFMT)
            Grid.TextMatrix(i, C_IVAIRREC) = Format(Abs(IVAIrrecuperable) * -1, NEGNUMFMT)
         Else
            Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA"))) * -1, NEGNUMFMT)
         End If
'         Grid.TextMatrix(i, C_OTROIMP) = Format(ValOtros * -1, NEGNUMFMT)
         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")) * -1, NEGNUMFMT)
      Else
         Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")), NEGNUMFMT)
         Grid.TextMatrix(i, C_EXENTO) = Format(vFld(Rs("Exento")), NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")), NEGNUMFMT)
'         Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - IVAActFijo - IVAIrrecuperable), NEGNUMFMT)
         If lSupermercado Then
            Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA")) - IVAIrrecuperable), NEGNUMFMT)
            Grid.TextMatrix(i, C_IVAIRREC) = Format(IVAIrrecuperable, NEGNUMFMT)
        Else
            Grid.TextMatrix(i, C_IVA) = Format((vFld(Rs("IVA"))), NEGNUMFMT)
         End If
'         Grid.TextMatrix(i, C_OTROIMP) = Format(ValOtros, NEGNUMFMT)
         Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NEGNUMFMT)
      End If
           
      Grid.TextMatrix(i, C_OTROIMP) = Format(ValOtros, NEGNUMFMT)
     
     
      For k = C_AFECTO To C_TOTAL
         Total(k) = Total(k) + vFmt(Grid.TextMatrix(i, k))
      Next k
      
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(i, C_ESTADO) = Left(gEstadoDoc(vFld(Rs("Estado"))), 9)
                   
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
      
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
   If Row = 0 Then
      Row = Grid.FixedRows
   End If
   
   Call FGrSelRow(Grid, Row)
   
   GridTot.TextMatrix(0, C_NUMDOC) = "Total"
   For k = C_AFECTO To C_TOTAL
      GridTot.TextMatrix(0, k) = Format(Total(k), NUMFMT)
   Next k
   
   Grid.Redraw = True
   

End Sub


Private Sub Form_Resize()

   Grid.Width = Me.Width - 500
   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 600
   GridTot.Top = Grid.Top + Grid.Height + 60
   GridTot.Width = Grid.Width - 60
   Call FGrVRows(Grid, 1)
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   
   PrtOrient = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   Dim Pag As Integer
      
   PrtOrient = Printer.Orientation
      
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Pag = gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   If Pag >= 0 Then
      Call Frm.FView(Caption)
   End If
   
   Set Frm = Nothing
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = LP_FGr2String(Grid, Me.Caption & vbTab & "Periodo:" & vbTab & Tx_Fecha(F_INICIO) & vbTab & Tx_Fecha(F_FIN), False, C_IDDOC)
   
   If Clip <> "" Then
      Clip = Clip & FGr2String(GridTot)
      
      Clipboard.Clear
      Clipboard.SetText Clip
   End If
   
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
   Dim Total(NCOLS)
   Dim Titulos(0) As String
   
   Printer.Orientation = ORIENT_HOR
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   ColWi(C_ESTADO) = 0
   ColWi(C_ENTIDAD) = ColWi(C_ENTIDAD) * 0.6
   ColWi(C_DESCRIP) = ColWi(C_DESCRIP) * 0.6
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDDOC
   
   For i = 0 To Grid.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   gPrtReportes.Total = Total
   
   
End Sub
Private Sub Bt_DetDoc_Click()
   Dim Frm As FrmDocLib
   Dim IdDoc As Long
   Dim Rc As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmDocLib
   Call Frm.FView(IdDoc)
   Set Frm = Nothing
      
End Sub

Private Sub Grid_DblClick()
   Call PostClick(Bt_DetDoc)
End Sub
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption & vbTab & "Periodo:" & vbTab & Tx_Fecha(F_INICIO) & vbTab & Tx_Fecha(F_FIN))
  End If
   
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub Tx_Fecha_GotFocus(Index As Integer)
   Call DtGotFocus(Tx_Fecha(Index))
End Sub

Private Sub Tx_Fecha_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_Fecha_LostFocus(Index As Integer)

   If Trim$(Tx_Fecha(Index)) = "" Then
      Exit Sub
   End If
   Call DtLostFocus(Tx_Fecha(Index))
      
End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha(Index))
   Set Frm = Nothing

End Sub
'obtiene los datos de otros impuestos para el documento (registro) indicado
Private Function GetDetOtroImp(ByVal TipoLib As Integer, ByVal IdDoc As Long, ByVal EsRebaja As Boolean, IVAActFijo As Double, IVAIrrecuperable As Double) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PrimerDetalle As Long
   Dim TotOtrosImp As Double
   Dim Col As Integer
   Dim Valor As Double
   
   GetDetOtroImp = 0
   
   If lVerOtrosImp = 0 Then
      Exit Function
   End If
   
   If IdDoc <= 0 Then
      Exit Function
   End If
   
   If TipoLib = LIB_COMPRAS Then
      PrimerDetalle = LIBCOMPRAS_IVAIRREC
   Else
      PrimerDetalle = LIBVENTAS_REBAJA65
   End If
   
   Q1 = "SELECT IdTipoValLib, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber FROM MovDocumento "
   Q1 = Q1 & " WHERE IdDoc = " & IdDoc & " AND IdTipoValLib >= " & LIBCOMPRAS_OTROSIMP
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdTipoValLib"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   TotOtrosImp = 0
   IVAActFijo = 0
   IVAIrrecuperable = 0
   
   Do While Not Rs.EOF
   
'      Valor = Abs(vFld(Rs("SumDebe")) - vFld(Rs("SumHaber")))      'FCA 29 nov 2017
      Valor = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
      
'      If (TipoLib = LIB_COMPRAS And EsRebaja) Or (TipoLib = LIB_VENTAS And Not EsRebaja) Then
'         Valor = Valor * -1
'      End If
      
      If TipoLib = LIB_COMPRAS Then
         
         If vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO Then
            IVAActFijo = IVAActFijo + Valor
         
         ElseIf vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC1 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC2 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC3 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC4 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC9 Then
            IVAIrrecuperable = IVAIrrecuperable + Valor
         
         Else
            TotOtrosImp = TotOtrosImp + Valor
         
         End If
     
      Else
         TotOtrosImp = TotOtrosImp + Valor
         
      End If
      
      Rs.MoveNext
   
   Loop
   
   Call CloseRs(Rs)
   
   GetDetOtroImp = TotOtrosImp
   
End Function

