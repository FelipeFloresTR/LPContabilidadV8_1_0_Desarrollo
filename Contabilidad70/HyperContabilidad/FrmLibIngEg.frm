VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmLibIngEg 
   Caption         =   "Libro de Ingresos e Egresos"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Cb_Ano 
      Height          =   315
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   660
      Width           =   975
   End
   Begin VB.ComboBox Cb_Mes 
      Height          =   315
      Left            =   11520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   -60
      Width           =   13875
      Begin VB.CommandButton Bt_ViewRes 
         Caption         =   "Vista Resumida"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   7980
         TabIndex        =   10
         ToolTipText     =   "Vista resumida para impresión"
         Top             =   180
         Width           =   1515
      End
      Begin VB.Frame Fr_BtGen 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   60
         TabIndex        =   17
         Top             =   180
         Width           =   3135
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
            Left            =   2640
            Picture         =   "FrmLibIngEg.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Calendario"
            Top             =   0
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
            Left            =   1740
            Picture         =   "FrmLibIngEg.frx":0429
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Convertir moneda"
            Top             =   0
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
            Left            =   2220
            Picture         =   "FrmLibIngEg.frx":07C7
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Calculadora"
            Top             =   0
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
            Left            =   0
            Picture         =   "FrmLibIngEg.frx":0B28
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Sumar datos seleccionados"
            Top             =   0
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
            Left            =   1320
            Picture         =   "FrmLibIngEg.frx":0BCC
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Copiar Excel"
            Top             =   0
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
            Left            =   900
            Picture         =   "FrmLibIngEg.frx":1011
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir"
            Top             =   0
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
            Left            =   480
            Picture         =   "FrmLibIngEg.frx":14CB
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Vista previa de la impresión"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   12480
         TabIndex        =   11
         Top             =   180
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab GridTab 
      Height          =   8175
      Left            =   120
      TabIndex        =   15
      Top             =   1140
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingresos"
      TabPicture(0)   =   "FrmLibIngEg.frx":1972
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Tx_CurrCell_I"
      Tab(0).Control(1)=   "Grid_I"
      Tab(0).Control(2)=   "GridTot_I"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Egresos"
      TabPicture(1)   =   "FrmLibIngEg.frx":198E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GridTot_E"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Grid_E"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Tx_CurrCell_E"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox Tx_CurrCell_I 
         Height          =   315
         Left            =   -74820
         TabIndex        =   21
         Top             =   7680
         Width           =   6735
      End
      Begin VB.TextBox Tx_CurrCell_E 
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   7680
         Width           =   6735
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_I 
         Height          =   6915
         Left            =   -74820
         TabIndex        =   12
         Top             =   420
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12197
         _Version        =   393216
         BackColorBkg    =   16777215
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_E 
         Height          =   6915
         Left            =   180
         TabIndex        =   0
         Top             =   420
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12197
         _Version        =   393216
         BackColorBkg    =   16777215
      End
      Begin MSFlexGridLib.MSFlexGrid GridTot_E 
         Height          =   315
         Left            =   180
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   7320
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   556
         _Version        =   393216
         Cols            =   11
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid GridTot_I 
         Height          =   315
         Left            =   -74820
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7320
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   556
         _Version        =   393216
         Cols            =   11
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periodo:"
      Height          =   195
      Left            =   10920
      TabIndex        =   14
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Lb_Contrib 
      AutoSize        =   -1  'True
      Caption         =   "Contribuyentes acogidos al Régimen del Artículo 14 Ter A) y no se encuentren obligados a llevar Libro de Compras y Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   10575
   End
End
Attribute VB_Name = "FrmLibIngEg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDLIBROCAJA = 0
Const C_NUMLIN = 1
Const C_NUMDOC = 2
Const C_TIPODOCEXT = 3
Const C_RUT = 4
Const C_FECHAOPER = 5
Const C_INGPERCIBIDOS = 6
Const C_INGDEVENGADOS = 7
Const C_DESCRIP = 8
Const C_CONENTREL = 9
Const C_OPERDEVENGADA = 10
Const C_TOTALINGEGMES = 11

Const C_EGPAGADOS = 7
Const C_EGADEUDADOS = 8

Const NCOLS = C_TOTALINGEGMES

Dim lInLoad As Boolean

Dim lOrientacion As Integer
Dim lInfoPreliminar As Boolean
Dim lPapelFoliado As Boolean
Dim lViewRes As Boolean

Private Sub Bt_Close_Click()
   Unload Me
End Sub


Private Sub Bt_ViewRes_Click()

   lViewRes = Not lViewRes

   Call SetUpGrid(TOPERCAJA_INGRESO, Grid_I, GridTot_I)
   Call SetUpGrid(TOPERCAJA_EGRESO, Grid_E, GridTot_E)
   
   If lViewRes Then
      Bt_ViewRes.Caption = "Vista Completa"
   Else
      Bt_ViewRes.Caption = "Vista Resumida"
   End If

End Sub

Private Sub Cb_Mes_Click()

   If lInLoad Then
      Exit Sub
   End If
      
   
   Me.MousePointer = vbHourglass
   
   Call LoadGrid(TOPERCAJA_INGRESO, Grid_I, GridTot_I)
   Call LoadGrid(TOPERCAJA_EGRESO, Grid_E, GridTot_E)
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim MesActual As Integer
   
   lInLoad = True

   MesActual = GetMesActual()
   lViewRes = False
   
   For i = 1 To 12
      Cb_Mes.AddItem gNomMes(i)
      Cb_Mes.ItemData(Cb_Mes.NewIndex) = i
   Next i
   
   Cb_Mes.ListIndex = 0
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   End If
   
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex
   
   Call SetUpGrid(TOPERCAJA_INGRESO, Grid_I, GridTot_I)
   Call SetUpGrid(TOPERCAJA_EGRESO, Grid_E, GridTot_E)
   
   Call LoadGrid(TOPERCAJA_INGRESO, Grid_I, GridTot_I)
   Call LoadGrid(TOPERCAJA_EGRESO, Grid_E, GridTot_E)

   lInLoad = False
End Sub

Private Sub SetUpGrid(ByVal TipoOper As Integer, Grid As MSFlexGrid, GridTot As MSFlexGrid)

   Grid.Cols = NCOLS + 1
   Grid.FixedCols = C_NUMLIN + 1
   Grid.rows = 10
   Grid.FixedRows = 2
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDLIBROCAJA) = 0
   Grid.ColWidth(C_NUMLIN) = 500
   Grid.ColWidth(C_NUMDOC) = 1200
   Grid.ColWidth(C_TIPODOCEXT) = 2000
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_FECHAOPER) = 900
   Grid.ColWidth(C_INGPERCIBIDOS) = 1200
   Grid.ColWidth(C_INGDEVENGADOS) = 1200
   If gEmpresa.Ano < 2021 Then
      Grid.ColWidth(C_CONENTREL) = IIf(TipoOper = TOPERCAJA_INGRESO, 1000, 0)
      Grid.ColWidth(C_OPERDEVENGADA) = 1000
      Grid.ColWidth(C_DESCRIP) = IIf(TipoOper = TOPERCAJA_INGRESO, 3000, 1000 + 3000)
   Else
      Grid.ColWidth(C_OPERDEVENGADA) = 0    '1000     'Se elimina esta columna. Solicitado por Victor Morales 31/5/2021
      Grid.ColWidth(C_CONENTREL) = 1000               'Se agrega esta columna. Solicitado por Victor Morales 31/5/2021
      Grid.ColWidth(C_DESCRIP) = 1000 + 3000
   End If
   Grid.ColWidth(C_TOTALINGEGMES) = 0
  
   If lViewRes Then
      Grid.ColWidth(C_DESCRIP) = 0
      Grid.ColWidth(C_OPERDEVENGADA) = 0
      Grid.ColWidth(C_TOTALINGEGMES) = 1200
   End If
   
   Grid.ColAlignment(C_NUMLIN) = flexAlignRightCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_INGPERCIBIDOS) = flexAlignRightCenter
   Grid.ColAlignment(C_INGDEVENGADOS) = flexAlignRightCenter
   Grid.ColAlignment(C_CONENTREL) = flexAlignCenterCenter
   Grid.ColAlignment(C_OPERDEVENGADA) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_NUMLIN) = "N°"
   Grid.TextMatrix(1, C_NUMLIN) = "Corr."
   Grid.TextMatrix(1, C_NUMDOC) = "N° Doc."
   Grid.TextMatrix(1, C_TIPODOCEXT) = "Tipo Documento"
   Grid.TextMatrix(0, C_FECHAOPER) = "Fecha"
   Grid.TextMatrix(1, C_FECHAOPER) = "Operación"
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(1, C_RUT) = "Receptor"

   Grid.TextMatrix(0, C_INGPERCIBIDOS) = "Ingresos $"
   Grid.TextMatrix(1, C_INGPERCIBIDOS) = "Percibidos"
   Grid.TextMatrix(0, C_INGDEVENGADOS) = "Ingresos $"
   Grid.TextMatrix(1, C_INGDEVENGADOS) = "Devengados"
   If Grid.Name = "Grid_E" Then
      Grid.TextMatrix(1, C_RUT) = "Emisor"
      Grid.TextMatrix(0, C_INGPERCIBIDOS) = "Egresos $"
      Grid.TextMatrix(1, C_INGPERCIBIDOS) = "Pagados"
      Grid.TextMatrix(0, C_INGDEVENGADOS) = "Egresos $"
      Grid.TextMatrix(1, C_INGDEVENGADOS) = "Adeudados"
   End If
   
   If Grid.ColWidth(C_DESCRIP) > 0 Then
      Grid.TextMatrix(0, C_DESCRIP) = "Glosa"
      Grid.TextMatrix(1, C_DESCRIP) = "Operación"
   Else
      Grid.TextMatrix(0, C_DESCRIP) = ""
      Grid.TextMatrix(1, C_DESCRIP) = ""
   End If
   
   If Grid.ColWidth(C_CONENTREL) > 0 Then
      Grid.TextMatrix(0, C_CONENTREL) = "Oper. Ent."
      Grid.TextMatrix(1, C_CONENTREL) = "Relacionada"
   End If
   
   If Grid.ColWidth(C_OPERDEVENGADA) > 0 Then
      Grid.TextMatrix(0, C_OPERDEVENGADA) = "Operación"
      Grid.TextMatrix(1, C_OPERDEVENGADA) = "Devengada"
   Else
      Grid.TextMatrix(0, C_OPERDEVENGADA) = ""
      Grid.TextMatrix(1, C_OPERDEVENGADA) = ""
   End If
   
   If Grid.ColWidth(C_TOTALINGEGMES) > 0 Then
      Grid.TextMatrix(0, C_TOTALINGEGMES) = "Total " & IIf(TipoOper = TOPERCAJA_INGRESO, "Ingresos", "Egresos")
      Grid.TextMatrix(1, C_TOTALINGEGMES) = "del Mes"
   Else
      Grid.TextMatrix(0, C_TOTALINGEGMES) = ""
      Grid.TextMatrix(1, C_TOTALINGEGMES) = ""
   End If
   
   Call FGrVRows(Grid, 1)
  
   Call FGrTotales(Grid, GridTot)
      
End Sub

Private Sub Form_Resize()
   Dim H As Integer
   
   GridTab.Tab = 0

   GridTab.Width = Me.Width - GridTab.Left - W.xFrame * 2 - 60
   Grid_I.Width = GridTab.Width - Grid_I.Left - 100
   Grid_E.Width = Grid_I.Width
   
   GridTab.Height = Me.Height - GridTab.Top - W.yFrame * 2 - W.YCaption - 100
   H = GridTab.Height - Grid_I.Top - GridTot_I.Height - Tx_CurrCell_I.Height - 230
   Grid_I.Height = H
   Grid_E.Height = H
   
   H = Grid_I.Top + Grid_I.Height + 30
   GridTot_I.Top = H
   GridTot_E.Top = H
   
   H = GridTot_I.Top + GridTot_I.Height + 60
   Tx_CurrCell_I.Top = H
   Tx_CurrCell_E.Top = H
   
   Call FGrVRows(Grid_I, 1)
   Call FGrVRows(Grid_E, 1)
   
End Sub

Private Function LoadGrid(ByVal TipoOper As Integer, Grid As MSFlexGrid, GridTot As MSFlexGrid)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim FirstDay As Long
   Dim LastDay As Long
   Dim i As Integer
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim NotValidRut As Boolean
   Dim EsRebaja As Boolean, Reb As Integer
   Dim TipoDoc As Integer
   Dim Total(NCOLS) As Double
   
   Grid.Redraw = False
   
   Call FirstLastMonthDay(DateSerial(Val(Cb_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   
   Q1 = "SELECT IdLibroCaja, LibroCaja.TipoDoc, LibroCaja.TipoLib, NumDoc, DTE, NumDocHasta, LibroCaja.IdEntidad, LibroCaja.RutEntidad "
   Q1 = Q1 & " , Entidades.Rut, Entidades.NotValidRut, FechaOperacion, Afecto, IVA, Exento, OtroImp "
   Q1 = Q1 & " , Total, Pagado, Descrip, ConEntRel, OperDevengada, LibroCaja.Estado "
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc )"
   Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Entidades.IdEmpresa = LibroCaja.IdEmpresa "
   Q1 = Q1 & " WHERE (FechaOperacion BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND TipoOper = " & TipoOper
'   Q1 = Q1 & " AND LibroCaja.TipoLib NOT IN (" & LIB_RETEN & "," & LIB_CAJAING & "," & LIB_CAJAEGR & ")"
   Q1 = Q1 & " AND LibroCaja.TipoLib NOT IN (" & LIB_CAJAING & "," & LIB_CAJAEGR & ")"
   Q1 = Q1 & " AND TipoDocs.Diminutivo NOT IN ( 'OTV', 'OTC')"
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
  
   
   Q1 = Q1 & " ORDER BY FechaOperacion, IdLibroCaja"
      
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows

   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
           
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If

      
      Grid.TextMatrix(i, C_IDLIBROCAJA) = vFld(Rs("IdLibroCaja"))
      Grid.TextMatrix(i, C_NUMLIN) = i - Grid.FixedRows + 1
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(i, C_TIPODOCEXT) = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Nombre & IIf(vFld(Rs("DTE")) <> 0, " E", "")
      
      If vFld(Rs("Estado")) <> ED_ANULADO Then
      
         If vFld(Rs("IdEntidad")) = 0 Then
            If vFld(Rs("RutEntidad")) <> "" And vFld(Rs("RutEntidad")) <> "0" Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RutEntidad")))
            End If
         Else
            If vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0" Then
               Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
            End If
         End If
      
      Else
         Grid.TextMatrix(i, C_RUT) = "NULO"
         
      End If
      
      Grid.TextMatrix(i, C_FECHAOPER) = Format(vFld(Rs("FechaOperacion")), SDATEFMT)
            
      EsRebaja = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).EsRebaja
      If EsRebaja Then
         Reb = -1
      Else
         Reb = 1
      End If
      
      'se usan estas columnas que son las mismas para ingresos y egresos
      Grid.TextMatrix(i, C_INGPERCIBIDOS) = Format(vFld(Rs("Pagado")) * Reb, NEGNUMFMT)
      Grid.TextMatrix(i, C_INGDEVENGADOS) = Format((vFld(Rs("Total")) - vFld(Rs("Pagado"))) * Reb, NEGNUMFMT)
      
      Grid.TextMatrix(i, C_TOTALINGEGMES) = Grid.TextMatrix(i, C_INGPERCIBIDOS)
      
      Total(C_INGPERCIBIDOS) = Total(C_INGPERCIBIDOS) + vFmt(Grid.TextMatrix(i, C_INGPERCIBIDOS))
      Total(C_INGDEVENGADOS) = Total(C_INGDEVENGADOS) + vFmt(Grid.TextMatrix(i, C_INGDEVENGADOS))
      
      Total(C_TOTALINGEGMES) = Total(C_TOTALINGEGMES) + vFmt(Grid.TextMatrix(i, C_TOTALINGEGMES))
      
      If vFld(Rs("ConEntRel")) <> 0 Then
         Grid.TextMatrix(i, C_CONENTREL) = "x"
      Else
         Grid.TextMatrix(i, C_CONENTREL) = ""
      End If
                
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
               
     
      If vFld(Rs("OperDevengada")) <> 0 Then
         Grid.TextMatrix(i, C_OPERDEVENGADA) = "x"
      Else
         Grid.TextMatrix(i, C_OPERDEVENGADA) = ""
      End If
                    
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   GridTot.TextMatrix(0, C_TIPODOCEXT) = "TOTAL"
   GridTot.TextMatrix(0, C_INGPERCIBIDOS) = Format(Total(C_INGPERCIBIDOS), NEGNUMFMT)
   GridTot.TextMatrix(0, C_INGDEVENGADOS) = Format(Total(C_INGDEVENGADOS), NEGNUMFMT)
   GridTot.TextMatrix(0, C_TOTALINGEGMES) = Format(Total(C_TOTALINGEGMES), NEGNUMFMT)
   
   Call FGrVRows(Grid, 1)
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = C_NUMDOC
   Grid.ColSel = Grid.Col
      
         
   'Call CalcTot
   
   Grid.Redraw = True
     
End Function

Private Sub Grid_I_Scroll()
   GridTot_I.LeftCol = Grid_I.LeftCol
End Sub
Private Sub Grid_I_SelChange()
   Tx_CurrCell_I = Grid_I.TextMatrix(Grid_I.Row, Grid_I.Col)
End Sub

Private Sub Grid_E_SelChange()
   Tx_CurrCell_E = Grid_E.TextMatrix(Grid_E.Row, Grid_E.Col)
End Sub
Private Sub Grid_E_Scroll()
   GridTot_E.LeftCol = Grid_E.LeftCol
End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   If GridTab.Tab = 0 Then
      Call Frm.FViewSum(Grid_I, Grid_I.Row, Grid_I.RowSel, Grid_I.Col, Grid_I.ColSel)
   Else
      Call Frm.FViewSum(Grid_E, Grid_I.Row, Grid_E.RowSel, Grid_E.Col, Grid_E.ColSel)
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   Dim FrmPrt As FrmPrtSetup
   Dim OldOrientacion As Integer
   Dim i As Integer
   Dim Total(NCOLS) As String
   Dim ColWi(NCOLS) As Integer
                  
   OldOrientacion = Printer.Orientation
   
   Me.MousePointer = vbHourglass
   
   'imprimimos los ingresos
   Set gPrtLibros.Grid = Grid_I
   Call SetUpPrtGrid
   
   For i = 0 To Grid_I.Cols - 1
      Total(i) = GridTot_I.TextMatrix(0, i)
   Next i
   gPrtLibros.Total = Total

   Set FrmPrt = Nothing
   
   Set Frm = New FrmPrintPreview
               
   gPrtLibros.PermitirMasDe1Franja = True
   gPrtLibros.FixedCols = C_TIPODOCEXT + 1
   
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   
   'imprimimos los egresos
   Set gPrtLibros.Grid = Grid_E
   For i = 0 To Grid_I.Cols - 1
      Total(i) = GridTot_E.TextMatrix(0, i)
   Next i
   gPrtLibros.Total = Total
   
   For i = 0 To Grid_I.Cols - 1
      ColWi(i) = Grid_E.ColWidth(i)
   Next i
      
   gPrtLibros.ColWi = ColWi

   gPrtLibros.EsContinuacion = True

   
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   
   gPrtLibros.CallEndDoc = True
   gPrtLibros.EsContinuacion = False
              
   Me.MousePointer = vbDefault
   
   Set Frm.PrtControl = Bt_Print
   
   Call Frm.FView(Caption)
   
   Set Frm = Nothing
         
   
   gPrtLibros.GrFontName = Grid_I.Font.Name
   gPrtLibros.GrFontSize = Grid_I.Font.Size
   gPrtLibros.TotFntBold = True
   gPrtLibros.PermitirMasDe1Franja = False

   Printer.Orientation = OldOrientacion
      
   Call ResetPrtBas(gPrtLibros)
  
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   Dim Pag As Integer
   Dim Total(NCOLS) As String
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   
   lOrientacion = ORIENT_HOR
   
   Set Frm = New FrmPrtSetup
   
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) <> vbOK Then
      Call ResetPrtBas(gPrtLibros)
      Set Frm = Nothing
      Exit Sub
   End If
  
   Set Frm = Nothing
  
   OldOrientacion = Printer.Orientation
   
   Call SetUpPrtGrid
         
   'Imprimimos los ingresos
   Set gPrtLibros.Grid = Grid_I
   For i = 0 To Grid_I.Cols - 1
      Total(i) = GridTot_I.TextMatrix(0, i)
   Next i
   gPrtLibros.Total = Total
   
   gPrtLibros.PermitirMasDe1Franja = True
   gPrtLibros.FixedCols = C_TIPODOCEXT + 1
   gPrtLibros.CallEndDoc = False
   
   Pag = gPrtLibros.PrtFlexGrid(Printer)
               
               
   'Imprimimos los egresos
   Set gPrtLibros.Grid = Grid_E
   For i = 0 To Grid_E.Cols - 1
      Total(i) = GridTot_E.TextMatrix(0, i)
   Next i
   gPrtLibros.Total = Total
         
   For i = 0 To Grid_I.Cols - 1
      ColWi(i) = Grid_E.ColWidth(i)
   Next i
      
   gPrtLibros.ColWi = ColWi
   gPrtLibros.EsContinuacion = True
   
   Pag = gPrtLibros.PrtFlexGrid(Printer)
               
   gPrtLibros.CallEndDoc = True
   Printer.EndDoc
   
   If lPapelFoliado Then
      Call AppendLogImpreso(LIBOF_INGEGR, 0, DateSerial(CbItemData(Cb_Ano), CbItemData(Cb_Mes), 1), DateSerial(CbItemData(Cb_Ano), CbItemData(Cb_Mes), 1))
   End If
   
   'Chequeo si debo actualizar folio ultimo usado
   Call UpdateUltUsado(lPapelFoliado, Pag)

   
   gPrtLibros.PermitirMasDe1Franja = False
   
   gPrtLibros.GrFontName = Grid_I.Font.Name
   gPrtLibros.GrFontSize = Grid_I.Font.Size
   gPrtLibros.TotFntBold = True

   lInfoPreliminar = False
   Printer.Orientation = OldOrientacion
   
   Call ResetPrtBas(gPrtLibros)

End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(6) As String
   Dim EncabezadosCont(6) As String
   Dim FontTit(2) As FontDef_t
   Dim OldOrient As Integer
   Dim Mes As String
   Dim Idx As Integer
   Dim FntEncabezados(0) As FontDef_t
     
   lOrientacion = ORIENT_HOR
  
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = "Libro de Ingresos y Egresos"
   
   FontTit(0).FontBold = True
         
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
      
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
'   Encabezados(0) = Lb_Contrib
   Encabezados(0) = "   "
   Encabezados(1) = "Periodo: " & Cb_Mes & " " & Val(Cb_Ano)
   Encabezados(2) = "RUT:     " & FmtCID(gEmpresa.Rut)
   Encabezados(3) = "Nombre/Razón Social: " & gEmpresa.RazonSocial
   Encabezados(4) = "   "
   Encabezados(5) = "SECCIÓN A INGRESOS"
   
   gPrtLibros.Encabezados = Encabezados
   
   EncabezadosCont(0) = "    "
   EncabezadosCont(1) = "    "
   EncabezadosCont(2) = "SECCIÓN B EGRESOS"
   
   gPrtLibros.EncabezadosCont = EncabezadosCont
   
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   FntEncabezados(0).FontBold = True
   FntEncabezados(0).FontUnderline = False

   Call gPrtLibros.FntEncabezados(FntEncabezados)
   
   gPrtLibros.GrFontName = Grid_I.Font.Name
   gPrtLibros.GrFontSize = Grid_I.Font.Size
   
   For i = 0 To Grid_I.Cols - 1
      ColWi(i) = Grid_I.ColWidth(i)
   Next i
      
   gPrtLibros.ColWi = ColWi
   gPrtLibros.NTotLines = 1
   gPrtLibros.ColObligatoria = C_IDLIBROCAJA
   
   gPrtLibros.Obs = ""   'para que no ponga las notas
   
   gPrtLibros.CallEndDoc = False
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = FGr2String(Grid_I, "Libro de Ingresos" & vbTab & Cb_Mes & " " & Val(Cb_Ano), False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTot_I)
   Clip = Clip & FGr2String(Grid_E, "Libro de Egresos" & vbTab & Cb_Mes & " " & Val(Cb_Ano), False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTot_E)
   
   Clipboard.Clear
   Clipboard.SetText Clip

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

