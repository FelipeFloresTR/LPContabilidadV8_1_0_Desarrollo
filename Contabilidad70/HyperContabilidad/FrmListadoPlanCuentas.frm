VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmLstPlanCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cuentas"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
   Icon            =   "FrmListadoPlanCuentas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   11456
      Cols            =   8
      Rows            =   20
      FixedCols       =   2
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
   Begin VB.CommandButton Bt_ViewLsAtrib 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   13380
      Picture         =   "FrmListadoPlanCuentas.frx":000C
      TabIndex        =   10
      ToolTipText     =   "Ver lista de claves para columnas de atributos"
      Top             =   780
      Width           =   255
   End
   Begin VB.TextBox Tx_Atributos 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   8490
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Atributos"
      Top             =   780
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   13515
      Begin VB.CommandButton Bt_DetCta 
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
         Left            =   60
         Picture         =   "FrmListadoPlanCuentas.frx":00A5
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Detalle cuenta seleccionada"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Grabar"
         Height          =   315
         Left            =   10920
         TabIndex        =   5
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_CopyExcel 
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
         Left            =   1440
         Picture         =   "FrmListadoPlanCuentas.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_View 
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
         Left            =   600
         Picture         =   "FrmListadoPlanCuentas.frx":094F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   795
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
         Picture         =   "FrmListadoPlanCuentas.frx":0DF6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   12180
         TabIndex        =   6
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de detalle:"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   8
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Label Lb_Nota 
      Caption         =   "Nota: para agregar o eliminar un atributo haga doble-click en la celda correspondiente"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   7680
      Width           =   11775
   End
End
Attribute VB_Name = "FrmLstPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CUENTA = 0
Const C_NOMBCORTO = 1
Const C_DESC = 2
Const C_IDCUENTA = 3
Const C_NIVEL = 4
Const C_CODF22 = 5
Const C_CODF22_14TER = 6
Const C_PARTIDA = 7
Const C_CODF29 = 8
Const C_UPDATE = 9
Const C_CLASCTA = 10
Const C_INIATRIB = 11
Const C_PERCEPCION = 12

Const NCOLS = C_PERCEPCION

Dim lEdAtrib As Boolean
Dim lCuenta As Cuenta_t


Private Sub SetUpGrid()
   Dim i As Integer
   
  ' Grid.AllowUserResizing = flexResizeColumns
  
  Grid.Cols = NCOLS + 1
  Call FGrSetup(Grid, False)
   
   Grid.ColWidth(C_CUENTA) = 1200
   Grid.ColWidth(C_NOMBCORTO) = 1000
   Grid.ColWidth(C_DESC) = 3600
   Grid.ColWidth(C_CODF22) = 760
   Grid.ColWidth(C_CODF22_14TER) = 1300
   Grid.ColWidth(C_PARTIDA) = 600
   Grid.ColWidth(C_CODF29) = 0     '800
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_CLASCTA) = 0
   
   
   For i = C_INIATRIB To MAX_ATRIB - 1 + C_INIATRIB
      Grid.Cols = i + 1
      
      Grid.ColWidth(i) = 600
      Grid.TextMatrix(0, i) = gAtribCuentas(i - C_INIATRIB + 1).NombreCorto
      Grid.ColAlignment(i) = flexAlignCenterCenter
   Next i
   
   Grid.ColWidth(MAX_ATRIB - 1 + C_INIATRIB) = 1000
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_NOMBCORTO) = "Nombre"
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   Grid.TextMatrix(0, C_CODF22) = "Cód. F22"
   Grid.TextMatrix(0, C_CODF22_14TER) = IIf(gEmpresa.Ano < 2020, "Cód. F22 " & gAtribCuentas(ATRIB_14TER).NombreCorto, "Äjuste 14D N3y8")
   Grid.TextMatrix(0, C_PARTIDA) = "Partida"
   Grid.TextMatrix(0, C_CODF29) = ""      '"Cód. F29"
   
End Sub


Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, "")
End Sub

Private Sub Bt_DetCta_Click()
   Dim Row As Integer
   Dim Frm As FrmCuenta
   Dim Rc As Integer
   
   Row = Grid.Row
   If Row < Grid.FixedRows Or Val(Grid.TextMatrix(Row, C_IDCUENTA)) = 0 Then
      Exit Sub
   End If

   lCuenta.Codigo = VFmtCodigoCta(Grid.TextMatrix(Row, C_CUENTA))
   lCuenta.Nivel = Val(Grid.TextMatrix(Row, C_NIVEL))
   lCuenta.NivelFather = Val(Grid.TextMatrix(Row, C_NIVEL)) - 1
   lCuenta.id = Val(Grid.TextMatrix(Row, C_IDCUENTA))
   lCuenta.Tipo = Val(Grid.TextMatrix(Row, C_CLASCTA))

   Set Frm = New FrmCuenta
   Rc = Frm.FEdit(lCuenta)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Me.MousePointer = vbHourglass
      Call LoadAll
      Me.MousePointer = vbDefault
   End If
   
End Sub

Private Sub Bt_OK_Click()
   Call SaveAll
   Unload Me
End Sub

Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
      
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub bt_View_Click()
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

Private Sub Bt_ViewLsAtrib_Click()
   FrmLstAtrib.Show vbModal
End Sub

Private Sub Cb_Nivel_Click()
     Call LoadAll
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   Call SetUpGrid
     
   For i = 1 To MAX_NIVELES
      Cb_Nivel.AddItem i
   Next i
   Cb_Nivel.ListIndex = gLastNivel - 1
   
   Call SetupPriv
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(C_INIATRIB + MAX_ATRIB) As Integer
   Dim Total(C_INIATRIB + MAX_ATRIB) As String
   Dim Titulos(0) As String
   Dim Encabezados(3) As String
   Dim FontTit(0) As FontDef_t
   Dim FontNom(0) As FontDef_t
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Listado de Plan de Cuentas"
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
      
   gPrtReportes.GrFontName = "Arial"
   gPrtReportes.GrFontSize = 8
   gPrtReportes.Encabezados = Encabezados
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDCUENTA
   gPrtReportes.NTotLines = 0
   
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer, i As Integer
   Dim Nivel As Integer, CurNiv As Integer
   Dim CodFather As String
   
   Grid.FlxGrid.Redraw = False
   
   Nivel = Val(Cb_Nivel)
   
   Q1 = "SELECT DISTINCT Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Nivel, CodF22, CodF22_14Ter, CodF29, TipoPartida,  "
   Q1 = Q1 & " Cuentas.Clasificacion , Cuentas.Descripcion"
   For i = 1 To MAX_ATRIB
      Q1 = Q1 & ", Atrib" & i
   Next i
   Q1 = Q1 & " ,Cuentas.percepcion"
   Q1 = Q1 & " FROM Cuentas  "
   Q1 = Q1 & " WHERE Cuentas.Nivel <= " & Nivel
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   Do While Rs.EOF = False
      CurNiv = vFld(Rs("Nivel"))
      
      If CodFather <> Left(vFld(Rs("Codigo")), gNiveles.Largo(CurNiv)) Then
         If Row > 1 Then
            'Salto una línea
            Grid.rows = Row + 1
            Grid.TextMatrix(Row, C_IDCUENTA) = "*******"
            
            Row = Row + 1
            
         End If
         
         CodFather = Left(vFld(Rs("Codigo")), gNiveles.Largo(vFld(Rs("Nivel"))))
         'NomFather = vFld(Rs("Descripcion"))
         
      End If
      
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_CUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (CurNiv - 1), " ") & FCase(vFld(Rs("Descripcion"), True))
      Grid.TextMatrix(Row, C_NOMBCORTO) = vFld(Rs("Nombre"))
      
      If vFld(Rs("Nivel")) = gLastNivel Then
         If vFld(Rs("CodF22")) <> 0 Then
            Grid.TextMatrix(Row, C_CODF22) = vFld(Rs("CodF22"))
         End If
         If vFld(Rs("CodF22_14TER")) <> 0 Then
            Grid.TextMatrix(Row, C_CODF22_14TER) = vFld(Rs("CodF22_14TER"))
         End If
         If vFld(Rs("CodF29")) <> 0 Then
            Grid.TextMatrix(Row, C_CODF29) = vFld(Rs("CodF29"))
         End If
         If vFld(Rs("TipoPartida")) <> 0 Then
            Grid.TextMatrix(Row, C_PARTIDA) = vFld(Rs("TipoPartida"))
         End If
         
      End If
      
      Grid.TextMatrix(Row, C_IDCUENTA) = vFld(Rs("idCuenta"))
      Grid.TextMatrix(Row, C_NIVEL) = vFld(Rs("Nivel"))
      Grid.TextMatrix(Row, C_CLASCTA) = vFld(Rs("Clasificacion"))
      
      
      For i = C_INIATRIB To C_INIATRIB + MAX_ATRIB - 1
         If vFld(Rs("Atrib" & i - C_INIATRIB + 1)) <> 0 Then
            Grid.TextMatrix(Row, i) = "x"
         End If
      Next i
      Grid.TextMatrix(Row, Grid.Cols - 1) = IIf(vFld(Rs("percepcion")) <> 0, "x", "")
      If vFld(Rs("Nivel")) = 1 Then
         Call FGrFontBold(Grid, Row, -1, True)
      End If
         
        
      Row = Row + 1
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   Grid.FlxGrid.Redraw = True
   
End Sub
Private Sub Form_Resize()
    Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - 2 * (Grid.Left + W.xFrame)
   If d > 1000 Then
      'Grid.Width = d
   End If
 
   d = Me.Height - Grid.Top - W.YCaption * 2 + 80 - Lb_Nota.Height - 100
   If d > 1000 Then
      Grid.Height = d
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   Lb_Nota.Top = Grid.Top + Grid.Height + 100
   'Tx_Atributos.Width = Me.Width - Tx_Atributos.Left - Bt_ViewLsAtrib.Width - 200
   'Bt_ViewLsAtrib.Left = Tx_Atributos.Left + Tx_Atributos.Width + 30
   
   Call FGrVRows(Grid)
   
End Sub



Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   Value = Trim(Value)
   
   If Col = C_CODF22 Then
      If Val(Value) < 0 Then
         MsgBox1 "Código inválido.", vbExclamation
         Action = vbCancel

      ElseIf gEmpresa.Ano = 2017 And Val(Value) > 0 And InStr("," & LSTCODF22_2017 & ",", "," & Value & ",") <= 0 Then    'es inválido
         MsgBox1 "Código Form 22 inválido.", vbExclamation
         Action = vbCancel

      Else
         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
         Action = vbOK
      End If
      
   ElseIf Col = C_CODF29 Or Col = C_CODF22_14TER Then
      If Val(Value) < 0 Then
         MsgBox1 "Código inválido.", vbExclamation
         Action = vbCancel
         
      Else
         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
         Action = vbOK
      End If
   
   ElseIf Col = C_PARTIDA Then
      If Val(Value) < 0 Or Val(Value) > MAX_TIPOPARTIDA Then
         MsgBox1 "Tipo partida inválido.", vbExclamation
         Action = vbCancel
      Else
         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
         Action = vbOK
      End If
   
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   EdType = FEG_None
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lEdAtrib = False Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_IDCUENTA) = "" Or Val(Grid.TextMatrix(Row, C_NIVEL)) <> gLastNivel Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Col = C_CODF22 Or Col = C_CODF29 Or Col = C_CODF22_14TER Then
      EdType = FEG_Edit
'      EdType = FEG_None  'por ahora no dejamos editar códigos formularios

   ElseIf Col = C_PARTIDA Then
      If Val(Grid.TextMatrix(Row, C_NIVEL)) = gLastNivel And Val(Grid.TextMatrix(Row, C_CLASCTA)) = CLASCTA_RESULTADO Then
         EdType = FEG_Edit
      Else
         MsgBeep vbExclamation
         Exit Sub
      End If
   
   Else
   
      If Col < C_INIATRIB Then
         MsgBeep vbExclamation
         Exit Sub
      End If
      
      If Col - C_INIATRIB + 1 = ATRIB_CAPITALPROPIO And Grid.TextMatrix(Row, C_CLASCTA) <> CLASCTA_ACTIVO And Grid.TextMatrix(Row, C_CLASCTA) <> CLASCTA_PASIVO And Grid.TextMatrix(Row, Col) = "" Then
         MsgBeep vbExclamation
         Exit Sub
      End If
   
      If Grid.TextMatrix(Row, Col) <> "" Then
         Grid.TextMatrix(Row, Col) = ""
      Else
         Grid.TextMatrix(Row, Col) = "x"
      End If
      
      Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)

   End If
   
End Sub

Private Sub Grid_DblClick_Old()
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lEdAtrib = False Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_IDCUENTA) = "" Or Val(Grid.TextMatrix(Row, C_NIVEL)) <> gLastNivel Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Col < C_INIATRIB Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Col - C_INIATRIB + 1 = ATRIB_CAPITALPROPIO And Grid.TextMatrix(Row, C_CLASCTA) <> CLASCTA_ACTIVO And Grid.TextMatrix(Row, C_CLASCTA) <> CLASCTA_PASIVO And Grid.TextMatrix(Row, Col) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
      
   If Grid.TextMatrix(Row, Col) <> "" Then
      Grid.TextMatrix(Row, Col) = ""
   Else
      Grid.TextMatrix(Row, Col) = "x"
   End If
   
   Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
End Sub

Private Sub SaveAll()
   Dim i As Integer, j As Integer
   Dim Q1 As String
   Dim Qcp As String
   Dim QCod As String
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_IDCUENTA) = "" Then
         Exit For
      End If
      
      Q1 = ""
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         For j = C_INIATRIB To Grid.Cols - 2
            Q1 = Q1 & ", Atrib" & j - C_INIATRIB + 1 & "=" & Abs(Grid.TextMatrix(i, j) <> "")
            
            If j - C_INIATRIB + 1 = ATRIB_CAPITALPROPIO Then
            
               If Abs(Grid.TextMatrix(i, j) <> "") Then
                  
                  If Val(Grid.TextMatrix(i, C_CLASCTA)) = CLASCTA_ACTIVO Then
                     Qcp = "UPDATE Cuentas SET TipoCapPropio = " & CAPPROPIO_ACTIVO_NORMAL
                     Qcp = Qcp & " WHERE IdCuenta=" & Grid.TextMatrix(i, C_IDCUENTA)
                     Qcp = Qcp & " AND (TipoCapPropio IS NULL OR NOT TipoCapPropio IN (" & CAPPROPIO_ACTIVO_NORMAL & "," & CAPPROPIO_ACTIVO_VALINTO & "," & CAPPROPIO_ACTIVO_COMPACTIVO & "))"
                     Qcp = Qcp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                  Else
                     Qcp = "UPDATE Cuentas SET TipoCapPropio = " & CAPPROPIO_PASIVO_EXIGIBLE
                     Qcp = Qcp & " WHERE IdCuenta=" & Grid.TextMatrix(i, C_IDCUENTA)
                     Qcp = Qcp & " AND (TipoCapPropio IS NULL OR NOT TipoCapPropio IN (" & CAPPROPIO_PASIVO_EXIGIBLE & "," & CAPPROPIO_PASIVO_NOEXIGIBLE & "))"
                     Qcp = Qcp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                  End If
                     
               Else
                  Qcp = "UPDATE Cuentas SET TipoCapPropio = 0"
                  Qcp = Qcp & " WHERE IdCuenta=" & Grid.TextMatrix(i, C_IDCUENTA)
                  Qcp = Qcp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                  
               End If
               
               Call ExecSQL(DbMain, Qcp)
            End If
         Next j
         
         QCod = ", CodF22 = " & Val(Grid.TextMatrix(i, C_CODF22))
         QCod = QCod & ", CodF22_14Ter = " & Val(Grid.TextMatrix(i, C_CODF22_14TER))
         QCod = QCod & ", CodF29 = " & Val(Grid.TextMatrix(i, C_CODF29))
         QCod = QCod & ", TipoPartida = " & Val(Grid.TextMatrix(i, C_PARTIDA))
         QCod = QCod & ", Percepcion = " & Abs(Grid.TextMatrix(i, C_PERCEPCION) <> "")
         
         
         Q1 = "UPDATE Cuentas SET " & Mid(Q1, 2) & QCod & " WHERE IdCuenta=" & Grid.TextMatrix(i, C_IDCUENTA)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next i

End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_CTAS) Then
      Bt_OK.Enabled = False
      lEdAtrib = False
   Else
      lEdAtrib = True
   End If
   
End Function


Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Dim Col As Integer
   
   Col = Grid.Col
   
   If Col = C_CODF22 Or Col = C_CODF29 Or Col = C_CODF22_14TER Or Col = C_PARTIDA Then
      Call KeyNumPos(KeyAscii)
   End If
   
   
End Sub
