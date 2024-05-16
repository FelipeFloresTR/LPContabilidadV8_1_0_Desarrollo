VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmSeguimientoComp 
   Caption         =   "Seguimiento de Comprobantes"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   18420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   -720
      TabIndex        =   0
      Top             =   -600
      Width           =   19215
      Begin VB.Frame Frame3 
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   17895
         Begin VB.CommandButton Bt_Search 
            Height          =   735
            Left            =   12480
            Picture         =   "FrmSeguimientoComp.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Tx_FechaOper 
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Bt_FechaOper 
            Height          =   315
            Index           =   1
            Left            =   5760
            Picture         =   "FrmSeguimientoComp.frx":0550
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Tx_NumComp 
            Height          =   315
            Left            =   9240
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   660
            Width           =   855
         End
         Begin VB.ComboBox Cb_Tipo 
            Height          =   315
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   660
            Width           =   1395
         End
         Begin VB.TextBox Tx_FechaOper 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton Bt_FechaOper 
            Height          =   315
            Index           =   0
            Left            =   3480
            Picture         =   "FrmSeguimientoComp.frx":05C5
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.ComboBox Cb_Oper 
            Height          =   315
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   240
            Width           =   1635
         End
         Begin VB.ComboBox Cb_Usuario 
            Height          =   315
            Left            =   10260
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Bt_FechaComp 
            Height          =   315
            Index           =   0
            Left            =   3480
            Picture         =   "FrmSeguimientoComp.frx":063A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   660
            Width           =   255
         End
         Begin VB.TextBox Tx_FechaComp 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   14
            Top             =   660
            Width           =   1275
         End
         Begin VB.CommandButton Bt_FechaComp 
            Height          =   315
            Index           =   1
            Left            =   5760
            Picture         =   "FrmSeguimientoComp.frx":06AF
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   660
            Width           =   255
         End
         Begin VB.TextBox Tx_FechaComp 
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   12
            Top             =   660
            Width           =   1335
         End
         Begin VB.ComboBox Cb_TipoAjuste 
            Height          =   315
            Left            =   10260
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   9
            Left            =   3840
            TabIndex        =   32
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comp.:"
            Height          =   195
            Index           =   5
            Left            =   6480
            TabIndex        =   31
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha operación desde:"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   30
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Operación:"
            Height          =   195
            Index           =   3
            Left            =   6480
            TabIndex        =   29
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Index           =   0
            Left            =   9540
            TabIndex        =   28
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha comprobante desde:"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   27
            Top             =   720
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   7
            Left            =   3840
            TabIndex        =   26
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Index           =   8
            Left            =   9000
            TabIndex        =   25
            Top             =   720
            Width           =   180
         End
      End
      Begin VB.Frame Fr_Botones 
         Height          =   555
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   17895
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
            Picture         =   "FrmSeguimientoComp.frx":0724
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Calendario"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Bt_DetComp 
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
            Left            =   1560
            Picture         =   "FrmSeguimientoComp.frx":0B4D
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Detalle comprobante seleccionado"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Bt_Orden 
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
            Left            =   2100
            Picture         =   "FrmSeguimientoComp.frx":0FB2
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ordenar listado por columna seleccionada"
            Top             =   120
            Width           =   375
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
            Left            =   960
            Picture         =   "FrmSeguimientoComp.frx":13A2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Copiar Excel"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Bt_Print 
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
            Left            =   540
            Picture         =   "FrmSeguimientoComp.frx":17E7
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir listado"
            Top             =   120
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
            Picture         =   "FrmSeguimientoComp.frx":1CA1
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Vista previa de la impresión"
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Bt_Close 
            Cancel          =   -1  'True
            Caption         =   "Cerrar"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   16560
            TabIndex        =   3
            Top             =   180
            Width           =   1215
         End
      End
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   3705
         Left            =   960
         TabIndex        =   1
         Top             =   2880
         Width           =   17865
         _ExtentX        =   31512
         _ExtentY        =   6535
         Cols            =   2
         Rows            =   3
         FixedCols       =   1
         FixedRows       =   1
         ScrollBars      =   3
         AllowUserResizing=   1
         HighLight       =   1
         SelectionMode   =   0
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   -1  'True
         Locked          =   0   'False
      End
   End
End
Attribute VB_Name = "FrmSeguimientoComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const C_IDCOMP = 0
Const C_FECHAHORA = 1
Const C_CORRELATIVO = 2
Const C_FECHA = 3
Const C_IDTIPO = 4
Const C_TIPO = 5
Const C_ESTADO = 6
Const C_GLOSA = 7
Const C_TOTALDEBE = 8
Const C_TOTALHABER = 9
Const C_IDUSUARIO = 10
Const C_USUARIO = 11
Const C_FECHACREACION = 12
Const C_IMPRESUMIDO = 13
Const C_ESCCMM = 14
Const C_FECHAIMPORT = 15
Const C_TIPOAJUSTE = 16
Const C_OTROSINGEG14TER = 17
Const C_ORIGEN = 18
Const C_QUERY = 19
Const C_VIGENTE = 20
Const C_FINGRESO = 21
Const C_AJUSTE = 22
Const NCOLS = C_AJUSTE

Dim G_IDDOC As Long
Dim lOrientacion As Integer

Const FI_MANUAL = 1
Const FI_IMPORTACION = 1

Const AJ_INSERTAR = 1
Const AJ_MODIFICAR = 2
Const AJ_ELIMINAR = 3

Dim lOrdenSel As Integer


Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Filtro As Boolean
   Dim F1 As Long, F2 As Long
   
   Filtro = False
   
   Q1 = "SELECT IdComp"
   Q1 = Q1 & " ,FechaHora"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Correlativo"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,TotalDebe"
   Q1 = Q1 & " ,TotalHaber"
   Q1 = Q1 & " ,TC.IdUsuario"
   Q1 = Q1 & " ,U.NombreLargo"
   Q1 = Q1 & " ,FechaCreacion"
   Q1 = Q1 & " ,ImpResumido"
   Q1 = Q1 & " ,EsCCMM"
   Q1 = Q1 & " ,FechaImport"
   Q1 = Q1 & " ,TipoAjuste"
   Q1 = Q1 & " ,OtrosIngEg14TER"
   Q1 = Q1 & " ,Origen"
   Q1 = Q1 & " ,Query"
   Q1 = Q1 & " ,IIF(Vigente IS NULL OR VIGENTE = 1, 'VIGENTE', 'ELIMINADO') AS Vigente2"
   Q1 = Q1 & " ,FormaIngreso"
   Q1 = Q1 & " ,Ajuste"
   Q1 = Q1 & " FROM Tracking_Comprobante TC"
   Q1 = Q1 & " INNER JOIN Usuarios U ON U.IdUsuario = TC.IdUsuario"
   Q1 = Q1 & " WHERE TC.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND TC.Ano = " & gEmpresa.Ano
   
   If Tx_NumComp.Text <> "" Then
    Q1 = Q1 & " AND TC.IdComp = " & Tx_NumComp.Text
    Filtro = True
   End If
   If ItemData(Cb_Tipo) > 0 Then
    Q1 = Q1 & " AND TC.TIPO = " & ItemData(Cb_Tipo)
    Filtro = True
   End If
   
   If ItemData(Cb_Usuario) > 0 Then
    Q1 = Q1 & " AND TC.IdUsuario = " & ItemData(Cb_Usuario)
    Filtro = True
   End If
   
   F1 = GetTxDate(Tx_FechaOper(0))
   F2 = GetTxDate(Tx_FechaOper(1))
   
   If F1 <> 0 And F2 <> 0 Then
     If gDbType = SQL_ACCESS Then
      Q1 = Q1 & " AND  TC.FechaHora  BETWEEN " & F1 & " AND " & F2
     Else
        Q1 = Q1 & " AND  Cast(fechahora as int) + 1  BETWEEN " & F1 & " AND " & F2
     End If
   End If
   
   F1 = GetTxDate(Tx_FechaComp(0))
   F2 = GetTxDate(Tx_FechaComp(1))
   
   If F1 <> 0 And F2 <> 0 Then
      Q1 = Q1 & " AND TC.FechaCreacion BETWEEN " & F1 & " AND " & F2
   End If
   
   If CbItemData(Cb_Oper) > 0 Then
     If CbItemData(Cb_Oper) > 3 Then
        Q1 = Q1 & " AND TC.FIngreso = " & CbItemData(Cb_Oper)
     Else
        Q1 = Q1 & " AND TC.Ajuste = " & CbItemData(Cb_Oper)
     End If
   End If
   If ItemData(Cb_TipoAjuste) > 0 Then
    Q1 = Q1 & " AND TC.TipoAjuste = " & ItemData(Cb_TipoAjuste)
    Filtro = True
   End If
'   If ItemData(Cb_Estado) > 0 Then
'    Q1 = Q1 & " AND TD.ESTADO = " & ItemData(Cb_Estado)
'    Filtro = True
'   End If
   Q1 = Q1 & " ORDER BY FechaHora DESC"
   
   If Not Filtro Then
    MsgBox "Favor ingresar al menos un filtro", vbInformation, "Seguimiento de Documento"
    Exit Sub
   End If
   
   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   If Rs.EOF = False Then
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
        
        Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
        Grid.TextMatrix(i, C_FECHAHORA) = vFld(Rs("FechaHora"))
        Grid.TextMatrix(i, C_CORRELATIVO) = vFld(Rs("Correlativo"))
        Grid.TextMatrix(i, C_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
        Grid.TextMatrix(i, C_TIPO) = gTipoComp(vFld(Rs("Tipo")))
        Grid.TextMatrix(i, C_ESTADO) = gEstadoComp(vFld(Rs("Estado"))) 'vFld(Rs("Estado"))
        Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"))
        Grid.TextMatrix(i, C_TOTALDEBE) = Format(vFld(Rs("TotalDebe")), NUMFMT)
        Grid.TextMatrix(i, C_TOTALHABER) = Format(vFld(Rs("TotalHaber")), NUMFMT)
        Grid.TextMatrix(i, C_IDUSUARIO) = vFld(Rs("IdUsuario"))
        Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("NombreLargo"))
        Grid.TextMatrix(i, C_FECHACREACION) = Format(vFld(Rs("FechaCreacion")), SDATEFMT)
        Grid.TextMatrix(i, C_IMPRESUMIDO) = vFld(Rs("ImpResumido"))
        Grid.TextMatrix(i, C_ESCCMM) = vFld(Rs("EsCCMM"))
        Grid.TextMatrix(i, C_FECHAIMPORT) = Format(vFld(Rs("FechaImport")), SDATEFMT)
        Grid.TextMatrix(i, C_TIPOAJUSTE) = Left(gTipoAjuste(vFld(Rs("TipoAjuste"))), 1) 'vFld(Rs("TipoAjuste"))
        Grid.TextMatrix(i, C_OTROSINGEG14TER) = Format(vFld(Rs("OtrosIngEg14TER")), NUMFMT)
        Grid.TextMatrix(i, C_VIGENTE) = vFld(Rs("Vigente2"))
        Grid.TextMatrix(i, C_FINGRESO) = FormaIngreso(vFld(Rs("FormaIngreso")))
        Grid.TextMatrix(i, C_AJUSTE) = Ajuste(vFld(Rs("Ajuste")))

    Rs.MoveNext
      i = i + 1
   Loop
   Else
        MsgBox "NO se encontraron documentos", vbInformation, "Seguimiento de Documento"
   End If
   
   lOrdenSel = C_FECHA
   
   
End Sub
Private Function FormaIngreso(valor As Long) As String

    Select Case valor
        Case FI_MANUAL
            FormaIngreso = "Manual"
            
        Case FI_IMPORTACION
            FormaIngreso = "Importado"
               
    End Select

End Function

Private Function Ajuste(valor As Long) As String

    Select Case valor
        Case AJ_INSERTAR
            Ajuste = "Creado"
            
        Case AJ_MODIFICAR
            Ajuste = "Modificado"
            
        Case AJ_ELIMINAR
            Ajuste = "Eliminado"
               
    End Select

End Function
   

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
End Sub

Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
Call FGr2Clip(Grid, Me.Caption)
End Sub

Private Sub Bt_DetComp_Click()
Call ViewDetComp(Grid.Row, Grid.Col)
End Sub
Private Sub ViewDetComp(ByVal Row As Integer, ByVal Col As Integer)
   Dim IdComp As Long
   Dim Frm As FrmComprobante

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   IdComp = Val(Grid.TextMatrix(Row, C_IDCOMP))

   If IdComp <> 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(IdComp, False)
      Set Frm = Nothing
      
   ElseIf Grid.TextMatrix(Row, C_IDCOMP) = "0" Then
      MsgBox1 "Este comprobante ha sido eliminado.", vbExclamation + vbOKOnly
      
   End If
            
End Sub

Private Sub Bt_FechaComp_Click(Index As Integer)
   Dim Frm As FrmCalendar
   Dim F1 As Long

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaComp(Index))
   Set Frm = Nothing
   
   If Index = 0 And GetTxDate(Tx_FechaComp(1)) < GetTxDate(Tx_FechaComp(0)) Then
      F1 = DateAdd("m", 1, GetTxDate(Tx_FechaComp(0)))
      Call SetTxDate(Tx_FechaComp(1), F1)
   End If
   
   Call EnableFrm(True)
End Sub

Private Sub Bt_FechaOper_Click(Index As Integer)
   Dim Frm As FrmCalendar
   Dim F1 As Long

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaOper(Index))
   Set Frm = Nothing
   
   If Index = 0 And GetTxDate(Tx_FechaOper(1)) < GetTxDate(Tx_FechaOper(0)) Then
      F1 = DateAdd("m", 1, GetTxDate(Tx_FechaOper(0)))
      Call SetTxDate(Tx_FechaOper(1), F1)
   End If

   Call EnableFrm(True)
End Sub

Private Sub Bt_Orden_Click()
Call OrdenaPorCol(Grid.Col)
End Sub
Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadGrid 'LoadAll
      
   Me.MousePointer = vbDefault
      
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
End Sub

Private Sub Bt_Print_Click()
Dim OldOrientation As Integer
   
'   If Bt_Search.Enabled = True Then
'      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
'      Exit Sub
'   End If
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
End Sub

Private Sub Bt_Search_Click()
Call SetUpGrid
Call LoadGrid
End Sub
Private Sub Form_Load()

   lOrientacion = ORIENT_HOR
   Cb_Tipo.Clear
   Cb_Tipo.AddItem ""
   Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = 0
   
   
   For i = 1 To N_TIPOCOMP
   
'      If Not (lOper = O_NEW And i = TC_APERTURA) Then
         Cb_Tipo.AddItem gTipoComp(i)
         Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
'      End If
'
'      If i <> TC_APERTURA Then
'         Cb_Tipo.AddItem gTipoComp(i)
'         Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
'      End If
         
   Next i
   Call FillCb

End Sub
Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   
End Sub

Private Sub FillCb()
   Dim i As Integer, Q1 As String
      
'   Call CbAddItem(Cb_Tipo, "(todos)", -1)
'   For i = 1 To N_TIPOCOMP
'      Call CbAddItem(Cb_Tipo, gTipoComp(i), i)
'   Next i
'   Cb_Tipo.ListIndex = 0
               
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_AMBOS), TAJUSTE_AMBOS)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_AMBOS)
               
   Call CbAddItem(Cb_Oper, "(todas)", -1)
   Call CbAddItem(Cb_Oper, "Crear", O_NEW)
   Call CbAddItem(Cb_Oper, "Modificar", O_EDIT)
   Call CbAddItem(Cb_Oper, "Eliminar", O_DELETE)
   Call CbAddItem(Cb_Oper, "Importar", O_IMPORT)
   Cb_Oper.ListIndex = 0
   
   Call CbAddItem(Cb_Usuario, "(todos)", -1)
   Q1 = "SELECT Usuario, IdUsuario FROM Usuarios ORDER BY Usuario"
   Call FillCombo(Cb_Usuario, DbMain, Q1, -1)
   Cb_Usuario.ListIndex = 0
   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    

    Grid.ColWidth(C_IDCOMP) = 0
    Grid.ColWidth(C_FECHAHORA) = 1750
    Grid.ColWidth(C_CORRELATIVO) = 900
    Grid.ColWidth(C_FECHA) = 900
    Grid.ColWidth(C_IDTIPO) = 0
    Grid.ColWidth(C_TIPO) = 900
    Grid.ColWidth(C_ESTADO) = 900
    Grid.ColWidth(C_GLOSA) = 2200
    Grid.ColWidth(C_TOTALDEBE) = 1200
    Grid.ColWidth(C_TOTALHABER) = 1200
    Grid.ColWidth(C_IDUSUARIO) = 0
    Grid.ColWidth(C_USUARIO) = 1200
    Grid.ColWidth(C_FECHACREACION) = 1300
    Grid.ColWidth(C_IMPRESUMIDO) = 1200
    Grid.ColWidth(C_ESCCMM) = 900
    Grid.ColWidth(C_FECHAIMPORT) = 1300
    Grid.ColWidth(C_TIPOAJUSTE) = 900
    Grid.ColWidth(C_OTROSINGEG14TER) = 1200
    Grid.ColWidth(C_ORIGEN) = 0
    Grid.ColWidth(C_QUERY) = 0
    Grid.ColWidth(C_VIGENTE) = 900
    Grid.ColWidth(C_FINGRESO) = 1200
    Grid.ColWidth(C_AJUSTE) = 900
   
   
    Grid.ColAlignment(C_IDCOMP) = flexAlignRightCenter
    Grid.ColAlignment(C_FECHAHORA) = flexAlignRightCenter
    Grid.ColAlignment(C_CORRELATIVO) = flexAlignRightCenter
    Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
    Grid.ColAlignment(C_TIPO) = flexAlignRightCenter
    Grid.ColAlignment(C_ESTADO) = flexAlignRightCenter
    Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
    Grid.ColAlignment(C_TOTALDEBE) = flexAlignLeftCenter
    Grid.ColAlignment(C_TOTALHABER) = flexAlignLeftCenter
    Grid.ColAlignment(C_USUARIO) = flexAlignRightCenter
    Grid.ColAlignment(C_FECHACREACION) = flexAlignRightCenter
    Grid.ColAlignment(C_IMPRESUMIDO) = flexAlignLeftCenter
    Grid.ColAlignment(C_ESCCMM) = flexAlignRightCenter
    Grid.ColAlignment(C_FECHAIMPORT) = flexAlignRightCenter
    Grid.ColAlignment(C_TIPOAJUSTE) = flexAlignRightCenter
    Grid.ColAlignment(C_OTROSINGEG14TER) = flexAlignLeftCenter
    Grid.ColAlignment(C_VIGENTE) = flexAlignLeftCenter
    Grid.ColAlignment(C_FINGRESO) = flexAlignLeftCenter
    Grid.ColAlignment(C_AJUSTE) = flexAlignLeftCenter
   
   
  
    Grid.TextMatrix(0, C_IDCOMP) = "Id Comprobante"
    Grid.TextMatrix(0, C_FECHAHORA) = "Fecha y Hora"
    Grid.TextMatrix(0, C_CORRELATIVO) = "Correlativo"
    Grid.TextMatrix(0, C_TIPO) = "Tipo"
    Grid.TextMatrix(0, C_FECHA) = "Fecha"
    Grid.TextMatrix(0, C_TIPO) = "Tipo"
    Grid.TextMatrix(0, C_ESTADO) = "Estado"
    Grid.TextMatrix(0, C_GLOSA) = "Glosa"
    Grid.TextMatrix(0, C_TOTALDEBE) = "Debe"
    Grid.TextMatrix(0, C_TOTALHABER) = "Haber"
    Grid.TextMatrix(0, C_USUARIO) = "Usuario"
    Grid.TextMatrix(0, C_FECHACREACION) = "Fecha Creacion"
    Grid.TextMatrix(0, C_IMPRESUMIDO) = "Imp Resumido"
    Grid.TextMatrix(0, C_ESCCMM) = "Esccmm"
    Grid.TextMatrix(0, C_FECHAIMPORT) = "Fecha Import"
    Grid.TextMatrix(0, C_TIPOAJUSTE) = "Tipo Ajuste"
    Grid.TextMatrix(0, C_OTROSINGEG14TER) = "Otros Ing 14T"
    Grid.TextMatrix(0, C_VIGENTE) = "Vigente"
    Grid.TextMatrix(0, C_FINGRESO) = "Forma Ingreso"
    Grid.TextMatrix(0, C_AJUSTE) = "Ajuste"

    
   Call FGrSetup(Grid)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub


Private Sub Grid_DblClick()
   Dim Frm As FrmSeguimientoMovComp
   
   Set Frm = New FrmSeguimientoMovComp
   Frm.FSearch (Val(Grid.TextMatrix(Grid.Row, C_ID)))
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   
   gPrtReportes.Titulos = Titulos
      
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
'   Total(C_DESC) = "Total"
'   Total(C_DEBE) = Tx_TotDebe
'   Total(C_HABER) = Tx_TotHaber
   
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   'gPrtReportes.ColObligatoria = C_CUENTA
   gPrtReportes.NTotLines = 1
   

End Sub



Private Sub Tx_FechaComp_Change(Index As Integer)
   Dim F1 As Long
   
   If Index = 0 And GetTxDate(Tx_FechaComp(1)) < GetTxDate(Tx_FechaComp(0)) Then
      F1 = DateAdd("m", 1, GetTxDate(Tx_FechaComp(0)))
      Call SetTxDate(Tx_FechaComp(1), F1)
   End If
   
   Call EnableFrm(True)
End Sub

Private Sub Tx_FechaComp_GotFocus(Index As Integer)
Call DtGotFocus(Tx_FechaComp(Index))
End Sub

Private Sub Tx_FechaComp_KeyPress(Index As Integer, KeyAscii As Integer)
Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_FechaComp_LostFocus(Index As Integer)
   If Trim$(Tx_FechaComp(Index)) = "" Then
      Exit Sub
   End If
   Call DtLostFocus(Tx_FechaComp(Index))
End Sub

Private Sub Tx_FechaOper_Change(Index As Integer)
   Dim F1 As Long
   
   If Index = 0 And GetTxDate(Tx_FechaOper(1)) < GetTxDate(Tx_FechaOper(0)) Then
      F1 = DateAdd("m", 1, GetTxDate(Tx_FechaOper(0)))
      Call SetTxDate(Tx_FechaOper(1), F1)
   End If
   
   Call EnableFrm(True)
End Sub

Private Sub Tx_FechaOper_GotFocus(Index As Integer)
 Call DtGotFocus(Tx_FechaOper(Index))
End Sub

Private Sub Tx_FechaOper_KeyPress(Index As Integer, KeyAscii As Integer)
Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_FechaOper_LostFocus(Index As Integer)
   If Trim$(Tx_FechaOper(Index)) = "" Then
      Exit Sub
   End If
   Call DtLostFocus(Tx_FechaOper(Index))
End Sub

