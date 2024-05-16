VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmSeguimientoDoc 
   Caption         =   "Seguimiento de Documentos"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
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
         TabIndex        =   7
         Top             =   1200
         Width           =   17895
         Begin VB.CommandButton Bt_Search 
            Caption         =   "&Listar"
            Height          =   675
            Left            =   15840
            Picture         =   "FrmSeguimientoDoc.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Tx_NumDoc 
            Height          =   315
            Left            =   13260
            MaxLength       =   15
            TabIndex        =   30
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox Cb_Estado 
            Height          =   315
            Left            =   13320
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox Cb_TipoDoc 
            Height          =   315
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox Cb_TipoLib 
            Height          =   315
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox Tx_FechaComp 
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   17
            Top             =   660
            Width           =   1335
         End
         Begin VB.CommandButton Bt_FechaComp 
            Height          =   315
            Index           =   1
            Left            =   5760
            Picture         =   "FrmSeguimientoDoc.frx":043E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   660
            Width           =   255
         End
         Begin VB.TextBox Tx_FechaComp 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   15
            Top             =   660
            Width           =   1275
         End
         Begin VB.CommandButton Bt_FechaComp 
            Height          =   315
            Index           =   0
            Left            =   3480
            Picture         =   "FrmSeguimientoDoc.frx":04B3
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   660
            Width           =   255
         End
         Begin VB.ComboBox Cb_Usuario 
            Height          =   315
            Left            =   10620
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox Cb_Oper 
            Height          =   315
            Left            =   10620
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   720
            Width           =   1635
         End
         Begin VB.CommandButton Bt_FechaOper 
            Height          =   315
            Index           =   0
            Left            =   3480
            Picture         =   "FrmSeguimientoDoc.frx":0528
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Tx_FechaOper 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   10
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton Bt_FechaOper 
            Height          =   315
            Index           =   1
            Left            =   5760
            Picture         =   "FrmSeguimientoDoc.frx":059D
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Tx_FechaOper 
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N° Doc.:"
            Height          =   195
            Index           =   1
            Left            =   12480
            TabIndex        =   31
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Index           =   6
            Left            =   12600
            TabIndex        =   29
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Doc.:"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   27
            Top             =   780
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Libro.:"
            Height          =   195
            Index           =   2
            Left            =   6120
            TabIndex        =   25
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   10
            Left            =   3840
            TabIndex        =   23
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha comprobante desde:"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   22
            Top             =   720
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Index           =   7
            Left            =   9780
            TabIndex        =   21
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Operación:"
            Height          =   195
            Index           =   3
            Left            =   9720
            TabIndex        =   20
            Top             =   780
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha operación desde:"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   19
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   9
            Left            =   3840
            TabIndex        =   18
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.Frame Fr_Botones 
         Height          =   555
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   17895
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
            Left            =   1980
            Picture         =   "FrmSeguimientoDoc.frx":0612
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Ordenar listado por columna seleccionada"
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
            Left            =   1440
            Picture         =   "FrmSeguimientoDoc.frx":0A02
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Detalle comprobante seleccionado"
            Top             =   120
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
            Left            =   2520
            Picture         =   "FrmSeguimientoDoc.frx":0E67
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Calendario"
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
            Picture         =   "FrmSeguimientoDoc.frx":1290
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
            Picture         =   "FrmSeguimientoDoc.frx":16D5
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
            Picture         =   "FrmSeguimientoDoc.frx":1B8F
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
Attribute VB_Name = "FrmSeguimientoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const C_IDDOC = 0
Const C_FECHAHORA = 1
Const C_IDTIPOLIB = 2
Const C_TIPOLIB = 3
Const C_IDTIPODOC = 4
Const C_TIPODOC = 5
Const C_NUMDOC = 6
Const C_DTE = 7
Const C_DELGIRO = 8
Const C_IDENTIDAD = 9
Const C_RUTENTIDAD = 10
Const C_ENTIDAD = 11
Const C_FEMISION = 12
Const C_FEVENC = 13
Const C_DESCRIPCION = 14
Const C_ESTADO = 15
Const C_EXCENTO = 16
Const C_CTAEXCENTO = 17
Const C_AFECTO = 18
Const C_CTAAFECTO = 19
Const C_IVA = 20
Const C_CTAIVA = 21
Const C_OTROIMP = 22
Const C_CTAOTROIMP = 23
Const C_TOTAL = 24
Const C_CTATOTAL = 25
Const C_USUARIO = 26
Const C_FCREACION = 27
Const C_SALDODOC = 28
Const C_GIRO = 29
Const C_VIGENTE = 30
Const C_FINGRESO = 31
Const C_AJUSTE = 32
Const NCOLS = C_AJUSTE

Dim G_IDDOC As Long
Dim lOrientacion As Integer

Const FI_MANUAL = 1
Const FI_IMPORTACION = 2

Const AJ_INSERTAR = 1
Const AJ_MODIFICAR = 2
Const AJ_ELIMINAR = 3


Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Filtro As Boolean
   Dim F1 As Long, F2 As Long
   
   Filtro = False
   
   Q1 = "SELECT IdDoc ,FechaHora ,TD.TipoLib ,TD.TipoDoc ,NumDoc, E.Rut ,E.Nombre as Entidad ,FEmision ,FVenc ,Descrip ,TD.Estado ,Exento ,C.Descripcion as CTAExento ,Afecto ,CA.Descripcion as CTAAfecto ,IVA ,CI.Descripcion as CTAIVA ,OtroImp ,CO.Descripcion as CTAOTROIMP ,Total ,CT.Descripcion as CTATOTAL ,U.NombreLargo AS USUARIO ,FechaCreacion ,FEmisionOri ,SaldoDoc ,PorcentRetencion ,TipoRetencion ,TD.Giro ,TD.PropIVA ,ValIVAIrrec ,IIF(Vigente IS NULL OR Vigente = 1, 'ACTIVO', 'ELIMINADO') AS Vigente2, FormaIngreso, Ajuste, DTE, TD.Giro "
   Q1 = Q1 & " FROM ((((((((Tracking_Documento AS TD"
   Q1 = Q1 & " LEFT JOIN PARAM as P ON   P.Codigo = TD.TipoLib)"
   'Q1 = Q1 & " LEFT JOIN TipoDocs TDO ON TDO.Id = TD.TipoDoc AND TDO.TipoLib = TD.TipoLib)"
   Q1 = Q1 & " LEFT JOIN Entidades E ON E.IdEntidad = TD.IdEntidad AND E.IdEmpresa = TD.IdEmpresa)"
   Q1 = Q1 & " LEFT JOIN Cuentas  C ON C.idCuenta = TD.IdCuentaExento AND C.IdEmpresa = TD.IdEmpresa AND C.Ano = TD.Ano)"
   Q1 = Q1 & " LEFT JOIN Cuentas  CA ON CA.idCuenta = TD.IdCuentaAfecto AND CA.IdEmpresa = TD.IdEmpresa AND CA.Ano = TD.Ano)"
   Q1 = Q1 & " LEFT JOIN Cuentas CI ON CI.idCuenta = TD.IdCuentaIVA AND CI.IdEmpresa = TD.IdEmpresa AND CI.Ano = TD.Ano)"
   Q1 = Q1 & " LEFT JOIN Cuentas CO ON CO.idCuenta = TD.IdCuentaOtroImp AND CO.IdEmpresa = TD.IdEmpresa AND CO.Ano = TD.Ano)"
   Q1 = Q1 & " LEFT JOIN Cuentas CT ON CT.idCuenta = TD.IdCuentaTotal AND CT.IdEmpresa = TD.IdEmpresa AND CT.Ano = TD.Ano)"
   Q1 = Q1 & " LEFT JOIN Usuarios U ON U.IdUsuario = TD.IdUsuario)"
   'Q1 = Q1 & " WHERE P.Tipo = 'TIPOLIB'"
   Q1 = Q1 & " WHERE TD.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND TD.Ano = " & gEmpresa.Ano
   
   If Tx_NumDoc.Text <> "" Then
    Q1 = Q1 & " AND NumDoc = '" & Tx_NumDoc.Text & "'"
    Filtro = True
   End If
   If ItemData(Cb_TipoLib) > 0 Then
    Q1 = Q1 & " AND TD.TIPOLIB = " & ItemData(Cb_TipoLib)
    Filtro = True
   End If
   If ItemData(Cb_TipoDoc) > 0 Then
    Q1 = Q1 & " AND TD.TIPODOC = " & ItemData(Cb_TipoDoc)
    Filtro = True
   End If
   If ItemData(Cb_Estado) > 0 Then
    Q1 = Q1 & " AND TD.ESTADO = " & ItemData(Cb_Estado)
    Filtro = True
   End If
   
   If ItemData(Cb_Usuario) > 0 Then
    Q1 = Q1 & " AND TD.IdUsuario = " & ItemData(Cb_Usuario)
    Filtro = True
   End If
   
   F1 = GetTxDate(Tx_FechaOper(0))
   F2 = GetTxDate(Tx_FechaOper(1))
   
   If F1 <> 0 And F2 <> 0 Then
     If gDbType = SQL_ACCESS Then
      Q1 = Q1 & " AND  TD.FechaHora  BETWEEN " & F1 & " AND " & F2
     Else
        Q1 = Q1 & " AND  Cast(TD.FechaHora as int) + 1  BETWEEN " & F1 & " AND " & F2
     End If
   End If
   
   F1 = GetTxDate(Tx_FechaComp(0))
   F2 = GetTxDate(Tx_FechaComp(1))
   
   If F1 <> 0 And F2 <> 0 Then
      Q1 = Q1 & " AND TD.Femision BETWEEN " & F1 & " AND " & F2
   End If
   
    If CbItemData(Cb_Oper) > 0 Then
     If CbItemData(Cb_Oper) > 3 Then
        Q1 = Q1 & " AND TD.FIngreso = " & CbItemData(Cb_Oper)
     Else
        Q1 = Q1 & " AND TD.Ajuste = " & CbItemData(Cb_Oper)
     End If
   End If
   
   
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
      
        
        Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
        Grid.TextMatrix(i, C_FECHAHORA) = vFld(Rs("FechaHora"))
        Grid.TextMatrix(i, C_TIPOLIB) = Left(ReplaceStr(gTipoLibNew(IIf(vFld(Rs("TipoLib")) = 8, 6, vFld(Rs("TipoLib")))).Nombre, "Libro de ", ""), 9) 'vFld(Rs("TipoLib"))
        Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) 'vFld(Rs("TipoDoc"))
        Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
        Grid.TextMatrix(i, C_DTE) = IIf(vFld(Rs("DTE")) <> 0, "Si", "No")
        Grid.TextMatrix(i, C_DELGIRO) = IIf(vFld(Rs("Giro")) <> 0, "Si", "No")
        Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Entidad"))
        Grid.TextMatrix(i, C_RUTENTIDAD) = FmtCID(vFld(Rs("Rut")), True)
        Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FEmision")), SDATEFMT)
        Grid.TextMatrix(i, C_FEVENC) = Format(vFld(Rs("FVenc")), SDATEFMT)
        Grid.TextMatrix(i, C_DESCRIPCION) = vFld(Rs("Descrip"))
        Grid.TextMatrix(i, C_ESTADO) = Left(gEstadoDoc(vFld(Rs("Estado"))), 9)
        Grid.TextMatrix(i, C_EXCENTO) = Format(vFld(Rs("Exento")), NUMFMT)
        Grid.TextMatrix(i, C_CTAEXCENTO) = vFld(Rs("CTAExento"))
        Grid.TextMatrix(i, C_AFECTO) = Format(vFld(Rs("Afecto")), NUMFMT)
        Grid.TextMatrix(i, C_CTAAFECTO) = vFld(Rs("CTAAfecto"))
        Grid.TextMatrix(i, C_IVA) = Format(vFld(Rs("IVA")), NUMFMT)
        Grid.TextMatrix(i, C_CTAIVA) = vFld(Rs("CTAIVA"))
        Grid.TextMatrix(i, C_OTROIMP) = Format(vFld(Rs("OtroImp")), NUMFMT)
        Grid.TextMatrix(i, C_CTAOTROIMP) = vFld(Rs("CTAOTROIMP"))
        Grid.TextMatrix(i, C_TOTAL) = Format(vFld(Rs("Total")), NUMFMT)
        Grid.TextMatrix(i, C_CTATOTAL) = vFld(Rs("CTATOTAL"))
        Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("USUARIO"))
        Grid.TextMatrix(i, C_FCREACION) = Format(vFld(Rs("FechaCreacion")), SDATEFMT)
        Grid.TextMatrix(i, C_SALDODOC) = Format(vFld(Rs("SaldoDoc")), NUMFMT)
        Grid.TextMatrix(i, C_GIRO) = vFld(Rs("Giro"))
        Grid.TextMatrix(i, C_VIGENTE) = vFld(Rs("Vigente2"))
        Grid.TextMatrix(i, C_FINGRESO) = FormaIngreso(vFld(Rs("FormaIngreso")))
        Grid.TextMatrix(i, C_AJUSTE) = Ajuste(vFld(Rs("Ajuste")))
      

      
    Rs.MoveNext
      i = i + 1
   Loop
   Else
        MsgBox "NO se encontraron documentos", vbInformation, "Seguimiento de Documento"
   End If
   
   
   
   
End Sub
   

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
Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   
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
Private Sub FillCb()
   Dim i As Integer, Q1 As String
      
'   Call CbAddItem(Cb_Tipo, "(todos)", -1)
'   For i = 1 To N_TIPOCOMP
'      Call CbAddItem(Cb_Tipo, gTipoComp(i), i)
'   Next i
'   Cb_Tipo.ListIndex = 0
               
               
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

Private Sub Cb_TipoLib_Click()
Dim i As Integer
   Dim TipoLib As Integer
   
   Cb_TipoDoc.Clear
   
   TipoLib = ItemData(Cb_TipoLib)
   
   If TipoLib > 0 Then
   
      Call FillTipoDoc(Cb_TipoDoc, TipoLib, True, True)
      Cb_TipoDoc.ListIndex = -1
      
      If (TipoLib = LIB_OTROS Or TipoLib = LIB_REMU) And Cb_Estado.ListCount > 0 Then    'dejamos sin selección de estado
         Cb_Estado.ListIndex = 0
      End If
   End If
End Sub

Private Sub Form_Load()

   lOrientacion = ORIENT_HOR
   Cb_TipoLib.AddItem ""
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = 0
   For i = 1 To UBound(gTipoLibNew)
      Cb_TipoLib.AddItem ReplaceStr(gTipoLibNew(i).Nombre, "Libro de ", "")
      Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).id 'i
   Next i
   
   Cb_Estado.AddItem ""
   Cb_Estado.ItemData(Cb_Estado.NewIndex) = 0
   
   For i = 1 To MAX_ESTADODOC
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i

   Call FillCb
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_FECHAHORA) = 1750
   Grid.ColWidth(C_IDTIPOLIB) = 0
   Grid.ColWidth(C_TIPOLIB) = 1600
   Grid.ColWidth(C_IDTIPODOC) = 0
   Grid.ColWidth(C_TIPODOC) = 1500
   Grid.ColWidth(C_NUMDOC) = 900
   Grid.ColWidth(C_DTE) = 500
   Grid.ColWidth(C_DELGIRO) = 800
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_RUTENTIDAD) = 1300
   Grid.ColWidth(C_ENTIDAD) = 1700
   Grid.ColWidth(C_FEMISION) = 900
   Grid.ColWidth(C_FEVENC) = 950
   Grid.ColWidth(C_DESCRIPCION) = 1700
   Grid.ColWidth(C_ESTADO) = 900
   Grid.ColWidth(C_EXCENTO) = 1200
   Grid.ColWidth(C_CTAEXCENTO) = 1200
   Grid.ColWidth(C_AFECTO) = 1200
   Grid.ColWidth(C_CTAAFECTO) = 1200
   Grid.ColWidth(C_IVA) = 1200
   Grid.ColWidth(C_CTAIVA) = 1200
   Grid.ColWidth(C_OTROIMP) = 1200
   Grid.ColWidth(C_CTAOTROIMP) = 1200
   Grid.ColWidth(C_TOTAL) = 1200
   Grid.ColWidth(C_CTATOTAL) = 1200
   Grid.ColWidth(C_USUARIO) = 700
   Grid.ColWidth(C_FCREACION) = 900
   Grid.ColWidth(C_SALDODOC) = 1200
   Grid.ColWidth(C_GIRO) = 0
   Grid.ColWidth(C_VIGENTE) = 900
   Grid.ColWidth(C_FINGRESO) = 900
   Grid.ColWidth(C_AJUSTE) = 900
   
   Grid.ColAlignment(C_IDDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHAHORA) = flexAlignRightCenter
   Grid.ColAlignment(C_IDTIPOLIB) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPOLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_IDTIPODOC) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_DTE) = flexAlignCenterCenter
   Grid.ColAlignment(C_DELGIRO) = flexAlignCenterCenter
   Grid.ColAlignment(C_IDENTIDAD) = flexAlignRightCenter
   Grid.ColAlignment(C_RUTENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
   Grid.ColAlignment(C_FEVENC) = flexAlignRightCenter
   Grid.ColAlignment(C_DESCRIPCION) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   Grid.ColAlignment(C_EXCENTO) = flexAlignRightCenter
   Grid.ColAlignment(C_CTAEXCENTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_AFECTO) = flexAlignRightCenter
   Grid.ColAlignment(C_CTAAFECTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_IVA) = flexAlignRightCenter
   Grid.ColAlignment(C_CTAIVA) = flexAlignLeftCenter
   Grid.ColAlignment(C_OTROIMP) = flexAlignRightCenter
   Grid.ColAlignment(C_CTAOTROIMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_CTATOTAL) = flexAlignLeftCenter
   Grid.ColAlignment(C_USUARIO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FCREACION) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDODOC) = flexAlignRightCenter
   Grid.ColAlignment(C_GIRO) = flexAlignLeftCenter
   Grid.ColAlignment(C_IVA) = flexAlignRightCenter
   Grid.ColAlignment(C_VIGENTE) = flexAlignLeftCenter
   Grid.ColAlignment(C_FINGRESO) = flexAlignLeftCenter
   Grid.ColAlignment(C_AJUSTE) = flexAlignLeftCenter
   
   
   Grid.TextMatrix(0, C_FECHAHORA) = "Fecha y Hora"
   Grid.TextMatrix(0, C_TIPOLIB) = "Tipo Libro"
   Grid.TextMatrix(0, C_TIPODOC) = "Tipo Doc."
   Grid.TextMatrix(0, C_NUMDOC) = "Num. Doc."
   Grid.TextMatrix(0, C_DTE) = "DTE."
   Grid.TextMatrix(0, C_DELGIRO) = "Del Giro."
   Grid.TextMatrix(0, C_RUTENTIDAD) = "Rut"
   Grid.TextMatrix(0, C_ENTIDAD) = "Entidad"
   Grid.TextMatrix(0, C_FEMISION) = "F Emision"
   Grid.TextMatrix(0, C_FEVENC) = "F Vencimiento"
   Grid.TextMatrix(0, C_DESCRIPCION) = "Descripcion"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"
   Grid.TextMatrix(0, C_EXCENTO) = "Exento"
   Grid.TextMatrix(0, C_CTAEXCENTO) = "Cta Exento"
   Grid.TextMatrix(0, C_AFECTO) = "Afecto"
   Grid.TextMatrix(0, C_CTAAFECTO) = "Cta Afecto"
   Grid.TextMatrix(0, C_IVA) = "IVA"
   Grid.TextMatrix(0, C_CTAIVA) = "Cta IVA"
   Grid.TextMatrix(0, C_OTROIMP) = "Otro Impu."
   Grid.TextMatrix(0, C_CTAOTROIMP) = "Cta Otro Impu."
   Grid.TextMatrix(0, C_TOTAL) = "Total"
   Grid.TextMatrix(0, C_CTATOTAL) = "Cta Total"
   Grid.TextMatrix(0, C_USUARIO) = "Usuario"
   Grid.TextMatrix(0, C_FCREACION) = "F. Creacion"
   Grid.TextMatrix(0, C_SALDODOC) = "Saldo"
   Grid.TextMatrix(0, C_GIRO) = "Giro"
   Grid.TextMatrix(0, C_VIGENTE) = "Vigente"
   Grid.TextMatrix(0, C_FINGRESO) = "Forma Ingreso"
   Grid.TextMatrix(0, C_AJUSTE) = "Ajuste"

    
   Call FGrSetup(Grid)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub


Private Sub Grid_DblClick()
   Dim Frm As FrmSeguimientoMovDoc
   
   Set Frm = New FrmSeguimientoMovDoc
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

