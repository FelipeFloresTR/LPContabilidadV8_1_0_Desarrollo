VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInfConciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Conciliación"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   Icon            =   "FrmInfConciliacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   10095
      Begin VB.CheckBox Ch_ChequesNulos 
         Caption         =   "Cheques Nulos"
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   8760
         Picture         =   "FrmInfConciliacion.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   1740
         Picture         =   "FrmInfConciliacion.frx":055C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   230
      End
      Begin VB.ComboBox cb_Banco 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   4275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.TextBox Tx_SaldoBanco 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6540
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox Tx_Concil 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Informe de Conciliaciï¿½n al ..."
      Top             =   7320
      Width           =   6315
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton Bt_VerComp 
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
         Picture         =   "FrmInfConciliacion.frx":0866
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerDoc 
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
         Picture         =   "FrmInfConciliacion.frx":0CDF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Detalle documento seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Width           =   1155
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
         Left            =   1500
         Picture         =   "FrmInfConciliacion.frx":1153
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1080
         Picture         =   "FrmInfConciliacion.frx":160D
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
         Left            =   1920
         Picture         =   "FrmInfConciliacion.frx":1AB4
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   3840
         Picture         =   "FrmInfConciliacion.frx":1EF9
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   3000
         Picture         =   "FrmInfConciliacion.frx":2322
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   3420
         Picture         =   "FrmInfConciliacion.frx":26C0
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2460
         Picture         =   "FrmInfConciliacion.frx":2A21
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   30
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      SelectionMode   =   1
   End
End
Attribute VB_Name = "FrmInfConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_DESC = 0
Const C_TIPODOC = 1
Const C_NRODOC = 2
Const C_TOTAL = 3
Const C_FMT = 4
Const C_OBLIGATORIA = 5
Const C_IDDOC = 6
Const C_IDCOMP = 7
Const C_IDDETCARTOLA = 8

Const NCOLS = C_IDDETCARTOLA

Private Sub Bt_Buscar_Click()

   MousePointer = vbHourglass
   DoEvents
   
   Call LoadAll
   
   MousePointer = vbDefault
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Tx_Concil)
End Sub


Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Hasta)
   
   Set Frm = Nothing

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

Private Sub Bt_VerComp_Click()
   Dim idcomp As Long, Frm As FrmComprobante
   
   idcomp = Val(Grid.TextMatrix(Grid.Row, C_IDCOMP))
   
   If idcomp <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmComprobante
   Call Frm.FView(idcomp, False)
   Set Frm = Nothing

End Sub

Private Sub Bt_VerDoc_Click()
   Dim IdDoc As Long, Frm As FrmDoc
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmDoc
   Frm.FView IdDoc
   Set Frm = Nothing

End Sub

Private Sub Cb_Banco_Click()
   Bt_Buscar.Enabled = True
End Sub

Private Sub Ch_ChequesNulos_Click()
   Bt_Buscar.Enabled = True

End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   Call BtFechaImg(Bt_Fecha)

   Call SetTxDate(Tx_Hasta, Now)
   
   Call SetUpGrid
   
   cb_Banco.AddItem "Todos"
   Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(cb_Banco, DbMain, Q1, -1)
   If cb_Banco.ListCount = 2 Then
      cb_Banco.ListIndex = 1
   End If
   
   Call LoadAll
   
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   
   Grid.ColWidth(C_DESC) = 6500
   Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_NRODOC) = 1130
   Grid.ColWidth(C_TOTAL) = 1600
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_IDDETCARTOLA) = 0
   
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_NRODOC) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim SubTot As Double
   Dim i As Integer, Rc As Long
   Dim Total As Double
   Dim Hasta As Long
   Dim QryCta As String
   Dim QryBanco As String
   Dim QName As String
   
   Grid.Redraw = False
   
   Total = 0
   Hasta = GetTxDate(Tx_Hasta)
   
   'PS 9 Jun 2006
   If cb_Banco.ListIndex = 0 Then
      QryCta = " AND Cuentas.Atrib" & ATRIB_CONCILIACION & "<>0"
      QryBanco = ""
   Else
      QryCta = " AND Cuentas.IdCuenta=" & ItemData(cb_Banco)
      QryBanco = " AND Cartola.IdCuentaBco = " & ItemData(cb_Banco)
   End If
   
   Grid.rows = 0
   
   Grid.rows = Grid.rows + 1
   Row = 0
   Grid.TextMatrix(0, C_FMT) = "              .FMT"
   Grid.RowHeight(0) = 0  'Row con el formateo
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   
   'QName = "tmp_qc_" & ReplaceStr(W.PcName, "-", "_") ' query temporal asociado al equipo
   QName = DbGenTmpName2(SQL_ACCESS, "qc_")       'Forzamos Access para que no le ponga # en el caso de SQL, ya que no se permite para vistas
   
   ' Los detalles de las cartolas conciliadas hasta la fecha indicada
   Q1 = "SELECT Cartola.IdCartola, DetCartola.IdDetCartola, Cartola.FDesde, DetCartola.Fecha, DetCartola.IdMov, DetCartola.IdEmpresa, DetCartola.Ano"
   Q1 = Q1 & " FROM Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cartola", "DetCartola")
   Q1 = Q1 & " WHERE Cartola.FDesde <=" & Hasta & " And DetCartola.Fecha <= " & Hasta
   Q1 = Q1 & " AND (NOT DetCartola.IdMov IS NULL AND DetCartola.IdMov <> 0)"
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   Rc = CreateQry(DbMain, QName, Q1)
   
   'Total Mayor Banco
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Saldo Mayor Banco"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   
   Q1 = "SELECT Sum( MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber"
   Q1 = Q1 & " FROM (((MovComprobante INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   '2907316
   'Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc  "
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc AND MovComprobante.IdEmpresa = Documento.IdEmpresa "
   '2907316
   
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad"
   Q1 = Q1 & "  AND Documento.IdEmpresa = Entidades.IdEmpresa "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   Q1 = Q1 & " WHERE Comprobante.Estado IN (" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & QryCta
   Q1 = Q1 & " AND Comprobante.Fecha <=" & Hasta
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Total = vFld(Rs("Debe")) - vFld(Rs("Haber"))
      Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NEGNUMFMT)
   End If
   
   Call CloseRs(Rs)
   
   'Cheques girados y no cobrados
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"
   Grid.TextMatrix(Row - 2, C_OBLIGATORIA) = "O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Documentos Girados y No Cobrados"
   Grid.TextMatrix(Row, C_TIPODOC) = "TD"
   Grid.TextMatrix(Row, C_NRODOC) = "N° Doc"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      
   Q1 = "SELECT Comprobante.Fecha, MovComprobante.Glosa, TipoLib, TipoDoc, Documento.NumDoc, MovComprobante.Debe, MovComprobante.Haber, MovComprobante.IdCartola, MovComprobante.idMov, MovComprobante.IdDoc, Comprobante.Tipo as TipoComp, Comprobante.Correlativo, Comprobante.IdComp, DetCartola.IdDetCartola "
   Q1 = Q1 & " FROM (((( MovComprobante INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.idCuenta=Cuentas.idCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   
    '2907316
   'Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc=Documento.IdDoc "
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc AND MovComprobante.IdEmpresa = Documento.IdEmpresa "
   '2907316
   
   
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN " & QName & " ON MovComprobante.idCartola=" & QName & ".idCartola AND MovComprobante.idMov=" & QName & ".idMov  " ' *** pam 29-nov-2006
'   Q1 = Q1 & "  AND " & QName & ".IdEmpresa = MovComprobante.IdEmpresa AND " & QName & ".Ano = MovComprobante.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, QName, "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN DetCartola ON Documento.NumDoc = Str(DetCartola.NumDoc) "    'FCA 21 mar 2013 se agrega para determinar si el documento aparece en alguna cartola (se hizo por consulta de cliente CPAIN)
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "Documento")
   
   Q1 = Q1 & " WHERE Comprobante.Estado IN (" & EC_APROBADO & ", " & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   
   Q1 = Q1 & " AND " & QName & ".idMov IS NULL"     ' *** pam 29-nov-2006
   'Q1 = Q1 & " AND (MovComprobante.IdCartola IS NULL OR MovComprobante.IdCartola = 0)"     'Fca 2 mar 2012 se elimina condición para que el informe sea a la fecha en que se pide, independiente de lo que pase el año siguiente - ' *** fca 9 ago 2011: cubrimos caso en que tenga un IdCartola < 0 que indica que fue conciliado en una cartola del año siguiente, por lo tanto está cobrado
   'Q1 = Q1 & " AND (" & QName & ".IdCartola IS NULL )"        'Fca: 21 ene 2013 esta condición debería reemplazar la anterior pero no estamos seguros ??
   Q1 = Q1 & QryCta
   Q1 = Q1 & " AND MovComprobante.Haber <> 0"
   Q1 = Q1 & " AND Comprobante.Fecha <=" & Hasta
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   'Q1 = Q1 & " ORDER BY Fecha, MovComprobante.Glosa"
   Q1 = Q1 & " ORDER BY Comprobante.Fecha, Documento.NumDoc, MovComprobante.Glosa, MovComprobante.idMov"
   
      
   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Row = Row + 1
      Grid.rows = Row + 1
   
      SubTot = SubTot + vFld(Rs("Haber"))
   
      Grid.TextMatrix(Row, C_DESC) = Format(vFld(Rs("Fecha")), EDATEFMT) & " - [" & Left(gTipoComp(vFld(Rs("TipoComp"))), 1) & "-" & vFld(Rs("Correlativo")) & "]   " & vFld(Rs("Glosa"))
      Grid.TextMatrix(Row, C_NRODOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(Row, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      Grid.TextMatrix(Row, C_IDDETCARTOLA) = vFld(Rs("IdDetCartola"))
      Grid.TextMatrix(Row, C_TOTAL) = Format(vFld(Rs("Haber")), NEGNUMFMT)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("idDoc"))
      Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("idComp"))
         
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   Total = Total + SubTot
   
   If Ch_ChequesNulos <> 0 Then
   
      'agregamos los cheques nulos a los Cheques girados y no cobrados
      Q1 = "SELECT IdDoc, FEmision, Documento.TipoLib, Documento.TipoDoc, NumDoc, Descrip "
      Q1 = Q1 & " FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
      Q1 = Q1 & " WHERE Documento.Estado = " & ED_ANULADO & " AND Documento.TipoLib = " & LIB_OTROS & " AND TipoDocs.Diminutivo IN ('CHE', 'CHF')"
      Q1 = Q1 & " AND Documento.FEmision <=" & Hasta
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY Documento.FEmision, Documento.NumDoc"
      
      Set Rs = OpenRs(DbMain, Q1)
         
      If Not Rs.EOF Then
         Row = Row + 1
         Grid.rows = Row + 1
      End If
   
      Do While Rs.EOF = False
      
         Row = Row + 1
         Grid.rows = Row + 1
         
         Grid.TextMatrix(Row, C_DESC) = Format(vFld(Rs("FEmision")), EDATEFMT) & " - [Cheque Nulo]   " & vFld(Rs("Descrip"))
         Grid.TextMatrix(Row, C_NRODOC) = vFld(Rs("NumDoc"))
         Grid.TextMatrix(Row, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
         Grid.TextMatrix(Row, C_IDDETCARTOLA) = " "
         Grid.TextMatrix(Row, C_TOTAL) = "0"
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("idDoc"))
         Grid.TextMatrix(Row, C_IDCOMP) = " "
            
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
   
   End If
   
   Row = Row + 1
   Grid.rows = Row + 1
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Subtotal"
   Grid.TextMatrix(Row, C_TOTAL) = Format(SubTot, NEGNUMFMT)
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"

   'Abonos bancarios no registrados por la empresa
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"
   Grid.TextMatrix(Row - 2, C_OBLIGATORIA) = "O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Abonos bancarios no registrados por la empresa"
   Grid.TextMatrix(Row, C_TIPODOC) = "TD"
   Grid.TextMatrix(Row, C_NRODOC) = "N° Doc"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   
   Q1 = "SELECT IdDetCartola, Fecha, Detalle, NumDoc, Cargo, Abono, IdMov "
   Q1 = Q1 & " FROM DetCartola INNER JOIN Cartola ON DetCartola.IdCartola = Cartola.IdCartola "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "Cartola")
   Q1 = Q1 & " WHERE IdMov = 0 AND DetCartola.Abono <> 0 "
   Q1 = Q1 & " AND DetCartola.Fecha <=" & Hasta
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   Q1 = Q1 & QryBanco
   Q1 = Q1 & " ORDER BY Fecha, NumDoc, Detalle, IdDetCartola "
  
   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Row = Row + 1
      Grid.rows = Row + 1
   
      SubTot = SubTot + vFld(Rs("Abono"))
   
      Grid.TextMatrix(Row, C_DESC) = Format(vFld(Rs("Fecha")), EDATEFMT) & " - " & vFld(Rs("Detalle"))
      Grid.TextMatrix(Row, C_NRODOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(Row, C_TIPODOC) = ""
      Grid.TextMatrix(Row, C_TOTAL) = Format(vFld(Rs("Abono")), NEGNUMFMT)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("IdDetCartola"))
         
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   Total = Total + SubTot
   
   Row = Row + 1
   Grid.rows = Row + 1
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Subtotal"
   Grid.TextMatrix(Row, C_TOTAL) = Format(SubTot, NEGNUMFMT)
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   

   'Cargos bancarios no registrados por la empresa
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"
   Grid.TextMatrix(Row - 2, C_OBLIGATORIA) = "O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Cargos bancarios no registrados por la empresa"
   Grid.TextMatrix(Row, C_TIPODOC) = "TD"
   Grid.TextMatrix(Row, C_NRODOC) = "N° Doc"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   
   Q1 = "SELECT IdDetCartola, Fecha, Detalle, NumDoc, Cargo, Abono, IdMov "
   Q1 = Q1 & " FROM DetCartola INNER JOIN Cartola ON DetCartola.IdCartola = Cartola.IdCartola "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "Cartola")
   Q1 = Q1 & " WHERE IdMov = 0 AND DetCartola.Cargo <> 0 "
   Q1 = Q1 & " AND DetCartola.Fecha <=" & Hasta
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   Q1 = Q1 & QryBanco
   Q1 = Q1 & " ORDER BY Fecha, NumDoc, Detalle, IdDetCartola "

   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Row = Row + 1
      Grid.rows = Row + 1
   
      SubTot = SubTot - vFld(Rs("Cargo"))
   
      Grid.TextMatrix(Row, C_DESC) = Format(vFld(Rs("Fecha")), EDATEFMT) & " - " & vFld(Rs("Detalle"))
      Grid.TextMatrix(Row, C_NRODOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(Row, C_TIPODOC) = ""
      Grid.TextMatrix(Row, C_TOTAL) = Format(-1 * vFld(Rs("Cargo")), NEGNUMFMT)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
      Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("IdDetCartola"))
         
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   Total = Total + SubTot
   
   Row = Row + 1
   Grid.rows = Row + 1
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Subtotal"
   Grid.TextMatrix(Row, C_TOTAL) = Format(SubTot, NEGNUMFMT)
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   

   'Cheques recibidos y no depositados
      
   Q1 = "SELECT Comprobante.Fecha,MovComprobante.Glosa,TipoLib,TipoDoc,NumDoc,MovComprobante.Debe,MovComprobante.Haber,MovComprobante.IdCartola,MovComprobante.idMov, MovComprobante.IdDoc, Comprobante.Tipo as TipoComp, Comprobante.Correlativo, Comprobante.IdComp"
   Q1 = Q1 & " FROM ((( MovComprobante INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.idCuenta = Cuentas.idCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN " & QName & " ON MovComprobante.idCartola = " & QName & ".idCartola AND MovComprobante.idMov = " & QName & ".idMov" ' *** pam 29-nov-2006
'   Q1 = Q1 & "  AND " & QName & ".IdEmpresa = MovComprobante.IdEmpresa AND " & QName & ".Ano = MovComprobante.Ano "
   Q1 = Q1 & JoinEmpAno(gDbType, QName, "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Estado IN (" & EC_APROBADO & ", " & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   
   Q1 = Q1 & " AND " & QName & ".idMov IS NULL"     ' *** pam 29-nov-2006
   'Q1 = Q1 & " AND (MovComprobante.IdCartola IS NULL OR MovComprobante.IdCartola = 0)"     'Fca 2 mar 2012 se elimina condición para que el informe sea a la fecha en que se pide, independiente de lo que pase el año siguiente - ' *** fca 9 ago 2011: se agrega condición para cubrir caso en que tenga un IdCartola < 0 que indica que fue conciliado en una cartola del año siguiente, por lo tanto está depositado
   Q1 = Q1 & QryCta
   Q1 = Q1 & " AND MovComprobante.Debe <> 0"
   Q1 = Q1 & " AND Comprobante.Fecha <=" & Hasta
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   '3021856
   Q1 = Q1 & " AND Comprobante.Tipo <> " & TC_APERTURA
   '3021856
   
   Q1 = Q1 & " ORDER BY Comprobante.Fecha, Documento.NumDoc, MovComprobante.Glosa, MovComprobante.idMov"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Row = Row + 2
      Grid.rows = Row + 1
      Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"
      Grid.TextMatrix(Row - 2, C_OBLIGATORIA) = "O"
      
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_DESC) = "Documentos Recibidos y No Depositados"
      Grid.TextMatrix(Row, C_TIPODOC) = "TD"
      Grid.TextMatrix(Row, C_NRODOC) = "N° Doc"
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   
      SubTot = 0
      Do While Rs.EOF = False
      
         Row = Row + 1
         Grid.rows = Row + 1
      
         SubTot = SubTot - vFld(Rs("Debe"))
      
         Grid.TextMatrix(Row, C_DESC) = Format(vFld(Rs("Fecha")), EDATEFMT) & " - [" & Left(gTipoComp(vFld(Rs("TipoComp"))), 1) & "-" & vFld(Rs("Correlativo")) & "]   " & vFld(Rs("Glosa"))
         Grid.TextMatrix(Row, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
         Grid.TextMatrix(Row, C_NRODOC) = vFld(Rs("NumDoc"))
         Grid.TextMatrix(Row, C_TOTAL) = Format(-vFld(Rs("Debe")), NEGNUMFMT)
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         Grid.TextMatrix(Row, C_IDDOC) = vFld(Rs("idDoc"))
         Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("idComp"))
            
         Rs.MoveNext
         
      Loop
      
      Total = Total + SubTot
      
      Row = Row + 1
      Grid.rows = Row + 1
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_DESC) = "Subtotal"
      Grid.TextMatrix(Row, C_TOTAL) = Format(SubTot, NEGNUMFMT)
      Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   End If
   Call CloseRs(Rs)
      
   ' Saldo final
   Debug.Print Format(Total, NUMFMT)

   Row = Row + 2
   Grid.rows = Row + 1
   Grid.TextMatrix(Row - 1, C_OBLIGATORIA) = "O"
   Grid.TextMatrix(Row - 2, C_OBLIGATORIA) = "O"
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Saldo Banco"
   Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NUMFMT)
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"

   Tx_SaldoBanco = Format(Total, NUMFMT)

'   Row = Row + 2
'   Grid.Rows = Row + 2

'   Grid.TextMatrix(Row, C_DESC) = "El Saldo Mayor Banco considera TODOS los documentos"
'   Grid.TextMatrix(Row + 1, C_DESC) = "que ya están conciliados a la fecha actual (" & Format(Now, "d mmm yyyy") & ")."

   If Grid.rows <= 1 Then
      Grid.rows = 2
   End If
   
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
         
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = 0
   Grid.ColSel = 0
   
   Grid.Redraw = True

   Tx_Concil = Me.Caption & " al " & FmtFecha(Hasta, True)

   Bt_Buscar.Enabled = False

End Sub
Public Sub FView()

   Me.Show vbModal

End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - Tx_Concil.Height - 500
   'Grid.Width = Me.Width - 230
   Tx_Concil.Top = Grid.Top + Grid.Height + 50
   Tx_SaldoBanco.Top = Tx_Concil.Top
   
   Call FGrLocateCntrl(Grid, Tx_SaldoBanco, C_TOTAL)
   Tx_Concil.Width = Tx_SaldoBanco.Left - Tx_Concil.Left - 50
  
   Call FGrVRows(Grid)

End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Informe de Conciliación"
   If CbItemData(cb_Banco) > 0 Then
      Titulos(1) = cb_Banco
   End If
   gPrtReportes.Titulos = Titulos
   Encabezados(0) = "Al " & FmtFecha(GetTxDate(Tx_Hasta), True)
   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   Total(C_DESC) = Tx_Concil
   Total(C_TOTAL) = Tx_SaldoBanco
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.NTotLines = 1

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
   
   If Bt_Buscar.Enabled Then
      MsgBox1 "Presione el botón buscar.", vbExclamation
      Exit Sub
   End If
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
End Sub

Private Sub Grid_DblClick()
   Call PostClick(Bt_VerDoc)
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCopy(KeyCode, Shift) Then
      Call Bt_CopyExcel_Click
   End If
   
End Sub

Private Sub tx_Hasta_Change()
   Bt_Buscar.Enabled = True
End Sub

Private Sub Tx_Hasta_GotFocus()

   If gEmpresa.FCierre <> 0 Then
      Exit Sub
   End If
   
   Call DtGotFocus(Tx_Hasta)
   
End Sub

Private Sub Tx_Hasta_LostFocus()
   Call DtLostFocus(Tx_Hasta)

End Sub

