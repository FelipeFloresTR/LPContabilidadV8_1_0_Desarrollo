VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmSeguimientoMovDoc 
   Caption         =   "Seguimiento de Movimiento de Documentos"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   18420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   -720
      TabIndex        =   0
      Top             =   -600
      Width           =   19215
      Begin VB.Frame Fr_Botones 
         Height          =   555
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   17895
         Begin VB.CommandButton Bt_Close 
            Cancel          =   -1  'True
            Caption         =   "Cerrar"
            CausesValidation=   0   'False
            Height          =   315
            Left            =   16560
            TabIndex        =   6
            Top             =   180
            Width           =   1215
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
            Picture         =   "FrmSeguimientoMovDoc.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Vista previa de la impresión"
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
            Picture         =   "FrmSeguimientoMovDoc.frx":04A7
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir listado"
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
            Picture         =   "FrmSeguimientoMovDoc.frx":0961
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Copiar Excel"
            Top             =   120
            Width           =   375
         End
      End
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   4200
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   17865
         _ExtentX        =   31512
         _ExtentY        =   7408
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
Attribute VB_Name = "FrmSeguimientoMovDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const C_IDMOVDOC = 0
Const C_FECHAHORA = 1
Const C_IDEMPRESA = 2
Const C_EMPRESA = 3
Const C_ANO = 4
Const C_IDDOC = 5
Const C_COMPCENT = 6
Const C_COMPPAGO = 7
Const C_ORDEN = 8
Const C_IDCUENTA = 9
Const C_DESCRIPCION = 10
Const C_DEBE = 11
Const C_HABER = 12
Const C_GLOSA = 13
Const C_IDTIPOVAL = 14
Const C_VALOR = 15
Const C_ESTOTALDOC = 16
Const C_IDAREANEG = 17
Const C_TASA = 18
Const C_ESRECUPERABLE = 19
Const C_CODSIIDTE = 20
Const C_CODCUENTAOLD = 21
Const C_ORIGEN = 22
Const C_QUERY = 23
Const C_IDVIGENTE = 24
Const C_VIGENTE = 25
Const C_FINGRESO = 26
Const C_AJUSTE = 27
Const NCOLS = C_AJUSTE

Dim lOrientacion As Integer

Const FI_MANUAL = 1
Const FI_IMPORTACION = 1

Const AJ_INSERTAR = 1
Const AJ_MODIFICAR = 2
Const AJ_ELIMINAR = 3


Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Filtro As Boolean
   Dim Compfecha As String
   Dim OtraFecha As Boolean
   
   Filtro = False
   
   Q1 = " SELECT IdMovDoc"
   Q1 = Q1 & "    ,FechaHora"
   Q1 = Q1 & "    ,TMD.IdEmpresa"
   Q1 = Q1 & "    ,E.NombreCorto"
   Q1 = Q1 & "    ,TMD.Ano"
   Q1 = Q1 & "    ,IdDoc"
   Q1 = Q1 & "    ,IdCompCent"
   Q1 = Q1 & "    ,IdCompPago"
   Q1 = Q1 & "    ,TMD.Orden"
   Q1 = Q1 & "    ,TMD.IdCuenta"
   Q1 = Q1 & "    ,C.Descripcion"
   Q1 = Q1 & "    ,TMD.Debe"
   Q1 = Q1 & "    ,TMD.Haber"
   Q1 = Q1 & "    ,Glosa"
   Q1 = Q1 & "    ,TMD.IdTipoValLib"
   Q1 = Q1 & "    ,T.Valor"
   Q1 = Q1 & "    ,EsTotalDoc"
   Q1 = Q1 & "    ,IdAreaNeg"
   Q1 = Q1 & "    ,TMD.Tasa"
   Q1 = Q1 & "    ,TMD.EsRecuperable"
   Q1 = Q1 & "    ,TMD.CodSIIDTE"
   Q1 = Q1 & "    ,CodCuentaOld"
   Q1 = Q1 & "    ,Origen"
   Q1 = Q1 & "    ,Query"
   Q1 = Q1 & "    ,Vigente AS IdVigente"
   Q1 = Q1 & "    ,IIF(Vigente IS NULL OR VIGENTE = 1, 'VIGENTE', 'ELIMINADO') AS Vigente2"
   Q1 = Q1 & "    ,FechaHora AS CompFecha"
   Q1 = Q1 & "    ,FormaIngreso AS Fingreso"
   Q1 = Q1 & "    ,Ajuste"
   Q1 = Q1 & " FROM (((Tracking_MovDocumento TMD"
   Q1 = Q1 & " LEFT JOIN Empresas E ON E.IdEmpresa = TMD.IdEmpresa)"
   Q1 = Q1 & " LEFT JOIN Cuentas C ON C.idCuenta = TMD.IdCuenta AND C.IdEmpresa = TMD.IdEmpresa AND C.Ano = TMD.Ano)"
   Q1 = Q1 & " LEFT JOIN TipoValor T ON T.idTValor = TMD.IdTipoValLib)"
   Q1 = Q1 & " WHERE TMD.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND TMD.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND TMD.IdDoc = " & IdDoc
   
   If gDbType = SQL_ACCESS Then
        Q1 = Q1 & " ORDER BY FechaHora DESC"
   Else
       Q1 = Q1 & " ORDER BY FechaHora DESC, Orden ASC"
   End If
'   If Tx_NumDoc.Text <> "" Then
'    Q1 = Q1 & " AND NumDoc = " & Tx_NumDoc.Text
'    Filtro = True
'   End If
'   If ItemData(Cb_TipoLib) > 0 Then
'    Q1 = Q1 & " AND TD.TIPOLIB = " & ItemData(Cb_TipoLib)
'    Filtro = True
'   End If
'   If ItemData(Cb_TipoDoc) > 0 Then
'    Q1 = Q1 & " AND TD.TIPODOC = " & ItemData(Cb_TipoDoc)
'    Filtro = True
'   End If
'   If ItemData(Cb_Estado) > 0 Then
'    Q1 = Q1 & " AND TD.ESTADO = " & ItemData(Cb_Estado)
'    Filtro = True
'   End If
'   Q1 = Q1 & " ORDER BY FechaHora DESC"
'
'   If Not Filtro Then
'    MsgBox "Favor ingresar al menos un filtro", vbInformation, "Seguimiento de Documento"
'    Exit Sub
'   End If
   
   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   If Rs.EOF = False Then
   Compfecha = Trim(vFld(Rs("CompFecha")))
   OtraFecha = True
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
        
        If Compfecha = Trim(vFld(Rs("CompFecha"))) And OtraFecha = True Then
            'Call FGrSetRowStyle(Grid, i, "BC", &HFFFF00)
            Call FGrSetRowStyle(Grid, i, "BC", GRAY)
        ElseIf Compfecha = Trim(vFld(Rs("CompFecha"))) And OtraFecha = False Then
            Call FGrSetRowStyle(Grid, i, "BC", &H80000005)
        Else
            If OtraFecha Then
              Call FGrSetRowStyle(Grid, i, "BC", &H80000005)
              OtraFecha = False
            Else
              Call FGrSetRowStyle(Grid, i, "BC", GRAY)
              OtraFecha = True
            End If
        End If
        Compfecha = Trim(vFld(Rs("CompFecha")))
        Grid.TextMatrix(i, C_IDMOVDOC) = vFld(Rs("IdMovDoc"))
        Grid.TextMatrix(i, C_FECHAHORA) = vFld(Rs("FechaHora"))
        Grid.TextMatrix(i, C_IDEMPRESA) = vFld(Rs("IdEmpresa"))
        Grid.TextMatrix(i, C_EMPRESA) = vFld(Rs("NombreCorto"))
        Grid.TextMatrix(i, C_ANO) = vFld(Rs("Ano"))
        Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
        Grid.TextMatrix(i, C_COMPCENT) = vFld(Rs("IdCompCent"))
        Grid.TextMatrix(i, C_COMPPAGO) = vFld(Rs("IdCompPago"))
        Grid.TextMatrix(i, C_ORDEN) = vFld(Rs("Orden"))
        Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
        Grid.TextMatrix(i, C_DESCRIPCION) = vFld(Rs("Descripcion"))
        Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
        Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
        Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"))
        Grid.TextMatrix(i, C_IDTIPOVAL) = vFld(Rs("IdTipoValLib"))
        Grid.TextMatrix(i, C_VALOR) = vFld(Rs("Valor"))
        Grid.TextMatrix(i, C_ESTOTALDOC) = vFld(Rs("EsTotalDoc"))
        Grid.TextMatrix(i, C_IDAREANEG) = vFld(Rs("IdAreaNeg"))
        Grid.TextMatrix(i, C_TASA) = vFld(Rs("Tasa"))
        Grid.TextMatrix(i, C_ESRECUPERABLE) = vFld(Rs("EsRecuperable"))
        Grid.TextMatrix(i, C_CODSIIDTE) = vFld(Rs("CodSIIDTE"))
        Grid.TextMatrix(i, C_CODCUENTAOLD) = vFld(Rs("CodCuentaOld"))
        Grid.TextMatrix(i, C_ORIGEN) = vFld(Rs("Origen"))
        Grid.TextMatrix(i, C_QUERY) = vFld(Rs("Query"))
        Grid.TextMatrix(i, C_IDVIGENTE) = vFld(Rs("IdVigente"))
        Grid.TextMatrix(i, C_VIGENTE) = vFld(Rs("Vigente2"))
        Grid.TextMatrix(i, C_FINGRESO) = FormaIngreso(vFld(Rs("Fingreso")))
        Grid.TextMatrix(i, C_AJUSTE) = Ajuste(vFld(Rs("Ajuste")))
      

      
    Rs.MoveNext
      i = i + 1
   Loop
   Else
        MsgBox "NO se encontraron documentos", vbInformation, "Seguimiento de Documento"
   End If
   
   
   
   
End Sub
   

Private Sub Bt_Search_Click()
Call SetUpGrid
Call LoadGrid
End Sub



Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   

   Grid.ColWidth(C_IDMOVDOC) = 0
   Grid.ColWidth(C_FECHAHORA) = 1750
   Grid.ColWidth(C_IDEMPRESA) = 0
   Grid.ColWidth(C_EMPRESA) = 1600
   Grid.ColWidth(C_ANO) = 500
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_COMPCENT) = 1600
   Grid.ColWidth(C_COMPPAGO) = 1200
   Grid.ColWidth(C_ORDEN) = 600
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_DESCRIPCION) = 2000
   Grid.ColWidth(C_DEBE) = 900
   Grid.ColWidth(C_HABER) = 900
   Grid.ColWidth(C_GLOSA) = 3000
   Grid.ColWidth(C_IDTIPOVAL) = 0
   Grid.ColWidth(C_VALOR) = 1800
   Grid.ColWidth(C_ESTOTALDOC) = 1200
   Grid.ColWidth(C_IDAREANEG) = 0
   Grid.ColWidth(C_TASA) = 800
   Grid.ColWidth(C_ESRECUPERABLE) = 1250
   Grid.ColWidth(C_CODSIIDTE) = 1000
   Grid.ColWidth(C_CODCUENTAOLD) = 0
   Grid.ColWidth(C_ORIGEN) = 1500
   Grid.ColWidth(C_QUERY) = 0
   Grid.ColWidth(C_IDVIGENTE) = 0
   Grid.ColWidth(C_VIGENTE) = 1200
   Grid.ColWidth(C_FINGRESO) = 1200
   Grid.ColWidth(C_AJUSTE) = 1200
   
   Grid.ColAlignment(C_IDMOVDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHAHORA) = flexAlignRightCenter
   Grid.ColAlignment(C_IDEMPRESA) = flexAlignRightCenter
   Grid.ColAlignment(C_EMPRESA) = flexAlignLeftCenter
   Grid.ColAlignment(C_ANO) = flexAlignRightCenter
   Grid.ColAlignment(C_IDDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_COMPCENT) = flexAlignRightCenter
   Grid.ColAlignment(C_COMPPAGO) = flexAlignRightCenter
   Grid.ColAlignment(C_ORDEN) = flexAlignRightCenter
   Grid.ColAlignment(C_IDCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_DESCRIPCION) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_IDTIPOVAL) = flexAlignRightCenter
   Grid.ColAlignment(C_VALOR) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTOTALDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_IDAREANEG) = flexAlignRightCenter
   Grid.ColAlignment(C_TASA) = flexAlignRightCenter
   Grid.ColAlignment(C_ESRECUPERABLE) = flexAlignRightCenter
   Grid.ColAlignment(C_CODSIIDTE) = flexAlignRightCenter
   Grid.ColAlignment(C_CODCUENTAOLD) = flexAlignRightCenter
   Grid.ColAlignment(C_ORIGEN) = flexAlignLeftCenter
   Grid.ColAlignment(C_QUERY) = flexAlignRightCenter
   Grid.ColAlignment(C_IDVIGENTE) = flexAlignRightCenter
   Grid.ColAlignment(C_VIGENTE) = flexAlignRightCenter
   Grid.ColAlignment(C_FINGRESO) = flexAlignRightCenter
   Grid.ColAlignment(C_AJUSTE) = flexAlignRightCenter

   Grid.TextMatrix(0, C_FECHAHORA) = "Fecha y Hora"
   Grid.TextMatrix(0, C_EMPRESA) = "Empresa"
   Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_COMPCENT) = "Comp. Centralizacion"
   Grid.TextMatrix(0, C_COMPPAGO) = "Comp. Pago"
   Grid.TextMatrix(0, C_ORDEN) = "Orden"
   Grid.TextMatrix(0, C_DESCRIPCION) = "Cuenta"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_VALOR) = "Tipo"
   Grid.TextMatrix(0, C_ESTOTALDOC) = "Es Total"
   Grid.TextMatrix(0, C_TASA) = "Tasa"
   Grid.TextMatrix(0, C_ESRECUPERABLE) = "Es Recuperable"
   Grid.TextMatrix(0, C_CODSIIDTE) = "Cod. SII"
   Grid.TextMatrix(0, C_ORIGEN) = "Origen"
   Grid.TextMatrix(0, C_VIGENTE) = "Estado"
   Grid.TextMatrix(0, C_FINGRESO) = "Forma Ingreso"
   Grid.TextMatrix(0, C_AJUSTE) = "Ajuste"

   Call FGrSetup(Grid)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub



Public Function FSearch(IdDoc As Long) As Integer
   G_IDDOC = IdDoc
   Call SetUpGrid
Call LoadGrid(IdDoc)
End Function

Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
Call FGr2Clip(Grid, Me.Caption)
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
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
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

Private Sub Form_Load()
lOrientacion = ORIENT_VER
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
