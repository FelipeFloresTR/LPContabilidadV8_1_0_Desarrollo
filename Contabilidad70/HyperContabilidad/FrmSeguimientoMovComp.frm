VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmSeguimientoMovComp 
   Caption         =   "Seguimiento de Movimiento de Comprobantes"
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
            Picture         =   "FrmSeguimientoMovComp.frx":0000
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
            Picture         =   "FrmSeguimientoMovComp.frx":04A7
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
            Picture         =   "FrmSeguimientoMovComp.frx":0961
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
Attribute VB_Name = "FrmSeguimientoMovComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const C_IDMOV = 0
Const C_FECHAHORA = 1
Const C_IDEMPRESA = 2
Const C_EMPRESA = 3
Const C_ANO = 4
Const C_IDCOMP = 5
Const C_IDDOC = 6
Const C_ORDEN = 7
Const C_IDCUENTA = 8
Const C_DESCRIPCION = 9
Const C_DEBE = 10
Const C_HABER = 11
Const C_GLOSA = 12
Const C_IDCCOSTO = 13
Const C_CCOSTO = 14
Const C_IDAREANEG = 15
Const C_AREANEG = 16
Const C_IDCARTOLA = 17
Const C_DECENTRA = 18
Const C_DEPAGO = 19
Const C_DEREMU = 20
Const C_NOTA = 21
Const C_IDDOCCUOTA = 22
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


Private Sub LoadGrid(Optional ByVal IdComp As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Filtro As Boolean
   Dim Compfecha As String
   Dim OtraFecha As Boolean
   
   Filtro = False
   
   Q1 = " SELECT IdMov"
   Q1 = Q1 & " ,FechaHora"
   Q1 = Q1 & " ,TMC.IdEmpresa"
   Q1 = Q1 & " ,E.NombreCorto"
   Q1 = Q1 & " ,TMC.Ano"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,IdDoc"
   Q1 = Q1 & " ,Orden"
   Q1 = Q1 & " ,TMC.IdCuenta"
   Q1 = Q1 & " ,C.Descripcion"
   Q1 = Q1 & " ,TMC.Debe"
   Q1 = Q1 & " ,TMC.Haber"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,idCCosto"
   Q1 = Q1 & " ,idAreaNeg"
   Q1 = Q1 & " ,IdCartola"
   Q1 = Q1 & " ,DeCentraliz"
   Q1 = Q1 & " ,DePago"
   Q1 = Q1 & " ,DeRemu"
   Q1 = Q1 & " ,Nota"
   Q1 = Q1 & " ,IdDocCuota"
   Q1 = Q1 & " ,Origen"
   Q1 = Q1 & " ,Query"
   Q1 = Q1 & " ,Vigente As IdVigente"
   Q1 = Q1 & " ,IIF(Vigente IS NULL OR VIGENTE = 1, 'VIGENTE', 'ELIMINADO') AS Vigente2"
   Q1 = Q1 & " ,FormaIngreso As Fingreso"
   Q1 = Q1 & " ,Ajuste"
   Q1 = Q1 & " FROM ((Tracking_MovComprobante TMC"
   Q1 = Q1 & " LEFT JOIN Empresas E ON E.IdEmpresa = TMC.IdEmpresa)"
   Q1 = Q1 & " LEFT JOIN Cuentas C ON C.idCuenta = TMC.IdCuenta AND C.IdEmpresa = TMC.IdEmpresa AND C.Ano = TMC.Ano)"
   Q1 = Q1 & " WHERE TMC.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND TMC.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND TMC.IdComp = " & IdComp
   Q1 = Q1 & " ORDER BY FechaHora DESC, Orden ASC"
   
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
   Compfecha = Trim(vFld(Rs("FechaHora")))
   OtraFecha = True
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
        
        If Compfecha = Trim(vFld(Rs("FechaHora"))) And OtraFecha = True Then
            'Call FGrSetRowStyle(Grid, i, "BC", &HFFFF00)
            Call FGrSetRowStyle(Grid, i, "BC", GRAY)
        ElseIf Compfecha = Trim(vFld(Rs("FechaHora"))) And OtraFecha = False Then
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
        
        Compfecha = Trim(vFld(Rs("FechaHora")))
        Grid.TextMatrix(i, C_IDMOV) = vFld(Rs("IdMov"))
        Grid.TextMatrix(i, C_FECHAHORA) = vFld(Rs("FechaHora"))
        Grid.TextMatrix(i, C_IDEMPRESA) = vFld(Rs("IdEmpresa"))
        Grid.TextMatrix(i, C_EMPRESA) = vFld(Rs("NombreCorto"))
        Grid.TextMatrix(i, C_ANO) = vFld(Rs("Ano"))
        Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
        Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
        Grid.TextMatrix(i, C_ORDEN) = vFld(Rs("Orden"))
        Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
        Grid.TextMatrix(i, C_DESCRIPCION) = vFld(Rs("Descripcion"))
        Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
        Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
        Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"))
        'Grid.TextMatrix(i, C_IDCCOSTO) = vFld(Rs("IdTipoValLib"))
        Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("idCCosto"))
        Grid.TextMatrix(i, C_IDAREANEG) = vFld(Rs("IdAreaNeg"))
        Grid.TextMatrix(i, C_IDCARTOLA) = vFld(Rs("IdCartola"))
        Grid.TextMatrix(i, C_DECENTRA) = vFld(Rs("DeCentraliz"))
        Grid.TextMatrix(i, C_DEPAGO) = vFld(Rs("DePago"))
        Grid.TextMatrix(i, C_DEREMU) = vFld(Rs("DeRemu"))
        Grid.TextMatrix(i, C_NOTA) = vFld(Rs("Nota"))
        Grid.TextMatrix(i, C_IDDOCCUOTA) = vFld(Rs("IdDocCuota"))
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
   
   
    Grid.ColWidth(C_IDMOV) = 0
    Grid.ColWidth(C_FECHAHORA) = 1750
    Grid.ColWidth(C_IDEMPRESA) = 0
    Grid.ColWidth(C_EMPRESA) = 1600
    Grid.ColWidth(C_ANO) = 500
    Grid.ColWidth(C_IDCOMP) = 1000
    Grid.ColWidth(C_IDDOC) = 1000
    Grid.ColWidth(C_ORDEN) = 600
    Grid.ColWidth(C_IDCUENTA) = 0
    Grid.ColWidth(C_DESCRIPCION) = 1800
    Grid.ColWidth(C_DEBE) = 900
    Grid.ColWidth(C_HABER) = 900
    Grid.ColWidth(C_GLOSA) = 3000
    Grid.ColWidth(C_IDCCOSTO) = 0
    Grid.ColWidth(C_CCOSTO) = 900
    Grid.ColWidth(C_IDAREANEG) = 900
    Grid.ColWidth(C_IDCARTOLA) = 900
    Grid.ColWidth(C_DECENTRA) = 900
    Grid.ColWidth(C_DEPAGO) = 900
    Grid.ColWidth(C_DEREMU) = 900
    Grid.ColWidth(C_NOTA) = 900
    Grid.ColWidth(C_IDDOCCUOTA) = 900
    Grid.ColWidth(C_QUERY) = 0
    Grid.ColWidth(C_IDVIGENTE) = 0
    Grid.ColWidth(C_VIGENTE) = 1200
    Grid.ColWidth(C_FINGRESO) = 1200
    Grid.ColWidth(C_AJUSTE) = 1200
   
   
   
    Grid.ColAlignment(C_IDMOV) = flexAlignRightCenter
    Grid.ColAlignment(C_FECHAHORA) = flexAlignLeftCenter
    Grid.ColAlignment(C_IDEMPRESA) = flexAlignRightCenter
    Grid.ColAlignment(C_EMPRESA) = flexAlignRightCenter
    Grid.ColAlignment(C_ANO) = flexAlignRightCenter
    Grid.ColAlignment(C_IDCOMP) = flexAlignRightCenter
    Grid.ColAlignment(C_IDDOC) = flexAlignRightCenter
    Grid.ColAlignment(C_ORDEN) = flexAlignRightCenter
    Grid.ColAlignment(C_DESCRIPCION) = flexAlignLeftCenter
    Grid.ColAlignment(C_DEBE) = flexAlignLeftCenter
    Grid.ColAlignment(C_HABER) = flexAlignLeftCenter
    Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
    Grid.ColAlignment(C_CCOSTO) = flexAlignLeftCenter
    Grid.ColAlignment(C_IDAREANEG) = flexAlignLeftCenter
    Grid.ColAlignment(C_IDCARTOLA) = flexAlignLeftCenter
    Grid.ColAlignment(C_DECENTRA) = flexAlignLeftCenter
    Grid.ColAlignment(C_DEPAGO) = flexAlignLeftCenter
    Grid.ColAlignment(C_DEREMU) = flexAlignLeftCenter
    Grid.ColAlignment(C_NOTA) = flexAlignRightCenter
    Grid.ColAlignment(C_IDDOCCUOTA) = flexAlignRightCenter
    Grid.ColAlignment(C_QUERY) = flexAlignRightCenter
    Grid.ColAlignment(C_VIGENTE) = flexAlignRightCenter
    Grid.ColAlignment(C_FINGRESO) = flexAlignRightCenter
    Grid.ColAlignment(C_AJUSTE) = flexAlignRightCenter

   
    'Grid.TextMatrix(0, C_IDMOV) = "Id Mov"
    Grid.TextMatrix(0, C_FECHAHORA) = "Fecha y Hora"
    'Grid.TextMatrix(0, C_IDEMPRESA) = "Id Empresa"
    Grid.TextMatrix(0, C_EMPRESA) = "Empresa"
    Grid.TextMatrix(0, C_ANO) = "Año"
    Grid.TextMatrix(0, C_IDCOMP) = "Comprobante"
    Grid.TextMatrix(0, C_IDDOC) = "Documento"
    Grid.TextMatrix(0, C_ORDEN) = "Orden"
    'Grid.TextMatrix(0, C_IDCUENTA) = ""
    Grid.TextMatrix(0, C_DESCRIPCION) = "Cuenta"
    Grid.TextMatrix(0, C_DEBE) = "Debe"
    Grid.TextMatrix(0, C_HABER) = "Haber"
    Grid.TextMatrix(0, C_GLOSA) = "Glosa"
    'Grid.TextMatrix(0, C_IDCCOSTO) = ""
    Grid.TextMatrix(0, C_CCOSTO) = "Centro Costo"
    Grid.TextMatrix(0, C_IDAREANEG) = "Area Negocio"
    Grid.TextMatrix(0, C_IDCARTOLA) = "Cartola"
    Grid.TextMatrix(0, C_DECENTRA) = "Centralizacion"
    Grid.TextMatrix(0, C_DEPAGO) = "Pago"
    Grid.TextMatrix(0, C_DEREMU) = "Remu"
    Grid.TextMatrix(0, C_NOTA) = "Nota"
    'Grid.TextMatrix(0, C_IDDOCCUOTA) = vFld(Rs("IdDocCuota"))
    'Grid.TextMatrix(0, C_QUERY) = vFld(Rs("Query"))
    Grid.TextMatrix(0, C_IDVIGENTE) = ""
    Grid.TextMatrix(0, C_VIGENTE) = "Vigente"
    Grid.TextMatrix(0, C_FINGRESO) = "Forma Ingreso"
    Grid.TextMatrix(0, C_AJUSTE) = "Ajuste"

   Call FGrSetup(Grid)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub
Public Function FSearch(IdComp As Long) As Integer
   G_IDDOC = IdComp
   Call SetUpGrid
Call LoadGrid(IdComp)
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
