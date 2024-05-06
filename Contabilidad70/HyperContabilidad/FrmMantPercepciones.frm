VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMantPercepciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Percepciones"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   13095
      Begin VB.Frame Frame 
         Height          =   735
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   12015
         Begin VB.Frame Fr_Doc 
            BorderStyle     =   0  'None
            Height          =   470
            Left            =   3660
            TabIndex        =   20
            Top             =   120
            Width           =   2055
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
            Left            =   1800
            Picture         =   "FrmMantPercepciones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Copiar Excel"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton Bt_Preview 
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
            Left            =   840
            Picture         =   "FrmMantPercepciones.frx":0445
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Vista previa de la impresión"
            Top             =   180
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
            Left            =   1320
            Picture         =   "FrmMantPercepciones.frx":08EC
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Imprimir"
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
            Left            =   360
            Picture         =   "FrmMantPercepciones.frx":0DA6
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Sumar movimientos seleccionados"
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
            Left            =   2280
            Picture         =   "FrmMantPercepciones.frx":0E4A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Calendario"
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox Tx_IdPerc 
            Height          =   375
            Left            =   3840
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Width           =   12015
         Begin VB.CommandButton Bt_Close 
            Caption         =   "Cerrar"
            Height          =   435
            Left            =   10680
            TabIndex        =   12
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton Bt_Buscar 
            Height          =   435
            Left            =   9000
            Picture         =   "FrmMantPercepciones.frx":1273
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox Tx_Rut 
            Height          =   315
            Left            =   6600
            MaxLength       =   12
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1755
         End
         Begin VB.TextBox Tx_Hasta 
            Height          =   315
            Left            =   3120
            TabIndex        =   6
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox Tx_Desde 
            Height          =   315
            Left            =   1020
            TabIndex        =   5
            Top             =   360
            Width           =   1035
         End
         Begin VB.CommandButton Bt_Fecha 
            Height          =   315
            Index           =   1
            Left            =   4140
            Picture         =   "FrmMantPercepciones.frx":17C3
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   230
         End
         Begin VB.CommandButton Bt_Fecha 
            Height          =   315
            Index           =   0
            Left            =   2040
            Picture         =   "FrmMantPercepciones.frx":1ACD
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   360
            Width           =   230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "RUT Empresa Fuente:"
            Height          =   195
            Left            =   4920
            TabIndex        =   10
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   8
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   7
            Left            =   2520
            TabIndex        =   7
            Top             =   420
            Width           =   465
         End
      End
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   4995
         Left            =   480
         TabIndex        =   1
         Top             =   2280
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   8811
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
   End
End
Attribute VB_Name = "FrmMantPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_IDPERC = 0
Const C_IDCUENTA = 1
Const C_CUENTA = 2
Const C_IDEMPRESA = 3
Const C_ANO = 4
Const C_FECHA = 5
Const C_NUMCERTIFICADO = 6
Const C_RUTEMPRESA = 7
Const C_REGIMEN = 8
Const C_DESREGIMEN = 9
Const C_CONTABILIZACION = 10
Const C_DESCONTABILIZACION = 11
Const C_TASATEF = 12
Const C_TASATEX = 13
Const C_PERCEPCION = 14
Const NCOLS = C_PERCEPCION

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean


Private Sub Bt_Buscar_Click()

If Tx_Desde.Text = "" And Tx_Hasta = "" And Tx_Rut.Text = "" Then
 MsgBox "Favor ingresar algun filtro a la busqueda.", vbExclamation, "Mantenedor Percepciones"
 Exit Sub
End If

Call SetUpGrid
Call LoadAll
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
'Call LP_FGr2Clip(Grid, "  Cuenta  " & vbTab & "  Fecha " & vbTab & "  N° Certificado " & vbTab & "  Rut Empresa " & vbTab & "  Regimen " & vbTab & "  Contabilizacion " & vbTab & "  Tasa TEF " & vbTab & "  Tasa TEX " & vbTab & "  Percepciones ")
Call LP_FGr2Clip(Grid, "")
End Sub

Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_Desde)
   Else
      Call Frm.TxSelDate(Tx_Hasta)
   End If
   
   Set Frm = Nothing
   
   'Call EnableFrm(True)
End Sub

Private Sub Bt_Preview_Click()
Dim Frm As FrmPrintPreview
   Dim Frmu As FrmSalyTotLibCajas
   Dim Pag As Integer
   Dim FrmPrt As FrmPrtSetup
   Dim FrmSald As FrmSalyTotLibCajas
   Dim OldOrientacion As Integer
lPapelFoliado = False
      
   lOrientacion = ORIENT_HOR
   
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, False, lInfoPreliminar, False) = vbOK Then
      
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      Set FrmPrt = Nothing
      
      Set Frm = New FrmPrintPreview

   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault

   Call Frm.FView(Caption)
   Set Frm = Nothing

   Printer.Orientation = lOrientacion
   Me.MousePointer = vbDefault
   Call ResetPrtBas(gPrtReportes)
   
   End If
End Sub

Private Sub Bt_Print_Click()
Dim Frm As FrmPrintPreview
   Dim Frmu As FrmSalyTotLibCajas
   Dim Pag As Integer
   Dim FrmPrt As FrmPrtSetup
   Dim FrmSald As FrmSalyTotLibCajas
   Dim OldOrientacion As Integer
lPapelFoliado = False
      
   lOrientacion = ORIENT_HOR
   
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, False, lInfoPreliminar, False) = vbOK Then


   OldOrientation = Printer.Orientation

   Call SetUpPrtGrid

   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault

   Printer.Orientation = lOrientacion
   
   End If
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   'Titulos(1) = "N° Certificado: " & Txt_NCertificado
   'Titulos(2) = "RUT Empresa Fuente: " & Tx_Rut
   gPrtReportes.Titulos = Titulos
'   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = 0
   Next i
   
   ColWi(C_IDPERC) = 0
   ColWi(C_IDCUENTA) = 0
   ColWi(C_CUENTA) = 1100
   ColWi(C_IDEMPRESA) = 0
   ColWi(C_ANO) = 0
   ColWi(C_FECHA) = 1500
   ColWi(C_NUMCERTIFICADO) = 1300
   ColWi(C_RUTEMPRESA) = 1300
   ColWi(C_REGIMEN) = 0
   ColWi(C_DESREGIMEN) = 1600
   ColWi(C_CONTABILIZACION) = 0
   ColWi(C_DESCONTABILIZACION) = 1500
   ColWi(C_TASATEF) = 1100
   ColWi(C_TASATEX) = 1100
   ColWi(C_PERCEPCION) = 1500
   
                  
   'Total(C_DESC) = "Capital Pripio Tributario"
   'Total(C_TOTAL) = ""
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   'gPrtReportes.ColObligatoria = C_REGIMEN
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 0

End Sub

Private Sub Bt_Sum_Click()
Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing
End Sub

Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   'Dim Frm As Form
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Grid.TextMatrix(Row, C_IDCUENTA) <> "" Then
    Dim Frm As FrmPercepciones
    Set Frm = New FrmPercepciones
    Frm.CodCta = Grid.TextMatrix(Row, C_IDCUENTA)
    Frm.GIdPerc = Grid.TextMatrix(Row, C_IDPERC)
    Frm.Fecha = ""
    Frm.Show vbModal
    Set Frm = Nothing
   End If
   
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not ValidRut(Me.Tx_Rut.Text) Then
      MsgBox1 "Rut No Válido, Favor volver a ingresar", vbInformation
      Me.Tx_Rut.Text = ""
      Exit Sub
   End If

End Sub

Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)


   Grid.ColWidth(C_IDPERC) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CUENTA) = 1100
   Grid.ColWidth(C_IDEMPRESA) = 0
   Grid.ColWidth(C_ANO) = 0
   Grid.ColWidth(C_FECHA) = 1500
   Grid.ColWidth(C_NUMCERTIFICADO) = 1200
   Grid.ColWidth(C_RUTEMPRESA) = 1300
   Grid.ColWidth(C_REGIMEN) = 0
   Grid.ColWidth(C_DESREGIMEN) = 1600
   Grid.ColWidth(C_CONTABILIZACION) = 0
   Grid.ColWidth(C_DESCONTABILIZACION) = 1500
   Grid.ColWidth(C_TASATEF) = 1100
   Grid.ColWidth(C_TASATEX) = 1100
   Grid.ColWidth(C_PERCEPCION) = 1500
   
'   Grid.ColWidth(C_UPDATE) = 0
'   Grid.ColWidth(C_FMT) = 0
'
'   Grid.ColAlignment(C_OPENCLOSE) = flexAlignCenterCenter
'   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
'
  ' Grid.TextMatrix(0, C_IDPERC) = ""
   Grid.TextMatrix(0, C_IDCUENTA) = "ID Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   'Grid.TextMatrix(0, C_IDEMPRESA) = "ID Empresa"
   'Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_NUMCERTIFICADO) = "N° Certificado"
   Grid.TextMatrix(0, C_RUTEMPRESA) = "Rut Empresa"
   Grid.TextMatrix(0, C_REGIMEN) = ""
   Grid.TextMatrix(0, C_DESREGIMEN) = "Regimen"
   Grid.TextMatrix(0, C_CONTABILIZACION) = ""
   Grid.TextMatrix(0, C_DESCONTABILIZACION) = "Contabilizacion"
   Grid.TextMatrix(0, C_TASATEF) = "Tasa TEF"
   Grid.TextMatrix(0, C_TASATEX) = "Tasa TEX"
   Grid.TextMatrix(0, C_PERCEPCION) = "Percepciones"
   

   Grid.ColAlignment(C_IDCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
   Grid.ColAlignment(C_NUMCERTIFICADO) = flexAlignRightCenter
   Grid.ColAlignment(C_RUTEMPRESA) = flexAlignRightCenter
   Grid.ColAlignment(C_REGIMEN) = flexAlignRightCenter
   Grid.ColAlignment(C_DESREGIMEN) = flexAlignCenterCenter
   Grid.ColAlignment(C_CONTABILIZACION) = flexAlignRightCenter
   Grid.ColAlignment(C_DESCONTABILIZACION) = flexAlignCenterCenter
   Grid.ColAlignment(C_TASATEF) = flexAlignRightCenter
   Grid.ColAlignment(C_TASATEX) = flexAlignRightCenter
   Grid.ColAlignment(C_PERCEPCION) = flexAlignRightCenter

   
      
End Sub


Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer

   
   'Q1 = "SELECT IDPerc, IdCuenta, IdEmpresa, Ano, Fecha, NumCertificado, RutEmpresa, Regimen, Contabilizacion, TasaTef, TasaTex, Percepciones "
   Q1 = "SELECT IDPerc, p.IdCuenta, c.codigo, p.IdEmpresa, p.Ano, Fecha, NumCertificado, RutEmpresa, Regimen,  "
   Q1 = Q1 & "(SELECT valor FROM Param WHERE TIPO = 'REGEMPREFUE' AND codigo = Regimen) as desreg "
   Q1 = Q1 & ", Contabilizacion, "
   Q1 = Q1 & "(SELECT valor FROM Param WHERE TIPO = 'CONTABILIZA' AND codigo = Contabilizacion) as descont "
   Q1 = Q1 & ", TasaTef, TasaTex, Percepciones "
   Q1 = Q1 & " FROM Percepciones p, cuentas c "
   Q1 = Q1 & " WHERE p.idcuenta = c.idcuenta "
   If Tx_Desde.Text <> "" Then
    Q1 = Q1 & " AND Fecha >= " & GetTxDate(Tx_Desde)
   End If
   If Tx_Hasta.Text <> "" Then
    Q1 = Q1 & " AND FECHA <= " & GetTxDate(Tx_Hasta)
   End If
   If Tx_Rut.Text <> "" Then
    Q1 = Q1 & " AND p.RutEmpresa = " & vFmtRut(Tx_Rut)
   End If
   Q1 = Q1 & " ORDER BY IDPerc, Fecha "
   Set Rs = OpenRs(DbMain, Q1)
   

   i = 1
   Grid.rows = i
   Do While Rs.EOF = False
      Grid.rows = i + 1

      
      Grid.TextMatrix(i, C_IDPERC) = vFld(Rs("IDPerc"), True)
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"), True)
      Grid.TextMatrix(i, C_CUENTA) = FmtCodCuenta(vFld(Rs("codigo"), True))
      Grid.TextMatrix(i, C_IDEMPRESA) = vFld(Rs("IdEmpresa"), True)
      Grid.TextMatrix(i, C_ANO) = vFld(Rs("Ano"), True)
      Grid.TextMatrix(i, C_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
      Grid.TextMatrix(i, C_NUMCERTIFICADO) = vFld(Rs("NumCertificado"), True)
      Grid.TextMatrix(i, C_RUTEMPRESA) = FmtCID(vFld(Rs("RutEmpresa"), True))
      Grid.TextMatrix(i, C_REGIMEN) = vFld(Rs("Regimen"), True)
      Grid.TextMatrix(i, C_DESREGIMEN) = vFld(Rs("desreg"), True)
      Grid.TextMatrix(i, C_CONTABILIZACION) = vFld(Rs("Contabilizacion"), True)
      Grid.TextMatrix(i, C_DESCONTABILIZACION) = vFld(Rs("descont"), True)
      Grid.TextMatrix(i, C_TASATEF) = vFld(Rs("TasaTef"), True)
      Grid.TextMatrix(i, C_TASATEX) = vFld(Rs("TasaTex"), True)
      Grid.TextMatrix(i, C_PERCEPCION) = Format(vFld(Rs("Percepciones")), NUMFMT)
      
      i = i + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   'Call FGrVRows(Grid)
   
End Sub

