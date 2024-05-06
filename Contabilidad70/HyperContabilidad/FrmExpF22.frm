VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmExpF22 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a HR-Form 22"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Notas2017 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   8640
      Width           =   12375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- El traspaso de la CCMM al formulario 22 debe ser en el  código 637 o 638 a valores neteados."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   780
         Width           =   6720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- Los valores son montos históricos, no actualizados."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- Sólo se consideran comprobantes en estado APROBADO."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   4230
      End
      Begin VB.Label Label3 
         Caption         =   "Notas:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   60
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"FrmExpF22.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   1020
         Width           =   9735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"FrmExpF22.frx":008B
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   1260
         Width           =   11010
      End
   End
   Begin VB.TextBox Tx_CurrCell 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   8220
      Width           =   12435
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10395
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
         Left            =   2640
         Picture         =   "FrmExpF22.frx":012B
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Copiar Excel"
         Top             =   240
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
         Left            =   2220
         Picture         =   "FrmExpF22.frx":0570
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Imprimir"
         Top             =   240
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
         Left            =   1800
         Picture         =   "FrmExpF22.frx":0A2A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   11340
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Exp 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   11340
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin TabDlg.SSTab Tab_Form22 
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12303
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Recuadro Nº2: Base Impon. Primera Categoría"
      TabPicture(0)   =   "FrmExpF22.frx":0ED1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(1)=   "Grid2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recuadro Nº3: Contable Balance 8 Columnas"
      TabPicture(1)   =   "FrmExpF22.frx":0EED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid3"
      Tab(1).Control(1)=   "Grid4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Recuadro N° 6: Datos Informativos"
      TabPicture(2)   =   "FrmExpF22.frx":0F09
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Grid5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin FlexEdGrid2.FEd2Grid Grid1 
         Height          =   6375
         Left            =   -74820
         TabIndex        =   6
         Top             =   420
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11245
         Cols            =   3
         Rows            =   2
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
      Begin FlexEdGrid2.FEd2Grid Grid2 
         Height          =   6375
         Left            =   -68700
         TabIndex        =   7
         Top             =   420
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11245
         Cols            =   3
         Rows            =   2
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
      Begin FlexEdGrid2.FEd2Grid Grid3 
         Height          =   6375
         Left            =   -74820
         TabIndex        =   8
         Top             =   420
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11245
         Cols            =   3
         Rows            =   2
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
      Begin FlexEdGrid2.FEd2Grid Grid4 
         Height          =   6375
         Left            =   -68700
         TabIndex        =   9
         Top             =   420
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11245
         Cols            =   3
         Rows            =   2
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
      Begin FlexEdGrid2.FEd2Grid Grid5 
         Height          =   6375
         Left            =   180
         TabIndex        =   23
         Top             =   420
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   11245
         Cols            =   3
         Rows            =   2
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
   End
   Begin VB.Frame Fr_Notas2020 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   180
      TabIndex        =   18
      Top             =   8700
      Width           =   12375
      Begin VB.Label Label3 
         Caption         =   "Notas:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   22
         Top             =   60
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- Sólo se consideran comprobantes en estado APROBADO."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   21
         Top             =   540
         Width           =   4230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- Los valores son montos históricos, no actualizados."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "- Mayores Instruccione Circ 13/21 SII"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   780
         Width           =   2640
      End
   End
End
Attribute VB_Name = "FrmExpF22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const C_GLOSA = 0
Const C_CODIGO = 1
Const C_VALOR = 2
Const C_SIGNO = 3
Const C_FMT = 4


Const nRows = 26

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
Call LP_FGr2Clip(Grid5, Me.Caption)
End Sub

Private Sub Bt_Exp_Click()
   Dim Msg As String
   
'   If gDbType = SQL_SERVER Then
'      MsgBox1 "Esta funcionalidad aún no está disponible para la versión SQL Server.", vbInformation
'      Exit Sub
'   End If

   
   Msg = "¡ATENCION!" & vbLf & "Esta exportación reemplazará los valores actuales en el producto HR-Form 22."
   Msg = Msg & vbLf & vbLf & "Antes de realizar la exportación, asegúrese que ningún usuario tenga abierto el producto HR-Form 22 con la empresa " & gEmpresa.RazonSocial & " para el año " & gEmpresa.Ano & "."
   Msg = Msg & vbLf & vbLf & "Recuerde que el sistema exportará a HR Impuestos Finales mientras no se encuentre disponible HR-Formulario 22."
   Msg = Msg & vbLf & vbLf & "¿ Desea continuar ?"
   If MsgBox1(Msg, vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   Bt_Exp.Enabled = False
   DoEvents
   
   If GenDB_F22 = True Then
   
'   If ExportF22() = True Then
   
      Msg = "¡ATENCIÓN!" & vbLf & "Para terminar el proceso, ahora debe abrir el producto HR-Form 22, capturar los datos desde LPContabilidad y realizar el Recálculo."
      MsgBox1 Msg, vbInformation
   
   Else
      MsgBox1 "No se pudo realizar la exportación.", vbExclamation
   
   End If
   
   Bt_Exp.Enabled = True
   MousePointer = vbDefault
   
   
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   'Dim GridNew As FEd3Grid
   
   Me.MousePointer = vbHourglass
   
   
'   Grid5.Row = Row
'   Grid5.Col = Col
'
'      Grid5.CellFontBold = Value

   'Set GridNew = Grid5
   'Call FGrFontBold(Grid5, 0, -1, True)
   'Grid.TextMatrix(0, C_FMT) = "B"
   Call SetUpPrtGrid(Grid5)
   
   
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

Private Sub SetUpPrtGrid(Grid As FEd2Grid)
   Dim i As Integer
   Dim ColWi(5) As Integer
   Dim Total(5) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   
   'GridNew = Grid.FlxGrid
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "HR-FORM 22"
   gPrtReportes.Titulos = Titulos
   Encabezados(0) = "Recuadro N°6: Datos Informativos"
   gPrtReportes.Encabezados = Encabezados
         
   'gPrtReportes.GrFontName = Grid.FontName
   'gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
'   Total(0) = "PR"
'   Total(1) = "UEBA"
'   Total(2) = "GRILLA"
'   Total(3) = "VISTA"
'   Total(4) = "OTRO"
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = 0
   gPrtReportes.NTotLines = 0
   gPrtReportes.FmtCol = C_FMT

End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid(Grid5)
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
End Sub

Private Sub Form_Activate()

   Tab_Form22.Tab = 0
   Tab_Form22.TabEnabled(2) = False
   Fr_Notas2020.visible = False
   
   If gEmpresa.Ano >= 2017 And gEmpresa.Ano < 2020 Then
      Tab_Form22.Tab = 1
      DoEvents
      Tab_Form22.TabEnabled(0) = False
   ElseIf gEmpresa.Ano >= 2020 Then
      Tab_Form22.Tab = 2
      DoEvents
      Tab_Form22.TabEnabled(0) = False
      Tab_Form22.TabEnabled(1) = False
      Tab_Form22.TabEnabled(2) = True
      Fr_Notas2017.visible = False
      Fr_Notas2020.visible = True
   End If

End Sub

Private Sub Form_Load()

   Call SetUpGrid
      
   Call LoadGridSinValor(Grid5)
   Call LoadAll

End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrSetup(Grid1)
   Grid1.Cols = 5
   Grid1.FixedCols = 2
   Grid1.FixedRows = 0
   
   Grid1.rows = nRows
            
   Grid1.ColWidth(C_GLOSA) = 4200
   Grid1.ColWidth(C_CODIGO) = 500
   Grid1.ColWidth(C_VALOR) = 1200
   Grid1.ColWidth(C_SIGNO) = 0
   Grid1.ColWidth(C_FMT) = 0
   

   Grid1.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid1.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid1.ColAlignment(C_SIGNO) = flexAlignCenterCenter
   Grid1.ColAlignment(C_FMT) = flexAlignCenterCenter
   
   Call FGrSetup(Grid2)
   Grid2.Cols = 5
   Grid2.FixedCols = 2
   Grid2.FixedRows = 0
   
   Grid2.rows = nRows
            
   Grid2.ColWidth(C_GLOSA) = 4200
   Grid2.ColWidth(C_CODIGO) = 500
   Grid2.ColWidth(C_VALOR) = 1200
   Grid2.ColWidth(C_SIGNO) = 0
   Grid2.ColWidth(C_FMT) = 0

   Grid2.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid2.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid2.ColAlignment(C_SIGNO) = flexAlignRightCenter
   Grid2.ColAlignment(C_FMT) = flexAlignCenterCenter
   
   Call FGrSetup(Grid3)
   Grid3.Cols = 5
   Grid3.FixedCols = 2
   Grid3.FixedRows = 0
   
   Grid3.rows = nRows
            
   Grid3.ColWidth(C_GLOSA) = 4200
   Grid3.ColWidth(C_CODIGO) = 500
   Grid3.ColWidth(C_VALOR) = 1200
   Grid3.ColWidth(C_SIGNO) = 0
   Grid3.ColWidth(C_FMT) = 0

   Grid3.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid3.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid3.ColAlignment(C_SIGNO) = flexAlignRightCenter
   Grid3.ColAlignment(C_FMT) = flexAlignCenterCenter
      
   Call FGrSetup(Grid4)
   Grid4.Cols = 5
   Grid4.FixedCols = 2
   Grid4.FixedRows = 0
   
   Grid4.rows = nRows
            
   Grid4.ColWidth(C_GLOSA) = 4200
   Grid4.ColWidth(C_CODIGO) = 500
   Grid4.ColWidth(C_VALOR) = 1200
   Grid4.ColWidth(C_SIGNO) = 0
   Grid4.ColWidth(C_FMT) = 0

   Grid4.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid4.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid4.ColAlignment(C_SIGNO) = flexAlignRightCenter
   Grid4.ColAlignment(C_FMT) = flexAlignCenterCenter
      
   Call FGrSetup(Grid5)
   Grid5.Cols = 5
   Grid5.FixedCols = 2
   Grid5.FixedRows = 0
   
   Grid5.rows = nRows
            
   Grid5.ColWidth(C_GLOSA) = 6300
   Grid5.ColWidth(C_CODIGO) = 500
   Grid5.ColWidth(C_VALOR) = 1200
   Grid5.ColWidth(C_SIGNO) = 0
   Grid5.ColWidth(C_FMT) = 0

   Grid5.ColAlignment(C_CODIGO) = flexAlignRightCenter
   Grid5.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid5.ColAlignment(C_SIGNO) = flexAlignRightCenter
   Grid5.ColAlignment(C_FMT) = flexAlignRightCenter
      
   Call FillGridGlosaCodigo
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim WhFecha As String
   Dim Valor As Double
   
   WhFecha = " AND Comprobante.Fecha BETWEEN " & CLng(Int(DateSerial(gEmpresa.Ano, 1, 1))) & " AND " & CLng(Int(DateSerial(gEmpresa.Ano, 12, 31)))
   
   
   Tx_Ano = gEmpresa.Ano
      
   
'   Q1 = "SELECT Cuentas.CodF22, Sum(MovComprobante.Debe-MovComprobante.Haber) As SumCta "    'FCA 17 abr 2013
   Q1 = "SELECT Cuentas.CodF22, Cuentas.Clasificacion, "
   Q1 = Q1 & " Sum(iif(Cuentas.Clasificacion = " & CLASCTA_ACTIVO & ", MovComprobante.Debe-MovComprobante.Haber, MovComprobante.Haber-MovComprobante.Debe))  As SumCta "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.CodF22 <> 0 AND NOT (Cuentas.CodF22 IS NULL) AND Comprobante.Estado = " & EC_APROBADO
   Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.CodF22, Cuentas.Clasificacion "
   Q1 = Q1 & " ORDER BY Cuentas.CodF22"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid1.FlxGrid.Redraw = False
   Grid2.FlxGrid.Redraw = False
   Grid3.FlxGrid.Redraw = False
   Grid4.FlxGrid.Redraw = False
   Grid5.FlxGrid.Redraw = False
   
  Do While Rs.EOF = False
   
'      If vFld(Rs("CodF22")) <> 636 And vFld(Rs("CodF22")) <> 643 Then    'estros
         
         If Not LoadGrid(Grid1, Rs) Then
            If Not LoadGrid(Grid2, Rs) Then
               If Not LoadGrid(Grid3, Rs) Then
                  Call LoadGrid(Grid4, Rs)
               End If
            End If
         End If
         
         Call LoadGrid(Grid5, Rs)    'hay códigos que se repiten en las otras gruilla, por lo que hay que llamarla siempre

         
'      End If
         
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   
   'Parche: Total activos a código 122       'FCA 8 abr 2015
   Q1 = "SELECT Sum(MovComprobante.Debe - MovComprobante.Haber) As TotalActivos "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion =" & CLASCTA_ACTIVO
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      If gEmpresa.Ano < 2020 Then
         For i = Grid3.FixedRows To Grid3.rows - 1
            If Val(Grid3.TextMatrix(i, C_CODIGO)) = 122 Then
               Grid3.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("TotalActivos"))), NUMFMT)
               Exit For
            End If
         Next i
      Else
         For i = Grid5.FixedRows To Grid5.rows - 1
            If Val(Grid5.TextMatrix(i, C_CODIGO)) = 122 Then
               Grid5.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("TotalActivos"))), NUMFMT)
               Exit For
            End If
         Next i
      End If
   End If
   
   Call CloseRs(Rs)
   
   'Parche: Total Pasivos a código 123       'FCA 8 abr 2015
   Q1 = "SELECT Sum(MovComprobante.Debe - MovComprobante.Haber) As TotalPasivos "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion =" & CLASCTA_PASIVO
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      If gEmpresa.Ano < 2020 Then
         For i = Grid3.FixedRows To Grid3.rows - 1
             If Val(Grid3.TextMatrix(i, C_CODIGO)) = 123 Then
                Grid3.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("TotalPasivos"))), NUMFMT)
                Exit For
             End If
          Next i
      Else
         For i = Grid5.FixedRows To Grid5.rows - 1
             If Val(Grid5.TextMatrix(i, C_CODIGO)) = 123 Then
                Grid5.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("TotalPasivos"))), NUMFMT)
                Exit For
             End If
          Next i
      End If
      
   End If
   
   Call CloseRs(Rs)
   
   'Parche: Si el año es 2017 o superior y el código es distinto de 101, 784, 129, 645, 646, 647, 648, 122, 123 o 844, dicho código será siempre igual a 0.   FCA: 10 ene 2018
   'Parche: Si el año es 2017 y el código es distinto de 101, 784, 129, 645, 646, 647, 648, 122, 123 o 844, dicho código será siempre igual a 0.   FCA: 28 abr 2020 por indicación de Víctor Morales
   If gEmpresa.Ano = 2017 Then
      For i = Grid3.FixedRows To Grid3.rows - 1
         If InStr("," & LSTCODF22_2017 & ",", "," & Trim(Grid3.TextMatrix(i, C_CODIGO)) & ",") <= 0 Then     'es inválido
            Grid3.TextMatrix(i, C_VALOR) = ""
         End If
      Next i
   End If
   
   'agregamos Capital Propio   FCA 23 abr 2020
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then

      Valor = Val(vFld(Rs("Valor")))
   
      For i = Grid3.FixedRows To Grid3.rows - 1
         If Val(Grid3.TextMatrix(i, C_CODIGO)) = 645 Then
            If Valor >= 0 Then
               Grid3.TextMatrix(i, C_VALOR) = Format(Valor, NUMFMT)
               Exit For
            End If
         End If
         If Val(Grid3.TextMatrix(i, C_CODIGO)) = 646 Then
            If Valor < 0 Then
               Grid3.TextMatrix(i, C_VALOR) = Format(Abs(Valor), NUMFMT)
               Exit For
            End If
         End If
      Next i
            
   End If
   
   Call CloseRs(Rs)

    Dim Frm As FrmCapitalPropio
    '3054289 se cambia long por double
    Dim Monto As Double
    '3054289
    Set Frm = New FrmCapitalPropio
    Frm.GetcapitalEfectivo
    Monto = Frm.capitalEfectivo
    Set Frm = Nothing

      If gEmpresa.Ano < 2020 Then
         For i = Grid3.FixedRows To Grid3.rows - 1
            If Val(Grid3.TextMatrix(i, C_CODIGO)) = 102 Then
               Grid3.TextMatrix(i, C_VALOR) = Format(Abs(Monto), NUMFMT)
               Exit For
            End If
         Next i
      Else
         For i = Grid5.FixedRows To Grid5.rows - 1
            If Val(Grid5.TextMatrix(i, C_CODIGO)) = 102 Then
               Grid5.TextMatrix(i, C_VALOR) = Format(Abs(Monto), NUMFMT)
               Exit For
            End If
         Next i
      End If
      
    Dim Frm1 As FrmBalTributario
    Set Frm1 = New FrmBalTributario
    Frm1.GetPatrimonio
    Monto = Frm1.Patrimonio
    Set Frm1 = Nothing

      If gEmpresa.Ano < 2020 Then
         For i = Grid3.FixedRows To Grid3.rows - 1
            If Val(Grid3.TextMatrix(i, C_CODIGO)) = 843 Then
               Grid3.TextMatrix(i, C_VALOR) = Format(Abs(Monto), NUMFMT)
               Exit For
            End If
         Next i
      Else
         For i = Grid5.FixedRows To Grid5.rows - 1
            If Val(Grid5.TextMatrix(i, C_CODIGO)) = 843 Then
               Grid5.TextMatrix(i, C_VALOR) = Format(Abs(Monto), NUMFMT)
               Exit For
            End If
         Next i
      End If
   

         
   Grid1.TopRow = Grid1.FixedRows + 1
   Grid2.TopRow = Grid2.FixedRows
   Grid3.TopRow = Grid3.FixedRows
   Grid4.TopRow = Grid3.FixedRows
   Grid5.TopRow = Grid3.FixedRows
         
   Call FGrSelRow(Grid1, Grid1.FixedRows)
   Call FGrSelRow(Grid2, Grid2.FixedRows)
   Call FGrSelRow(Grid3, Grid3.FixedRows)
   Call FGrSelRow(Grid4, Grid4.FixedRows)
   Call FGrSelRow(Grid5, Grid4.FixedRows)
   
   Grid1.FlxGrid.Redraw = True
   Grid2.FlxGrid.Redraw = True
   Grid3.FlxGrid.Redraw = True
   Grid4.FlxGrid.Redraw = True
   Grid5.FlxGrid.Redraw = True
   
   Tab_Form22.Tab = 0
      

End Sub

Private Function LoadGridSinValor(Grid As FEd2Grid) As Boolean
   Dim i As Integer
   
   LoadGridSinValor = False
            
'   If vFld(Rs("CodF22")) <> 636 And vFld(Rs("CodF22")) <> 643 Then    'estos se calculan en Form22
      
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_GLOSA) <> "" And Grid.TextMatrix(i, C_CODIGO) <> "" Then
            Grid.TextMatrix(i, C_VALOR) = Format(0, NEGNUMFMT)
         End If
         LoadGridSinValor = True
      Next i
      
'   End If
                           
End Function

Private Function LoadGrid(Grid As FEd2Grid, Rs As Recordset) As Boolean
   Dim i As Integer
   
   LoadGrid = False
            
'   If vFld(Rs("CodF22")) <> 636 And vFld(Rs("CodF22")) <> 643 Then    'estos se calculan en Form22
      
      For i = Grid.FixedRows To Grid.rows - 1
        
         If Val(Grid.TextMatrix(i, C_CODIGO)) = vFld(Rs("CodF22")) Then
                           
            If vFld(Rs("CodF22")) <> 304 And vFld(Rs("CodF22")) <> 305 And vFld(Rs("CodF22")) <> 618 And vFld(Rs("CodF22")) <> 843 And vFld(Rs("CodF22")) <> 844 Then   'esto se modificó por solicitud de Victor Morales 24 abr 2009
               Grid.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("SumCta"))), NEGNUMFMT)
            Else
               Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("SumCta")), NEGNUMFMT)
            End If
            
            LoadGrid = True
            Exit Function
'         Else
'            If Grid.TextMatrix(i, C_GLOSA) = "" Then
'                Grid.TextMatrix(i, C_VALOR) = Format(0, NEGNUMFMT)
'            End If
         End If
         
      Next i
      
'   End If
                           
End Function

Private Function GenDB_F22() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rs As Recordset
   Dim Q1 As String
   Dim TblName As String
   Dim i As Integer
   Dim UpdOK As Boolean
   Dim Rc As Integer
   Dim DbF22 As Database
   Dim DbF22Path As String, sAno As String
   Dim DbExpHR As String
   
   GenDB_F22 = False
   
   If Not gFunciones.NuevoTraspasoForm22 Then
      Exit Function
   End If
         
   'creamos la db
      
   
   On Error Resume Next
   
   DbF22Path = gHRPath & "\RUTS\" & Right("000000000" & gEmpresa.Rut, 8)
   MkDir DbF22Path
   DbF22Path = DbF22Path & "\ImpConta"
   MkDir DbF22Path
   
   DbF22Path = DbF22Path & "\F22_" & Right(gEmpresa.Ano, 2) & ".mdb"
           
   ERR.Clear
         
   DbExpHR = gDbPath & "\ExpHR_F22_F29.mdb"           'dado que usamos DAO 3.6 y HR no, usamos una base vacia creada en DAO 3.5 como base para exportar a HR
   Call CopyFile(DbExpHR, DbF22Path, True)
   
   If ERR Then
      Call MsgErr(DbExpHR)
      Exit Function
   End If
   
   Set DbF22 = OpenDatabase(DbF22Path)
   
   'creamos la tabla con la fecha y version de exportacion
   sAno = Right("0" & gEmpresa.Ano, 2)
   TblName = "Param_" & sAno
   
   'vemos si la tabla existe
   Q1 = "SELECT Id FROM " & TblName
   Set Rs = OpenRsDao(DbF22, Q1, False)
   
   If Rs Is Nothing Then
      'no existe
      
      'Creamos la tabla de parametros
   
      Set Tbl = DbF22.CreateTableDef(TblName)
      
      ERR.Clear
    
      Tbl.Fields.Append Tbl.CreateField("Id", dbLong)
      Tbl.Fields("Id").Attributes = dbAutoIncrField
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Id", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Codigo", dbText, 15)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Valor", dbText, 30)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Valor", vbExclamation
         UpdOK = False
      End If
                 
      DbF22.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF22.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (Id)"
         Rc = ExecSQLDao(DbF22, Q1, False)
         
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF22, "DELETE * FROM " & TblName)
      
   End If
     
   'insertamos registros de la exportación
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Version','" & W.Version & "')"
   Call ExecSQLDao(DbF22, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Fecha Version','" & Format(W.FVersion, "dd mmm yy") & "')"
   Call ExecSQLDao(DbF22, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Fecha Export','" & Format(Now, "dd mmm yy hh:mm") & "')"
   Call ExecSQLDao(DbF22, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'RUT','" & FmtCID(gEmpresa.Rut) & "')"
   Call ExecSQLDao(DbF22, Q1)

   
   'creamos la tabla con el código y el valor
   TblName = "ExpF22_" & sAno
   
   'vemos si la tabla existe
   Q1 = "SELECT Id FROM " & TblName
   Set Rs = OpenRsDao(DbF22, Q1, False)
   
   If Rs Is Nothing Then
      'no existe
      
      'Creamos la tabla ExpF22_mmyy
   
      Set Tbl = DbF22.CreateTableDef(TblName)
      
      ERR.Clear
    
      Tbl.Fields.Append Tbl.CreateField("Id", dbLong)
      Tbl.Fields("Id").Attributes = dbAutoIncrField
      
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Id", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Codigo", dbText, 15)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Valor", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Valor", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Descripcion", dbText, 50)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Descripcion", vbExclamation
         UpdOK = False
      End If
           
      DbF22.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF22.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Codigo ON " & TblName & " (Codigo)"
         Rc = ExecSQLDao(DbF22, Q1, False)
         
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
   
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF22, "DELETE * FROM " & TblName)
      
   End If
    
   If Tab_Form22.TabEnabled(0) Then
      Call ExpGrid(DbF22, Grid1, TblName)
      Call ExpGrid(DbF22, Grid2, TblName)
   End If
   If Tab_Form22.TabEnabled(1) Then
      Call ExpGrid(DbF22, Grid3, TblName)
      Call ExpGrid(DbF22, Grid4, TblName)
   End If
   If Tab_Form22.TabEnabled(2) Then
      Call ExpGrid(DbF22, Grid5, TblName)
   End If
  
   Call CloseDb(DbF22)
   
   GenDB_F22 = True
      
End Function
'#If DATACON = 1 Then
Private Sub ExpGrid(DbF22 As Database, Grid As FEd2Grid, TblName As String)
'#Else
'Private Sub ExpGrid(DbF22 As Connection, Grid As FEd2Grid, TblName As String)
'#End If
   Dim i As Integer, Q1 As String
   
   Q1 = "INSERT INTO " & TblName & " (Codigo, Valor, Descripcion) VALUES("
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Trim(Grid.TextMatrix(i, C_GLOSA)) = "" Then
         Exit For
      End If
      
      If vFmt(Grid.TextMatrix(i, C_CODIGO)) <> 0 Then
         Call ExecSQLDao(DbF22, Q1 & "'" & Grid.TextMatrix(i, C_CODIGO) & "', " & vFmt(Grid.TextMatrix(i, C_VALOR)) & ", ' ' )")
      End If
                  
   Next i
    

End Sub


Private Sub Grid1_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Grid1.TextMatrix(Row, C_CODIGO) = "637" Or Grid1.TextMatrix(Row, C_CODIGO) = "638" Then
      Value = Format(vFmt(Value), NEGNUMFMT)
   End If

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)

   If Grid1.TextMatrix(Row, C_CODIGO) = "637" Or Grid1.TextMatrix(Row, C_CODIGO) = "638" Then
      EdType = FEG_Edit
   End If
   
End Sub

Private Sub Grid1_EditKeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub FillGridGlosaCodigo()
   Dim i As Integer
   
   'Recuadro Nº2 -Izq
   'Grid1
   
   i = 0
   Grid1.TextMatrix(i, C_CODIGO) = "628"
   Grid1.TextMatrix(i, C_GLOSA) = "Ingresos del Giro Percibidos o Devengados"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "851"
   Grid1.TextMatrix(i, C_GLOSA) = "Rentas de Fuente Extranjera"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "629"
   Grid1.TextMatrix(i, C_GLOSA) = "Intereses Percibidos o Devengados"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "651"
   Grid1.TextMatrix(i, C_GLOSA) = "Otros Ingresos Percibidos o Devengados"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "630"
   Grid1.TextMatrix(i, C_GLOSA) = "Costo Directo de los Bienes y Servicios"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "631"
   Grid1.TextMatrix(i, C_GLOSA) = "Remuneraciones"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "632"
   Grid1.TextMatrix(i, C_GLOSA) = "Depreciación Financiera del Ejercicio"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "633"
   Grid1.TextMatrix(i, C_GLOSA) = "Intereses Pagados o Adeudados"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "966"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Donaciones"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "967"
   Grid1.TextMatrix(i, C_GLOSA) = "Otros Gastos Financieros"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
      
'   Grid1.TextMatrix(i, C_CODIGO) = "792"                'se elimina en mar 2016
'   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Donaciones para fines Sociales"
'   Grid1.TextMatrix(i, C_SIGNO) = "-"
'   i = i + 1
   
'   Grid1.TextMatrix(i, C_CODIGO) = "793"                'se elimina en mar 2016
'   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Donaciones para fines Políticos"
'   Grid1.TextMatrix(i, C_SIGNO) = "-"
'   i = i + 1
   
'   Grid1.TextMatrix(i, C_CODIGO) = "772"                'se elimina en mar 2016
'   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por otras Donaciones según Art. N° 10, Ley N° 19.885"
'   Grid1.TextMatrix(i, C_SIGNO) = "-"
'   i = i + 1
      
'   Grid1.TextMatrix(i, C_CODIGO) = "873"                  'se elimina en mar 2016
'   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por donaciones según Art. 7° Ley N° 16.282/1965."
'   Grid1.TextMatrix(i, C_SIGNO) = "-"
'   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "852"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Inversión en Investigación y Desarrollo certificados por Corfo"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "897"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Inversión en Investigación y Desarrollo no certificados por Corfo"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "853"
   Grid1.TextMatrix(i, C_GLOSA) = "Costos y  Gastos necesarios para producir las Rentas de Fuente Extranjera"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "941"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Responsabilidad Social"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "968"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por Impuesto Renta e Impuesto Diferido"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "969"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos por adquisición en supermercados y negocios similares"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "635"
   Grid1.TextMatrix(i, C_GLOSA) = "Otros Gastos Deducidos de los Ingresos Brutos"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "636"
   Grid1.TextMatrix(i, C_GLOSA) = "Renta Líquida (o Pérdida)"
   Grid1.TextMatrix(i, C_SIGNO) = "="
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "637"
   Grid1.TextMatrix(i, C_GLOSA) = "Corrección Monetaria Saldo Deudor (Art. 32)"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid1, i, "BC", COLOR_EDITCELL, C_VALOR, C_VALOR)
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "638"
   Grid1.TextMatrix(i, C_GLOSA) = "Corrección Monetaria Saldo Acreedor (Art. 32)"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   Call FGrSetRowStyle(Grid1, i, "BC", COLOR_EDITCELL, C_VALOR, C_VALOR)
   i = i + 1
           
   Grid1.TextMatrix(i, C_CODIGO) = "926"
   Grid1.TextMatrix(i, C_GLOSA) = "Depreciación Financiera del Ejercicio"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid1.TextMatrix(i, C_CODIGO) = "927"
   Grid1.TextMatrix(i, C_GLOSA) = "Depreciación Tributaria del Ejercicio"
   Grid1.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
           
   Grid1.TextMatrix(i, C_CODIGO) = "970"
   Grid1.TextMatrix(i, C_GLOSA) = "Rentas tributables no reconocidas financieramente"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
           
   Grid1.TextMatrix(i, C_CODIGO) = "971"
   Grid1.TextMatrix(i, C_GLOSA) = "Gastos agregados por donaciones"
   Grid1.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
  
         
            
   'Recuadro Nº2 -Der
   'Grid(2)
   
            
    i = 0
   
   Grid2.TextMatrix(i, C_CODIGO) = "639"
   Grid2.TextMatrix(i, C_GLOSA) = "Gastos que se deben agregar a la RLI según el Nº1 del Art. 33"
   Grid2.TextMatrix(i, C_SIGNO) = "+"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "1000"
   Grid2.TextMatrix(i, C_GLOSA) = " Gasto Goodwill Tributario del ejercicio"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
'   Grid2.TextMatrix(i, C_CODIGO) = "794"                   'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Gastos Rechazados por Donaciones para fines Sociales"
'   Grid2.TextMatrix(i, C_SIGNO) = "+"
'   i = i + 1
   
'   Grid2.TextMatrix(i, C_CODIGO) = "812"                   'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Gastos Rechazados por Donaciones para fines Políticos"
'   Grid2.TextMatrix(i, C_SIGNO) = "+"
'   i = i + 1
   
'   Grid2.TextMatrix(i, C_CODIGO) = "811"                   'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Gastos Rechazados por otras Donaciones según Art. N° 10, Ley N° 19.885"
'   Grid2.TextMatrix(i, C_SIGNO) = "+"
'   i = i + 1
   
'   Grid2.TextMatrix(i, C_CODIGO) = "876"                   'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Gastos rechazados por donaciones al FNR (Art. 4° y 9° Ley N° 20.444/2010)."
'   Grid2.TextMatrix(i, C_SIGNO) = "+"
'   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "827"
   Grid2.TextMatrix(i, C_GLOSA) = "Impuesto Específico a la Actividad Minera"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "634"
   Grid2.TextMatrix(i, C_GLOSA) = "Pérdidas de Ejercicios Anteriores (Art. 31 N°3)"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "928"
   Grid2.TextMatrix(i, C_GLOSA) = "Gastos Rechazados afectos a la tributación del Inc. 1º Art. 21"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "929"
   Grid2.TextMatrix(i, C_GLOSA) = "Gastos Rechazados afectos a la tributación del Inc. 3º Art. 21"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
  
   Grid2.TextMatrix(i, C_CODIGO) = "640"
   Grid2.TextMatrix(i, C_GLOSA) = "Ingresos No Renta (Art. 17) "
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "807"
   Grid2.TextMatrix(i, C_GLOSA) = "Otras Partidas"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "641"
   Grid2.TextMatrix(i, C_GLOSA) = "Rentas Exentas Impto. 1ª Categoría (Art. 33 N°2)"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "642"
   Grid2.TextMatrix(i, C_GLOSA) = "Dividendos y/o Utilidades Sociales (Art.33 N°2)"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "972"
   Grid2.TextMatrix(i, C_GLOSA) = "Renta Líquida (ó Pérdida) antes de rebajar como gasto donaciones"
   Grid2.TextMatrix(i, C_SIGNO) = "="
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "973"
   Grid2.TextMatrix(i, C_GLOSA) = "Gastos aceptados por donaciones"
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
'   Grid2.TextMatrix(i, C_CODIGO) = "874"                'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Renta Líquida (o Pérdida) antes de rebajar como gasto las donaciones al FNR."
'   Grid2.TextMatrix(i, C_SIGNO) = "="
'   i = i + 1
         
'   Grid2.TextMatrix(i, C_CODIGO) = "875"                  'se elimina en mar 2016
'   Grid2.TextMatrix(i, C_GLOSA) = "Gastos aceptados por donaciones al FNR (Art. 4° y 9° Ley N° 20.444/2010)."
'   Grid2.TextMatrix(i, C_SIGNO) = "-"
'   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "868"
   Grid2.TextMatrix(i, C_GLOSA) = "Rentas Exentas de Impuesto de Primera Categoría (Art. 14 quáter y Art. 40 N° 7)."
   Grid2.TextMatrix(i, C_SIGNO) = "-"
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "643"
   Grid2.TextMatrix(i, C_GLOSA) = "Renta Líquida Imponible (o Pérdida)"
   Grid2.TextMatrix(i, C_SIGNO) = "="
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "808"
   Grid2.TextMatrix(i, C_GLOSA) = "Base imponible Renta Presunta"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "758"
   Grid2.TextMatrix(i, C_GLOSA) = "Rentas afectas al Impuesto Único de Primera Categoría"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "809"
   Grid2.TextMatrix(i, C_GLOSA) = "Rentas por arriendos de Bienes Raíces Agrícolas"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "759"
   Grid2.TextMatrix(i, C_GLOSA) = "Rentas por arriendos de Bienes Raíces No Agrícolas"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid2.TextMatrix(i, C_CODIGO) = "760"
   Grid2.TextMatrix(i, C_GLOSA) = "Otras rentas afectas al Impuesto de Primera Categoría"
'   Grid2.TextMatrix(i, C_GLOSA) = "Renta Neta de Fuente Extranjera (art. 41 A letra E N° 6)"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid2.TextMatrix(i, C_CODIGO) = "974"
   Grid2.TextMatrix(i, C_GLOSA) = "Renta Neta de Fuente Extranjera (art. 41 A letra D N° 6)"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid2.TextMatrix(i, C_CODIGO) = "975"
   Grid2.TextMatrix(i, C_GLOSA) = "Gastos adeudados o pagados por cuotas  de bienes en leasing"
'   Grid2.TextMatrix(i, C_GLOSA) = "Total cant. contab. en gasto contraídas con relacionados en el exterior(Arts.31 inc. 3 y 59)"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid2.TextMatrix(i, C_CODIGO) = "976"
'   Grid2.TextMatrix(i, C_GLOSA) = "Total cant. adeud., pag. o abon. a relacionados en el exterior(Arts.31 inc. 3 y 59 LIR)"
   Grid2.TextMatrix(i, C_GLOSA) = "Total cant. contab. en gasto contraídas con relacionados en el exterior (Arts. 31 inc. 3 y 59)."
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid2.TextMatrix(i, C_CODIGO) = "1019"
   Grid2.TextMatrix(i, C_GLOSA) = "Beneficio antes de Gastos Financieros (EBITDA)"
   Grid2.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            

   'Recuadro Nº3
   'Grid(3)
   
            
    i = 0
           
   
   Grid3.TextMatrix(i, C_CODIGO) = "101"
   Grid3.TextMatrix(i, C_GLOSA) = "Saldo de Caja (sólo dinero en efectivo y documentos al día según arqueo)"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "784"
   Grid3.TextMatrix(i, C_GLOSA) = "Saldo cuenta corriente bancaria según conciliación"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "783"
   Grid3.TextMatrix(i, C_GLOSA) = "Préstamos efectuados a propietarios, socios o accionistas en el ejercicio"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "977"
   Grid3.TextMatrix(i, C_GLOSA) = "Cuentas por Cobrar  (por Ventas o Servicios)"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "129"
   Grid3.TextMatrix(i, C_GLOSA) = "Existencia Final"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "647"
   Grid3.TextMatrix(i, C_GLOSA) = "Activo Inmovilizado"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "940"
   Grid3.TextMatrix(i, C_GLOSA) = "Cantidad de Bienes del Activo Inmovilizado"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   'estos códigos ya no están para el 2014
   
'   Grid3.TextMatrix(i, C_CODIGO) = "785"
'   Grid3.TextMatrix(i, C_GLOSA) = "Depreciación tributaria normal del ejercicio"
'   Grid3.TextMatrix(i, C_SIGNO) = ""
'   i = i + 1
'
'   Grid3.TextMatrix(i, C_CODIGO) = "938"
'   Grid3.TextMatrix(i, C_GLOSA) = "Depreciación tributaria acelerada del ejercicio"
'   Grid3.TextMatrix(i, C_SIGNO) = ""
'   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "648"
   Grid3.TextMatrix(i, C_GLOSA) = "Bienes Adquiridos Contrato Leasing"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "978"
   Grid3.TextMatrix(i, C_GLOSA) = "Cant. adeud. a relacionados en el exterior, o pag. cuyo impto. adic. no ha sido enterado (Arts.31 inc. 3 y 59)"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "815"
   Grid3.TextMatrix(i, C_GLOSA) = "Monto inversión Ley Arica"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "741"
   Grid3.TextMatrix(i, C_GLOSA) = "Monto inversión Ley Austral"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "122"
   Grid3.TextMatrix(i, C_GLOSA) = "Total del Activo"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "1020"
   Grid3.TextMatrix(i, C_GLOSA) = "Total Pasivos Contraídos en Chile"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "123"
   Grid3.TextMatrix(i, C_GLOSA) = "Total del Pasivo"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "102"
   Grid3.TextMatrix(i, C_GLOSA) = "Capital Efectivo"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "645"
   Grid3.TextMatrix(i, C_GLOSA) = "Capital Propio Tributario Positivo"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "1023"
   Grid3.TextMatrix(i, C_GLOSA) = "Diferencia entre el CPT y el capital aportado, FUT, FUNT y FUR"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "893"
   Grid3.TextMatrix(i, C_GLOSA) = "Aumentos de Capital"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "894"
   Grid3.TextMatrix(i, C_GLOSA) = "Disminuciones de Capital"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "646"
   Grid3.TextMatrix(i, C_GLOSA) = "Capital Propio Tributario Negativo"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "1021"
   Grid3.TextMatrix(i, C_GLOSA) = "Monto del capital directa o indirectamente financiado por partes relacionadas"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "843"
   Grid3.TextMatrix(i, C_GLOSA) = "Patrimonio Financiero"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "844"
   Grid3.TextMatrix(i, C_GLOSA) = "Total Capital Enterado"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid3.TextMatrix(i, C_CODIGO) = "1003"
   Grid3.TextMatrix(i, C_GLOSA) = "Activo Gasto Diferido Goodwill Tributario"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid3.TextMatrix(i, C_CODIGO) = "1004"
   Grid3.TextMatrix(i, C_GLOSA) = "Activo Intangible Goodwill Tributario (Ley N° 20.780)"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   Grid3.TextMatrix(i, C_CODIGO) = "1005"
   Grid3.TextMatrix(i, C_GLOSA) = "Utilidades Financieras Capitalizadas y Sobreprecio en Colocación de Acciones"
   Grid3.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
            
   'Recuadro Nº3, parte inferior
   'Grid4
   
            
   i = 0
   
   Grid4.TextMatrix(i, C_CODIGO) = ""
   Grid4.TextMatrix(i, C_GLOSA) = "Depreciación Tributaria"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   Grid4.TextMatrix(i, C_FMT) = "B"
   Call FGrSetRowStyle(Grid4, i, "B", 0, C_GLOSA, C_GLOSA)
   Call FGrSetRowStyle(Grid4, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
   i = i + 1

   Grid4.TextMatrix(i, C_CODIGO) = "785"
   Grid4.TextMatrix(i, C_GLOSA) = "Total depreciación normal de los bienes en el ejercicio"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid4.TextMatrix(i, C_CODIGO) = "950"
   Grid4.TextMatrix(i, C_GLOSA) = "Total dep. normal de los bienes con dep. acelerada informada en los cód. 938, 942 y/o 949"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid4.TextMatrix(i, C_CODIGO) = "938"
   Grid4.TextMatrix(i, C_GLOSA) = "Depreciación tributaria acelerada del ejercicio"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid4.TextMatrix(i, C_CODIGO) = "942"
   Grid4.TextMatrix(i, C_GLOSA) = "Depreciación acelerada en 1 año (Art. 31 N° 5 bis)"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid4.TextMatrix(i, C_CODIGO) = "949"
   Grid4.TextMatrix(i, C_GLOSA) = "Depreciación acelerada en 1/10 de la vida útil normal (Art. 31 N° 5 bis)"
   Grid4.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   
   
   'Recuadro Nº6
   'Grid5
   i = 0
   
   Grid5.TextMatrix(i, C_CODIGO) = ""
   Grid5.TextMatrix(i, C_GLOSA) = "Operaciones Internacionales"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   Grid5.TextMatrix(i, C_FMT) = "B"
   Call FGrSetRowStyle(Grid5, i, "B", 0, C_GLOSA, C_GLOSA)
   Call FGrSetRowStyle(Grid5, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
   'Call FGrFontBold(Grid5, i, 0, True)
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "783"
   Grid5.TextMatrix(i, C_GLOSA) = "Préstamos efectuados a propietarios, socios o accionistas en el ejercicio"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "978"
   Grid5.TextMatrix(i, C_GLOSA) = "Cantidades adeudadas, pagadas, abonadas en cuenta o puestas a disposición de relacionados en el exterior, cuyo IA no ha sido enterado (arts. 31 inc.  3° y 59 LIR)"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "1020"
   Grid5.TextMatrix(i, C_GLOSA) = "Total pasivos contraídos en Chile"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = ""
   Grid5.TextMatrix(i, C_GLOSA) = "Datos de Balance"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   Grid5.TextMatrix(i, C_FMT) = "B"
   Call FGrSetRowStyle(Grid5, i, "B", 0, C_GLOSA, C_GLOSA)
   Call FGrSetRowStyle(Grid5, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
   i = i + 1
   
   Grid5.TextMatrix(i, C_CODIGO) = "122"
   Grid5.TextMatrix(i, C_GLOSA) = "Total del activo"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1
   
   Grid5.TextMatrix(i, C_CODIGO) = "123"
   Grid5.TextMatrix(i, C_GLOSA) = "Total del pasivo"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "101"
   Grid5.TextMatrix(i, C_GLOSA) = "Saldo de caja (sólo dinero en efectivo y documentos al día, según arqueo)"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "102"
   Grid5.TextMatrix(i, C_GLOSA) = "Capital efectivo"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "784"
   Grid5.TextMatrix(i, C_GLOSA) = "Saldo cuenta corriente bancaria según, conciliación"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "129"
   Grid5.TextMatrix(i, C_GLOSA) = "Existencia final"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "648"
   Grid5.TextMatrix(i, C_GLOSA) = "Bienes adquiridos contrato leasing"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "647"
   Grid5.TextMatrix(i, C_GLOSA) = "Activo inmovilizado"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "1003"
   Grid5.TextMatrix(i, C_GLOSA) = "Activo gasto diferido goodwill tributario"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "1004"
   Grid5.TextMatrix(i, C_GLOSA) = "Activo intangible goodwill tributario (Ley N° 20.780)"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "843"
   Grid5.TextMatrix(i, C_GLOSA) = "Patrimonio financiero"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = ""
   Grid5.TextMatrix(i, C_GLOSA) = "Otros Antecedentes"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   Grid5.TextMatrix(i, C_FMT) = "B"
   Call FGrSetRowStyle(Grid5, i, "B", 0, C_GLOSA, C_GLOSA)
   Call FGrSetRowStyle(Grid5, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
   i = i + 1
   
   Grid5.TextMatrix(i, C_CODIGO) = "1005"
   Grid5.TextMatrix(i, C_GLOSA) = "Utilidades financieras capitalizadas "
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "975"
   Grid5.TextMatrix(i, C_GLOSA) = "Gastos adeudados o pagados por cuotas de bienes en leasing"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1

   Grid5.TextMatrix(i, C_CODIGO) = "1021"
   Grid5.TextMatrix(i, C_GLOSA) = "Monto del capital  directa o indirectamente financiado por partes relacionadas"
   Grid5.TextMatrix(i, C_SIGNO) = ""
   i = i + 1


End Sub
Private Sub Grid1_SelChange()
   Tx_CurrCell = Grid1.TextMatrix(Grid1.Row, C_GLOSA)
   
End Sub
Private Sub Grid2_SelChange()
   Tx_CurrCell = Grid2.TextMatrix(Grid2.Row, C_GLOSA)
   
End Sub
Private Sub Grid3_SelChange()
   Tx_CurrCell = Grid3.TextMatrix(Grid3.Row, C_GLOSA)
   
End Sub
Private Sub Grid4_SelChange()
   Tx_CurrCell = Grid4.TextMatrix(Grid4.Row, C_GLOSA)
   
End Sub

Private Sub Grid5_SelChange()
   Tx_CurrCell = Grid4.TextMatrix(Grid5.Row, C_GLOSA)
   
End Sub


