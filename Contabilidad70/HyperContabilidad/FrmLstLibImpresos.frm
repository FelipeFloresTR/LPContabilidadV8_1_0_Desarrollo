VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstLibImpresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Libros Oficiales Impresos"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "FrmLstLibImpresos.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   8599
      _Version        =   393216
      Rows            =   20
      Cols            =   10
      FixedRows       =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   9675
      Begin VB.CommandButton Bt_AnulaLibImp 
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
         Picture         =   "FrmLstLibImpresos.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anular última impresión de libro oficial"
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
         Left            =   1200
         Picture         =   "FrmLstLibImpresos.frx":0487
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8400
         TabIndex        =   4
         Top             =   180
         Width           =   1095
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
         Left            =   780
         Picture         =   "FrmLstLibImpresos.frx":0941
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   1620
         Picture         =   "FrmLstLibImpresos.frx":0DE8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmLstLibImpresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDLOG = 0
Const C_IDLIBOF = 1
Const C_LIBRO = 2
Const C_MES = 3
Const C_FDESDE = 4
Const C_FHASTA = 5
Const C_FECHA = 6
Const C_USUARIO = 7
Const C_IDESTADO = 8
Const C_ESTADO = 9

Const NCOLS = C_ESTADO

Private Sub Bt_AnulaLibImp_Click()
   Dim i As Integer
   Dim MaxIdLog As Long
   Dim RowMax As Integer
   Dim Q1 As String
   
   MaxIdLog = 0
   RowMax = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_IDLOG)) = 0 Then
         Exit For
      
      ElseIf Val(Grid.TextMatrix(i, C_IDLOG)) > MaxIdLog And Val(Grid.TextMatrix(i, C_IDESTADO)) <> EL_ANULADO Then
         MaxIdLog = Val(Grid.TextMatrix(i, C_IDLOG))
         RowMax = i
         
      End If
      
   Next i
      
   If MaxIdLog > 0 And RowMax > 0 Then
   
      If MsgBox1("Se anulará el registro de impresión oficial de: " & vbNewLine & vbNewLine & Grid.TextMatrix(RowMax, C_LIBRO) & " entre el " & Grid.TextMatrix(RowMax, C_FDESDE) & " y el " & Grid.TextMatrix(RowMax, C_FHASTA) & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      Q1 = "UPDATE LogImpreso SET Estado = " & EL_ANULADO & " WHERE IdLog = " & MaxIdLog
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      Grid.TextMatrix(RowMax, C_ESTADO) = gEstadoLibImp(EL_ANULADO)
      Grid.TextMatrix(RowMax, C_IDESTADO) = EL_ANULADO
      
   End If


End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   Call LoadAll

End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.ColWidth(C_IDLOG) = 0
   Grid.ColWidth(C_IDLIBOF) = 0
   Grid.ColWidth(C_LIBRO) = 3000
   Grid.ColWidth(C_MES) = 500
   Grid.ColWidth(C_FDESDE) = 1200
   Grid.ColWidth(C_FHASTA) = 1200
   Grid.ColWidth(C_FECHA) = 1200
   Grid.ColWidth(C_USUARIO) = 1300
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_ESTADO) = 900
   
   For i = 0 To Grid.Cols - 1
      Grid.ColAlignment(i) = flexAlignLeftCenter
      Grid.FixedAlignment(i) = flexAlignCenterCenter
   Next i
   
   Grid.TextMatrix(1, C_LIBRO) = "Libro"
   Grid.TextMatrix(1, C_MES) = "Mes"
   Grid.TextMatrix(0, C_FDESDE) = "Fecha"
   Grid.TextMatrix(1, C_FDESDE) = "Inicio"
   Grid.TextMatrix(0, C_FHASTA) = "Fecha"
   Grid.TextMatrix(1, C_FHASTA) = "Fin"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(1, C_FECHA) = "Impresión"
   Grid.TextMatrix(1, C_USUARIO) = "Usuario"
   Grid.TextMatrix(1, C_ESTADO) = "Estado"

End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim IdInf As Long
   Dim Mes As Integer
   
   Q1 = "SELECT IdLog, IdInforme, Mes, FDesde, FHasta, Fecha, Usuarios.Usuario, Estado "
   Q1 = Q1 & " FROM LogImpreso LEFT JOIN Usuarios ON LogImpreso.IdUsuario = Usuarios.IdUsuario"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY IdInforme, Mes, FDesde, Fecha  DESC "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   Row = Grid.FixedRows
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      
      If IdInf <> vFld(Rs("IdInforme")) Then
      
         IdInf = vFld(Rs("IdInforme"))
         Grid.TextMatrix(Row, C_IDLIBOF) = vFld(Rs("IdInforme"))
         Grid.TextMatrix(Row, C_LIBRO) = gLibroOficial(vFld(Rs("IdInforme")))
         
      End If
      
      Grid.TextMatrix(Row, C_IDLOG) = vFld(Rs("IdLog"))
      Mes = vFld(Rs("Mes"))
      If vFld(Rs("Mes")) > 0 Then
         Grid.TextMatrix(Row, C_MES) = Left(gNomMes(vFld(Rs("Mes"))), 3)
      End If
      If vFld(Rs("FDesde")) > 0 Then
         Grid.TextMatrix(Row, C_FDESDE) = Format(vFld(Rs("FDesde")), DATEFMT)
      End If
      If vFld(Rs("FHasta")) > 0 Then
         Grid.TextMatrix(Row, C_FHASTA) = Format(vFld(Rs("FHasta")), DATEFMT)
      End If
      Grid.TextMatrix(Row, C_FECHA) = Format(vFld(Rs("Fecha")), DATEFMT)
      Grid.TextMatrix(Row, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.TextMatrix(Row, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(Row, C_ESTADO) = gEstadoLibImp(vFld(Rs("Estado")))

      Row = Row + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 500
   'Grid.Width = Me.Width - 230
   
   Call FGrVRows(Grid)

End Sub
Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
End Sub

Private Sub Bt_Preview_Click()
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
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(0) As String
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Listado de Libros Oficiales Impresos"
   gPrtReportes.Titulos = Titulos
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDLOG
   gPrtReportes.NTotLines = 0
   
   
End Sub

