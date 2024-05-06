VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAuditoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Auditoria de Comprobantes"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   60
      TabIndex        =   23
      Top             =   780
      Width           =   11775
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   660
         Width           =   1635
      End
      Begin VB.TextBox Tx_FechaComp 
         Height          =   315
         Index           =   1
         Left            =   4380
         TabIndex        =   9
         Top             =   660
         Width           =   1335
      End
      Begin VB.CommandButton Bt_FechaComp 
         Height          =   315
         Index           =   1
         Left            =   5760
         Picture         =   "FrmAuditoria.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   660
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaComp 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   660
         Width           =   1275
      End
      Begin VB.CommandButton Bt_FechaComp 
         Height          =   315
         Index           =   0
         Left            =   3480
         Picture         =   "FrmAuditoria.frx":0075
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   660
         Width           =   255
      End
      Begin VB.ComboBox Cb_Usuario 
         Height          =   315
         Left            =   9900
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Cb_Oper 
         Height          =   315
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton Bt_FechaOper 
         Height          =   315
         Index           =   0
         Left            =   3480
         Picture         =   "FrmAuditoria.frx":00EA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaOper 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox Tx_IdComp 
         Height          =   315
         Left            =   9060
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton Bt_FechaOper 
         Height          =   315
         Index           =   1
         Left            =   5760
         Picture         =   "FrmAuditoria.frx":015F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Tx_FechaOper 
         Height          =   315
         Index           =   1
         Left            =   4380
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Search 
         Height          =   375
         Left            =   10500
         Picture         =   "FrmAuditoria.frx":01D4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         Height          =   195
         Index           =   8
         Left            =   8820
         TabIndex        =   33
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   3840
         TabIndex        =   32
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha comprobante desde:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   31
         Top             =   720
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Index           =   1
         Left            =   9180
         TabIndex        =   28
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operación:"
         Height          =   195
         Index           =   0
         Left            =   6480
         TabIndex        =   27
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha operación desde:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comp.:"
         Height          =   195
         Index           =   2
         Left            =   6480
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   24
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   11775
      Begin VB.CommandButton Bt_DelImport 
         Caption         =   "Eliminar Comp. Importado"
         Height          =   315
         Left            =   4740
         TabIndex        =   30
         ToolTipText     =   "Eliminar comprobante seleccionado sólo si la operación es ""Importar"""
         Top             =   180
         Width           =   2055
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
         Left            =   660
         Picture         =   "FrmAuditoria.frx":0724
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ordenar listado por columna seleccionada"
         Top             =   180
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
         Left            =   120
         Picture         =   "FrmAuditoria.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   1620
         Picture         =   "FrmAuditoria.frx":0F79
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10260
         TabIndex        =   20
         Top             =   180
         Width           =   1275
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
         Left            =   1200
         Picture         =   "FrmAuditoria.frx":1433
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   2040
         Picture         =   "FrmAuditoria.frx":18DA
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   2580
         Picture         =   "FrmAuditoria.frx":1D1F
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   25
      Cols            =   11
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Lb_Nota 
      Caption         =   "Estado Oper.: estado del comprobante después de la operación"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   34
      Top             =   8880
      Width           =   4635
   End
   Begin VB.Label Lb_Nota 
      Caption         =   "Nota: en azul se muestran los comprobantes Eliminados"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   29
      Top             =   8880
      Width           =   4335
   End
End
Attribute VB_Name = "FrmAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDCOMP = 0
Const C_FECHAOPER = 1
Const C_LNGFECHAOPER = 2
Const C_CORRCOMP = 3
Const C_TIPOCOMP = 4
Const C_TAJUSTE = 5
Const C_IDTAJUSTE = 6
Const C_IDTIPOCOMP = 7
Const C_ESTADOCOMP = 8
Const C_IDESTADOCOMP = 9
Const C_FEMISION = 10
Const C_LNGFEMISION = 11
Const C_USUARIO = 12
Const C_OPER = 13
Const C_IDOPER = 14
Const C_ESTADOOPER = 15


Const NCOLS = C_ESTADOOPER

Dim lOrdenGr(NCOLS) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Const F_INICIO = 0
Const F_FIN = 1

Dim lOper As Integer

Dim lOrientacion As Integer


Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
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

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub
Private Sub Bt_CopyExcel_Click()
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, Me.Caption & vbTab & "Fecha Inicio: " & Tx_FechaOper(0) & " Fecha Término: " & Tx_FechaOper(1))
End Sub

Private Sub Bt_DelImport_Click()
   Dim Row As Integer
   Dim idcomp As Long
   Dim CorrComp As Long
   Dim FechaComp As Long
   Dim EstadoComp As Integer
   Dim TipoComp As Integer
   Dim PcName As String
   Dim TipoAjuste As Integer
   
   Row = Grid.Row
   
   idcomp = Val(Grid.TextMatrix(Row, C_IDCOMP))
   
   If vFmt(Grid.TextMatrix(Row, C_IDESTADOCOMP)) = EC_ELIMINADO Then
      MsgBox1 "Este comprobante ya ha sido eliminado.", vbExclamation
      Exit Sub
   End If
      
   If idcomp = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDOPER)) <> O_IMPORT Then
      MsgBox1 "Esta operación es sólo para comprobantes importados.", vbExclamation
      Exit Sub
   End If
   
   PcName = IsLockedAction(DbMain, LK_COMPROBANTE, idcomp)
   If PcName <> "" Then
      MsgBox1 "Este comprobante se está editando en el equipo '" & PcName & "'. No puede ser eliminado.", vbInformation
      Exit Sub
   End If

   If vFmt(Grid.TextMatrix(Row, C_LNGFECHAOPER)) < DateAdd("m", -1, Now) Then
      If MsgBox1("Este comprobante fue importado hace más de un mes atrás." & vbNewLine & vbNewLine & "¿Está seguro que desea eliminarlo?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
         Exit Sub
      End If
      'esto se pregunta en la función DeleteComprobante
'   ElseIf MsgBox1("¿Está seguro que desea eliminar este comprobante?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
'      Exit Sub
   End If
      
   CorrComp = vFmt(Grid.TextMatrix(Grid.Row, C_CORRCOMP))
   FechaComp = vFmt(Grid.TextMatrix(Grid.Row, C_LNGFEMISION))
   EstadoComp = vFmt(Grid.TextMatrix(Grid.Row, C_IDESTADOCOMP))
   TipoComp = vFmt(Grid.TextMatrix(Grid.Row, C_IDTIPOCOMP))
   TipoAjuste = vFmt(Grid.TextMatrix(Grid.Row, C_IDTAJUSTE))
   
   If DeleteComprobante(idcomp) = True Then

      Call AddLogComprobantes(idcomp, gUsuario.IdUsuario, O_DELETE, Now, EC_ELIMINADO, CorrComp, FechaComp, TipoComp, EC_ELIMINADO, TipoAjuste)
      
      MousePointer = vbHourglass
      Call LoadAll
      MousePointer = vbDefault
      
      MsgBox1 "El comprobante ha sido eliminado.", vbInformation + vbOKOnly
   
   End If
   
   
End Sub

Private Sub Bt_DetComp_Click()

   Call ViewDetComp(Grid.Row, Grid.Col)
   
End Sub

Private Sub Bt_Orden_Click()
   
   Call OrdenaPorCol(Grid.Col)
         
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar al vista previa.", vbExclamation
      Exit Sub
   End If
   
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
   
   Call ResetPrtBas(gPrtReportes)
   
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
   
   Call ResetPrtBas(gPrtReportes)
   
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
      
   If Trim(Tx_FechaOper(0)) <> "" Then
      Encabezados(0) = "Fecha: " & vbTab & Tx_FechaOper(0) & " - " & Tx_FechaOper(1)
   End If
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDCOMP
   gPrtReportes.NTotLines = 0
   

End Sub

Private Sub Bt_Search_Click()

   If valida() = False Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
      
   Call LoadAll
   MousePointer = vbDefault
   
End Sub



Private Sub Cb_TipoAjuste_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim F1 As Long
   Dim F2 As Long
   
   lOrientacion = ORIENT_VER
         
   MesActual = GetMesActual()
   If MesActual = 0 Then
      MesActual = GetUltimoMesConComps()
   End If
   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
   Call FirstLastMonthDay(DateSerial(Year(Now), month(Now), 1), F1, F2)
'   Call SetTxDate(Tx_FechaOper(F_INICIO), F1)
'   Call SetTxDate(Tx_FechaOper(F_FIN), F2)
      
   Call FillCb
   Call SetUpGrid
   
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_CORRCOMP) = "CorrComp, TipoComp, FechaComp, LogComprobantes.Fecha"
   lOrdenGr(C_TIPOCOMP) = "TipoComp, CorrComp, FechaComp, LogComprobantes.Fecha"
   lOrdenGr(C_TAJUSTE) = "TipoAjusteComp, FechaComp, LogComprobantes.Fecha"
   lOrdenGr(C_ESTADOCOMP) = "Comprobante.Estado, EstadoComp, CorrComp, TipoComp, LogComprobantes.Fecha"
   lOrdenGr(C_FEMISION) = "Comprobante.Fecha, FechaComp, CorrComp, TipoComp, LogComprobantes.Fecha "
   lOrdenGr(C_USUARIO) = "Usuarios.Usuario, LogComprobantes.Fecha"
   lOrdenGr(C_OPER) = "LogComprobantes.IdOper, LogComprobantes.Fecha"
   lOrdenGr(C_ESTADOOPER) = "LogComprobantes.Estado, LogComprobantes.Fecha"
   lOrdenGr(C_FECHAOPER) = "LogComprobantes.Fecha, CorrComp, TipoComp "
   
   
   lOrdenSel = C_FECHAOPER
      
   Call LoadAll
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   Grid.FixedRows = 1
   Grid.FixedCols = 2
   
   Call FGrSetup(Grid)
    
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_CORRCOMP) = 1200
   Grid.ColWidth(C_TIPOCOMP) = 1000
   Grid.ColWidth(C_IDTIPOCOMP) = 0
   Grid.ColWidth(C_TAJUSTE) = 400
   Grid.ColWidth(C_IDTAJUSTE) = 0
   Grid.ColWidth(C_ESTADOCOMP) = 1300
   Grid.ColWidth(C_IDESTADOCOMP) = 0
   Grid.ColWidth(C_FEMISION) = FW_FECHA + 300
   Grid.ColWidth(C_LNGFEMISION) = 0
   Grid.ColWidth(C_USUARIO) = 1400
   Grid.ColWidth(C_OPER) = 1400
   Grid.ColWidth(C_IDOPER) = 0
   Grid.ColWidth(C_FECHAOPER) = FW_FECHA + 960
   Grid.ColWidth(C_LNGFECHAOPER) = 0
   Grid.ColWidth(C_ESTADOOPER) = 1300
      
   Grid.ColAlignment(C_IDCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_CORRCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPOCOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADOCOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADOOPER) = flexAlignLeftCenter
   Grid.ColAlignment(C_USUARIO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
   Grid.ColAlignment(C_OPER) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHAOPER) = flexAlignLeftCenter
   Grid.ColAlignment(C_TAJUSTE) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_CORRCOMP) = "N° Comp."
   Grid.TextMatrix(0, C_TIPOCOMP) = "Tipo"
   Grid.TextMatrix(0, C_TAJUSTE) = "Ajus"
   Grid.TextMatrix(0, C_ESTADOCOMP) = "Estado Comp."
   Grid.TextMatrix(0, C_USUARIO) = "Usuario"
   Grid.TextMatrix(0, C_FEMISION) = "Fecha Comp."
   Grid.TextMatrix(0, C_OPER) = "Operación"
   Grid.TextMatrix(0, C_FECHAOPER) = "Fecha Operación"
   Grid.TextMatrix(0, C_ESTADOOPER) = "Estado Oper."
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Wh As String
   Dim UsrJoin As String

   Wh = CreateWhere()
      
   If Wh = "" Then
      Wh = " WHERE "
   Else
      Wh = Wh & " AND "
   End If
   
   Wh = Wh & " LogComprobantes.IdEmpresa = " & gEmpresa.id & " AND LogComprobantes.Ano = " & gEmpresa.Ano

      
   Q1 = "SELECT Comprobante.IdComp, LogComprobantes.IdUsuario, LogComprobantes.Fecha as FechaOper, IdOper, LogComprobantes.Estado as EstadoOper, "
   Q1 = Q1 & " CorrComp, FechaComp, "
   Q1 = Q1 & " TipoComp, EstadoComp, TipoAjusteComp, Usuarios.Usuario,"
   Q1 = Q1 & " Comprobante.Correlativo as CurrCorrComp, Comprobante.Fecha as CurrFechaComp, "
   Q1 = Q1 & " Comprobante.Tipo as CurrTipoComp, Comprobante.Estado as CurrEstadoComp, Comprobante.TipoAjuste "
   Q1 = Q1 & " FROM (LogComprobantes "
   Q1 = Q1 & " LEFT JOIN Usuarios ON LogComprobantes.IdUsuario = Usuarios.IdUsuario) "
   Q1 = Q1 & " LEFT JOIN Comprobante ON LogComprobantes.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "LogComprobantes")
      
   Q1 = Q1 & Wh
   'Q1 = Q1 & " ORDER BY IdOper Desc " '& lOrdenGr(lOrdenSel)
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
   
      Grid.rows = i + 1
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If
      
      
      Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"))
      Grid.TextMatrix(i, C_OPER) = Cb_Oper.list(CbFindItem(Cb_Oper, vFld(Rs("IdOper"))))
      Grid.TextMatrix(i, C_IDOPER) = vFld(Rs("IdOper"))
      Grid.TextMatrix(i, C_FECHAOPER) = Format(vFld(Rs("FechaOper")), EDATEFMT & "      hh:mm")
      Grid.TextMatrix(i, C_LNGFECHAOPER) = vFld(Rs("FechaOper"))
      Grid.TextMatrix(i, C_TAJUSTE) = Left(gTipoAjuste(vFld(Rs("TipoAjusteComp"))), 1)
      Grid.TextMatrix(i, C_IDTAJUSTE) = vFld(Rs("TipoAjusteComp"))
      
      If vFld(Rs("EstadoOper")) = EC_ELIMINADO Then
         Grid.TextMatrix(i, C_ESTADOOPER) = "Eliminado"
         Grid.TextMatrix(i, C_TAJUSTE) = Left(gTipoAjuste(vFld(Rs("TipoAjusteComp"))), 1)
         Grid.TextMatrix(i, C_IDTAJUSTE) = vFld(Rs("TipoAjusteComp"))
      Else
         Grid.TextMatrix(i, C_ESTADOOPER) = gEstadoComp(vFld(Rs("EstadoOper")))
         Grid.TextMatrix(i, C_TAJUSTE) = Left(gTipoAjuste(vFld(Rs("TipoAjuste"))), 1)
         Grid.TextMatrix(i, C_IDTAJUSTE) = vFld(Rs("TipoAjuste"))
      End If
      
      If vFld(Rs("CurrCorrComp")) = 0 Then    'comprobante eliminado
            
         Grid.TextMatrix(i, C_CORRCOMP) = vFld(Rs("CorrComp"))
         Grid.TextMatrix(i, C_TIPOCOMP) = gTipoComp(vFld(Rs("TipoComp")))
         Grid.TextMatrix(i, C_IDTIPOCOMP) = vFld(Rs("TipoComp"))
         
         Grid.TextMatrix(i, C_ESTADOCOMP) = "Eliminado"
         Grid.TextMatrix(i, C_IDESTADOCOMP) = EC_ELIMINADO
         
         Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FechaComp")), EDATEFMT & " ")
         Grid.TextMatrix(i, C_LNGFEMISION) = vFld(Rs("FechaComp"))
         
         Call FGrSetRowStyle(Grid, i, "FC", vbBlue)
         
      Else
      
         Grid.TextMatrix(i, C_CORRCOMP) = vFld(Rs("CurrCorrComp"))
         Grid.TextMatrix(i, C_TIPOCOMP) = gTipoComp(vFld(Rs("CurrTipoComp")))
         Grid.TextMatrix(i, C_IDTIPOCOMP) = vFld(Rs("CurrTipoComp"))
         
         If vFld(Rs("CurrEstadoComp")) <= UBound(gEstadoComp) Then
            Grid.TextMatrix(i, C_ESTADOCOMP) = gEstadoComp(vFld(Rs("CurrEstadoComp")))
            Grid.TextMatrix(i, C_IDESTADOCOMP) = vFld(Rs("CurrEstadoComp"))
         Else
            Grid.TextMatrix(i, C_ESTADOCOMP) = gEstadoComp(EC_PENDIENTE)
            Grid.TextMatrix(i, C_IDESTADOCOMP) = EC_PENDIENTE
         End If
         
         Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("CurrFechaComp")), EDATEFMT & " ")
         Grid.TextMatrix(i, C_LNGFEMISION) = vFld(Rs("CurrFechaComp"))
         
      End If
         
      Rs.MoveNext

      i = i + 1
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
   Grid.TopRow = Grid.FixedRows
   
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = FrmMain.Pc_Flecha

   Grid.Col = C_CORRCOMP
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col

   Grid.Redraw = True
   
   Call EnableFrm(False)
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadAll
      
   Me.MousePointer = vbDefault
      
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - Lb_Nota(0).Height - 600
   
   Lb_Nota(0).Top = Grid.Top + Grid.Height + 100
   Lb_Nota(1).Top = Lb_Nota(0).Top
   
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_Click()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If

   Call OrdenaPorCol(Col)
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   Dim i As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   Call ViewDetComp(Row, Col)
         
End Sub
Private Sub ViewDetComp(ByVal Row As Integer, ByVal Col As Integer)
   Dim idcomp As Long
   Dim Frm As FrmComprobante

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   idcomp = Val(Grid.TextMatrix(Row, C_IDCOMP))

   If idcomp <> 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(idcomp, False)
      Set Frm = Nothing
      
   ElseIf Grid.TextMatrix(Row, C_IDCOMP) = "0" Then
      MsgBox1 "Este comprobante ha sido eliminado.", vbExclamation + vbOKOnly
      
   End If
            
End Sub
Public Sub FView()
   lOper = O_VIEW
   
   Me.Show vbModal
End Sub

Private Function CreateWhere() As String
   Dim Wh As String
   Dim F1 As Long, F2 As Long
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Idx As Integer
   Dim CodCuenta As String
   Dim NotValidRut As Boolean

   Wh = ""
   
   If Val(Tx_IdComp) <> 0 Then
      'Wh = Wh & " AND LogComprobantes.CorrComp=" & Val(Tx_IdComp)
      Wh = Wh & " AND Comprobante.Correlativo=" & Val(Tx_IdComp)
   End If
      
   If Cb_Tipo.ListIndex > 0 Then
      Wh = Wh & " AND LogComprobantes.TipoComp=" & ItemData(Cb_Tipo)
   End If
   
   If ItemData(Cb_TipoAjuste) > 0 Then
      If ItemData(Cb_TipoAjuste) = TAJUSTE_FINANCIERO Then
         Wh = Wh & " AND (LogComprobantes.TipoAjusteComp IS NULL OR LogComprobantes.TipoAjusteComp IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
      Else
         Wh = Wh & " AND (LogComprobantes.TipoAjusteComp IS NULL OR LogComprobantes.TipoAjusteComp =" & ItemData(Cb_TipoAjuste) & ")"
      End If
   End If

   F1 = GetTxDate(Tx_FechaOper(0))
   F2 = GetTxDate(Tx_FechaOper(1))
   
   If F1 <> 0 And F2 <> 0 Then
      Wh = Wh & " AND (" & SqlInt("LogComprobantes.Fecha") & " BETWEEN " & F1 & " AND " & F2 & ")"
      'Wh = Wh & " AND (" & SqlInt("Comprobante.Fecha") & " BETWEEN " & F1 & " AND " & F2 & ")"
   End If
   
   F1 = GetTxDate(Tx_FechaComp(0))
   F2 = GetTxDate(Tx_FechaComp(1))
   
'   If F1 <> 0 And F2 <> 0 Then
'      Wh = Wh & " AND (Comprobante.Fecha BETWEEN " & F1 & " AND " & F2 & ")"
'   End If
   If F1 <> 0 And F2 <> 0 Then
      'Wh = Wh & " AND (LogComprobantes.FechaComp BETWEEN " & F1 & " AND " & F2 & ")"
      Wh = Wh & " AND (Comprobante.Fecha BETWEEN " & F1 & " AND " & F2 & ")"
   End If
   
   If CbItemData(Cb_Oper) > 0 Then
      Wh = Wh & " AND LogComprobantes.IdOper = " & CbItemData(Cb_Oper)
   End If
            
   If Wh <> "" Then
      Wh = " WHERE " & Mid(Wh, 5)
   End If
   
   CreateWhere = Wh
   
End Function

Private Sub Tx_FechaComp_Change(Index As Integer)
   Dim F1 As Long
   
   If Index = 0 And GetTxDate(Tx_FechaComp(1)) < GetTxDate(Tx_FechaComp(0)) Then
      F1 = DateAdd("m", 1, GetTxDate(Tx_FechaComp(0)))
      Call SetTxDate(Tx_FechaComp(1), F1)
   End If
   
   Call EnableFrm(True)

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
Private Sub FillCb()
   Dim i As Integer, Q1 As String
      
   Call CbAddItem(Cb_Tipo, "(todos)", -1)
   For i = 1 To N_TIPOCOMP
      Call CbAddItem(Cb_Tipo, gTipoComp(i), i)
   Next i
   Cb_Tipo.ListIndex = 0
               
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)
               
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
Private Function valida() As Boolean
   Dim F1 As Long, F2 As Long
   '2953007
   Dim MesActual As Integer
   '2953007
   valida = False
      
   F1 = GetTxDate(Tx_FechaOper(0))
   F2 = GetTxDate(Tx_FechaOper(1))
   
   If F1 = 0 And F2 <> 0 Or F1 <> 0 And F2 = 0 Or F1 > F2 Then
      MsgBox1 "Rango de fechas de operación inválido.", vbExclamation + vbOKOnly
      Tx_FechaOper(0).SetFocus
      Exit Function
   End If
   
   '2953007
   If F1 <> 0 Then
    If Year(Tx_FechaOper(0)) < gEmpresa.Ano Then
         MsgBox1 "Rango de fechas de operación inválido, Solo es posible visualizar año actual .", vbExclamation + vbOKOnly
          
         MesActual = GetMesActual()
         If MesActual = 0 Then
            MesActual = GetUltimoMesConComps()
         End If
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
         Call FirstLastMonthDay(DateSerial(Year(Now), month(Now), 1), F1, F2)
         Call SetTxDate(Tx_FechaOper(F_INICIO), F1)
         Call SetTxDate(Tx_FechaOper(F_FIN), F2)
           
         Tx_FechaOper(0).SetFocus
       
       Exit Function
    End If
   End If
   
   If F2 <> 0 Then
    If Year(Tx_FechaOper(1)) > gEmpresa.Ano Then
         MsgBox1 "Rango de fechas de operación inválido, Solo es posible visualizar año actual .", vbExclamation + vbOKOnly
          
         MesActual = GetMesActual()
         If MesActual = 0 Then
            MesActual = GetUltimoMesConComps()
         End If
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
         Call FirstLastMonthDay(DateSerial(Year(Now), month(Now), 1), F1, F2)
         Call SetTxDate(Tx_FechaOper(F_INICIO), F1)
         Call SetTxDate(Tx_FechaOper(F_FIN), F2)
           
         Tx_FechaOper(1).SetFocus
       
       Exit Function
    End If
   End If
   '2953007
   
      
   F1 = GetTxDate(Tx_FechaComp(0))
   F2 = GetTxDate(Tx_FechaComp(1))
   
   If F1 = 0 And F2 <> 0 Or F1 <> 0 And F2 = 0 Or F1 > F2 Then
      MsgBox1 "Rango de fechas de comprobante inválido.", vbExclamation + vbOKOnly
      Tx_FechaOper(0).SetFocus
      Exit Function
   End If
   
   '2953007
   If F1 <> 0 Then
    If Year(Tx_FechaComp(0)) < gEmpresa.Ano Then
         MsgBox1 "Rango de fechas de comprobante inválido, Solo es posible visualizar año actual .", vbExclamation + vbOKOnly
          
         MesActual = GetMesActual()
         If MesActual = 0 Then
            MesActual = GetUltimoMesConComps()
         End If
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
         Call FirstLastMonthDay(DateSerial(Year(Now), month(Now), 1), F1, F2)
         Call SetTxDate(Tx_FechaComp(F_INICIO), F1)
         Call SetTxDate(Tx_FechaComp(F_FIN), F2)
           
         Tx_FechaComp(0).SetFocus
       
       Exit Function
    End If
   End If
      
   If F2 <> 0 Then
    If Year(Tx_FechaComp(1)) > gEmpresa.Ano Then
         MsgBox1 "Rango de fechas de comprobante inválido, Solo es posible visualizar año actual .", vbExclamation + vbOKOnly
          
         MesActual = GetMesActual()
         If MesActual = 0 Then
            MesActual = GetUltimoMesConComps()
         End If
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
         Call FirstLastMonthDay(DateSerial(Year(Now), month(Now), 1), F1, F2)
         Call SetTxDate(Tx_FechaComp(F_INICIO), F1)
         Call SetTxDate(Tx_FechaComp(F_FIN), F2)
           
         Tx_FechaComp(1).SetFocus
       
       Exit Function
    End If
   End If
      
    '2953007
      
   valida = True
End Function

Private Sub Tx_IdComp_Change()
   Call EnableFrm(True)

End Sub

Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   
End Sub
Private Sub Cb_Tipo_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Usuario_Click()
   Call EnableFrm(True)

End Sub
Private Sub Cb_Oper_Click()
   Call EnableFrm(True)

End Sub

Private Sub Tx_IdComp_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub


