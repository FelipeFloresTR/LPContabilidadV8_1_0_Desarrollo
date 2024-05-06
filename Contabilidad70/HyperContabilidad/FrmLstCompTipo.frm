VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmLstCompTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Comprobantes Tipo"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "FrmLstCompTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bt_Recuperar 
      Caption         =   "Recuperar Comp. Tipo"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Recupera  los Comprobantes Tipo que el sistema ofrece por omisión"
      Top             =   3600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Cm_ComDlg 
      Left            =   7740
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Impor&tar"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":061D
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir comprobantes tipo"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7980
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Bt_New 
      Caption         =   "&Agregar"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":0C90
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo comprobante tipo"
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Edit 
      Caption         =   "Edi&tar"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":1222
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Modificar comprobante ipo"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":17F5
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar comprobante tipo"
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   870
      Left            =   7980
      Picture         =   "FrmLstCompTipo.frx":1E57
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir listado de comprobantes tipo"
      Top             =   5400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmLstCompTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NOMBRE = 0
Const C_TIPO = 1
Const C_DESCRIP = 2
Const C_TDEBE = 3
Const C_THABER = 4
Const C_IDTIPO = 5
Const C_IDCOMP = 6
Dim lCompTipo As CompTipo_t

Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.ColWidth(C_NOMBRE) = 2300
   Grid.ColWidth(C_TIPO) = 1200
   Grid.ColWidth(C_TDEBE) = 0
   Grid.ColWidth(C_THABER) = 0
   Grid.ColWidth(C_DESCRIP) = 3400
   Grid.ColWidth(C_IDTIPO) = 0
   Grid.ColWidth(C_IDCOMP) = 0
   
   For Col = 0 To Grid.Cols - 1
      Grid.FixedAlignment(Col) = flexAlignCenterCenter
   Next Col
   
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPO) = flexAlignLeftCenter
   Grid.ColAlignment(C_TDEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_THABER) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_TIPO) = "Tipo"
   Grid.TextMatrix(0, C_DESCRIP) = "Glosa"
   Grid.TextMatrix(0, C_TDEBE) = "Total Debe"
   Grid.TextMatrix(0, C_THABER) = "Total Haber"
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Q1 As String
   Dim Row As Integer
      
   Row = Grid.Row
   
   If Trim(Grid.TextMatrix(Row, C_NOMBRE)) = "" Then
      Exit Sub
   End If
   If MsgBox1("¿Está seguro de eliminar este comprobante tipo?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
      Exit Sub
   End If
   
'   Q1 = "DELETE * FROM CT_Comprobante WHERE idComp=" & Grid.TextMatrix(Row, C_IDCOMP)
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE idComp=" & Grid.TextMatrix(Row, C_IDCOMP)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call DeleteSQL(DbMain, "CT_Comprobante", Q1)
   
'   Q1 = "DELETE * FROM CT_MovComprobante WHERE idComp=" & Grid.TextMatrix(Row, C_IDCOMP)
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE idComp=" & Grid.TextMatrix(Row, C_IDCOMP)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call DeleteSQL(DbMain, "CT_MovComprobante", Q1)
   
   Grid.RowHeight(Row) = 0
   Grid.rows = Grid.rows + 1
   
End Sub

Private Sub Bt_Edit_Click()
   Dim Row As Integer
   Dim Frm As FrmComprobante
   
   Row = Grid.Row
   If Trim(Grid.TextMatrix(Row, C_NOMBRE)) = "" Then
      Exit Sub
   End If
    
   Set Frm = New FrmComprobante
   If Frm.FEdit(Grid.TextMatrix(Row, C_IDCOMP), True) = vbOK Then
      Call LoadAll
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Importar_Click()
   Dim i As Integer
        
   Cm_ComDlg.CancelError = True
   Cm_ComDlg.Filename = ""
   Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   MousePointer = vbHourglass
   DoEvents
   
   If ImportarComprobante(Cm_ComDlg.Filename, gNiveles.nNiveles, gNiveles.Largo()) > 0 Then
      Call LoadAll
   End If
   
   MousePointer = vbDefault


End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmComprobante
    
   Set Frm = New FrmComprobante
   If Frm.FNew(True) = vbOK Then
      Call LoadAll
   End If
      
   Set Frm = Nothing
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation

   Call SetUpPrtGrid
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub bt_Recuperar_Click()
   Dim Q1 As String
   Dim Rc As Long
   Dim Rs As Recordset
   
#If DATACON = 1 Then
   Dim DbVacia As String
   Dim FldName As String
   
   DbVacia = gDbPath & "\" & BD_VACIA
   Q1 = "¡ADVERTENCIA!" & vbNewLine & "Esta opción recupera todos los comprobantes tipo que el sistema ofrece por omisión, perderá cualquier cambio que usted haya realizado." & vbNewLine & "¿ Desea continuar ?"
   If MsgBox1(Q1, vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
   'ELIMINO LOS COMPROBANTES TIPOS
'   Call ExecSQL(DbMain, "DELETE * FROM CT_Comprobante")
'   Call ExecSQL(DbMain, "DELETE * FROM CT_MovComprobante")
   Call DeleteSQL(DbMain, "CT_Comprobante", " WHERE IdEmpresa = " & gEmpresa.id)
   Call DeleteSQL(DbMain, "CT_MovComprobante", " WHERE IdEmpresa = " & gEmpresa.id)
   
   '*****AHORA HAGO EL LINK CON LOS COMPROBANTES TIPO DE EMPRESA VACIA (EmpresaVacia.mdb no tiene password)
   Call LinkMdbTable(DbMain, DbVacia, "CT_Comprobante", "CT_ComprobanteCopy", True)
   Call LinkMdbTable(DbMain, DbVacia, "CT_MovComprobante", "CT_MovComprobanteCopy", True)
   
   '****INSERT
   FldName = " Correlativo, Nombre, Descrip, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber "
   Q1 = "INSERT INTO CT_Comprobante "
   Q1 = Q1 & " ( " & FldName & ", IdCompOld, IdEmpresa ) "
   Q1 = Q1 & " SELECT " & FldName & ", IdComp, " & gEmpresa.id & " As IdEmpresa FROM CT_ComprobanteCopy "
   Rc = ExecSQL(DbMain, Q1)
   FldName = "  IdComp, Orden, IdCuenta, CodCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, Conciliado "
   Q1 = "INSERT INTO CT_MovComprobante "
   Q1 = Q1 & " ( " & FldName & ", IdEmpresa ) "
   Q1 = Q1 & " SELECT " & FldName & ", " & gEmpresa.id & " As IdEmpresa FROM CT_MovComprobanteCopy "
   Call ExecSQL(DbMain, Q1)
   
   'reenlazamos los movimientos de comprobantes
   Q1 = "UPDATE CT_MovComprobante INNER JOIN CT_Comprobante"
   Q1 = Q1 & " ON CT_MovComprobante.IdComp = CT_Comprobante.IdCompOld AND CT_MovComprobante.IdEmpresa = CT_Comprobante.IdEmpresa "
   Q1 = Q1 & " SET CT_MovComprobante.IdComp = CT_Comprobante.IdComp "
   Q1 = Q1 & " WHERE CT_MovComprobante.IdEmpresa = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)

   
   'actualizamos las cuentas con el nuevo plan, si es que hay
   Q1 = "UPDATE CT_MovComprobante INNER JOIN Cuentas "
   Q1 = Q1 & " ON CT_MovComprobante.CodCuenta = Cuentas.Codigo AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
   Q1 = Q1 & " SET CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE CT_MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   '*****DESLINKEO
   Call ExecSQL(DbMain, "Drop Table " & "CT_ComprobanteCopy")
   Call ExecSQL(DbMain, "Drop Table " & "CT_MovComprobanteCopy")
   
#Else

   Call ResetCompTipoEmpJuntas(gEmpresa.id, True)
      
#End If

   'REORDENO LOS COMPROBANTES FUNCION FRANCA
   Call UpdateComprobantesTipo
   Call LoadAll
   
   MsgBox1 "La recuperación de comprobantes tipo por omisión ha sido realizada exitosamente.", vbExclamation

End Sub

Private Sub Form_Load()
   
   Call SetUpGrid
   Call LoadAll
   Call SetupPriv
   
   Bt_Importar.visible = W.InDesign
      
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim n As Integer
   
'   Q1 = "SELECT Count(*) as N FROM CT_Comprobante WHERE IdEmpresa = " & gEmpresa.id
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Not Rs.EOF Then
'
'      If vFld(Rs("N")) = 0 Then    'no tiene comp tipo para esta emnpresa, agregamos los base
'
'         Call ResetCompTipoEmpJuntas(gempresa.id)
'
'      End If
'   End If
'
'   Call CloseRs(Rs)

   
   Grid.Redraw = False
   
   Q1 = "SELECT Nombre, IdComp, Tipo, TotalDebe, TotalHaber, Descrip FROM CT_Comprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY Nombre"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 1
   Grid.rows = 1
   Do While Rs.EOF = False
      Grid.rows = Row + 1
      
      Grid.TextMatrix(Row, C_NOMBRE) = vFld(Rs("Nombre"), True)
      Grid.TextMatrix(Row, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(Row, C_TIPO) = gTipoComp(vFld(Rs("Tipo")))
      Grid.TextMatrix(Row, C_TDEBE) = Format(vFld(Rs("TotalDebe")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_THABER) = Format(vFld(Rs("TotalHaber")), BL_NUMFMT)
      Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(Row, C_IDTIPO) = vFld(Rs("Tipo"))
      
      Row = Row + 1
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   Grid.Redraw = True
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(C_IDCOMP) As Integer
   Dim Total(C_IDCOMP) As String
   Dim Titulos(0) As String
   Dim Encabezados(3) As String
   Dim FontTit(0) As FontDef_t
   Dim FontNom(0) As FontDef_t
   
   Set gPrtReportes.Grid = Grid
   
   Printer.Orientation = ORIENT_VER
   
   Titulos(0) = "LISTADO DE COMPROBANTES TIPO"
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
   gPrtReportes.NTotLines = 0
   
End Sub

Private Sub Grid_DblClick()
   Call PostClick(Bt_Edit)
   
End Sub
Private Sub SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Bt_New.Enabled = False
      Bt_Edit.Caption = "Ver"
      Bt_Del.Enabled = False
      bt_Recuperar.Enabled = False
      Bt_Importar.Enabled = False
   End If
End Sub

