VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSucursales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sucursales"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "FrmSucursales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5640
      TabIndex        =   6
      Top             =   540
      Width           =   1095
   End
   Begin VB.Frame Fr_Edit 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5640
      TabIndex        =   7
      Top             =   960
      Width           =   1095
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmSucursales.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modificar Centro de costo"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmSucursales.frx":05DF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar Centro de costo"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmSucursales.frx":0C41
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo Centro de costo"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   0
         Picture         =   "FrmSucursales.frx":11D3
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Empleado"
         Top             =   2700
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Sel 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   1155
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmSucursales.frx":1802
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3975
      Left            =   1320
      TabIndex        =   0
      Top             =   540
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "FrmSucursales.frx":1B0C
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "FrmSucursales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_DESCRIP = 1
Const C_ID = 2

Dim lSucursal As Sucursal_t
Dim lRc As Integer
Dim lOper As Integer

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
   End If
   
'   Q1 = "SELECT Count(*) as n FROM MovComprobante WHERE idCCosto=" & vFmt(Grid.TextMatrix(Row, C_ID))
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If vFld(Rs("n")) <> 0 Then
'      MsgBox1 "No puede borrar la sucursal " & Grid.TextMatrix(Row, C_DESCRIP) & ", existe un movimiento asociado.", vbExclamation
'      Call CloseRs(Rs)
'      Exit Sub
'   End If
'   Call CloseRs(Rs)
   
   If MsgBox1("¿Está seguro de eliminar la sucursal " & Grid.TextMatrix(Row, C_DESCRIP) & "?", vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
   Grid.RowHeight(Row) = 0
   Grid.rows = Grid.rows + 1
   Q1 = " WHERE IdSucursal=" & vFmt(Grid.TextMatrix(Row, C_ID))
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call DeleteSQL(DbMain, "Sucursales", Q1)
   
End Sub

Private Sub Bt_Edit_Click()
   Dim Frm As FrmSucursal
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Beep
      Exit Sub
      
   End If
   
   Set Frm = New FrmSucursal
   
   lSucursal.id = Grid.TextMatrix(Row, C_ID)
   If Frm.FEdit(lSucursal) = vbOK Then
      Call UpDateGrid(Row)
     
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmSucursal
   Dim Row As Integer
   
   Set Frm = New FrmSucursal
   If Frm.FNew(lSucursal) = vbOK Then
      Row = FGrAddRow(Grid)
      Call UpDateGrid(Row)
      
   End If
   Set Frm = Nothing
   
End Sub
Friend Function FSelect(Sucursal As Sucursal_t) As Integer
   lOper = O_VIEW
   
   Me.Show vbModal
   
   FSelect = lRc
   Sucursal = lSucursal
   
End Function
Public Sub FEdit()
   lOper = O_EDIT
   Me.Show vbModal
      
End Sub

Private Sub Bt_Print_Click()
   Dim ColWi(C_ID) As Integer
   Dim Total(C_ID) As String
   Dim i As Integer
   
   If Grid.TextMatrix(1, C_CODIGO) = "" Then
      Exit Sub
   End If
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
     
   Next i
   
   Total(0) = ""
   Call PrtFlexGrid(Grid, "", "LISTADO DE SUCURSALES", "", "", ColWi, Total, False, , , , , , , , , , , True)
   
End Sub

Private Sub Bt_Sel_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
   End If
   
   lSucursal.id = Grid.TextMatrix(Row, C_ID)
   lSucursal.Codigo = Grid.TextMatrix(Row, C_CODIGO)
   lSucursal.Descrip = Grid.TextMatrix(Row, C_DESCRIP)
   lRc = vbOK
   
   Unload Me
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   
   Fr_Edit.visible = (lOper = O_EDIT)
   Fr_Sel.visible = (lOper = O_VIEW)
      
   Call SetUpGrid
   Call LoadAll
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv

End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim n As Integer
   
   
   Q1 = "SELECT count(*) FROM Sucursales "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   n = 0
   If Not Rs.EOF Then
      n = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   If n = 0 Then  'no hay sucursales definidas
      Q1 = "INSERT INTO Sucursales (Codigo, Descripcion, IdEmpresa, Vigente) VALUES('CASMAT', 'Casa Matriz', " & gEmpresa.id & ", -1 )"
      Call ExecSQL(DbMain, Q1)
   End If
      
   Q1 = "SELECT Codigo, IdSucursal, Descripcion FROM Sucursales "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   Grid.rows = i
   Do While Rs.EOF = False
      Grid.rows = i + 1
      
      Grid.TextMatrix(i, C_CODIGO) = vFld(Rs("Codigo"), True)
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descripcion"), True)
      
      Grid.TextMatrix(i, C_ID) = vFld(Rs("IdSucursal"))
      
      i = i + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.ColWidth(C_CODIGO) = 1500
   Grid.ColWidth(C_DESCRIP) = 2000
   Grid.ColWidth(C_ID) = 0
      
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
      Grid.ColAlignment(i) = flexAlignLeftCenter
      
   Next i
   
   Grid.TextMatrix(0, C_CODIGO) = "Código"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   
End Sub
Private Sub UpDateGrid(Row As Integer)
   
    Grid.TextMatrix(Row, C_CODIGO) = lSucursal.Codigo
    Grid.TextMatrix(Row, C_DESCRIP) = lSucursal.Descrip
    Grid.TextMatrix(Row, C_ID) = lSucursal.id
    
End Sub

Private Sub Grid_DblClick()
   If Grid.TextMatrix(Grid.Row, C_CODIGO) = "" And Val(Grid.TextMatrix(Grid.Row, C_ID)) = 0 Then
      Call Bt_New_Click
   Else
      Call Bt_Edit_Click
   End If
   
End Sub
Private Function SetupPriv()
   
   If lOper = O_EDIT Then
   
      If Not ChkPriv(PRV_ADM_DEF) Then
         Call EnableForm(Me, False)
      End If
   
   End If
   
End Function

