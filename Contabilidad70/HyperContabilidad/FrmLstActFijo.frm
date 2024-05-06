VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstActFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Activos Fijos"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   Icon            =   "FrmLstActFijo.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Del_All 
      Caption         =   "&Eliminar Todo"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox Pc_HdCheck 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   7440
      Picture         =   "FrmLstActFijo.frx":0316
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "Copiar a Excel"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":067B
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Copiar Excel"
      Top             =   6600
      Width           =   1155
   End
   Begin VB.CommandButton Bt_MarcarExport 
      Caption         =   "Marcar para Exportar a año siguiente"
      Height          =   315
      Left            =   7560
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Marcar Activo Fijo seleccionado para ser exportado al año siguiente"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":0C30
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":10FE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   3600
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   11160
      TabIndex        =   15
      Top             =   240
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   11
      BackColorBkg    =   16777215
      SelectionMode   =   1
   End
   Begin VB.CommandButton Bt_Edit 
      Caption         =   "Edi&tar"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":1408
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Modificar cuenta seleccionada"
      Top             =   2640
      Width           =   1155
   End
   Begin VB.CommandButton Bt_ViewDoc 
      Caption         =   "Ver D&ocumento"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":19DB
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Nueva cuenta"
      Top             =   2640
      Width           =   1155
   End
   Begin VB.CommandButton Bt_New 
      Caption         =   "&Agregar"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":1D89
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Nueva cuenta"
      Top             =   1680
      Width           =   1155
   End
   Begin VB.CommandButton Bt_ViewDet 
      Caption         =   "Ver D&etalle"
      Height          =   855
      Left            =   11160
      Picture         =   "FrmLstActFijo.frx":231B
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Nueva cuenta"
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Frame Fr_List 
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10635
      Begin VB.PictureBox Pc_Check 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   7680
         Picture         =   "FrmLstActFijo.frx":2717
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Ch_SortDesc 
         Caption         =   "Ordenar por Descripción"
         Height          =   195
         Left            =   6420
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Index           =   1
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   215
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Index           =   0
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   215
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Search 
         Caption         =   "&Listar"
         Height          =   855
         Left            =   8760
         Picture         =   "FrmLstActFijo.frx":278E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_Descrip 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         ToolTipText     =   "Ingrese parte de la descripción para encontrar los activos fijos relacionados"
         Top             =   780
         Width           =   4875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-->"
         Height          =   195
         Index           =   10
         Left            =   2700
         TabIndex        =   19
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha compra:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   885
      End
   End
   Begin VB.Frame Fr_Doc 
      Caption         =   "Documento"
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   9735
      Begin VB.TextBox Tx_Doc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   2115
      End
      Begin VB.TextBox Tx_NumDoc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Tx_FechaDoc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Tx_Rut 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox Tx_Nombre 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   780
         Width           =   4815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   39
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Index           =   2
         Left            =   7500
         TabIndex        =   38
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Entidad:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   37
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Doc:"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   27
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   1
         Left            =   7500
         TabIndex        =   26
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Frame Fr_Comp 
      Caption         =   "Movimiento Comprobante"
      Height          =   1335
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   9735
      Begin VB.TextBox Tx_NumMov 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Tx_GlosaMov 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   780
         Width           =   8535
      End
      Begin VB.TextBox Tx_FechaComp 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Tx_NumComp 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Línea:"
         Height          =   195
         Index           =   6
         Left            =   2460
         TabIndex        =   36
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   34
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   4
         Left            =   7680
         TabIndex        =   32
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Comp.:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   420
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmLstActFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDACTFIJO = 0
Const C_IDDOC = 1
Const C_IDCOMP = 2
Const C_TIPOMOVAF = 3
Const C_FECHA = 4
Const C_CANTIDAD = 5
Const C_DESCRIP = 6
Const C_NETO = 7
Const C_DOCCOMP = 8
Const C_IDCUENTA = 9
Const C_IMPFILE = 10
Const C_FECHAVENTA = 11
Const C_DEANOANT = 12
'2861733 tema 1
Const C_CHECK = 13
Const NCOLS = C_CHECK
'Const NCOLS = C_DEANOANT
'2861733 tema 1

Const F_INICIO = 0
Const F_FIN = 1

Const O_LIST = -1  'oper list

Dim lOper As Integer
Dim lIdDoc As Long
Dim lidComp As Long
Dim lIdMov As Long
Dim lFecha As Long

Dim lIdActFijo As Long
Dim lTipoLib As Integer

Dim lMsgFiltro As Boolean

Dim vSel As Long

'2861733
Dim lIdArea As Long
Dim lIdCentro As Long
'2861733

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
End Sub

'2861733 tema 1
Private Sub Bt_Del_All_Click()
Dim LstComp As String
 Dim i As Integer

If MsgBox1("¿Está seguro que desea borrar los Activos Fijos Seleccionados?" & vbCrLf, vbYesNo + vbQuestion) = vbNo Then
      Exit Sub
End If

  For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDACTFIJO)) <= 0 Then
         Exit For
      End If

      Grid.Row = i
      Grid.Col = C_CHECK

      If Grid.CellPicture <> 0 Then

         LstComp = LstComp & ", " & Grid.TextMatrix(i, C_IDACTFIJO)

      End If
   Next i

   LstComp = Mid(LstComp, 2)


   Call DeleteSQL(DbMain, "MovActivoFijo", " WHERE IdActFijo in (" & LstComp & ")")
   Call DeleteSQL(DbMain, "ActFijoFicha", " WHERE IdActFijo in (" & LstComp & ")")
   Call DeleteSQL(DbMain, "ActFijoCompsFicha", " WHERE IdActFijo in (" & LstComp & ")")

   Call LoadGrid
End Sub
'2861733 tema 1

Private Sub Bt_Del_Click()
   Dim IdActFijo As Long
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdActFijo = Val(Grid.TextMatrix(Row, C_IDACTFIJO))
   
   If IdActFijo = 0 Then
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea borrar este Activo Fijo?" & vbCrLf & vbCrLf & "Activo fijo: " & Grid.TextMatrix(Row, C_DESCRIP), vbYesNo + vbQuestion) = vbNo Then
      Exit Sub
   End If
   
   Call DeleteSQL(DbMain, "MovActivoFijo", " WHERE IdActFijo = " & IdActFijo)
   Call DeleteSQL(DbMain, "ActFijoFicha", " WHERE IdActFijo = " & IdActFijo)
   Call DeleteSQL(DbMain, "ActFijoCompsFicha", " WHERE IdActFijo = " & IdActFijo)
   
   Call LoadGrid
   
   
End Sub

Private Sub Bt_DetDocComp_Click()
   Dim Row As Integer
   Dim FrmDoc As FrmDocLib
   Dim FrmComp As FrmComprobante
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      Set FrmDoc = New FrmDocLib
      Call FrmDoc.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
      Set FrmDoc = Nothing
   
   ElseIf Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
      Set FrmComp = New FrmComprobante
      Call FrmComp.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
      Set FrmComp = Nothing
   
   End If

End Sub

Private Sub Bt_Edit_Click()
   Dim Frm As FrmActivoFijo
   Dim IdActFijo As Long
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdActFijo = Val(Grid.TextMatrix(Row, C_IDACTFIJO))
   
   If IdActFijo = 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmActivoFijo
   If Frm.FEdit(IdActFijo) = vbOK Then
      Call LoadGrid
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_MarcarExport_Click()
   Dim Frm As FrmMarkActFijo
   Dim IdActivo As Long
   
   If Grid.Row < Grid.FixedRows Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   IdActivo = Val(Grid.TextMatrix(Grid.Row, C_IDACTFIJO))
   
   If IdActivo <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Set Frm = New FrmMarkActFijo
   Call Frm.FEdit(IdActivo)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmActivoFijo
   Dim Rc As Integer
   Dim IdActFijo As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim IdCuenta As Long
   
   'obtenemos la cuenta contable del activo fijo anterior en la lista
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDACTFIJO)) = 0 Then
         Exit For
      End If
      
      If Val(Grid.TextMatrix(i, C_IDCUENTA)) > 0 Then
         IdCuenta = Val(Grid.TextMatrix(i, C_IDCUENTA))
      End If
   Next i
   
   Set Frm = New FrmActivoFijo
   If lIdDoc <> 0 Then
   
      If lTipoLib = LIB_COMPRAS Then
      '2861733
         'Rc = Frm.FNewFromDoc(lIdDoc, lFecha, 0, 0, "", lTipoLib, IdCuenta, 0)
         Rc = Frm.FNewFromDocActFijo(lIdDoc, lFecha, 0, 0, "", lTipoLib, IdCuenta, 0, lIdArea, lIdCentro)
       '2861733
      End If
   Else
      Rc = Frm.FNewFromComp(lidComp, lIdMov, lFecha, 0, 0, "", IdCuenta)
        
   End If
   
   If Rc = vbOK Then
      Call LoadGrid
      
      If Not lMsgFiltro Then
         MsgBox "Si no se muestra el Activo Fijo recién agregado, ajuste el Filtro de Fecha de Compra y presione el botón Listar.", vbInformation + vbOKOnly
         lMsgFiltro = True
      End If
      
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_OK_Click()
   Unload Me
End Sub

Private Sub Bt_Search_Click()
   
   Call LoadGrid
   
End Sub

Private Sub Bt_Sel_Click()
   Dim IdActFijo As Long
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdActFijo = Val(Grid.TextMatrix(Row, C_IDACTFIJO))
   
   If IdActFijo = 0 Then
      Exit Sub
   End If
   
   lIdActFijo = Grid.TextMatrix(Grid.Row, C_IDACTFIJO)
   Unload Me
   
End Sub

Private Sub Bt_ViewDet_Click()
   Dim Frm As FrmActivoFijo
   Dim IdActFijo As Long
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdActFijo = Val(Grid.TextMatrix(Row, C_IDACTFIJO))
   
   If IdActFijo = 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmActivoFijo
   Call Frm.FView(IdActFijo)
   Set Frm = Nothing


End Sub

Private Sub Bt_ViewDoc_Click()
   Dim id As Long
   Dim Row As Integer
   Dim Frm As Form
   
   Row = Grid.Row
   
   id = Val(Grid.TextMatrix(Row, C_IDDOC))
   
   If id > 0 Then
      Set Frm = New FrmDocLib
      Call Frm.FView(id)
      Set Frm = Nothing
   Else
      id = Val(Grid.TextMatrix(Row, C_IDCOMP))
      If id > 0 Then
         Set Frm = New FrmComprobante
         Call Frm.FView(id, False)
         Set Frm = Nothing
      End If
   End If
   
End Sub

Private Sub Form_Load()
   
   Call SetUpGrid
   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))
   
   '2861733 tema 1
  vSel = 0
   '2861733 tema 1
   
   If lOper = O_LIST Or (lOper = O_EDIT And lidComp = 0 And lIdDoc = 0) Then
      Fr_Doc.visible = False
      Fr_Comp.visible = False
   Else
      Fr_List.visible = False
      Call LoadDatos     'oculta frames si corresponde
   End If
   
   If lOper = O_EDIT Then
   
      Bt_ViewDoc.Top = Grid.Top + Grid.Height - Bt_ViewDoc.Height - 600
      
      Bt_ViewDet.visible = False
      Bt_Sel.visible = False
      
      Bt_ViewDet.Enabled = False
      Bt_Sel.Enabled = False
      
      If lTipoLib = LIB_VENTAS Then
         Bt_New.Enabled = False
      End If
      
      Bt_MarcarExport.visible = True
      
      
   ElseIf lOper = O_SELECT Then
   
      Bt_New.visible = False
      Bt_Edit.visible = False
      Bt_Del.visible = False
      
      Bt_New.Enabled = False
      Bt_Edit.Enabled = False
      Bt_Del.Enabled = False
      
   Else  'O_VIEW o O_LIST
   
      Bt_New.visible = False
      Bt_Edit.visible = False
      Bt_Del.visible = False
      Bt_Sel.visible = False
      
      Bt_New.Enabled = False
      Bt_Edit.Enabled = False
      Bt_Del.Enabled = False
      Bt_Sel.Enabled = False
      
   End If
         

   Call SetTxDate(Tx_Fecha(F_INICIO), DateSerial(gEmpresa.Ano, 1, 1))
   Call SetTxDate(Tx_Fecha(F_FIN), DateSerial(gEmpresa.Ano, 12, 31))
   
   Call LoadGrid
   
   Call SetupPriv
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDACTFIJO) = 0
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_TIPOMOVAF) = 0
   Grid.ColWidth(C_FECHA) = 800
   Grid.ColWidth(C_CANTIDAD) = 900
   Grid.ColWidth(C_DESCRIP) = 4000
   Grid.ColWidth(C_NETO) = 1200
   Grid.ColWidth(C_DOCCOMP) = 1470
   Grid.ColWidth(C_IMPFILE) = 600
   Grid.ColWidth(C_FECHAVENTA) = 1000
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_DEANOANT) = 0   '1200
   
   '2861733 tema 1
   Grid.ColWidth(C_CHECK) = 300
   '2861733 tema 1
   
   Grid.ColAlignment(C_TIPOMOVAF) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CANTIDAD) = flexAlignRightCenter
   Grid.ColAlignment(C_DESCRIP) = flexAlignLeftCenter
   Grid.ColAlignment(C_NETO) = flexAlignRightCenter
   Grid.ColAlignment(C_DOCCOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHAVENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_DEANOANT) = flexAlignCenterCenter
   Grid.ColAlignment(C_IMPFILE) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_TIPOMOVAF) = "Tipo"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_CANTIDAD) = "Cantidad"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(0, C_NETO) = "Neto"
   Grid.TextMatrix(0, C_DOCCOMP) = "Doc./Comp."
   Grid.TextMatrix(0, C_IMPFILE) = "Imp.Txt"
   Grid.TextMatrix(0, C_FECHAVENTA) = "Venta o Baja"
   'Grid.TextMatrix(0, C_DEANOANT) = "De Año Anterior"
   
   '2861733 tema 1
   Grid.Row = 0
   Grid.Col = C_CHECK
   'Set Grid.CellPicture = Pc_Prt
   Set Grid.CellPicture = Pc_HdCheck
   Grid.CellPictureAlignment = flexAlignCenterCenter
   '2861733 tema 1
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
   Next i
   
   Call FGrVRows(Grid)

End Sub
'edita la lista de Movs Activo Fijo asociados a un doc
Public Sub FEditFromDoc(ByVal IdDoc As Long, ByVal TipoLib As Integer, ByVal Fecha As Long)
   
   lOper = O_EDIT
   
   lTipoLib = TipoLib
   lIdDoc = 0
   If TipoLib = LIB_COMPRAS Then
      lIdDoc = IdDoc
   End If
   lidComp = 0
   lIdMov = 0
      
   lFecha = Fecha
      
   Me.Show vbModal
   
End Sub

'2861733
'edita la lista de Movs Activo Fijo asociados a un doc
Public Sub FEditFromDocActiFijo(ByVal IdDoc As Long, ByVal TipoLib As Integer, ByVal Fecha As Long, ByVal AreaN As Long, ByVal CentroG As Long)

   lOper = O_EDIT

   lTipoLib = TipoLib
   lIdDoc = 0
   If TipoLib = LIB_COMPRAS Then
      lIdDoc = IdDoc
   End If
   lidComp = 0
   lIdMov = 0

   lIdArea = AreaN
   lIdCentro = CentroG

   lFecha = Fecha

   Me.Show vbModal
   
End Sub
'2861733

'edita la lista de Movs Activo Fijo asociados a un comp
Public Sub FEditFromComp(ByVal idcomp As Long, ByVal idMov As Long, ByVal Fecha As Long)
   
   lOper = O_EDIT
   
   lIdDoc = 0
   lidComp = idcomp
   lIdMov = idMov
      
   lFecha = Fecha
      
   Me.Show vbModal
End Sub
'edita la lista de Activos Fijos
Public Sub FEdit()
   
   lOper = O_EDIT
   
   lIdDoc = 0
   lidComp = 0
   lIdMov = 0
   
   lFecha = 0
      
   Me.Show vbModal
End Sub
'ver la lista de Movs Activo Fijo asociados a un doc
Public Sub FViewFromDoc(ByVal IdDoc As Long, ByVal TipoLib As Integer)
   
   lOper = O_VIEW
   
   lTipoLib = TipoLib
   lIdDoc = IdDoc
   lidComp = 0
   lIdMov = 0
         
   Me.Show vbModal
End Sub
'retorna activo fijo seleccionado
Public Function FSelect() As Long
   lOper = O_SELECT
            
   Me.Show vbModal
   
   FSelect = lIdActFijo

End Function
'ver la lista de Movs Activo Fijo asociados a un comp
Public Sub FViewFromComp(ByVal idcomp As Long, ByVal idMov As Long)
   
   lOper = O_VIEW
   
   lIdDoc = 0
   lidComp = idcomp
   lIdMov = idMov
         
   Me.Show vbModal
End Sub
'muestra la lista de Mov Activo Fijo y permite filtrar por algunos campos
Public Sub FList()
   
   lOper = O_LIST
   lIdDoc = 0
   lidComp = 0
      
   Me.Show vbModal
   
End Sub

Private Sub LoadGrid()
   Dim Q1 As String
   Dim Where As String
   Dim Rs As Recordset
   Dim i As Integer

   If lIdDoc <> 0 Then
      Where = " WHERE MovActivoFijo.IdDoc = " & lIdDoc
   
   ElseIf lidComp <> 0 Then
      Where = " WHERE MovActivoFijo.IdComp = " & lidComp
   
   Else  'Filtros en Fr_List
   
      If Trim(Tx_Fecha(F_INICIO)) <> "" And Trim(Tx_Fecha(F_FIN)) <> "" Then
         Where = " WHERE (MovActivoFijo.Fecha BETWEEN " & GetTxDate(Tx_Fecha(F_INICIO)) & " AND " & GetTxDate(Tx_Fecha(F_FIN)) & ")"
      End If
      
      If Trim(Tx_Descrip) <> "" Then
         If Where <> "" Then
            Where = Where & " AND "
         Else
            Where = " WHERE "
         End If
         
         Where = Where & GenLike(DbMain, Tx_Descrip, "MovActivoFijo.Descrip")
      End If
   
   End If
   
   If Where <> "" Then
      Where = Where & " AND "
   Else
      Where = " WHERE "
   End If
      
   Where = Where & " MovActivoFijo.IdEmpresa = " & gEmpresa.id & " AND MovActivoFijo.Ano = " & gEmpresa.Ano
 
   Q1 = "SELECT IdActFijo, MovActivoFijo.IdDoc, Comprobante.IdComp, TipoMovAF, MovActivoFijo.Fecha, "
   Q1 = Q1 & " Cantidad, MovActivoFijo.Descrip, MovActivoFijo.Neto, "
   Q1 = Q1 & " Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc, Comprobante.Tipo, "
   Q1 = Q1 & " Comprobante.Correlativo, MovActivoFijo.IdCuenta, FechaVentaBaja, MovActivoFijo.FImported,"
   Q1 = Q1 & " MovActivoFijo.FechaImportFile "
   Q1 = Q1 & " FROM (MovActivoFijo "
   Q1 = Q1 & " LEFT JOIN Documento ON MovActivoFijo.IdDoc = Documento.IdDoc) "
   Q1 = Q1 & " LEFT JOIN Comprobante ON MovActivoFijo.IdComp = Comprobante.IdComp "
   
   If Ch_SortDesc <> 0 Then
      Q1 = Q1 & Where & " ORDER BY MovActivoFijo.Descrip, MovActivoFijo.Fecha "
   Else
      Q1 = Q1 & Where & " ORDER BY MovActivoFijo.Fecha, MovActivoFijo.Descrip "
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDACTFIJO) = vFld(Rs("IdActFijo"))
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(i, C_TIPOMOVAF) = gMovActivoFijo(vFld(Rs("TipoMovAF")))
      Grid.TextMatrix(i, C_FECHA) = Format(vFld(Rs("Fecha")), SDATEFMT)
      Grid.TextMatrix(i, C_CANTIDAD) = Format(vFld(Rs("Cantidad")), NUMFMT)
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(i, C_NETO) = Format(vFld(Rs("Neto")), NUMFMT)
      
      If vFld(Rs("IdDoc")) <> 0 Then
         Grid.TextMatrix(i, C_DOCCOMP) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc"))
      ElseIf vFld(Rs("IdComp")) <> 0 Then
         Grid.TextMatrix(i, C_DOCCOMP) = UCase(Left(gTipoComp(vFld(Rs("Tipo"))), 3)) & " " & vFld(Rs("Correlativo"))
      End If
      
      If vFld(Rs("FechaVentaBaja")) <> 0 Then
         Grid.TextMatrix(i, C_FECHAVENTA) = Format(vFld(Rs("FechaVentaBaja")), SDATEFMT)
      End If
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      
      If vFld(Rs("FImported")) <> 0 Then    'viene del año anterior
         Grid.TextMatrix(i, C_DEANOANT) = "Si"
      End If
     
      If vFld(Rs("FechaImportFile")) <> 0 Then    'fue importado desde archivo txt
         Grid.TextMatrix(i, C_IMPFILE) = "Si"
      End If
     
      
      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   Grid.Col = C_FECHA
   Grid.ColSel = Grid.Col
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 500
   
   If Bt_MarcarExport.visible Then
      Grid.Height = Grid.Height - Bt_MarcarExport - 500
      Bt_MarcarExport.Top = Grid.Top + Grid.Height + 100
   End If
   
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   
   Col = Grid.MouseCol
   Row = Grid.MouseRow
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   '2861733 tema 1
   If Col = C_CHECK Then

      If Val(Grid.TextMatrix(Row, C_IDACTFIJO)) <> 0 Then
      Grid.Row = Row
      Grid.Col = Col
         If Grid.CellPicture = 0 Then
            Call FGrSetPicture(Grid, Row, Col, Pc_Check, 0)
            Bt_Del_All.visible = True
            vSel = vSel + 1
         Else
            Set Grid.CellPicture = LoadPicture()
            vSel = vSel - 1

            If vSel = 0 Then
            Bt_Del_All.visible = False
            End If

         End If
      End If
'
''2861733    tema 1
   ElseIf Col <> C_DOCCOMP Or Bt_ViewDoc.visible = False Then
   
      If lOper = O_EDIT Then
         Call PostClick(Bt_Edit)
      ElseIf lOper = O_SELECT Then
         Call PostClick(Bt_Sel)
      Else
         Call PostClick(Bt_ViewDet)
      End If
      
   Else
      Call PostClick(Bt_ViewDoc)
   End If
         
End Sub

Private Sub Tx_Fecha_GotFocus(Index As Integer)
   Call DtGotFocus(Tx_Fecha(Index))
End Sub

Private Sub Tx_Fecha_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_Fecha_LostFocus(Index As Integer)
   
   If Trim$(Tx_Fecha(Index)) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Fecha(Index))
   
End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha(Index))
   
   Set Frm = Nothing
End Sub

Private Sub LoadDatos()
   Dim Rs As Recordset
   Dim Q1 As String

   If lIdDoc <> 0 Then
      
      Q1 = "SELECT TipoLib, TipoDoc, NumDoc, FEmision, Entidades.Rut, Entidades.Nombre "
      Q1 = Q1 & " FROM Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " WHERE Documento.idDoc = " & lIdDoc
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
      
         Tx_Doc = gTipoLib(vFld(Rs("TipoLib"))) & " - " & GetNombreTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
         Tx_NumDoc = vFld(Rs("NumDoc"))
         Tx_FechaDoc = Format(vFld(Rs("FEmision")), DATEFMT)
         lFecha = vFld(Rs("FEmision"))
         Tx_Rut = FmtCID(vFld(Rs("Rut")))
         Tx_Nombre = vFld(Rs("Nombre"), True)
         
         Me.Caption = "Activos Fijos asociados a un Documento"
         
      Else
         Fr_Doc.visible = False
         
      End If
      
      Fr_Comp.visible = False
    
      Call CloseRs(Rs)

   ElseIf lidComp <> 0 And lIdMov <> 0 Then
      
      Q1 = "SELECT Comprobante.Correlativo, MovComprobante.Orden, Comprobante.Fecha, MovComprobante.Glosa "
      Q1 = Q1 & " FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
      Q1 = Q1 & " WHERE MovComprobante.IdComp = " & lidComp & " AND " & "MovComprobante.IdMov = " & lIdMov
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
      
         Tx_NumComp = vFld(Rs("Correlativo"))
         Tx_NumMov = vFld(Rs("Orden"))
         Tx_FechaComp = Format(vFld(Rs("Fecha")), DATEFMT)
         lFecha = vFld(Rs("Fecha"))
         Tx_GlosaMov = vFld(Rs("Glosa"), True)
      
         Me.Caption = "Activos Fijos asociados a una Línea de Comprobante"
         
      Else
         Fr_Comp.visible = False
         
      End If
      
      Call CloseRs(Rs)
      
      Fr_Doc.visible = False
      
   Else
      
      Fr_Doc.visible = False
      Fr_Comp.visible = False
      
   End If

End Sub

Private Sub SetupPriv()

   If Not ChkPriv(PRV_ING_DOCS) Then
      Bt_New.Enabled = False
      Bt_Edit.Caption = "Ver"
      Bt_Del.Enabled = False
   End If
   
End Sub
