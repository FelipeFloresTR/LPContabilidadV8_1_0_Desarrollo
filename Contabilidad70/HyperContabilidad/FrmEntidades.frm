VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEntidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entidades"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "FrmEntidades.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8760
      TabIndex        =   14
      Top             =   540
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   480
      TabIndex        =   17
      Top             =   420
      Width           =   7815
      Begin VB.ComboBox Cb_OrdenarPor 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1875
      End
      Begin VB.ComboBox Cb_Clasif 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenar por:"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   21
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   915
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5415
      Left            =   450
      TabIndex        =   1
      Top             =   1140
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame Fr_SelEdit 
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   8760
      TabIndex        =   19
      Top             =   1260
      Width           =   1095
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   4500
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":05C1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Seleccionar"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":0C03
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nueva Entidad"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":1195
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar Entidad"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":17F7
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modificar Entidad"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Index           =   1
         Left            =   0
         Picture         =   "FrmEntidades.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir Entidad"
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Edit 
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   8760
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
      Begin VB.CommandButton Bt_CopyExcel 
         Caption         =   "&Copiar Excel"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":23F9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Copiar datos a Excel"
         Top             =   3645
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":29AE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir Entidad"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":2FDD
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Modificar Entidad"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":35B0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Eliminar Entidad"
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":3C12
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Nueva Entidad"
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Sel 
      BorderStyle     =   0  'None
      Height          =   5265
      Left            =   8715
      TabIndex        =   16
      Top             =   1200
      Width           =   1155
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Index           =   0
         Left            =   0
         Picture         =   "FrmEntidades.frx":41A4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Seleccionar"
         Top             =   60
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_CODIGO = 1
Const C_NOMBRE = 2
Const C_ESTADO = 3
Const C_DIRECCION = 4
Const C_TELEFONO = 5
Const C_FAX = 6
Const C_EMAIL = 7
Const C_WEB = 8
Const C_ID = 9
Const C_IDESTADO = 10
Const C_NOTVALIDRUT = 11

Dim lEntidad As Entidad_t
Dim lTipoEntidad As Integer
Dim lRc As Integer

Dim InLoad As Boolean

Dim lOper As Integer

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click(Index As Integer)
   MousePointer = vbHourglass
   Call FGr2Clip(Grid, "Listado de " & Cb_Clasif)
   MousePointer = vbDefault
End Sub

Private Sub Bt_Del_Click(Index As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_CODIGO) = "" Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_RUT) = ENTIMP_RUT Then
      MsgBox1 "Este RUT corresponde a la entidad especial para Formulario de Importaciones y no puede ser eliminada.", vbExclamation
      Exit Sub
   End If
   
   Q1 = "SELECT Count(*) as n FROM Documento WHERE idEntidad=" & vFmt(Grid.TextMatrix(Row, C_ID))
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If vFld(Rs("n")) <> 0 Then
      MsgBox1 "No puede borrar la entidad " & Grid.TextMatrix(Row, C_NOMBRE) & ", existe un movimiento asociado.", vbExclamation
      Call CloseRs(Rs)
      Exit Sub
   End If
   Call CloseRs(Rs)
   
   If MsgBox1("¿Está seguro de eliminar la entidad " & Grid.TextMatrix(Row, C_NOMBRE) & "?", vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
   Grid.RowHeight(Row) = 0
   Grid.rows = Grid.rows + 1
   
   Q1 = " WHERE idEntidad = " & vFmt(Grid.TextMatrix(Row, C_ID))
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call DeleteSQL(DbMain, "Entidades", Q1)
   
End Sub

Private Sub Bt_Edit_Click(Index As Integer)
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Rc As Integer
      
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   Call FillStruct(Row, Cb_Clasif)
   Set Frm = New FrmEntidad
   Rc = Frm.FEdit(lEntidad)
   If Rc = vbOK Or Rc = vbRetry Then
      Call UpDateGrid(Row)
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_New_Click(Index As Integer)
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Rc As Integer
 
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   lEntidad.Clasif = ItemData(Cb_Clasif)
   Rc = Frm.FNew(lEntidad)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      If lEntidad.Clasif = ItemData(Cb_Clasif) Or ItemData(Cb_Clasif) < 0 Then
         Row = FGrAddRow(Grid)
         Call UpDateGrid(Row)
      End If
      
   ElseIf Rc = vbRetry Then ' ya existe
      If lEntidad.Clasif = ItemData(Cb_Clasif) Then
         ' si ya existe lo buscamos para actualizarlo
         For Row = Grid.FixedRows To Grid.rows - 1
            If Val(Grid.TextMatrix(Row, C_ID)) = lEntidad.id Then
               Call UpDateGrid(Row)
               Exit For
            End If
         Next Row
      End If
      
   End If
   
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_Print_Click(Index As Integer)
   Dim ColWi(C_NOTVALIDRUT) As Integer
   Dim Total(C_NOTVALIDRUT) As String
   Dim i As Integer
   Dim OldOrient As Integer
      
   If Grid.TextMatrix(1, C_RUT) = "" Then
      Exit Sub
   End If
      
   MousePointer = vbHourglass
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   Total(0) = ""
   
   OldOrient = Printer.Orientation
   Printer.Orientation = ORIENT_HOR
   Call PrtFlexGrid(Grid, "", "LISTADO DE ENTIDADES", "", Cb_Clasif, ColWi, Total, False, , , , , , , , , , , True)
   Printer.Orientation = OldOrient
   
   MousePointer = vbDefault
   
End Sub


Private Sub Bt_Sel_Click(Index As Integer)
   Dim Row As Integer
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
   End If
   
'   Call FillStruct(Row, cb_ClasifSel)
   Call FillStruct(Row, Cb_Clasif)
   
   lRc = vbOK
   Unload Me
End Sub

Private Sub cb_Clasif_Click()

   Me.MousePointer = vbHourglass
   Call LoadAll(Cb_Clasif)
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Cb_OrdenarPor_Click()
   
   If Not InLoad Then
      Me.MousePointer = vbHourglass
      Call LoadAll(Cb_Clasif)
      Me.MousePointer = vbDefault
   End If

End Sub

Private Sub Form_Load()
   
   lRc = vbCancel
   
   InLoad = True
   
   Call CbAddItem(Cb_OrdenarPor, "Nombre", 1)
   Call CbAddItem(Cb_OrdenarPor, "RUT", 2)
   Cb_OrdenarPor.ListIndex = 0   'nombre
   
   Call CbAddItem(Cb_Clasif, " ", -1)
   
   Call FillCbClasifEnt(Cb_Clasif, lTipoEntidad)
   Cb_Clasif.ListIndex = 1 'clientes
   
   Fr_Edit.visible = lOper = O_EDIT
   Fr_Sel.visible = lOper = O_VIEW
   Fr_SelEdit.visible = lOper = O_SELEDIT
   
   Call SetUpGrid
   
   Call FrmEnab(gEmpresa.FCierre = 0)
   InLoad = False
      
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
      
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_NOMBRE) = 2800
   Grid.ColWidth(C_ESTADO) = 0
   
   Grid.ColWidth(C_DIRECCION) = 2400
   Grid.ColWidth(C_TELEFONO) = 1100
   Grid.ColWidth(C_FAX) = 1000
   Grid.ColWidth(C_EMAIL) = 2000
   Grid.ColWidth(C_WEB) = 2000
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_NOTVALIDRUT) = 0
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
      Grid.ColAlignment(i) = flexAlignLeftCenter
   Next i
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter

   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_CODIGO) = "Nombre Corto"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"
   Grid.TextMatrix(0, C_DIRECCION) = "Dirección"
   Grid.TextMatrix(0, C_TELEFONO) = "Teléfonos"
   Grid.TextMatrix(0, C_FAX) = "Fax"
   Grid.TextMatrix(0, C_EMAIL) = "email"
   Grid.TextMatrix(0, C_WEB) = "WEB"
   
   
End Sub
Public Function FEdit() As Integer
   lOper = O_EDIT
   Me.Show vbModal
   
   FEdit = lRc
   
End Function
Friend Function FSelect(Entidad As Entidad_t, Optional ByVal TipoEntidad As Integer = ENT_CLIENTE) As Integer
   lOper = O_VIEW
   lTipoEntidad = TipoEntidad
   
   Me.Show vbModal
   
   FSelect = lRc
   Entidad = lEntidad
   
End Function
Friend Function FSelEdit(Entidad As Entidad_t, Optional ByVal TipoEntidad As Integer = ENT_CLIENTE) As Integer
   lOper = O_SELEDIT
   lTipoEntidad = TipoEntidad
   
   Me.Show vbModal
   
   FSelEdit = lRc
   Entidad = lEntidad
   
End Function

Private Sub UpDateGrid(Row As Integer)
   
   Grid.TextMatrix(Row, C_RUT) = lEntidad.Rut
   'Grid.TextMatrix(Row, C_NOMBRE) = lEntidad.Nombre
   Grid.TextMatrix(Row, C_CODIGO) = lEntidad.Codigo
   Grid.TextMatrix(Row, C_NOMBRE) = lEntidad.Nombre
   Grid.TextMatrix(Row, C_ESTADO) = gEstadoEntidad(lEntidad.Estado)
   Grid.TextMatrix(Row, C_IDESTADO) = lEntidad.Estado
   Grid.TextMatrix(Row, C_ID) = lEntidad.id
   
End Sub
Private Sub LoadAll(Cb As ComboBox)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Clasif As Integer
   
   Clasif = ItemData(Cb)
   
   Q1 = "SELECT idEntidad, Rut, Codigo, Nombre, Estado, NotValidRut, Direccion,Ciudad,"
   Q1 = Q1 & "Telefonos,Fax,email,Web"
   Q1 = Q1 & " FROM Entidades"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   If Clasif >= 0 Then
      Q1 = Q1 & " AND Clasif" & Clasif & "=" & CON_CLASIF
   End If
   
   If LCase(Cb_OrdenarPor) = "rut" Then
      Q1 = Q1 & " ORDER BY right( " & SqlConcat(gDbType, "'0'", "RUT") & ", 8)"
   Else
      Q1 = Q1 & " ORDER BY " & Cb_OrdenarPor
   
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   i = Grid.FixedRows
   Grid.rows = i
   Do While Rs.EOF = False
      Grid.rows = i + 1
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If

      Grid.TextMatrix(i, C_RUT) = FmtStRut(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)       'FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)
      Grid.TextMatrix(i, C_CODIGO) = vFld(Rs("Codigo"))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
      Grid.TextMatrix(i, C_ESTADO) = gEstadoEntidad(vFld(Rs("Estado")))
      Grid.TextMatrix(i, C_DIRECCION) = vFld(Rs("Direccion")) & " " & vFld(Rs("Ciudad"))
      Grid.TextMatrix(i, C_TELEFONO) = vFld(Rs("Telefonos"))
      Grid.TextMatrix(i, C_FAX) = vFld(Rs("Fax"), True)
      Grid.TextMatrix(i, C_EMAIL) = vFld(Rs("Email"), True)
      Grid.TextMatrix(i, C_WEB) = vFld(Rs("Web"), True)
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(i, C_ID) = vFld(Rs("idEntidad"))
      Grid.TextMatrix(i, C_NOTVALIDRUT) = vFld(Rs("NotValidRut"))
      
      i = i + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   Grid.Redraw = True

End Sub
Private Sub FillStruct(Row As Integer, Cb As ComboBox)

   lEntidad.Rut = Grid.TextMatrix(Row, C_RUT)
   lEntidad.Codigo = Grid.TextMatrix(Row, C_CODIGO)
   lEntidad.Nombre = Grid.TextMatrix(Row, C_NOMBRE)
   lEntidad.Estado = Val(Grid.TextMatrix(Row, C_IDESTADO))
   lEntidad.id = Grid.TextMatrix(Row, C_ID)
   lEntidad.Clasif = ItemData(Cb)
   lEntidad.NotValidRut = Val(Grid.TextMatrix(Row, C_NOTVALIDRUT))
   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - W.YCaption - Grid.Top - 435 - W.yFrame * 2
   Grid.Width = Me.Width - Grid.Left - Fr_Edit.Width - 435 * 2 - W.xFrame * 2
   Fr_Edit.Left = Grid.Left + Grid.Width + 435
   
   Call FGrVRows(Grid)
 
End Sub

Private Sub Grid_DblClick()

   If lOper = O_VIEW Then
      Call Bt_Sel_Click(0)
   ElseIf lOper = O_EDIT Then
      Call Bt_Edit_Click(0)
   ElseIf lOper = O_SELEDIT Then
      Call Bt_Sel_Click(1)
   End If

End Sub
Private Sub FrmEnab(ByVal bool As Boolean)
   Dim i As Integer

   If Not ChkPriv(PRV_ADM_DEF) Then
      bool = False
   End If
   
   For i = 0 To 1
      Bt_New(i).Enabled = bool
      Bt_Edit(i).Enabled = bool
      Bt_Del(i).Enabled = bool
      Bt_CopyExcel(i).Enabled = bool
   Next i

End Sub
