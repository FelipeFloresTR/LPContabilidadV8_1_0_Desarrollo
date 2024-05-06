VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSelCompTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Comprobante Tipo"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   Icon            =   "FrmSelCompTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   7575
      Begin VB.TextBox Tx_Glosa 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   5295
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   6300
         Picture         =   "FrmSelCompTipo.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Para listar Comprobantes Tipo, basta ingresar parte de una palabra en el nombre o en la glosa."
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1020
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   9
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   11100
      TabIndex        =   6
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   735
      Left            =   11100
      Picture         =   "FrmSelCompTipo.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   660
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5355
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "FrmSelCompTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NOMBRE = 0
Const C_DESCRIP = 1
Const C_GLOSA = 2
Const C_TIPO = 3
Const C_IDTIPO = 4
Const C_IDCOMP = 5

Const NCOLS = C_IDCOMP

Dim lRc As Integer
Dim lidComp As Long
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_NOMBRE) = 1500
   Grid.ColWidth(C_DESCRIP) = 2500
   Grid.ColWidth(C_GLOSA) = 6000
   Grid.ColWidth(C_TIPO) = 1700
   Grid.ColWidth(C_IDTIPO) = 0
   Grid.ColWidth(C_IDCOMP) = 0
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
      Grid.ColAlignment(i) = flexAlignLeftCenter
   Next i
   
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_TIPO) = "Tipo"
   
   Call FGrVRows(Grid)
End Sub
Private Sub Bt_Buscar_Click()
   Dim Q1 As String
   Dim Wh As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Wh = ""
   If Trim(Tx_Nombre) <> "" Then
      Wh = GenLike(DbMain, Tx_Nombre, "Nombre", 3)
      
   End If
   
   If Trim(Tx_Glosa) <> "" Then
      If Wh <> "" Then
         Wh = Wh & " AND "
      End If
      Wh = Wh & GenLike(DbMain, Tx_Glosa, "Glosa", 3)
   End If
   
   If Cb_Tipo.ListIndex <> 0 Then
      If Wh <> "" Then
         Wh = Wh & " AND "
      End If
      Wh = Wh & " Tipo=" & ItemData(Cb_Tipo)
   End If
   
   If Wh <> "" Then
      Wh = " WHERE " & Wh
      Wh = Wh & " AND IdEmpresa = " & gEmpresa.id
      Wh = Wh & " AND (Nombre IS NOT NULL OR GLOSA IS NOT NULL) "
   
   Else
      Wh = " WHERE IdEmpresa = " & gEmpresa.id
      Wh = Wh & " AND (Nombre IS NOT NULL OR GLOSA IS NOT NULL) "
   End If
   
   Q1 = "SELECT Nombre, Descrip, Glosa, Tipo, IdComp FROM CT_Comprobante " & Wh
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 1
   Grid.rows = 1
   Do While Rs.EOF = False
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(Row, C_NOMBRE) = vFld(Rs("Nombre"), True)
      Grid.TextMatrix(Row, C_DESCRIP) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(Row, C_GLOSA) = vFld(Rs("Glosa"), True)
      Grid.TextMatrix(Row, C_TIPO) = gTipoComp(vFld(Rs("Tipo")))
      Grid.TextMatrix(Row, C_IDTIPO) = vFld(Rs("Tipo"))
      Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("IdComp"))
      
      Row = Row + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row = 0 Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_IDCOMP) = "" Then
      Exit Sub
   End If
   
   lidComp = Grid.TextMatrix(Row, C_IDCOMP)
   lRc = vbOK
   
   Unload Me
   
End Sub
Public Function FSelect(idcomp As Long) As Integer
   Me.Show vbModal
   
   idcomp = lidComp
   FSelect = lRc
   
End Function

Private Sub Form_Load()
   Dim i As Integer
   
   lRc = vbCancel
   
   '3133472
   Call UpdateComprobantesTipo
   '3133472

   Call AddItem(Cb_Tipo, "", 0)
   'Cb_Tipo.AddItem ""
   For i = 1 To N_TIPOCOMP
      Call AddItem(Cb_Tipo, gTipoComp(i), i)
      'Cb_Tipo.AddItem gTipoComp(i)
      'Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
      
   Next i
   Cb_Tipo.ListIndex = 0
   
   
   
   Call SetUpGrid
   
   Call Bt_Buscar_Click
End Sub

Private Sub Grid_DblClick()
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Grid.Row, C_IDCOMP) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Call Bt_Sel_Click
End Sub

Private Sub Tx_Nombre_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
   Call KeyUpper(KeyAscii)
End Sub
