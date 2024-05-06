VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMonedas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monedas"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "FrmMonedas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bt_Print 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmMonedas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Cuentas"
      Top             =   3720
      Width           =   1155
   End
   Begin VB.CommandButton bt_Equivalencia 
      Caption         =   "Equi&valencia"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmMonedas.frx":063B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7860
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4035
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7117
      _Version        =   393216
      Rows            =   10
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton Bt_New 
      Caption         =   "&Agregar"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmMonedas.frx":0C54
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Edit 
      Caption         =   "Edi&tar"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmMonedas.frx":11E6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmMonedas.frx":17B9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   420
      Picture         =   "FrmMonedas.frx":1E1B
      Top             =   420
      Width           =   720
   End
End
Attribute VB_Name = "FrmMonedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_DESC = 1
Const C_SIMBOLO = 2
Const C_DECINF = 3
Const C_DECVENTA = 4
Const C_CARACT = 5

Const LASTCOL = C_CARACT

Dim lMoneda As Monedas_t

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Trim(Grid.TextMatrix(Row, 0)) = "" Then
      MsgBox1 "No hay moneda seleccionada.", vbExclamation
      Exit Sub
      
   End If
   
   MsgBox1 "FALTA VER CUAL ES SU RESTRICCION", vbExclamation
   
   If MsgBox1("¿Está seguro que desea eliminar moneda " & Grid.TextMatrix(Row, C_DESC) & " ?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   Call DeleteSQL(DbMain, "Monedas", " WHERE idMoneda = " & vFmt(Grid.TextMatrix(Row, C_CODIGO)))
   Grid.RowHeight(Row) = 0
   Grid.rows = Grid.rows + 1
      
   MousePointer = vbDefault
End Sub

Private Sub Bt_Edit_Click()
   Dim Moneda As Monedas_t
   Dim Frm As FrmMantMoneda
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Trim(Grid.TextMatrix(Row, C_CODIGO)) = "" Then
      Exit Sub
   End If
      
   Set Frm = New FrmMantMoneda
   Moneda.id = Grid.TextMatrix(Row, C_CODIGO)
   If Frm.FEdit(Moneda) = vbOK Then
      Grid.TextMatrix(Row, C_DESC) = Moneda.Descrip
      Grid.TextMatrix(Row, C_SIMBOLO) = Moneda.Simbolo
      Grid.TextMatrix(Row, C_DECINF) = Moneda.DecInf
      Grid.TextMatrix(Row, C_DECVENTA) = Moneda.DecVenta
      Grid.TextMatrix(Row, C_CARACT) = Moneda.Caract
      
   End If
   
   Set Frm = Nothing
  
End Sub

Private Sub bt_Equivalencia_Click()
   Dim Frm As FrmEquivalencias
   Dim Row As Integer
   
   Row = Grid.Row
   If Trim(Grid.TextMatrix(Row, C_CODIGO)) = "" Then
      Exit Sub
   End If
   
   If vFmt(Grid.TextMatrix(Row, C_CARACT)) = MON_NACION Then
      MsgBox1 "Moneda nacional no requiere equivalencia", vbExclamation
      Exit Sub
   End If
   
   Set Frm = New FrmEquivalencias
   Call Frm.FEdit(vFmt(Grid.TextMatrix(Row, C_CODIGO)))
   Set Frm = Nothing
   
End Sub

Private Sub Bt_New_Click()
   Dim Moneda As Monedas_t
   Dim Frm As FrmMantMoneda
   Dim Row As Integer
   
   Set Frm = New FrmMantMoneda
   If Frm.FNew(Moneda) = vbOK Then
      Row = FGrAddRow(Grid)
      
      Grid.TextMatrix(Row, C_CODIGO) = Moneda.id
      Grid.TextMatrix(Row, C_DESC) = Moneda.Descrip
      Grid.TextMatrix(Row, C_SIMBOLO) = Moneda.Simbolo
      Grid.TextMatrix(Row, C_DECINF) = Moneda.DecInf
      Grid.TextMatrix(Row, C_DECVENTA) = Moneda.DecVenta
      Grid.TextMatrix(Row, C_CARACT) = Moneda.Caract
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Print_Click()
   Dim ColWi(C_CARACT) As Integer
   Dim Total(C_CARACT) As String
   Dim i As Integer
   
   If Grid.TextMatrix(Grid.FixedRows, C_CODIGO) = "" Then
      Exit Sub
   End If
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
     
   Next i
   
   Total(0) = ""
   Call PrtFlexGrid(Grid, "", "Informe de Monedas", "", "", ColWi, Total, False, , , , , , , , , , , True)
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   Call LoadAll
   
   Call SetupPriv
   
End Sub
Private Sub SetUpGrid()
   Dim Col As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_CODIGO) = 600
   Grid.ColWidth(C_DESC) = 2680
   Grid.ColWidth(C_SIMBOLO) = 700
   Grid.ColWidth(C_DECINF) = 900
   Grid.ColWidth(C_DECVENTA) = 900
   Grid.ColWidth(C_CARACT) = 0
      
   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_SIMBOLO) = flexAlignLeftCenter
   Grid.ColAlignment(C_DECINF) = flexAlignRightCenter
   Grid.ColAlignment(C_DECVENTA) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_CODIGO) = "Índice"
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   Grid.TextMatrix(0, C_SIMBOLO) = "Símbolo"
   Grid.TextMatrix(0, C_DECINF) = "Decimales"
   Grid.TextMatrix(1, C_DECINF) = "Ingreso"
   Grid.TextMatrix(0, C_DECVENTA) = "Decimales"
   Grid.TextMatrix(1, C_DECVENTA) = "Salida"
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Q1 = "SELECT * FROM Monedas ORDER BY idMoneda"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   
   Do While Rs.EOF = False
      Grid.rows = Row + 1
         
      Grid.TextMatrix(Row, C_CODIGO) = vFld(Rs("idMoneda"))
      Grid.TextMatrix(Row, C_DESC) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(Row, C_SIMBOLO) = vFld(Rs("Simbolo"))
      Grid.TextMatrix(Row, C_DECINF) = vFld(Rs("DecInf"))
      Grid.TextMatrix(Row, C_DECVENTA) = vFld(Rs("DecVenta"))
      Grid.TextMatrix(Row, C_CARACT) = vFld(Rs("Caracteristica"))
      
      Row = Row + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
End Sub

Private Sub Grid_DblClick()
   Call Bt_Edit_Click
End Sub
Private Sub SetupPriv()

   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
      bt_Equivalencia.Enabled = True
   End If

End Sub

