VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmEquivalencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equivalencias de Monedas"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "FrmEquivalencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.Frame Fr_ValUnico 
      Height          =   735
      Left            =   900
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Valor único:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   3660
         Picture         =   "FrmEquivalencias.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   660
         Width           =   915
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox Cb_Moneda 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   255
         Index           =   3
         Left            =   2100
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Fr_ValMes 
      Height          =   3735
      Left            =   900
      TabIndex        =   6
      Top             =   1380
      Width           =   4815
      Begin VB.CommandButton Bt_Copy 
         Caption         =   "&Copiar "
         Height          =   675
         Left            =   3600
         Picture         =   "FrmEquivalencias.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Copia el primer valor en todos los días del mes"
         Top             =   2820
         Width           =   975
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   3195
         Left            =   840
         TabIndex        =   14
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   5636
         Cols            =   4
         Rows            =   12
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "FrmEquivalencias.frx":0464
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "FrmEquivalencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_DIA = 0
Const C_VALOR = 1
Const C_ID = 2
Const C_ESTADO = 3

Dim lIdMoneda As Integer
Dim lIdxMoneda As Integer
Dim lCaract As Byte
Dim cbMoneda As ClsCombo
Dim lOper As Integer

Private Sub Bt_Buscar_Click()
   Call SetUpGridFecha
   Call LoadAll
End Sub

Private Sub bt_Cancel_Click()
   Unload Me
End Sub


Private Sub Bt_Copy_Click()
   Dim i As Integer
   
   If vFmt(Grid.TextMatrix(Grid.FixedRows, C_VALOR)) <> 0 Then
   
      For i = Grid.FixedRows To Grid.Rows - 1
         If Grid.TextMatrix(i, C_DIA) = "" Then
            Exit For
         End If
      
         Grid.TextMatrix(i, C_VALOR) = Grid.TextMatrix(Grid.FixedRows, C_VALOR)
         
         Call FGrModRow(Grid, i, FGR_U, C_ID, C_ESTADO)
      Next i
      
      Cb_Mes.Enabled = False
      Cb_Ano.Enabled = False
      Cb_Moneda.Enabled = False
      Bt_Buscar.Enabled = False
      
   End If

End Sub

Private Sub bt_OK_Click()
   Dim Row As Integer
   Dim Q1 As String
   Dim F1 As Long
   
   If Val(cbMoneda.Matrix(2)) = MON_VUNICO Then
      If Trim(vFmt(Tx_Valor)) = "" Then
         MsgBox1 "No ha ingresado valor", vbExclamation
         Exit Sub
      End If
      
      If lOper = O_EDIT Then
         Q1 = "UPDATE Equivalencia SET Valor=" & vFmt(Tx_Valor)
         Q1 = Q1 & " WHERE idMoneda=" & cbMoneda.ItemData
      Else
         Q1 = "INSERT INTO Equivalencia (idMoneda,Valor) "
         Q1 = Q1 & " VALUES (" & cbMoneda.ItemData & "," & vFmt(Tx_Valor) & ")"
         
      End If
      Call ExecSQL(DbMain, Q1)
      
   Else
      For Row = 1 To Grid.Rows - 1
         If Trim(Grid.TextMatrix(Row, C_VALOR)) <> "" And Trim(Grid.TextMatrix(Row, C_ESTADO)) <> "" Then
         
            If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
               F1 = DateSerial(Cb_Ano, Cb_Mes.ListIndex + 1, Grid.TextMatrix(Row, C_DIA))
            Else
               F1 = DateSerial(Cb_Ano, Row, 1)
            End If
         
            If Grid.TextMatrix(Row, C_ESTADO) = FGR_U Then
               Q1 = "UPDATE Equivalencia SET Valor=" & Str(vFmt(Grid.TextMatrix(Row, C_VALOR)))
               Q1 = Q1 & " WHERE idMoneda=" & cbMoneda.ItemData
               Q1 = Q1 & " AND Fecha= " & F1
               
            ElseIf Grid.TextMatrix(Row, C_ESTADO) = FGR_I Then
               Q1 = "INSERT INTO Equivalencia (idMoneda,Fecha,Valor) "
               Q1 = Q1 & " VALUES (" & cbMoneda.ItemData & "," & F1
               Q1 = Q1 & "," & Str(vFmt(Grid.TextMatrix(Row, C_VALOR))) & ")"
               
            End If
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      Next Row
   End If
   Unload Me
   
End Sub

Private Sub Cb_Ano_Click()
   Call EnableFrm(True)
   
End Sub

Private Sub Cb_Mes_Click()
  Call EnableFrm(True)
End Sub

Private Sub cb_Moneda_Click()
   Dim i As Integer
   
   Cb_Mes.Enabled = (Val(cbMoneda.Matrix(2)) = MON_VDIA)
   Cb_Ano.Enabled = (Val(cbMoneda.Matrix(2)) <> MON_VUNICO)
   
   lIdMoneda = cbMoneda.ItemData
   
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).Id = lIdMoneda Then
         lIdxMoneda = i
         Exit For
      End If
   Next i
   
   Call EnableFrm(True)
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim Wh As String
   
   Call FillMes(Cb_Mes)
   Cb_Mes.ListIndex = Month(Int(Now)) - 1
   
   Call FillCbAno(Cb_Ano)
   Call SelItem(Cb_Ano, Year(Now))
   
   If lIdMoneda <> 0 Then
      Wh = " AND idMoneda=" & lIdMoneda
   End If
   
   Set cbMoneda = New ClsCombo
   Call cbMoneda.SetControl(Cb_Moneda)
   
   Q1 = "SELECT Descrip,idMoneda,Caracteristica FROM Monedas WHERE Caracteristica<>" & MON_NACION
   Q1 = Q1 & Wh
   Call cbMoneda.FillCombo(DbMain, Q1, -1)
   
   Call SetUpGrid
   Call LoadAll
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
End Sub

Public Function FEdit(idMoneda As Long) As Integer
   lIdMoneda = idMoneda
   Me.Show vbModal
   
End Function
Private Sub SetUpGridFecha()
   Dim F1 As Long, F2 As Long
   Dim Row As Integer
   
   If Val(cbMoneda.Matrix(2)) = MON_VUNICO Then
      Me.Height = 2580
   Else
      Me.Height = 5475
   End If
   
   Fr_ValMes.Visible = Val(cbMoneda.Matrix(2)) <> MON_VUNICO
   Fr_ValUnico.Visible = Val(cbMoneda.Matrix(2)) = MON_VUNICO
   
   For Row = Grid.FixedRows To Grid.Rows - 1
      Grid.TextMatrix(Row, C_VALOR) = ""
      Grid.TextMatrix(Row, C_ID) = ""
      
   Next Row
   
   If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
      
      Grid.ColAlignment(C_DIA) = flexAlignRightCenter
      Grid.TextMatrix(0, C_DIA) = "Día mes"
      
      Grid.ColWidth(C_DIA) = 800
      Grid.ColWidth(C_VALOR) = 1450
      
      F1 = DateSerial(Cb_Ano, Cb_Mes.ListIndex + 1, 1)
      Call FirstLastMonthDay(F1, F1, F2)
      
      Row = Grid.FixedRows
      Grid.Rows = Grid.FixedRows
      
      Do While F1 <= F2
         Grid.Rows = Row + 1
         
         Grid.TextMatrix(Row, C_DIA) = Day(F1)
         F1 = F1 + 1
         
         Row = Row + 1
         
      Loop
      
   ElseIf Val(cbMoneda.Matrix(2)) = MON_VMES Then
   
      Grid.ColAlignment(C_DIA) = flexAlignLeftCenter
      Grid.TextMatrix(0, C_DIA) = "Mes"
      
      Grid.ColWidth(C_DIA) = 1100
      Grid.ColWidth(C_VALOR) = 1400
      
      F1 = 1
      Row = Grid.FixedRows
      Grid.Rows = Grid.FixedRows
      Do While F1 <= 12
         Grid.Rows = Row + 1
         
         Grid.TextMatrix(Row, C_DIA) = gNomMes(F1)
         F1 = F1 + 1
         
         Row = Row + 1
         
      Loop
      
   End If
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim F1 As Long, F2 As Long
      
   Q1 = " SELECT Fecha,Valor FROM Equivalencia WHERE idMoneda=" & cbMoneda.ItemData
   If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
      F1 = DateSerial(Cb_Ano, Cb_Mes.ListIndex + 1, 1)
      Call FirstLastMonthDay(F1, F1, F2)
      Q1 = Q1 & " AND Fecha BETWEEN " & F1 & " AND " & F2
      
   ElseIf Val(cbMoneda.Matrix(2)) = MON_VMES Then
      Q1 = Q1 & " AND YEAR(Fecha)=" & Cb_Ano
      
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
      Row = 1
      Do While Rs.EOF = False
         Grid.TextMatrix(Day(vFld(Rs("Fecha"))), C_VALOR) = Format(vFld(Rs("Valor")), gMonedas(lIdxMoneda).FormatInf)
         Grid.TextMatrix(Day(vFld(Rs("Fecha"))), C_ID) = vFld(Rs("Valor"))
         Rs.MoveNext
         
      Loop
   ElseIf Val(cbMoneda.Matrix(2)) = MON_VMES Then
      Row = 1
      Do While Rs.EOF = False
         Grid.TextMatrix(Month(vFld(Rs("Fecha"))), C_VALOR) = Format(vFld(Rs("Valor")), gMonedas(lIdxMoneda).FormatInf)
         Grid.TextMatrix(Month(vFld(Rs("Fecha"))), C_ID) = vFld(Rs("Valor"))
         Rs.MoveNext
         
      Loop
   Else
      If Rs.EOF = False Then
         Tx_Valor = Format(vFld(Rs("Valor")), gMonedas(lIdxMoneda).FormatInf)
         lOper = O_EDIT
      Else
         lOper = O_NEW
      End If
      
   End If
   Call CloseRs(Rs)
 '  Call EnableFrm(False)
   
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Trim(Value) <> "" Then
      Value = Format(Value, gMonedas(lIdxMoneda).FormatInf)
      
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_ESTADO)
      Cb_Mes.Enabled = False
      Cb_Ano.Enabled = False
      Cb_Moneda.Enabled = False
      Bt_Buscar.Enabled = False
      
   End If
      
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
 
   EdType = FEG_Edit
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyDec(KeyAscii)
End Sub
Private Sub SetUpGrid()
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESTADO) = 0
   
   Grid.ColWidth(C_DIA) = 1200
   Grid.ColWidth(C_VALOR) = 1200
   
   Grid.FixedAlignment(C_DIA) = flexAlignCenterCenter
   Grid.FixedAlignment(C_VALOR) = flexAlignCenterCenter
   
   Grid.ColAlignment(C_DIA) = flexAlignRightCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_DIA) = "Fecha"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call SetUpGridFecha
End Sub
Private Sub EnableFrm(bool As Boolean)
   Grid.Locked = Not bool
End Sub
