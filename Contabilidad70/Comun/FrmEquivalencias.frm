VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmEquivalencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equivalencias de Monedas"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "FrmEquivalencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6600
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Top             =   300
      Width           =   4815
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   3660
         Picture         =   "FrmEquivalencias.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   660
         Width           =   915
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Moneda 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   255
         Index           =   3
         Left            =   2220
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Fr_ValMes 
      Height          =   5235
      Left            =   1440
      TabIndex        =   14
      Top             =   1620
      Width           =   4815
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   4635
         Left            =   780
         TabIndex        =   4
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   8176
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
      Begin VB.CommandButton Bt_Valor 
         Caption         =   "Obtener Valor"
         Height          =   615
         Left            =   3600
         TabIndex        =   6
         ToolTipText     =   "Obtener valor del mes seleccionado"
         Top             =   4260
         Width           =   975
      End
      Begin VB.CommandButton Bt_Copy 
         Caption         =   "&Copiar "
         Height          =   675
         Left            =   3600
         Picture         =   "FrmEquivalencias.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Copia el primer valor en todos los días del mes"
         Top             =   3420
         Width           =   975
      End
   End
   Begin VB.Frame Fr_ValUnico 
      Height          =   735
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   4815
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Valor único:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   360
      Picture         =   "FrmEquivalencias.frx":0A72
      Top             =   420
      Width           =   780
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
Dim lEditNew As Integer
Dim lHeight As Long
Dim lOper As Integer
Dim lValor As Double
Dim lRc As Integer

Dim lSoloMonedas As Boolean   'no UF ni UTM

Private Sub Bt_Buscar_Click()
   Call SetUpGridFecha
   Call LoadAll
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub


Private Sub bt_Copy_Click()
   Dim i As Integer
   
   If vFmt(Grid.TextMatrix(Grid.FixedRows, C_VALOR)) <> 0 Then
   
      For i = Grid.FixedRows To Grid.rows - 1
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

Private Sub Bt_Valor_Click()
   Dim Value As Double, r As Integer, Dia As Integer, Dt As Long, bFound As Boolean
   
   If Grid.Locked Then
      Exit Sub
   End If
   
   r = Grid.Row
   Dia = r
   
   If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
      
      For r = Grid.Row To Grid.RowSel
         Dt = DateSerial(CbItemData(Cb_Ano), CbItemData(Cb_Mes), r)
         Value = LPGetValorDia(cbMoneda.Matrix(3), Dt)
         If Value <> -7777 Then
      
            Grid.TextMatrix(r, C_VALOR) = Format(Value, gMonedas(lIdxMoneda).FormatInf)
            Call FGrModRow(Grid, r, FGR_U, C_ID, C_ESTADO)
            bFound = True
         Else
            Exit For
         End If
      Next r
         
   Else
      If Len(Trim(Grid.TextMatrix(r, C_DIA))) <= 2 Then    'es valor mensual y aún dice día del mes
         MsgBox1 "Presione el botón Listar antes de continuar.", vbExclamation
         Exit Sub
      End If
      
      Value = LPGetValorMes(cbMoneda.Matrix(3), CbItemData(Cb_Ano), r)
      If Value <> -7777 Then
   
         Grid.TextMatrix(r, C_VALOR) = Format(Value, gMonedas(lIdxMoneda).FormatInf)
         
         Call FGrModRow(Grid, r, FGR_U, C_ID, C_ESTADO)
         bFound = True
      End If
   End If
      
   If bFound Then

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
   Dim ValUTM As Double
      
   If Val(cbMoneda.Matrix(2)) = MON_VUNICO Then
   
      If Trim(vFmt(Tx_Valor)) = "" Then
         MsgBox1 "No ha ingresado valor", vbExclamation
         Exit Sub
      End If
      
      If lOper = O_SELECT Then
         lValor = vFmt(Tx_Valor)
         lRc = vbOK
         Unload Me
      End If
         
      If lEditNew = O_EDIT Then
         Q1 = "UPDATE Equivalencia SET Valor=" & vFmt(Tx_Valor)
         Q1 = Q1 & " WHERE idMoneda=" & cbMoneda.ItemData
      Else
         Q1 = "INSERT INTO Equivalencia (idMoneda,Valor) "
         Q1 = Q1 & " VALUES (" & cbMoneda.ItemData & "," & vFmt(Tx_Valor) & ")"
         
      End If
      Call ExecSQL(DbMain, Q1)
      
   Else
   
      If lOper = O_SELECT Then
      
         If Grid.Row < Grid.FixedRows Or vFmt(Grid.TextMatrix(Grid.Row, C_VALOR)) = 0 Then
            Exit Sub
         End If
         
         lValor = vFmt(Grid.TextMatrix(Grid.Row, C_VALOR))
         lRc = vbOK
         Unload Me
         
      End If

      For Row = 1 To Grid.rows - 1
         If Trim(Grid.TextMatrix(Row, C_VALOR)) <> "" And Trim(Grid.TextMatrix(Row, C_ESTADO)) <> "" Then
         
            If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
               F1 = DateSerial(Cb_Ano, Cb_Mes.ListIndex + 1, Grid.TextMatrix(Row, C_DIA))
            Else
               F1 = DateSerial(Cb_Ano, Row, 1)
            End If
         
            If Grid.TextMatrix(Row, C_ESTADO) = FGR_U Then
               Q1 = "UPDATE Equivalencia SET Valor=" & str(vFmt(Grid.TextMatrix(Row, C_VALOR)))
               Q1 = Q1 & " WHERE idMoneda=" & cbMoneda.ItemData
               Q1 = Q1 & " AND Fecha= " & F1
               
            ElseIf Grid.TextMatrix(Row, C_ESTADO) = FGR_I Then
               Q1 = "INSERT INTO Equivalencia (idMoneda,Fecha,Valor) "
               Q1 = Q1 & " VALUES (" & cbMoneda.ItemData & "," & F1
               Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(Row, C_VALOR))) & ")"
               
            End If
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      Next Row
   End If
   
   'recalculamos valor de MaxCred33 en UTM por si cambió valor UTM del mes
   If GetValMoneda("UTM", ValUTM, DateSerial(gEmpresa.Ano, 12, 1)) = False Then
      gMaxUTMCred33_Pesos = 0
   Else
      gMaxUTMCred33_Pesos = gMaxUTMCred33 * ValUTM
   End If
   
   lRc = vbOK

   Unload Me
   
End Sub

Private Sub Cb_Ano_Click()
   Call SetupPriv
   
End Sub

Private Sub Cb_Mes_Click()
  Call SetupPriv
End Sub

Private Sub cb_Moneda_Click()
   Dim i As Integer
   
   Cb_Mes.Enabled = (Val(cbMoneda.Matrix(2)) = MON_VDIA)
   Cb_Ano.Enabled = (Val(cbMoneda.Matrix(2)) <> MON_VUNICO)
   
   lIdMoneda = cbMoneda.ItemData
   
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).id = lIdMoneda Then
         lIdxMoneda = i
         Exit For
      End If
   Next i
   
   Call SetupPriv
   
   If gMonedas(i).EsFijo Then
      Bt_Valor.Enabled = True
   Else
      Bt_Valor.Enabled = False
   End If
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim Wh As String
   
   lHeight = Me.Height
   
   Call FillMes(Cb_Mes)
   Cb_Mes.ListIndex = Month(Int(Now)) - 1
   
   Call FillCbAno(Cb_Ano)
   Call SelItem(Cb_Ano, Year(Now))
   
   If lIdMoneda <> 0 Then
      Wh = " AND idMoneda=" & lIdMoneda
   End If
   
   If lSoloMonedas Then
      Wh = " AND Simbolo NOT IN ('UF', 'UTM')"
   End If
   
   Set cbMoneda = New ClsCombo
   Call cbMoneda.SetControl(Cb_Moneda)
   
   Q1 = "SELECT Descrip, idMoneda, Caracteristica, Simbolo FROM Monedas WHERE Caracteristica<>" & MON_NACION
   Q1 = Q1 & Wh
   Call cbMoneda.FillCombo(DbMain, Q1, -1)
   
   Call SetUpGrid
   Call LoadAll
      
   Call EnableForm(Me, gEmpresa.FCierre = 0)    'OJO: si el año está cerrado, el form queda deshabilitado
   
   Call SetupPriv
   
   If lOper = O_SELECT Then    'tiene que estar después de SetupPriv por el grid.locked
      Bt_Copy.visible = False
      Bt_Valor.visible = False
      Bt_Ok.Caption = "Seleccionar"
      Grid.Locked = True
   End If
   
End Sub

Public Function FEdit(IdMoneda As Long) As Integer
   lIdMoneda = IdMoneda
   lOper = O_EDIT
   Me.Show vbModal
   
End Function

Public Function FSelect(IdMoneda As Long, Valor As Double, Optional ByVal SoloMonedas As Boolean = False) As Integer
   lIdMoneda = IdMoneda
   lOper = O_SELECT
   lSoloMonedas = SoloMonedas
   Me.Show vbModal
   Valor = lValor
   FSelect = lRc
   
End Function
Private Sub SetUpGridFecha()
   Dim F1 As Long, F2 As Long
   Dim Row As Integer
   
   If Val(cbMoneda.Matrix(2)) = MON_VUNICO Then
      Me.Height = 2580
   Else
      Me.Height = lHeight
   End If
   
   Fr_ValMes.visible = Val(cbMoneda.Matrix(2)) <> MON_VUNICO
   Fr_ValUnico.visible = Val(cbMoneda.Matrix(2)) = MON_VUNICO
   
   For Row = Grid.FixedRows To Grid.rows - 1
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
      Grid.rows = Grid.FixedRows
      
      Do While F1 <= F2
         Grid.rows = Row + 1
         
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
      Grid.rows = Grid.FixedRows
      Do While F1 <= 12
         Grid.rows = Row + 1
         
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
      
   Q1 = " SELECT Fecha, Valor FROM Equivalencia WHERE idMoneda=" & cbMoneda.ItemData
   If Val(cbMoneda.Matrix(2)) = MON_VDIA Then
      F1 = DateSerial(Cb_Ano, Cb_Mes.ListIndex + 1, 1)
      Call FirstLastMonthDay(F1, F1, F2)
      Q1 = Q1 & " AND (Fecha BETWEEN " & F1 & " AND " & F2 & ")"
      
   ElseIf Val(cbMoneda.Matrix(2)) = MON_VMES Then
      Q1 = Q1 & " AND " & SqlYearLng("Fecha") & " = " & Cb_Ano
      
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
         lEditNew = O_EDIT
      Else
         lEditNew = O_NEW
      End If
      
   End If
   Call CloseRs(Rs)
'   Call EnableFrm(False)   'deshabilita botón buscar y habilita grilla para ingreso de datos
   Call SetupPriv
End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Trim(Value) <> "" Then
      Value = Format(vFmt(Value), gMonedas(lIdxMoneda).FormatInf)
      
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_ESTADO)
      Cb_Mes.Enabled = False
      Cb_Ano.Enabled = False
      Cb_Moneda.Enabled = False
      Bt_Buscar.Enabled = False
      
   End If
      
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG3_EdType)
 
   EdType = FEG_Edit
End Sub

Private Sub Grid_DblClick()
   If lOper = O_SELECT Then
      Call bt_OK_Click
   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyDec(KeyAscii)
End Sub
Private Sub SetUpGrid()
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESTADO) = 0
   
   Grid.ColWidth(C_DIA) = 1200
   Grid.ColWidth(C_VALOR) = 1200
      
   Grid.ColAlignment(C_DIA) = flexAlignRightCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_DIA) = "Fecha"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call SetUpGridFecha
End Sub
Private Sub EnableFrm(bool As Boolean)
   Grid.Locked = bool
   Bt_Buscar.Enabled = bool
   Bt_Valor.Enabled = Not bool
   Bt_Ok.Enabled = Not bool
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Cb_Moneda)
   End If
End Sub

Private Sub SetupPriv()

   If Not ChkPriv(PRV_CFG_EMP) Or gEmpresa.FCierre <> 0 Then
      EnableFrm (True)
   End If

End Sub
