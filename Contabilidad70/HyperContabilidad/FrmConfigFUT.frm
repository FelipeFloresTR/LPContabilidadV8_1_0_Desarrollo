VERSION 5.00
Begin VB.Form FrmConfigFUT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Cuentas FUT"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Contribuyente"
      Height          =   2655
      Left            =   1380
      TabIndex        =   17
      Top             =   360
      Width           =   3315
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Sociedad Anónima Abierta"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2475
      End
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Sociedad Anónima Cerrada"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2475
      End
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Sociedad por Acción"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Soc. Personas 1ª Categoría"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   2895
      End
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Empresario Individual (EIRL)"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   2895
      End
      Begin VB.OptionButton Op_TipoContrib 
         Caption         =   "Empresario Individual"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   2895
      End
   End
   Begin VB.CommandButton Bt_List 
      Caption         =   "Ver Listado..."
      Height          =   960
      Left            =   6180
      Picture         =   "FrmConfigFUT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   300
      Picture         =   "FrmConfigFUT.frx":05A3
      ScaleHeight     =   690
      ScaleWidth      =   720
      TabIndex        =   15
      Top             =   480
      Width           =   720
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   1380
      TabIndex        =   11
      Top             =   3240
      Width           =   6060
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Ingreso o Gasto:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   450
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   1380
      TabIndex        =   8
      Top             =   4260
      Width           =   6045
      Begin VB.CommandButton Bt_DelCta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         Picture         =   "FrmConfigFUT.frx":0BEC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Borrar la cuenta asociada a item"
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox Tx_Cuenta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton Bt_Cuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Picture         =   "FrmConfigFUT.frx":0FE8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Plan de Cuentas"
         Top             =   1320
         Width           =   315
      End
      Begin VB.ComboBox Cb_Item 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   4575
      End
      Begin VB.ComboBox Cb_Grupo 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox Tx_NoItem 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "(No tiene item asociado)"
         Top             =   780
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1395
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   450
         Width           =   480
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6180
      TabIndex        =   6
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   6180
      TabIndex        =   7
      Top             =   780
      Width           =   1275
   End
End
Attribute VB_Name = "FrmConfigFUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'agregar a SourceSafe

Dim lTContribFUT As Integer

Dim lCuentasFUT() As CuentaFUT_t

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub
Private Sub Bt_Cuentas_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim CodCuenta As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, CodCuenta, Descrip, Nombre, False) = vbOK Then
      Tx_Cuenta = Descrip
      Call ModCuenta(IdCuenta, CodCuenta, Descrip)
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_DelCta_Click()
   
   If Tx_Cuenta = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea borrar la cuenta asociada a este item?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   Call DelCuenta
   Tx_Cuenta = ""
   
End Sub

Private Sub Bt_List_Click()
   Dim Frm As FrmLstConfigFUT
   Dim Q1 As String
   
   If Not Valida() Then
      Exit Sub
   End If
      
   If UBound(lCuentasFUT) <> 0 Or lCuentasFUT(0).IdItemFUT <> 0 Then  'hay cambios no grabados
      If MsgBox1("Antes de ver el listado es necesario grabar los cambios realizados en esta configuración" & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   Call SaveAll
   ReDim lCuentasFUT(0)
   
   Set Frm = New FrmLstConfigFUT
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()
   
   If Valida() Then
      Call SaveAll
      Unload Me
   End If
   
End Sub

Private Sub Cb_Item_Click()
   Call SelCuenta
End Sub

Private Sub Cb_Tipo_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim WhTipo As String
   Dim WhTContrib As String
   
   If Cb_Tipo.ListIndex < 0 Then
      Exit Sub
   End If
   
   If Not Valida() Then
      Exit Sub
   End If
      
   Call GenWhere(WhTipo, WhTContrib)
   
   If gLinkParFUT Then
      Q1 = "SELECT Descripci, IdItem FROM HR_FutGrItems "
      Q1 = Q1 & " WHERE GrpOIte IN ('GRP', 'GIT')"
      Q1 = Q1 & " AND " & WhTipo
      Q1 = Q1 & " AND " & WhTContrib
      Q1 = Q1 & " ORDER BY COrden"
   
      Cb_Grupo.Clear
      
      Call FillCombo(Cb_Grupo, DbMain, Q1, -1)
   End If

End Sub
Private Sub Cb_Grupo_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim WhTipo As String
   Dim InItem As Boolean
   Dim WhTContrib As String
   
   If Cb_Grupo.ListIndex < 0 Then
      Exit Sub
   End If
      
   If Not Valida() Then
      Exit Sub
   End If
      
   Call GenWhere(WhTipo, WhTContrib)
   
   Q1 = "SELECT GrpOIte, Descripci, IdItem FROM HR_FutGrItems "
   Q1 = Q1 & " WHERE IdItem >= '" & Right(String(3, "0") & ItemData(Cb_Grupo), 3) & "'"
   Q1 = Q1 & " AND " & WhTipo
   Q1 = Q1 & " AND " & WhTContrib
   Q1 = Q1 & " ORDER BY COrden"
      
   Set Rs = OpenRs(DbMain, Q1)
   
   Cb_Item.Clear
   
   Do While Rs.EOF = False
   
      If Not InItem And vFld(Rs("GrpOIte")) = "GRP" And Val(vFld(Rs("IdItem"))) = ItemData(Cb_Grupo) Then
         InItem = True
      ElseIf InItem Then
         If vFld(Rs("GrpOIte")) = "ITM" Then
            Call AddItem(Cb_Item, vFld(Rs("Descripci")), Val(vFld(Rs("IdItem"))))
         Else
            InItem = False
            Exit Do
         End If
      End If
      
      Rs.MoveNext
      
   Loop
      
   Call CloseRs(Rs)
   
   If Cb_Item.ListCount > 0 Then
      Cb_Item.ListIndex = 0
      Cb_Item.Visible = True
      Tx_NoItem.Visible = False
   Else
      Cb_Item.Visible = False
      Tx_NoItem.Visible = True
   End If
   
   Call SelCuenta

End Sub

Private Sub Form_Activate()
   
   If Not gLinkParFUT Then
      MsgBox1 "No se ha encontrado la tabla de configuración de ítemes de HR-FUT. Es posible que HR-FUT no esté correctamente instalado en el sistema. No es posible realizar esta operación.", vbExclamation + vbOKOnly
      Call PostClick(Bt_Cancel)
   End If

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   ReDim lCuentasFUT(0)

   Q1 = "SELECT TipoContrib FROM Empresa "
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      lTContribFUT = vFld(Rs("TipoContrib"))
      If lTContribFUT > 0 Then
         Op_TipoContrib(lTContribFUT) = True
      End If
   End If

   Call CloseRs(Rs)
   
   For i = 1 To UBound(gTipoIngGasFUT)
      Cb_Tipo.AddItem gTipoIngGasFUT(i)
      Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
   Next i
   
   Cb_Tipo.ListIndex = 0      'debe estar después del SELECT del tipo de contribuyente

End Sub

Private Sub Op_TContrib_Click(Index As Integer)

End Sub

Private Sub Op_TipoContrib_Click(Index As Integer)
   Static InOpClick As Boolean
   Dim Rc As Integer

   If InOpClick = True Then
      Exit Sub
   End If
   
   InOpClick = True
   
   If lTContribFUT <> 0 And lTContribFUT <> Index Then
   
      Rc = MsgBox1("Si cambia el tipo de contribuyente, es necesario eliminar todas las asignaciones a cuentas hechas previamente, dado que algunos itemes sólo se encuentran disponibles para algunos tipos de contribuyentes." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo)
      If Rc = vbNo Then
         Op_TipoContrib(lTContribFUT) = True
      Else
         ReDim lCuentasFUT(0)
         Call ExecSQL(DbMain, "DELETE * FROM CuentasFUT")
         lTContribFUT = Index
         Tx_Cuenta = ""
      End If
   Else
      lTContribFUT = Index
   End If
   
   If Cb_Tipo.ListCount > 0 Then
      Cb_Tipo_Click
   End If
   
   InOpClick = False
      
End Sub

Private Function Valida() As Boolean

   Valida = False
   
   If lTContribFUT = 0 Then
      If Me.Visible Then
         MsgBox1 "Debe elegir un tipo de contribuyente.", vbExclamation
      End If
      Exit Function
   End If
   
   Valida = True
End Function

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "UPDATE Empresa SET TipoContrib=" & lTContribFUT & ", TContribFUT = " & lTContribFUT
   Call ExecSQL(DbMain, Q1)
   
   For i = 0 To UBound(lCuentasFUT)
      
      If lCuentasFUT(i).IdItemFUT <> 0 Then
      
         If lCuentasFUT(i).id <> 0 Then
            If lCuentasFUT(i).IdCuenta <> 0 Then
               Q1 = "UPDATE CuentasFUT SET "
               Q1 = Q1 & " IdCuenta = " & lCuentasFUT(i).IdCuenta
               Q1 = Q1 & ", CodCuenta = '" & lCuentasFUT(i).CodCuenta & "'"
               Q1 = Q1 & " WHERE Id = " & lCuentasFUT(i).id
            Else
               Q1 = "DELETE * FROM CuentasFUT "
               Q1 = Q1 & " WHERE Id = " & lCuentasFUT(i).id
            End If
         ElseIf lCuentasFUT(i).IdCuenta <> 0 Then
            Q1 = "INSERT INTO CuentasFUT "
            Q1 = Q1 & "(TipoIngGas, IdItem, IdCuenta, CodCuenta)"
            Q1 = Q1 & "VALUES(" & lCuentasFUT(i).TipoIngGas & "," & lCuentasFUT(i).IdItemFUT & "," & lCuentasFUT(i).IdCuenta & ",'" & lCuentasFUT(i).CodCuenta & "')"
         End If
         
         Call ExecSQL(DbMain, Q1)
   
      End If
   Next i

End Sub
Private Sub ModCuenta(ByVal IdCuenta As Long, ByVal CodCuenta As String, ByVal Descrip As String, Optional ByVal DelCta As Boolean = False)
   Dim i As Integer
   Dim idItem As Integer
   Dim Idx As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   idItem = ItemData(Cb_Item)
   
   If idItem < 0 Then
      idItem = ItemData(Cb_Grupo)
   End If
   
   For i = 0 To UBound(lCuentasFUT)
      If lCuentasFUT(i).TipoIngGas = ItemData(Cb_Tipo) And lCuentasFUT(i).IdItemFUT = idItem Then  'ya está en la lista, actualizamos
         If DelCta Then
            lCuentasFUT(i).IdCuenta = 0
            lCuentasFUT(i).CodCuenta = ""
            lCuentasFUT(i).Descrip = ""
         Else
            lCuentasFUT(i).IdCuenta = IdCuenta
            lCuentasFUT(i).CodCuenta = CodCuenta
            lCuentasFUT(i).Descrip = Descrip
         End If
         Idx = i
      End If
   Next i
         
   If Idx = 0 Then   'no lo encontramos en la lista, lo agregamos
      Idx = UBound(lCuentasFUT)
      
      If lCuentasFUT(Idx).IdItemFUT <> 0 Then   'arreglo vacío
         ReDim Preserve lCuentasFUT(Idx + 1)
         Idx = Idx + 1
      End If
      
      lCuentasFUT(Idx).TipoIngGas = ItemData(Cb_Tipo)
      lCuentasFUT(Idx).IdItemFUT = idItem
      If DelCta Then
         lCuentasFUT(Idx).IdCuenta = 0
         lCuentasFUT(Idx).CodCuenta = ""
         lCuentasFUT(Idx).Descrip = ""
      Else
         lCuentasFUT(Idx).IdCuenta = IdCuenta
         lCuentasFUT(Idx).CodCuenta = CodCuenta
         lCuentasFUT(Idx).Descrip = Descrip
      End If
      
      'vemos si está en la base de datos, para asignar el Id y usarlo en el Update
      Q1 = "SELECT Id FROM CuentasFUT "
      Q1 = Q1 & " WHERE TipoIngGas = " & ItemData(Cb_Tipo) & " AND IdItem = " & idItem
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         lCuentasFUT(Idx).id = vFld(Rs("Id"))
      End If
      
      Call CloseRs(Rs)
      
   End If
   
End Sub
Private Sub DelCuenta()

   Call ModCuenta(0, "", "", True)
   
End Sub
Private Sub SelCuenta()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim idItem As Integer
   Dim i As Integer
   
   idItem = ItemData(Cb_Item)
   
   If idItem < 0 Then
      idItem = ItemData(Cb_Grupo)
   End If
   
   Tx_Cuenta = ""
   
   For i = 0 To UBound(lCuentasFUT)
      If lCuentasFUT(i).TipoIngGas = ItemData(Cb_Tipo) And lCuentasFUT(i).IdItemFUT = idItem Then  'ya está en la lista, actualizamos
         Tx_Cuenta = lCuentasFUT(i).Descrip
         Exit Sub
      End If
   Next i
   
   Q1 = "SELECT Descripcion "
   Q1 = Q1 & " FROM Cuentas INNER JOIN CuentasFUT ON Cuentas.IdCuenta = CuentasFUT.IdCuenta "
   Q1 = Q1 & " WHERE TipoIngGas = " & ItemData(Cb_Tipo)
   Q1 = Q1 & " AND CuentasFUT.IdItem = " & idItem
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Tx_Cuenta = vFld(Rs("Descripcion"), True)
   End If
   
   Call CloseRs(Rs)

End Sub
Private Sub GenWhere(WhTipo As String, WhTContrib As String)

   Select Case lTContribFUT
      Case CONTRIB_EMPINDIVIDUALEIRL, CONTRIB_PRIMCAT
         WhTContrib = " SAuOtra IN( 'AMB', 'OTR')"
      Case CONTRIB_EMPINDIVIDUAL
         WhTContrib = " SAuOtra IN( 'AMB')"
      Case CONTRIB_SAABIERTA, CONTRIB_SACERRADA, CONTRIB_SPORACCION
         WhTContrib = " SAuOtra IN( 'AMB', 'SA')"
      
   End Select

   Select Case ItemData(Cb_Tipo)
      Case FUT_AGRPAG
         WhTipo = "AgreDedu = 'AGR' AND PerDev IN('PAG', 'AMB')"
      Case FUT_AGRADE
         WhTipo = "AgreDedu = 'AGR' AND PerDev IN('ADE', 'AMB')"
      Case FUT_DEDPER
         WhTipo = "AgreDedu = 'DED' AND PerDev IN('PER', 'AMB')"
      Case FUT_DEDDEV
         WhTipo = "AgreDedu = 'DED' AND PerDev IN('DEV', 'AMB')"
   End Select
   
End Sub
