VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPlanCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "FrmPlanCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   300
      Picture         =   "FrmPlanCuentas.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   600
      TabIndex        =   34
      Top             =   420
      Width           =   600
   End
   Begin VB.CheckBox Ch_VerNombCorto 
      Caption         =   "Ver nombre corto cuentas último nivel"
      Height          =   255
      Left            =   1380
      TabIndex        =   29
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8760
      TabIndex        =   13
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9960
      TabIndex        =   14
      Top             =   540
      Width           =   1095
   End
   Begin VB.Frame Fr_Search 
      Height          =   615
      Left            =   1260
      TabIndex        =   20
      Top             =   360
      Width           =   7335
      Begin VB.TextBox Tx_Item 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   2
         Top             =   180
         Width           =   3915
      End
      Begin VB.CommandButton Bt_Search 
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
         Left            =   6840
         Picture         =   "FrmPlanCuentas.frx":05AF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar una cuenta"
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox Cb_Buscar 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.TreeView Tr_Plan 
      Height          =   5055
      Left            =   1260
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   512
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   6660
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanCuentas.frx":0932
            Key             =   "CarpetaCerrada"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanCuentas.frx":0CCE
            Key             =   "CarpetaAsterisco"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanCuentas.frx":1074
            Key             =   "CarpetaConMano"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanCuentas.frx":1437
            Key             =   "CarpetaAbierta"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanCuentas.frx":17DC
            Key             =   "CarpetaAbiertaMano"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fr_Sel 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   8760
      TabIndex        =   22
      Top             =   1020
      Width           =   1095
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Imprimir"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":1BA6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Imprimir Cuentas"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":1EB0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Edit 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   8760
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
      Begin VB.CommandButton Bt_SubirCta 
         Height          =   435
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":21BA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Mover cuenta hacia arriba"
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton Bt_BajarCta 
         Height          =   435
         Left            =   660
         Picture         =   "FrmPlanCuentas.frx":25A1
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Mover cuenta hacia abajo"
         Top             =   4260
         Width           =   435
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "&Listar/Imprimir"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":297D
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprimir Cuentas"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Refresh 
         Caption         =   "&Refrescar"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":2FAC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Refrescar cuenta"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":3492
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar cuenta seleccionada"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":3AF4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Modificar cuenta seleccionada"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":40C7
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Nueva cuenta"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_CtasBas 
      Height          =   5115
      Left            =   8625
      TabIndex        =   26
      Top             =   990
      Width           =   2475
      Begin VB.TextBox Tx_TipoLib 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox Tx_TipoValor 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   540
         Width           =   2235
      End
      Begin VB.ListBox Ls_Cuentas 
         Height          =   2595
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2235
      End
      Begin VB.CommandButton Bt_DelCtaLst 
         Caption         =   "Eliminar Cuenta de Lista"
         Height          =   615
         Left            =   180
         Picture         =   "FrmPlanCuentas.frx":4659
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4380
         Width           =   2115
      End
      Begin VB.CommandButton Bt_AddCtaLst 
         Caption         =   "Agregar Cuenta a Lista"
         Height          =   615
         Left            =   180
         Picture         =   "FrmPlanCuentas.frx":4A55
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3660
         Width           =   2115
      End
   End
   Begin VB.Frame Fr_SelEdit 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   8760
      TabIndex        =   31
      Top             =   1080
      Width           =   1095
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":4F23
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":522D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Nueva cuenta"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":5537
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modificar cuenta seleccionada"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":5841
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar cuenta seleccionada"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Refresh 
         Caption         =   "&Refrescar"
         Height          =   800
         Index           =   1
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":5B4B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Refrescar cuenta"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Print 
         Caption         =   "Listar/Imprimir"
         Height          =   800
         Index           =   2
         Left            =   0
         Picture         =   "FrmPlanCuentas.frx":5E55
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir Cuentas"
         Top             =   4200
         Width           =   1095
      End
   End
   Begin VB.Label Lb_CtaOmision 
      AutoSize        =   -1  'True
      Caption         =   "Nota: la primera cuenta en la lista es la que se toma por omisión."
      Height          =   195
      Left            =   6720
      TabIndex        =   30
      Top             =   6360
      Width           =   4530
   End
End
Attribute VB_Name = "FrmPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCuenta As Cuenta_t

'SF 14808914
Dim lCuentaNew As Cuenta_t
'SF 14808914
Dim lRc As Integer
Dim lOper As Integer

Dim lTipoLib As Integer
Dim lTipoVal As Integer

Dim lTblCuentas As String
Dim lNombrePlan As String

Dim lNotSelUltNivel As Boolean

'Dim lBtDelEnabled As BackStyleConstants
Dim lBtDelEnabled As Boolean

Const W_SMALL = 10470
Const W_LARGE = 11610

Const O_SELCTAS = -1


Public Function FSelect(IdCuenta As Long, Codigo As String, Descrip As String, Nombre As String, Optional ByVal SelUltNivel As Boolean = True) As Integer
   lOper = O_SELECT
   
   lNotSelUltNivel = Not SelUltNivel
   
   Me.Show vbModal
   
   'Debo entregar el idCuenta, la descripción y el nombre corto del árbol
   FSelect = lRc
   IdCuenta = lCuenta.id
   CodCuentaSelec = lCuenta.id
   Codigo = lCuenta.Codigo
   Descrip = lCuenta.Descripcion
   Nombre = lCuenta.Nombre
   
End Function

Public Function FSelEdit(IdCuenta As Long, Codigo As String, Descrip As String, Nombre As String, Optional ByVal BtDelEnabled As Boolean = True) As Integer
   lOper = O_SELEDIT
   lBtDelEnabled = BtDelEnabled
   
   Me.Show vbModal
   
   'Debo entregar el idCuenta, la descripción y el nombre corto del árbol
   FSelEdit = lRc
   IdCuenta = lCuenta.id
   CodCuentaSelec = lCuenta.id
   Codigo = lCuenta.Codigo
   Descrip = lCuenta.Descripcion
   Nombre = lCuenta.Nombre
   
End Function
Public Function FEdit(Optional ByVal BtDelEnabled As Boolean = True) As String
   lOper = O_EDIT
   lBtDelEnabled = BtDelEnabled
   
   Me.Show vbModal
   
End Function
Public Function FViewPlan(Optional ByVal PlanCuentas As String = "Cuentas", Optional ByVal NombrePlan As String = "")
   lTblCuentas = PlanCuentas
   lNombrePlan = NombrePlan
   lOper = O_VIEW
   Me.Show vbModal
   
End Function

Private Sub Bt_AddCtaLst_Click()
   Dim i As Integer
   
   lRc = vbCancel
   Call Bt_Sel_Click(0)
   
   If lRc = vbOK Then    'seleccionó una cuenta
   
      'vemos si ya está en la lista
      For i = 0 To Ls_Cuentas.ListCount - 1
         If Ls_Cuentas.ItemData(i) = lCuenta.id Then
            MsgBox1 "Esta cuenta ya está en la lista.", vbExclamation + vbOKOnly
            Exit Sub
         End If
      Next i
      
      'la agregamos
      Ls_Cuentas.AddItem Format(lCuenta.Codigo, gFmtCodigoCta) & " " & lCuenta.Descripcion
      Ls_Cuentas.ItemData(Ls_Cuentas.NewIndex) = lCuenta.id
      Ls_Cuentas.ListIndex = Ls_Cuentas.NewIndex
      
   End If
   
End Sub

Private Sub Bt_BajarCta_Click()
   Dim OldTipo As Integer
   Dim NextNode As Node
   Dim NextNiv As Integer
   Dim NextId As Long
   Dim NextTipo As Integer
   Dim AuxKey As String
   Dim AuxTag As String
   Dim AuxText As String
   Dim Q1 As String
   Dim LenCod As Integer
   Dim PrefixCod1 As String
   Dim PrefixCod2 As String
   Dim i As Integer
   
   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
      
   End If
   
   lCuenta.Nivel = Left(Tr_Plan.SelectedItem.Tag, InStr(Tr_Plan.SelectedItem.Tag, "&") - 1)
   If lCuenta.Nivel = 0 Then
      Exit Sub
   End If
   
'   If lCuenta.Nivel < gLastNivel Then
'      MsgBox1 "Esta funcionalidad sólo está disponible para el último nivel del plan de cuentas.", vbInformation + vbOKOnly
'      Exit Sub
'   End If
      
   lCuenta.Codigo = VFmtCodigoCta(Tr_Plan.SelectedItem.Key)
   LenCod = 0
   
   For i = 1 To lCuenta.Nivel       'gNiveles.nNiveles
      LenCod = LenCod + gNiveles.Largo(i)
   Next i
   
   lCuenta.NivelFather = lCuenta.Nivel - 1
   Call ReadTag(Tr_Plan.SelectedItem.Tag, lCuenta.id, lCuenta.Tipo)
   
   Set NextNode = Tr_Plan.SelectedItem.Next
   
   If NextNode Is Nothing Then
      Exit Sub
   End If
   
   NextNiv = Left(NextNode.Tag, InStr(NextNode.Tag, "&") - 1)
   Call ReadTag(NextNode.Tag, NextId, NextTipo)
      
   If NextNiv = lCuenta.Nivel Then
      AuxKey = NextNode.Key
      AuxTag = NextNode.Tag
      AuxText = NextNode.Text
            
      NextNode.Tag = Tr_Plan.SelectedItem.Tag
      NextNode.Text = ReplaceStr(Tr_Plan.SelectedItem.Text, Tr_Plan.SelectedItem.Key, NextNode.Key)
      
      Tr_Plan.SelectedItem.Tag = AuxTag
      Tr_Plan.SelectedItem.Text = ReplaceStr(AuxText, AuxKey, Tr_Plan.SelectedItem.Key)
      
      
      PrefixCod1 = Left(VFmtCodigoCta(NextNode.Key), LenCod)
      PrefixCod2 = Left(VFmtCodigoCta(Tr_Plan.SelectedItem.Key), LenCod)
     
  'SF 14481830
    If gDbType = SQL_ACCESS Then
        Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'$'", "Codigo") & " WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
 '      Q1 = "UPDATE Cuentas SET Codigo = '$' & Codigo WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
        Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
        Call ExecSQL(DbMain, Q1)

       Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'" & PrefixCod1 & "'", "Right(Codigo, Len(Codigo)- " & LenCod) & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
     ' Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "' & Right(Codigo, Len(Codigo)- " & LenCod & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
       Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
       Call ExecSQL(DbMain, Q1)

       Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'" & PrefixCod2 & "'", "Right(Codigo, Len(Codigo)- 1 - " & LenCod) & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
''     Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "' & Right(Codigo, Len(Codigo)- 1 - " & LenCod & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
       Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
       Call ExecSQL(DbMain, Q1)
      
     Else
     
      Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'$'", "Codigo") & " WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
'      Q1 = "UPDATE Cuentas SET Codigo = '$' & Codigo WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)

       Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "'+ SUBSTRING(Codigo," & LenCod + 1 & ", LEN(Codigo))  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
'      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "' & Right(Codigo, Len(Codigo)- " & LenCod & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "'+SUBSTRING(Codigo," & LenCod + 2 & ", LEN(Codigo))  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
     ''Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "' & Right(Codigo, Len(Codigo)- 1 - " & LenCod & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)
      
      End If
  'SF 14481830
     
'      Q1 = "UPDATE Cuentas SET Codigo = '$" & VFmtCodigoCta(Tr_Plan.SelectedItem.Key) & "' WHERE idCuenta = " & NextId
'      Call ExecSQL(DbMain, Q1)
'
'      Q1 = "UPDATE Cuentas SET Codigo = '" & VFmtCodigoCta(NextNode.Key) & "' WHERE idCuenta = " & lCuenta.Id
'      Call ExecSQL(DbMain, Q1)
'
'      Q1 = "UPDATE Cuentas SET Codigo = '" & VFmtCodigoCta(Tr_Plan.SelectedItem.Key) & "' WHERE idCuenta = " & NextId
'      Call ExecSQL(DbMain, Q1)
      
      Set Tr_Plan.SelectedItem = NextNode
      
      If lCuenta.Nivel < gLastNivel Then    'no es último nivel
         Call Bt_Refresh_Click(0)
      End If
      
      
   Else
      MsgBeep vbExclamation
   End If
   
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
   
End Sub
Private Sub Bt_Del_Click(Index As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Nivel As Byte
   Dim Indice As Integer
   Dim Hijos As Integer
   Dim CodCuenta As String
   
   Nivel = Left(Tr_Plan.SelectedItem.Tag, InStr(Tr_Plan.SelectedItem.Tag, "&") - 1)
   
   If Nivel = 0 Then 'Para q no elimine el texto Plan de Cuentas
      Exit Sub
   End If
   
   Call ReadTag(Tr_Plan.SelectedItem.Tag, lCuenta.id, lCuenta.Tipo)
   
   If Nivel < gLastNivel Then 'And Nivel < (gLastNivel - 1) Then
   
      'vemos si tiene hijos. Si no tiene, podemos eliminarla
      Q1 = "SELECT Count(idPadre) as Hijos FROM Cuentas "
      Q1 = Q1 & " WHERE idPadre=" & lCuenta.id
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      Hijos = vFld(Rs("Hijos"))
      Call CloseRs(Rs)

      If Hijos > 0 Then   'tiene hijos => podría haber movimientos asociados a los hijos
   
         'Si hay al menos un comprobante ingresado, no permitimos eliminar, ya que
         'no podemos verificar que se haga referencia a cuentas que son hijas de esta
         
         Q1 = "SELECT Count(*) as n FROM Comprobante"
         Q1 = Q1 & " INNER JOIN MovComprobante ON Comprobante.idComp = MovComprobante.idComp"
         Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
         Q1 = Q1 & " WHERE Estado<>" & EC_ANULADO
         Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If vFld(Rs("n")) > 0 Then
            MsgBox1 "Para eliminar una cuenta que no es de último nivel, debe antes eliminar las cuentas que están bajo ella.", vbExclamation
            Call CloseRs(Rs)
            Exit Sub
         End If
         Call CloseRs(Rs)
         
         'Si hay al menos un documento ingresado, no permitimos eliminar, ya que
         'no podemos verificar que se haga referencia a cuentas que son hijas de esta
         
         Q1 = "SELECT Count(*) as n FROM Documento"
         Q1 = Q1 & " INNER JOIN MovDocumento ON Documento.IdDoc=MovDocumento.IdDoc"
         Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento")
         Q1 = Q1 & " WHERE Estado<>" & ED_ANULADO
         Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If vFld(Rs("n")) > 0 Then
            MsgBox1 "Para eliminar una cuenta que no es de último nivel, debe antes eliminar las cuentas que están bajo ella.", vbExclamation
            Call CloseRs(Rs)
            Exit Sub
         End If
         Call CloseRs(Rs)

         
      End If
   
   End If
   
   Q1 = "SELECT Count(*) as n FROM MovComprobante WHERE idCuenta=" & lCuenta.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If vFld(Rs("n")) > 0 Then
      MsgBox1 "No se puede eliminar esta cuenta. Existen movimientos asociados a ella.", vbExclamation
      Call CloseRs(Rs)
      Exit Sub
   End If
   Call CloseRs(Rs)
   
   Q1 = "SELECT Count(*) as n FROM MovDocumento WHERE IdCuenta=" & lCuenta.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If vFld(Rs("n")) > 0 Then
      MsgBox1 "No se puede eliminar esta cuenta. Existen documentos que hacen referencia a ella.", vbExclamation
      Call CloseRs(Rs)
      Exit Sub
   End If
   Call CloseRs(Rs)
   
   If Hijos > 0 Then
      Q1 = "¡ ATENCION !, recuerde que al eliminar la cuenta '" & Tr_Plan.SelectedItem.Text & "' se eliminarán todas las cuentas que estén bajo ella. ¿ Desea continuar ?"
   Else
      Q1 = "¿ Está seguro de eliminar la cuenta '" & Tr_Plan.SelectedItem.Text & "' ?"
   End If
   
   If MsgBox1(Q1, vbYesNo Or vbDefaultButton2 Or vbExclamation) <> vbYes Then
      Exit Sub
      
   End If
      
'   Q1 = "DELETE FROM Cuentas WHERE Nivel=" & Nivel & " AND idCuenta=" & lCuenta.Id
'   Q1 = Q1 & " OR idPadre=" & lCuenta.Id
   
   CodCuenta = GetCodCuenta(lCuenta.id)
'   Q1 = "DELETE FROM Cuentas WHERE " & GenWhereCuentas(FmtCodCuenta(CodCuenta))
'   Call ExecSQL(DbMain, Q1)
   
   Q1 = " WHERE " & GenWhereCuentas(FmtCodCuenta(CodCuenta))
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "Cuentas", Q1)
   
   Tr_Plan.Nodes.Remove Tr_Plan.SelectedItem.Index
   
End Sub

Private Sub Bt_DelCtaLst_Click()
   
   If Ls_Cuentas.ListIndex < 0 Then
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar la cuenta '" & Ls_Cuentas.Text & "' de la lista de cuentas definida para el " & GetNombreTipoValLib(lTipoLib, lTipoVal) & " del " & gTipoLib(lTipoLib) & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Ls_Cuentas.RemoveItem Ls_Cuentas.ListIndex
   
End Sub

Private Sub Bt_Edit_Click(Index As Integer)
   Dim Frm As FrmCuenta
   Dim Key As String
   Dim Nd As Node
   Dim Nivel As Integer
   Dim Tx As String
   Dim OldTipo As Integer
   
   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
      
   End If
   
   lCuenta.Nivel = Left(Tr_Plan.SelectedItem.Tag, InStr(Tr_Plan.SelectedItem.Tag, "&") - 1)
   If lCuenta.Nivel = 0 Then
      Exit Sub
   End If
   
   lCuenta.Codigo = VFmtCodigoCta(Tr_Plan.SelectedItem.Key)
   lCuenta.NivelFather = lCuenta.Nivel - 1
   Call ReadTag(Tr_Plan.SelectedItem.Tag, lCuenta.id, lCuenta.Tipo)
   OldTipo = lCuenta.Tipo
   
   Set Frm = New FrmCuenta
   If Frm.FEdit(lCuenta) = vbOK Then

'      If gOrdPlan = ORDPLAN_COD Then
'         Tr_Plan.SelectedItem.Text = Format(lCuenta.Codigo, gFmtCodigoCta) & " " & lCuenta.Descripcion
'      Else
'         Tr_Plan.SelectedItem.Text = lCuenta.Nombre & IIf(lCuenta.Nombre <> "", " - ", " ") & lCuenta.Descripcion
'      End If
      
      If Ch_VerNombCorto = 0 Or lCuenta.Nivel <> gLastNivel Then
         Tr_Plan.SelectedItem.Text = FmtCodCuenta(lCuenta.Codigo) & " " & lCuenta.Descripcion
      ElseIf lCuenta.Nombre <> "" Then
         Tr_Plan.SelectedItem.Text = FmtCodCuenta(lCuenta.Codigo) & " [" & lCuenta.Nombre & "] " & lCuenta.Descripcion
      Else
         Tr_Plan.SelectedItem.Text = FmtCodCuenta(lCuenta.Codigo) & " " & lCuenta.Descripcion
      End If
   
      If lCuenta.Nivel = NIVEL_1 And lCuenta.Tipo <> OldTipo Then   'cambio tipo de cuenta
         'actualizamos el tag y repintamos árbol para que se actualice tag en los hijos
         Tr_Plan.SelectedItem.Tag = lCuenta.Nivel & "&" & lCuenta.id & "&" & lCuenta.Tipo & "&0&"
         Call Bt_Refresh_Click(0)
      End If

   End If
   
End Sub

Private Sub Bt_New_Click(Index As Integer)
   Dim Frm As FrmCuenta
   Dim Key As String
   Dim Nd As Node
   Dim NivelFather As Integer
   Dim Tx As String
   
   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
      
   End If
   
   'OJO CUANDO HAGO UN NEW SIEMPRE Tr_Plan.SelectedItem.Tag
   'SON LOS DATOS DEL PADRE
   NivelFather = Left(Tr_Plan.SelectedItem.Tag, InStr(Tr_Plan.SelectedItem.Tag, "&") - 1)
   
   If NivelFather <> gLastNivel Then
      'LLENO ESTRUCTURA
      lCuenta.Codigo = VFmtCodigoCta(Tr_Plan.SelectedItem.Key)
      lCuenta.NivelFather = NivelFather
      lCuenta.Nivel = NivelFather + 1
      Call ReadTag(Tr_Plan.SelectedItem.Tag, lCuenta.IdPadre, lCuenta.Tipo)
      
      If TieneMovimientos(lCuenta.IdPadre) = True Then
         MsgBox1 "No es posible agregar nuevas cuentas bajo una cuenta que tiene movimientos contables asociados.", vbOKOnly + vbExclamation
         Exit Sub
      End If
      
      Set Frm = New FrmCuenta
      If Frm.FNew(lCuenta) = vbOK Then
         'AGREGO RAMITA
         If Tr_Plan.SelectedItem.Expanded = True Then
            Key = lCuenta.Codigo
            
'            If gOrdPlan = ORDPLAN_COD Then
'               Tx = lCuenta.Codigo & " " & lCuenta.Descripcion
'            Else
'               Tx = lCuenta.Nombre & IIf(Trim(lCuenta.Nombre) <> "", " - ", " ") & lCuenta.Descripcion
'            End If

            If Ch_VerNombCorto = 0 Or lCuenta.Nivel <> gLastNivel Then
               Tx = lCuenta.Codigo & " " & lCuenta.Descripcion
            ElseIf lCuenta.Nombre <> "" Then
               Tx = lCuenta.Codigo & " [" & lCuenta.Nombre & "] " & lCuenta.Descripcion
            Else
               Tx = lCuenta.Codigo & " " & lCuenta.Descripcion
            End If
            
            Set Nd = Tr_Plan.Nodes.Add(Tr_Plan.SelectedItem.Index, tvwChild, Key, Tx)
            Nd.Tag = lCuenta.Nivel & "&" & lCuenta.id & "&" & lCuenta.Tipo & "&"
            Nd.ExpandedImage = "CarpetaAbierta"
            Nd.Image = "CarpetaCerrada"
            
            If lCuenta.Nivel < gLastNivel Then
               Call Tr_Plan.Nodes.Add(Nd.Index, tvwChild, , "*")
               Nd.ExpandedImage = "CarpetaAbierta"
               Nd.Image = "CarpetaCerrada"
               
            End If
            
         Else
            Tr_Plan.SelectedItem.Expanded = True
         End If
         
      End If
      Set Frm = Nothing
      
   Else
      MsgBox1 "Está en el último nivel, no puede crear otra cuenta bajo éste.", vbExclamation
      
   End If
End Sub

Private Sub Bt_OK_Click()

   Call SaveLstCtas
   lRc = vbOK
   
   Unload Me
End Sub

Private Sub Bt_Print_Click(Index As Integer)
   Dim Frm As FrmLstPlanCuentas
   
   Set Frm = New FrmLstPlanCuentas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Refresh_Click(Index As Integer)
   Dim CodCta As String
   Dim SelNode As Node
   
   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   CodCta = VFmtCodigoCta(Tr_Plan.SelectedItem.Key)
   
   Tr_Plan.Nodes.Remove Tr_Plan.Nodes(1).Index
   Call FillRoot
   Call ShowCta(Trim(CodCta), Tr_Plan.Nodes(1), SelNode)
   MousePointer = vbDefault
   
End Sub
Private Sub Bt_Search_Click()
   Dim Q1 As String
   Dim Wh As String, CodCta As String
   Dim Item As Integer, Largo As Integer
   Dim Rs As Recordset
   Dim SelCta As Boolean
   Dim SelNode As Node
   Dim FirstNode As Node
   Dim n As Integer
   
   If Tx_Item = "" Then
      MsgBox1 "Debe ingresar " & Cb_Buscar & ".", vbExclamation
      Cb_Buscar.SetFocus
      Exit Sub
      
   End If
   
   Tr_Plan.visible = False
   
   Item = ItemData(Cb_Buscar)
   If Item = ORDPLAN_COD Then
      CodCta = Replace(Tx_Item, "-", "")
      Largo = Len(Replace(gFmtCodigoCta, "-", "")) 'Completo el largo del codigo
      CodCta = Left(CodCta & String(Largo, "0"), Largo)
      Wh = "Codigo='" & CodCta & "'"
      
   ElseIf Item = ORDPLAN_NOM Then
      Wh = "Nombre='" & ParaSQL(Tx_Item) & "'"
   Else
      Wh = GenLike(DbMain, Tx_Item, "Descripcion")
      
   End If
   
   Q1 = "SELECT Codigo FROM  " & lTblCuentas
   Q1 = Q1 & " WHERE " & Wh
   If lTblCuentas = "Cuentas" Then
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   End If
   Q1 = Q1 & " ORDER BY Codigo "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   n = 0
   
   If Rs.EOF = False Then
      
      Tr_Plan.Nodes.Remove (Tr_Plan.Nodes(1).Index)
      Call FillRoot
      
      SelCta = True
      
      Do While Rs.EOF = False
         CodCta = Trim(vFld(Rs("Codigo")))
         
         'sólo seleccionamos el primero pero los marcamos todos
         Call ShowCta(CodCta, Tr_Plan.Nodes(1), SelNode, False, True)
         If SelCta Then
            Set FirstNode = SelNode
         End If
         SelCta = False
         n = n + 1
         
         Rs.MoveNext
      Loop
      
      If Not FirstNode Is Nothing Then
         FirstNode.Selected = True
         Call FirstNode.EnsureVisible
      End If
      
      Tr_Plan.visible = True
      
      If n = 1 Then
         MsgBox1 "Se encontró una cuenta que calza con la búsqueda realizada.", vbExclamation
      Else
         MsgBox1 "Se encontraron " & n & " cuentas que calzan con la búsqueda realizada.", vbExclamation
      End If
   Else
      Tr_Plan.visible = True
      MsgBox1 "No se encontraron cuentas que calzan con la búsqueda realizada.", vbExclamation
      
   End If
   
   Call CloseRs(Rs)
   
End Sub

Private Sub Bt_Sel_Click(Index As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Nivel As Integer
   Dim Nd As Node

   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   Set Nd = Tr_Plan.SelectedItem
   
   Call ReadTag(Nd.Tag, lCuenta.id, lCuenta.Tipo)
   
   If lCuenta.id > 0 Then
   
      Nivel = Left(Nd.Tag, InStr(Nd.Tag, "&") - 1)
      
      If Nivel <> gLastNivel And lNotSelUltNivel = False Then
         MsgBox1 "Debe seleccionar una cuenta de último nivel.", vbExclamation
         Exit Sub
      End If
            
      'vemos si está activa
      
      Q1 = "SELECT Estado "
      Q1 = Q1 & " FROM " & lTblCuentas
      Q1 = Q1 & " WHERE IdCuenta = " & lCuenta.id
      If lTblCuentas = "Cuentas" Then
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      End If
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = False Then
         If vFld(Rs("Estado")) = 0 Then
            MsgBox1 ("Esta cuenta no está activa.")
            Call CloseRs(Rs)
            Exit Sub
         End If
      End If

      Call CloseRs(Rs)
   
      Call TxCuentas(lCuenta.id, lCuenta.Codigo, lCuenta.Descripcion, lCuenta.Nombre)
      
      lRc = vbOK
      
      If Bt_Sel(Index).visible = True Then
         Unload Me
      End If
   End If
      
End Sub

Private Sub Bt_SubirCta_Click()
   Dim OldTipo As Integer
   Dim PrevNode As Node
   Dim PrevNiv As Integer
   Dim PrevId As Long
   Dim PrevTipo As Integer
   Dim AuxKey As String
   Dim AuxTag As String
   Dim AuxText As String
   Dim Q1 As String
   Dim LenCod As Integer
   Dim PrefixCod1 As String
   Dim PrefixCod2 As String
   Dim i As Integer
   
   If Tr_Plan.SelectedItem Is Nothing Then
      Exit Sub
      
   End If
   
   lCuenta.Nivel = Left(Tr_Plan.SelectedItem.Tag, InStr(Tr_Plan.SelectedItem.Tag, "&") - 1)
   If lCuenta.Nivel = 0 Then
      Exit Sub
   End If
   
'   If lCuenta.Nivel < gLastNivel Then
'      MsgBox1 "Esta funcionalidad sólo está disponible para el último nivel del plan de cuentas.", vbInformation + vbOKOnly
'      Exit Sub
'   End If
   
   lCuenta.Codigo = VFmtCodigoCta(Tr_Plan.SelectedItem.Key)
   
   LenCod = 0
   
   For i = 1 To lCuenta.Nivel       'gNiveles.nNiveles
      LenCod = LenCod + gNiveles.Largo(i)
   Next i
      
   lCuenta.NivelFather = lCuenta.Nivel - 1
   Call ReadTag(Tr_Plan.SelectedItem.Tag, lCuenta.id, lCuenta.Tipo)
   
   Set PrevNode = Tr_Plan.SelectedItem.Previous
   
   If PrevNode Is Nothing Then
      Exit Sub
   End If
   
   PrevNiv = Left(PrevNode.Tag, InStr(PrevNode.Tag, "&") - 1)
   Call ReadTag(PrevNode.Tag, PrevId, PrevTipo)
   
   If PrevNiv = lCuenta.Nivel Then
      AuxKey = PrevNode.Key
      AuxTag = PrevNode.Tag
      AuxText = PrevNode.Text
            
      PrevNode.Tag = Tr_Plan.SelectedItem.Tag
      PrevNode.Text = ReplaceStr(Tr_Plan.SelectedItem.Text, Tr_Plan.SelectedItem.Key, PrevNode.Key)
      
      Tr_Plan.SelectedItem.Tag = AuxTag
      Tr_Plan.SelectedItem.Text = ReplaceStr(AuxText, AuxKey, Tr_Plan.SelectedItem.Key)
      
      PrefixCod1 = Left(VFmtCodigoCta(PrevNode.Key), LenCod)
      PrefixCod2 = Left(VFmtCodigoCta(Tr_Plan.SelectedItem.Key), LenCod)
   'SF 14481830
    If gDbType = SQL_ACCESS Then
      
              Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'$'", "Codigo") & " WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
              Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
              Call ExecSQL(DbMain, Q1)
              
        '      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "' & Right(Codigo, Len(Codigo)- " & LenCod & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
              Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'" & PrefixCod1 & "'", "Right(Codigo, Len(Codigo)- " & LenCod) & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
              Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
              Call ExecSQL(DbMain, Q1)
              
        '      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "' & Right(Codigo, Len(Codigo)- 1 - " & LenCod & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
              Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'" & PrefixCod2 & "'", "Right(Codigo, Len(Codigo)- 1 - " & LenCod) & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
              Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
              Call ExecSQL(DbMain, Q1)
     Else
            
     Q1 = "UPDATE Cuentas SET Codigo = " & SqlConcat(gDbType, "'$'", "Codigo") & " WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
'      Q1 = "UPDATE Cuentas SET Codigo = '$' & Codigo WHERE Left(Codigo," & LenCod & ") = '" & PrefixCod1 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)

       Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "'+ SUBSTRING(Codigo," & LenCod + 1 & ", LEN(Codigo))  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
'      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod1 & "' & Right(Codigo, Len(Codigo)- " & LenCod & ")  WHERE Left(Codigo, " & LenCod & ") = '" & PrefixCod2 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "'+SUBSTRING(Codigo," & LenCod + 2 & ", LEN(Codigo))  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
     ''Q1 = "UPDATE Cuentas SET Codigo = '" & PrefixCod2 & "' & Right(Codigo, Len(Codigo)- 1 - " & LenCod & ")  WHERE Left(Codigo, " & LenCod + 1 & ") = '$" & PrefixCod1 & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Call ExecSQL(DbMain, Q1)
     
     
     End If
'SF 14481830
              
'      Q1 = "UPDATE Cuentas SET Codigo = '$" & VFmtCodigoCta(Tr_Plan.SelectedItem.Key) & "' WHERE idCuenta = " & PrevId
'      Call ExecSQL(DbMain, Q1)
'
'      Q1 = "UPDATE Cuentas SET Codigo = '" & VFmtCodigoCta(PrevNode.Key) & "' WHERE idCuenta = " & lCuenta.Id
'      Call ExecSQL(DbMain, Q1)
'
'      Q1 = "UPDATE Cuentas SET Codigo = '" & VFmtCodigoCta(Tr_Plan.SelectedItem.Key) & "' WHERE idCuenta = " & PrevId
'      Call ExecSQL(DbMain, Q1)
      
      Set Tr_Plan.SelectedItem = PrevNode
      
      If lCuenta.Nivel < gLastNivel Then    'no es último nivel
         Call Bt_Refresh_Click(0)
      End If
            
   Else
      MsgBeep vbExclamation
   End If
   
End Sub

Private Sub Cb_Buscar_Click()
   Tx_Item = ""
   Select Case ItemData(Cb_Buscar)
      Case ORDPLAN_COD
         Tx_Item.MaxLength = 15
      Case ORDPLAN_NOM
         Tx_Item.MaxLength = 10
      Case ORDPLAN_NOM
         Tx_Item.MaxLength = 50
    End Select
    
End Sub

Private Sub Ch_VerNombCorto_Click()
   
   Call SetIniString(gIniFile, "Config", "VerNombreCorto", Abs(Ch_VerNombCorto.Value))

   'cargamos de nuevo el árbol
   Call FillRoot
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   If lTblCuentas = "" Then
      lTblCuentas = "Cuentas"
   End If
   

   Call EnableForm(Me, gEmpresa.FCierre = 0)
   Fr_Search.Enabled = True
   Bt_Search.Enabled = True
   Cb_Buscar.Locked = False
   Call SetTxRO(Tx_Item, False)
   
   Fr_Edit.visible = (lOper = O_EDIT)
   Fr_Sel.visible = (lOper = O_SELECT Or lOper = O_VIEW)
   'Fr_Search.Visible = (lOper = O_VIEW)
   Fr_CtasBas.visible = (lOper = O_SELCTAS)
   Fr_SelEdit.visible = (lOper = O_SELEDIT)
   
   If lOper <> O_SELCTAS Then
      Bt_OK.visible = False
      Bt_Cancel.Left = Bt_OK.Left
   End If
   
   If lOper = O_VIEW Then
      Bt_Sel(0).visible = False
      Bt_Print(1).Top = Bt_Sel(0).Top
   End If
      
   'If lOper = O_EDIT Or lOper = O_SELCTAS Then
   '   Tr_Plan.Top = Bt_Cancel.Top
   '   Tr_Plan.Height = 5595
   'End If
   
   'SF 14481830
   If gDbType = SQL_SERVER Then
     Call DeleteSignoPeso
   End If
   'SF 14481830
   
      
   If lOper = O_SELCTAS Then
      Me.Width = W_LARGE
      Tx_TipoLib = gTipoLib(lTipoLib)
      Tx_TipoValor = " Cuentas " & GetNombreTipoValLib(lTipoLib, lTipoVal)
      
      Call SetTxRO(Tx_TipoLib, True)
      Call SetTxRO(Tx_TipoValor, True)
      Bt_Cancel.Caption = "Cancelar"
      Call LoadLstCtas
      
      Me.Caption = "Cuentas Definidas para " & gTipoLib(lTipoLib) & " - " & GetNombreTipoValLib(lTipoLib, lTipoVal)
   
   Else
   
      Me.Width = W_SMALL
      Lb_CtaOmision.visible = False
      
      If lOper = O_VIEW Then
         If lTblCuentas <> "Cuentas" Then
            Me.Caption = Me.Caption & " " & lNombrePlan
         End If
      
      Else
         Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            If LCase(vFld(Rs("Valor"))) = "ifrs" Then
               Me.Caption = Me.Caption & " " & UCase(vFld(Rs("Valor")))
            Else
               Me.Caption = Me.Caption & " " & FCase(vFld(Rs("Valor")))
            End If
         End If
         
         Call CloseRs(Rs)
      End If
      
   End If
   
   Ch_VerNombCorto.Value = Abs(Val(GetIniString(gIniFile, "Config", "VerNombreCorto", "1")) <> 0)
   
   Call FillRoot
   
   For i = ORDPLAN_DESC To ORDPLAN_COD Step -1
      Cb_Buscar.AddItem gFindPlan(i)
      Cb_Buscar.ItemData(Cb_Buscar.NewIndex) = i
   Next i
   Cb_Buscar.ListIndex = 0
   
   Call SetupPriv
   
End Sub
Private Sub FillRoot()
   Dim Nd As Node
   
   Tr_Plan.Nodes.Clear
   
   'ESTO ES FIJO TITULO PLAN DE CUENTA
   Set Nd = Tr_Plan.Nodes.Add(, , "*", "Plan de Cuentas")
   Nd.ExpandedImage = "CarpetaAbierta"
   Nd.Image = "CarpetaCerrada"
   Nd.Tag = "0&0&0&"
   
   'CREO HIJO
   Set Nd = Tr_Plan.Nodes.Add(Nd.Index, tvwChild, , "*")
   Nd.ExpandedImage = "CarpetaAbierta"
   Nd.Image = "CarpetaCerrada"
   
   'expandimos el primer nodo "Plan de cuentas" para que se vea el primer nivel al entrar
   Tr_Plan.Nodes.Item(1).Expanded = True
   
End Sub

Private Sub Tr_Plan_DblClick()
   
   If lOper = O_SELECT Then
      Call PostClick(Bt_Sel(0))
   ElseIf lOper = O_SELCTAS Then
      Call PostClick(Bt_AddCtaLst)
   ElseIf lOper = O_EDIT Then
      Call PostClick(Bt_Edit(0))
   ElseIf lOper = O_SELEDIT Then
      Call PostClick(Bt_Sel(1))
   End If
   
End Sub

Private Sub Tr_Plan_Expand(ByVal Node As MSComctlLib.Node)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Nd As Node
   Dim Key As String
   Dim Nivel As Integer
   Dim IdPadre As Long
   Dim Tipo As Integer
   Dim Tx As String
   
   'SF 14808914
   Dim Q2 As String
   Dim Rs2 As Recordset
   'SF 14808914
   
   Nivel = Left(Node.Tag, InStr(Node.Tag, "&") - 1)
   Call ReadTag(Node.Tag, IdPadre, Tipo)
   
   If Node.Child.Text <> "*" Or Nivel >= gLastNivel Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
      

   Q1 = "SELECT Codigo, Nombre, Descripcion, idCuenta, Nivel, Clasificacion"
   Q1 = Q1 & " FROM  " & lTblCuentas
   Q1 = Q1 & " WHERE Nivel = " & Nivel + 1 & " AND IdPadre = " & IdPadre
   If lTblCuentas = "Cuentas" Then
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   End If
   
   Q1 = Q1 & " ORDER BY " & ReplaceStr(gFldOrdPlan(gOrdPlan), "Cuentas.", "") & ",idCuenta"
   Set Rs = OpenRs(DbMain, Q1)
   
   'SF 14808914
   If gDbType = SQL_ACCESS Then
   Q2 = "SELECT len(trim(max(Codigo)))  as dig "
   Else
   Q2 = "SELECT len(Ltrim(Rtrim(max(Codigo))))  as dig "
   End If
   Q2 = Q2 & " FROM  " & lTblCuentas
   Q2 = Q2 & " WHERE Nivel = " & 1 & " AND IdPadre = 0"
   If lTblCuentas = "Cuentas" Then
      Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   End If
'
   Set Rs2 = OpenRs(DbMain, Q2)
'
   If Rs2.EOF = False Then
   Dim canDig As Integer

   canDig = vFld(Rs2("dig"))

   End If

   Call CloseRs(Rs2)

   Q2 = "SELECT max(Nivel)  as MaxNiv "
   Q2 = Q2 & " FROM  " & lTblCuentas
   
   If lTblCuentas = "Cuentas" Then
      Q2 = Q2 & " WHERE "
      Q2 = Q2 & " IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   End If

   Set Rs2 = OpenRs(DbMain, Q2)

   If Rs2.EOF = False Then
   Dim NivelMax As Integer

   NivelMax = vFld(Rs2("MaxNiv"))

   End If
'
   Call CloseRs(Rs2)
   
   'SF 14808914
   
   Do Until Rs.EOF
   'SF 14481830
   Dim vCodigo As String
   
   If Nivel + 1 < NivelMax Then
   vCodigo = Trim(vFld(Rs("Codigo"))) & String(canDig - Len(Trim(vFld(Rs("Codigo")))), "0")
   Else
     If Trim(vFld(Rs("Codigo"))) < GetCodCuenta(IdPadre) Then
         vCodigo = 0
     Else
       vCodigo = Trim(vFld(Rs("Codigo"))) & String(canDig - Len(Trim(vFld(Rs("Codigo")))), "0")
     End If
   End If

   'SF 14481830
    'Key = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
    Key = Format(vCodigo, gFmtCodigoCta)
   'SF 14481830
   
      If gOrdPlan = ORDPLAN_COD Then
      '14481830
         'Tx = Format(vFld(Rs("Codigo")), gFmtCodigoCta) & " " & vFld(Rs("Descripcion"), True)
          Tx = Format(vCodigo, gFmtCodigoCta) & " " & vFld(Rs("Descripcion"), True)
       '14481830
      Else
         Tx = vFld(Rs("Nombre")) & IIf(Trim(vFld(Rs("Nombre"))) <> "", " - ", " ") & vFld(Rs("Descripcion"), True)
      End If
      
      
      If Ch_VerNombCorto = 0 Or vFld(Rs("Nivel")) <> gLastNivel Then
      ''SF 14481830
         'Tx = FmtCodCuenta(vFld(Rs("Codigo"))) & " " & vFld(Rs("Descripcion"), True)
         Tx = FmtCodCuenta(vCodigo) & " " & vFld(Rs("Descripcion"), True)
       ''SF 14481830
      ElseIf vFld(Rs("Nombre")) <> "" Then
      'SF 14481830
         'Tx = FmtCodCuenta(vFld(Rs("Codigo"))) & " [" & vFld(Rs("Nombre"), True) & "] " & vFld(Rs("Descripcion"), True)
         Tx = FmtCodCuenta(vCodigo) & " [" & vFld(Rs("Nombre"), True) & "] " & vFld(Rs("Descripcion"), True)
      'SF 14481830
      Else
      'SF 14481830
         'Tx = FmtCodCuenta(vFld(Rs("Codigo"))) & " " & vFld(Rs("Descripcion"), True)
         Tx = FmtCodCuenta(vCodigo) & " " & vFld(Rs("Descripcion"), True)
       'SF 14481830
      End If
      
       'SF 14808914
      If VFmtCodigoCta(vCodigo) > 0 Then

          If gDbType = SQL_SERVER Then
            If Len(Trim(vFld(Rs("Codigo")))) < canDig Then
             Call UpdateCuentasMenosCaracteres(vCodigo, vFld(Rs("Codigo")))
            End If
          End If
      'SF 14808914
      
      Set Nd = Tr_Plan.Nodes.Add(Node.Index, tvwChild, Key, Tx)
       
      Nd.Tag = vFld(Rs("Nivel")) & "&" & vFld(Rs("idCuenta")) & "&" & vFld(Rs("Clasificacion")) & "&0&"
      Nd.ExpandedImage = "CarpetaAbierta"
      Nd.Image = "CarpetaCerrada"
      
      End If
      
      If vFld(Rs("Nivel")) < gLastNivel Then
         Call Tr_Plan.Nodes.Add(Nd.Index, tvwChild, , "*")
         Nd.ExpandedImage = "CarpetaAbierta"
         Nd.Image = "CarpetaCerrada"
         
      End If
      
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
     
   Tr_Plan.Nodes.Remove Node.Child.Index
   MousePointer = vbDefault
   
End Sub
'LEO EL TAG QUE TIENE EL NIVEL,IDPADRE y CLASIFICACION O TIPO
'VIENE POR EJEMPLO: 1&1&1&
Private Sub ReadTag(ByVal TagTree As String, IdPadre As Long, Tipo As Integer)
   Dim St As String
   
   St = TagTree
   St = Mid(St, InStr(St, "&") + 1)
   IdPadre = Left(St, InStr(St, "&") - 1)
   St = Mid(St, InStr(St, "&") + 1)
   Tipo = Left(St, InStr(St, "&") - 1)
   
End Sub
Private Function ShowCta(cKey As String, Nd As Node, SelNode As Node, Optional ByVal SelCta As Boolean = True, Optional ByVal MarkCta As Boolean = False) As Boolean
   Dim Key As String
   Dim Nivel As Integer
   Dim Nd1 As Node
   Dim i As Integer
   Dim Largo As Integer
   
   ShowCta = False
             
   If Nd Is Nothing Then
      Set Nd = Nothing
      Exit Function
      
   End If
   
   Nd.Expanded = True
   
   Set Nd1 = Nd.Child
   Do While Not (Nd1 Is Nothing)
   
      Key = VFmtCodigoCta(Nd1.Key)
      
      Nivel = Left(Nd1.Tag, InStr(Nd1.Tag, "&") - 1)
     
      If Key = cKey Then
      
         Set SelNode = Nd1
         
         If SelCta Then
            Nd1.Selected = True
            Nd1.EnsureVisible
         End If
         
         If MarkCta And Nivel = gLastNivel Then
            Nd1.ForeColor = vbBlue
         End If
         
         Set Nd1 = Nothing
         ShowCta = True
         Exit Function
         
      End If

      Largo = gNiveles.Inicio(Nivel) + gNiveles.Largo(Nivel) - 1
      
      'For i = 1 To Nivel
      '   Largo = Largo + gNiveles.Largo(i)
      'Next i
      
      If Left(Key, Largo) = Left(cKey, Largo) Then
         If ShowCta(cKey, Nd1, SelNode, SelCta, MarkCta) Then
            Exit Function
         End If
         
      End If

      Set Nd1 = Nd1.Next
   Loop
   
End Function

Private Sub Tx_Item_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Bt_Search_Click
      Exit Sub
   End If
   
   If ItemData(Cb_Buscar) = ORDPLAN_COD Then
      Call KeyCodCta(KeyAscii)
   ElseIf ItemData(Cb_Buscar) = ORDPLAN_NOM Then
      Call KeyUpper(KeyAscii)
   End If
   
End Sub
'Obtengo el nombre corto y descripcion
Private Sub TxCuentas(ByVal id As Long, Codigo As String, Descripcion As String, NombreCorto As String)
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Codigo, Nombre, Descripcion FROM  " & lTblCuentas
   Q1 = Q1 & " WHERE idCuenta=" & id
   If lTblCuentas = "Cuentas" Then
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   End If
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Codigo = vFld(Rs("Codigo"))
      NombreCorto = vFld(Rs("Nombre"), True)
      Descripcion = vFld(Rs("Descripcion"), True)
   End If
   Call CloseRs(Rs)
   
End Sub
Private Function TieneMovimientos(ByVal IdCuenta As Long) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT IdMov FROM MovComprobante WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   TieneMovimientos = Not Rs.EOF
   
   Call CloseRs(Rs)

End Function

Public Function FSelCuentasBasicas(ByVal TipoLib As Integer, ByVal TipoVal As Integer)
   
   lTipoLib = TipoLib
   lTipoVal = TipoVal
   
   lOper = O_SELCTAS
   
   Me.Show vbModal

End Function

Private Sub LoadLstCtas()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion "
   Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
   Q1 = Q1 & " WHERE Tipo = 0 AND TipoLib = " & lTipoLib & " AND TipoValor = " & lTipoVal
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY CuentasBasicas.Id"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Ls_Cuentas.AddItem Format(vFld(Rs("Codigo")), gFmtCodigoCta) & " " & vFld(Rs("Descripcion"), True)
      Ls_Cuentas.ItemData(Ls_Cuentas.NewIndex) = vFld(Rs("IdCuenta"))
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

End Sub

Private Sub SaveLstCtas()
   Dim Q1 As String
   Dim i As Integer
   
'   Q1 = "DELETE * FROM CuentasBasicas WHERE Tipo=0 AND TipoLib=" & lTipoLib & " AND TipoValor=" & lTipoVal
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE Tipo=0 AND TipoLib=" & lTipoLib & " AND TipoValor=" & lTipoVal
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "CuentasBasicas", Q1)
   
   For i = 0 To Ls_Cuentas.ListCount - 1
      Q1 = "INSERT INTO CuentasBasicas "
      Q1 = Q1 & "(Tipo, TipoLib, TipoValor, IdCuenta, IdEmpresa, Ano)"
      Q1 = Q1 & " VALUES( 0, " & lTipoLib & "," & lTipoVal & "," & Ls_Cuentas.ItemData(i) & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
      
      Call ExecSQL(DbMain, Q1)
      
   Next i

End Sub
Private Sub SetupPriv()
   Dim i As Integer
   Dim bool As Boolean

   If ChkPriv(PRV_ADM_CTAS) And gEmpresa.FCierre = 0 Then
      bool = True
   End If
   
   For i = 0 To 1
      Bt_New(i).Enabled = bool
      If bool = False Then
         Bt_Edit(i).Caption = "Ver"
         Bt_Edit(i).Enabled = True
      End If
      
      If lBtDelEnabled Then
         Bt_Del(i).Enabled = bool
      Else
         Bt_Del(i).Enabled = False
      End If
      
   Next i
   
   Bt_AddCtaLst.Enabled = bool
   Bt_DelCtaLst.Enabled = bool

End Sub

'SF 14481830 solo se utilizara en sql server
Private Sub DeleteSignoPeso()
   Dim Q1 As String
   Dim Rs As Recordset, Rs2 As Recordset

   Q1 = ""
   Q1 = "select codigo from cuentas "
   Q1 = Q1 & " where codigo is not null and codigo like '%$%'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
'
   Do While Rs.EOF = False
       
           Q1 = ""
           Q1 = "select codigo from cuentas "
           Q1 = Q1 & " where codigo is not null and codigo = '" & Replace(Trim(vFld(Rs("codigo"))), "$", "") & "'"
           Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
        
           Set Rs2 = OpenRs(DbMain, Q1)
        '
           Do While Rs2.EOF = False
              
                Q1 = " "
                Q1 = " WHERE codigo = '" & Trim(vFld(Rs2("codigo"))) & "'"
                Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                Call DeleteSQL(DbMain, "cuentas", Q1)
                
            Rs2.MoveNext
           Loop
            
           Call CloseRs(Rs2)
  
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

      Q1 = ""
      Q1 = "UPDATE Cuentas "
      Q1 = Q1 & " set codigo = SUBSTRING(codigo,2,len(codigo)) "
      Q1 = Q1 & " Where IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " and codigo like '%$%' "

      Call ExecSQL(DbMain, Q1)
       
'                Q1 = " "
'                Q1 = " WHERE idCuenta = '3611'"
'                Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'                Call DeleteSQL(DbMain, "cuentas", Q1)
'
End Sub
'SF 14481830 solo se utilizara en sql server


'SF 14808914 solo se utilizara en sql server
Private Sub UpdateCuentasMenosCaracteres(ByVal vCuentaSet As String, ByVal vCuentaWhere As String)
   Dim Q1 As String
   Dim Rs As Recordset, Rs2 As Recordset

      Q1 = ""
      Q1 = "UPDATE Cuentas "
      Q1 = Q1 & " set codigo = '" & vCuentaSet & "'"
      Q1 = Q1 & " Where IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " and codigo = '" & vCuentaWhere & "' "

      Call ExecSQL(DbMain, Q1)
       
End Sub
'SF 14808914 solo se utilizara en sql server

'SF 14808914
'Private Sub EliminaCuentasErroneas(ByVal vIdCuenta As Long)
'
'   Q1 = "SELECT Count(*) as n FROM MovComprobante WHERE idCuenta=" & lCuenta.id
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Set Rs = OpenRs(DbMain, Q1)
'   If vFld(Rs("n")) > 0 Then
'      MsgBox1 "No se puede eliminar esta cuenta. Existen movimientos asociados a ella.", vbExclamation
'      Call CloseRs(Rs)
'      Exit Sub
'   End If
'   Call CloseRs(Rs)
'
'   Q1 = "SELECT Count(*) as n FROM MovDocumento WHERE IdCuenta=" & lCuenta.id
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Set Rs = OpenRs(DbMain, Q1)
'   If vFld(Rs("n")) > 0 Then
'      MsgBox1 "No se puede eliminar esta cuenta. Existen documentos que hacen referencia a ella.", vbExclamation
'      Call CloseRs(Rs)
'      Exit Sub
'   End If
'   Call CloseRs(Rs)
'
'End Sub
'SF 14808914
