VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConfigCtasAjustesRLI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración  Cuentas Ajustes Extra - Contables RLI HR RAB"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   300
      Picture         =   "FrmConfigCtasAjustesRLI.frx":0000
      ScaleHeight     =   600
      ScaleWidth      =   630
      TabIndex        =   9
      Top             =   420
      Width           =   690
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   420
      Width           =   1035
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      Top             =   840
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignación de Cuentas"
      Height          =   6555
      Left            =   1140
      TabIndex        =   8
      Top             =   300
      Width           =   6735
      Begin VB.ComboBox Cb_Grupo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1020
         Width           =   4875
      End
      Begin VB.CommandButton Bt_Del 
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
         Left            =   6120
         Picture         =   "FrmConfigCtasAjustesRLI.frx":0733
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar documento seleccionado"
         Top             =   1980
         Width           =   375
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
         Height          =   375
         Left            =   5640
         Picture         =   "FrmConfigCtasAjustesRLI.frx":0B2F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Plan de Cuentas"
         Top             =   1980
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3855
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   6800
         _Version        =   393216
      End
      Begin VB.ComboBox Cb_Item 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1500
         Width           =   4875
      End
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   4875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Cuentas:"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Ajuste:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   600
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmConfigCtasAjustesRLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDCUENTA = 0
Const C_CODCUENTA = 1
Const C_CUENTA = 2
Const C_SELCTA = 3
Const C_IDCTAAJUSTES = 4
Const C_UPD = 5

Const NCOLS = C_UPD

Dim lTipoAjuste As Integer
Dim lIdGrupo As Integer
Dim lIdItem As Integer
Dim lInLoad As Boolean
Dim lModCuentas As Boolean

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Cuentas_Click()
   Dim Row As Integer
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call Grid_DblClick

End Sub
Private Sub Cb_Grupo_Click()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   If Not lInLoad And lModCuentas Then
      If MsgBox1("¿Desea grabar los cambios realizados a las cuentas asociadas a este ítem?", vbQuestion + vbYesNo) = vbYes Then
         If valida() Then
            SaveAll
         End If
      End If
   
   End If
   
   lIdGrupo = CbItemData(Cb_Grupo)
   
   Cb_Item.Clear
   i = 1
   For j = 1 To MAX_ITEMAJUSTESECRLI
      For i = 1 To MAX_ITEMAJUSTESECRLI
         If gAjustesExtraContRLI(lTipoAjuste, lIdGrupo, i).Nombre <> "" And gAjustesExtraContRLI(lTipoAjuste, lIdGrupo, i).orden = j Then
            
      '      If gAjustesExtraContRli(lTipoAjuste, i).TipoIngresoAjuste = TIA_CTASASOCIADAS Then
               Call CbAddItem(Cb_Item, gAjustesExtraContRLI(lTipoAjuste, lIdGrupo, i).Nombre, i)
      '      End If
         
         End If
         
      Next i
   Next j

   Cb_Item.ListIndex = 0

   lIdItem = CbItemData(Cb_Item)
   
   Call LoadAll

End Sub

Private Sub Cb_Item_Click()
   
   If Not lInLoad And lModCuentas Then
      If MsgBox1("¿Desea grabar los cambios realizados a las cuentas asociadas a este ítem?", vbQuestion + vbYesNo) = vbYes Then
         If valida() Then
            SaveAll
         End If
      End If
   End If
   
   lIdItem = CbItemData(Cb_Item)
      
   Call LoadAll
End Sub

Private Sub Grid_DblClick()
   Dim FrmPlan As FrmPlanCuentas
   Dim DescCta As String
   Dim CodCta As String
   Dim NombCuenta As String
   Dim Row As Integer, i As Integer
   Dim IdCuenta As Long
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Row > Grid.FixedRows And Val(Grid.TextMatrix(Row - 1, C_IDCUENTA)) = 0 Then
      Exit Sub
   End If
   
   Set FrmPlan = New FrmPlanCuentas

   If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta, True) = vbOK Then
      If DescCta <> "" Then
      
         For i = Grid.FixedRows To Grid.rows - 1
            If Grid.TextMatrix(i, C_IDCUENTA) = "" Then
               Exit For
            End If
            If Val(Grid.TextMatrix(i, C_IDCUENTA)) = IdCuenta And Row <> i Then
               MsgBox1 "Esta cuenta ya ha sido seleccionada para este ítem.", vbExclamation
               Exit Sub
            End If
         Next i
         
         Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
         Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
         Grid.TextMatrix(Row, C_CUENTA) = DescCta
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(Row + 1, C_SELCTA) = ">>"

         Call FGrModRow(Grid, Row, FGR_U, C_IDCTAAJUSTES)
         
         lModCuentas = True
         
     End If

   End If
   Set FrmPlan = Nothing

End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer

   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_CUENTA) <> "" Then
      If MsgBox1("¿Está seguro que desea eliminar esta cuenta asociada al ítem seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDCTAAJUSTES, C_UPD)
   
   lModCuentas = True

End Sub

Private Sub Bt_OK_Click()
   If valida() Then
      SaveAll
      Unload Me
   End If
End Sub

Private Sub Cb_TipoAjuste_Click()
   Dim i As Integer
   
   If Not lInLoad And lModCuentas Then
      If MsgBox1("¿Desea grabar los cambios realizados a las cuentas asociadas a este ítem?", vbQuestion + vbYesNo) = vbYes Then
         If valida() Then
            SaveAll
         End If
      End If
   
   End If
   
   lTipoAjuste = CbItemData(Cb_TipoAjuste)
   
   Cb_Grupo.Clear
   i = 1
   For i = 1 To MAX_GRUPOAJUSTESECRLI
      If gGrupoAjustesECRLI(lTipoAjuste, i) = "" Then
         Exit For
      End If
      
      Call CbAddItem(Cb_Grupo, gGrupoAjustesECRLI(lTipoAjuste, i), i)
      
   Next i
   
   Cb_Grupo.ListIndex = 0
   
   lIdGrupo = CbItemData(Cb_Grupo)
   
   Call LoadAll

End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   lInLoad = True
   lModCuentas = False
   
   Call SetUpGrid
   
   For i = 1 To MAX_TIPOAJUSTESECRLI
      Call CbAddItem(Cb_TipoAjuste, gTipoAjustesECRLI(i), i)
   Next i
   
   Cb_TipoAjuste.ListIndex = 0
   
   lTipoAjuste = CbItemData(Cb_TipoAjuste)
   
   lInLoad = False
   '3402617
   If gEmpresa.R14ASemiIntegrado = True And gEmpresa.Ano >= 2023 Then
    Me.Caption = "Configuración  Cuentas Ajustes Extra - Contables RLI HR RAD"
   Else
    Me.Caption = "Configuración  Cuentas Ajustes Extra - Contables RLI HR RAB"
   End If
   '3402617
   
End Sub


Private Function SetUpGrid()
   
   Grid.Cols = NCOLS + 1
   Grid.rows = 4
   Grid.FixedRows = 1
   Grid.FixedCols = 0
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_IDCTAAJUSTES) = 0
   Grid.ColWidth(C_CODCUENTA) = 1500
   Grid.ColWidth(C_CUENTA) = 4100
   Grid.ColWidth(C_SELCTA) = 300
   Grid.ColWidth(C_UPD) = 0
   
   Grid.ColAlignment(C_SELCTA) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.Col = C_SELCTA
   Grid.Row = 0
   Set Grid.CellPicture = Bt_Cuentas.Picture
   
   Call FGrVRows(Grid, 1)
   
   
End Function
Private Sub LoadAll()
   Dim i As Integer
   Dim IdCuenta As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim LstCuentas As String
   
   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
      
   Q1 = "SELECT CtasAjustesExContRLI.IdCtaAjustesRLI, CtasAjustesExContRLI.IdCuenta, Codigo, Descripcion "
   Q1 = Q1 & " FROM CtasAjustesExContRLI INNER JOIN Cuentas ON CtasAjustesExContRLI.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "CtasAjustesExContRLI", "Cuentas")
   Q1 = Q1 & " WHERE CtasAjustesExContRLI.TipoAjuste = " & CbItemData(Cb_TipoAjuste) & " AND CtasAjustesExContRLI.IdGrupo = " & CbItemData(Cb_Grupo)
   Q1 = Q1 & " AND CtasAjustesExContRLI.IdItem = " & CbItemData(Cb_Item)
   Q1 = Q1 & " AND CtasAjustesExContRLI.IdEmpresa = " & gEmpresa.id & " AND CtasAjustesExContRLI.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(i, C_IDCTAAJUSTES) = vFld(Rs("IdCtaAjustesRLI"))
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Descripcion"))
      Grid.TextMatrix(i, C_SELCTA) = ">>"
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   Grid.Col = C_CODCUENTA
   Grid.Row = Grid.FixedRows
   Grid.TextMatrix(Grid.Row, C_SELCTA) = ">>"
   
   Grid.Redraw = True
End Sub
Private Function SaveAll()
   Dim i As Integer
   Dim Q1 As String
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_IDCUENTA) = "" Then
         Exit For
      End If
   
      Select Case Grid.TextMatrix(i, C_UPD)
      
         Case FGR_I
            Q1 = "INSERT INTO CtasAjustesExContRLI (TipoAjuste, IdGrupo, IdItem, IdCuenta, CodCuenta, IdEmpresa, Ano) "
            Q1 = Q1 & "VALUES( " & lTipoAjuste & ", " & lIdGrupo & ", " & lIdItem & ", " & Grid.TextMatrix(i, C_IDCUENTA) & ", '" & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA)) & "',"
            Q1 = Q1 & gEmpresa.id & "," & gEmpresa.Ano & ")"
            Call ExecSQL(DbMain, Q1)
         
         Case FGR_U
            Q1 = "UPDATE CtasAjustesExContRLI SET IdCuenta =" & Grid.TextMatrix(i, C_IDCUENTA) & ",CodCuenta = '" & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA)) & "'"
            Q1 = Q1 & " WHERE IdCtaAjustesRLI = " & Grid.TextMatrix(i, C_IDCTAAJUSTES)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
         
         Case FGR_D
'            Q1 = "DELETE * FROM CtasAjustesExCont"
            Q1 = " WHERE IdCtaAjustesRLI = " & Grid.TextMatrix(i, C_IDCTAAJUSTES)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call DeleteSQL(DbMain, "CtasAjustesExContRLI", Q1)
         
      End Select
      
      
   Next i
   
   lModCuentas = False
   
   Call ReadCtasAjustesExtraContRLI

End Function
Private Function valida() As Boolean
   valida = True
End Function

