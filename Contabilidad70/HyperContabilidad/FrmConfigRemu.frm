VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmConfigRemu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Traspaso desde Remuneraciones"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   6315
      Begin VB.OptionButton Op_VerSQLServer 
         Caption         =   "Versión SQL Server"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1755
      End
      Begin VB.OptionButton Op_VerAccess 
         Caption         =   "Versión Access"
         Height          =   255
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base de datos Remuneraciones:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   2310
      End
   End
   Begin VB.CommandButton Bt_CopyDesdeOtraEmp 
      Caption         =   "Copiar Configuración de otra Empresa..."
      Height          =   435
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copiar la configuración de remuneraciones desde otra empresa"
      Top             =   5940
      Width           =   2955
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   9869
      Cols            =   2
      Rows            =   2
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
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   800
      Left            =   10740
      Picture         =   "FrmConfigRemu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   3060
      Width           =   1155
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "Copiar a Excel"
      Height          =   795
      Left            =   10740
      Picture         =   "FrmConfigRemu.frx":0662
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Copiar Excel"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton Bt_SelCuenta 
      Caption         =   "Cuentas"
      Height          =   795
      Left            =   10740
      Picture         =   "FrmConfigRemu.frx":0C17
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1260
      Width           =   1155
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
      Left            =   10860
      Picture         =   "FrmConfigRemu.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Plan de Cuentas"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localización del Sistema de Remuneraciones"
      Height          =   915
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   11835
      Begin VB.CommandButton Bt_Browse 
         Height          =   435
         Left            =   10500
         Picture         =   "FrmConfigRemu.frx":1573
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox Tx_DbRemu 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Lb_File 
         Caption         =   "Base de datos Remuneraciones"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   2475
      End
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   10740
      TabIndex        =   6
      Top             =   300
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10740
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "FrmConfigRemu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDTIPO = 0
Const C_IDCUENTA = 1
Const C_DESC = 2
Const C_CODCUENTA = 3
Const C_CUENTA = 4
Const C_SELCTA = 5
Const C_UPD = 6

Const NCOLS = C_UPD



Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Grid.FixedCols = C_DESC + 1
   
   Grid.ColWidth(C_IDTIPO) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_DESC) = 4300
   Grid.ColWidth(C_CODCUENTA) = 1400
   Grid.ColWidth(C_CUENTA) = 4060
   Grid.ColWidth(C_SELCTA) = 300
   Grid.ColWidth(C_UPD) = 0
   
   Call FGrSetup(Grid, True)
   
   Grid.TextMatrix(0, C_DESC) = "Tipo de Dato"
   Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.Col = C_SELCTA
   Grid.row = 0
   Set Grid.CellPicture = Bt_Cuentas.Picture
   

End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyDesdeOtraEmp_Click()
   Dim Frm As FrmCopyPlan

   Set Frm = New FrmCopyPlan
   Call Frm.FCopyConfigRemu
   Set Frm = Nothing

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_CopyExcel_Click()

   Call FGr2Clip(Grid, Me.Caption)

End Sub

Private Sub Bt_Del_Click()
   Dim row As Integer

   row = Grid.row
   
   If Grid.TextMatrix(row, C_CUENTA) <> "" Then
      If MsgBox1("¿Está seguro que desea eliminar la cuenta asociada a este concepto?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   Grid.TextMatrix(row, C_IDCUENTA) = ""
   Grid.TextMatrix(row, C_CODCUENTA) = ""
   Grid.TextMatrix(row, C_CUENTA) = ""

   Grid.TextMatrix(row, C_UPD) = "1"
         
End Sub

Private Sub Bt_OK_Click()
   Call SaveAll
   Unload Me
   
End Sub

Private Sub Bt_Browse_Click()
   
   gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = ""
   If Op_VerAccess Then
      gFrmMain.Cm_ComDlg.Filter = "LPRemu.mdb|LPRemu.mdb|FairPay2.mdb|FairPay2.mdb"
      gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Base de Datos Remuneraciones"
   Else
      gFrmMain.Cm_ComDlg.Filter = "LPRemu.cfg|LPRemu.cfg|FairPay.cfg|FairPay.cfg"
      gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Configuración Remuneraciones"
   End If
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   
   ERR.Clear
   
   If Op_VerAccess Then
      If LCase(FrmMain.Cm_ComDlg.FileTitle) <> "fairpay2.mdb" And LCase(FrmMain.Cm_ComDlg.FileTitle) <> "lpremu.mdb" Then
         MsgBox1 "Nombre de archivo invalido.", vbExclamation
         Exit Sub
      End If
   Else
      If LCase(FrmMain.Cm_ComDlg.FileTitle) <> "fairpay.cfg" And LCase(FrmMain.Cm_ComDlg.FileTitle) <> "lpremu.cfg" Then
         MsgBox1 "Nombre de archivo invalido.", vbExclamation
         Exit Sub
      End If
   End If
   
   Tx_DbRemu = FrmMain.Cm_ComDlg.Filename

End Sub

Private Sub Bt_SelCuenta_Click()
   Dim row As Integer
   
   If Grid.row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call Grid_DblClick
End Sub


Private Sub Form_Load()

   Call SetUpGrid

   Call LoadBase

   Call LoadAll

End Sub
Private Sub LoadBase()
   Dim row As Integer
   Dim i As Integer
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   row = Grid.rows - 1
   
   For i = 1 To UBound(gTipoDatosRemu)
      If gTipoDatosRemu(i) <> "" Then
         If InStr(gTipoDatosRemu(i), "21227") <= 0 Or gEmpresa.Ano >= 2020 Then
            Grid.rows = Grid.rows + 1
            row = row + 1
            Grid.TextMatrix(row, C_IDTIPO) = i
            Grid.TextMatrix(row, C_DESC) = gTipoDatosRemu(i)
         End If
      End If
   Next i
      
   Grid.FlxGrid.Redraw = True
      
End Sub
Private Sub LoadAll()
   Dim Buf As String
   Dim i As Integer
   Dim IdCuenta As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tipo As Integer
   
   Grid.FlxGrid.Redraw = False
   
   'limpiamos todo por si viene de copia de otra empresa
   For i = Grid.FixedRows To Grid.rows - 1
      Grid.TextMatrix(i, C_IDCUENTA) = ""
      Grid.TextMatrix(i, C_CODCUENTA) = ""
      Grid.TextMatrix(i, C_CUENTA) = ""
      Grid.TextMatrix(i, C_UPD) = ""
   Next i
      
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      Tipo = Grid.TextMatrix(i, C_IDTIPO)
      IdCuenta = Val(GetParamEmpresa("CTASREMU", Tipo))
      If IdCuenta > 0 Then
         Grid.TextMatrix(i, C_IDCUENTA) = IdCuenta
      
         Q1 = "SELECT Codigo, Descripcion FROM Cuentas WHERE IdCuenta=" & IdCuenta
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Descripcion"))
         End If
         Call CloseRs(Rs)
      End If
      
      Grid.TextMatrix(i, C_SELCTA) = ">>"
   Next i
      
   Op_VerAccess = True
   Buf = UCase(GetIniString(gIniFile, "Config", "VersionRemu", ""))
   If Buf <> "" Then
      Op_VerAccess = IIf(Buf = "ACCESS", True, False)
      Op_VerSQLServer = IIf(Buf = "SQLSERVER", True, False)
   End If
   
   Buf = GetIniString(gIniFile, "Config", "PathRemu", "")
   Tx_DbRemu = Buf
   
   Grid.FlxGrid.Redraw = True
End Sub

Private Sub SaveAll()
   Dim Txt As TextBox
   Dim i As Integer
   Dim Tipo As Integer
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Trim(Grid.TextMatrix(i, C_UPD)) <> "" Then
         Tipo = Val(Grid.TextMatrix(i, C_IDTIPO))
         Call UpdParamEmpresa("CTASREMU", Tipo, Val(Grid.TextMatrix(i, C_IDCUENTA)))
      End If
   Next i
   
   Call SetIniString(gIniFile, "Config", "VersionRemu", IIf(Op_VerAccess <> 0, "ACCESS", "SQLSERVER"))
   Call SetIniString(gIniFile, "Config", "PathRemu", Tx_DbRemu)

End Sub

Private Sub Grid_AcceptValue(ByVal row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim cod As String
   Dim UltimoNivel As Boolean
   Dim NombCta As String, DescCta As String
   Dim IdCuenta As Long
   
   Value = Trim(Value)
   
   cod = Trim(ReplaceStr(Value, "-", ""))
   If Len(cod) < Len(VFmtCodigoCta(gFmtCodigoCta)) Then   'asumimos que está usando nombre corto
      NombCta = UCase(Trim(Value))
      cod = ""
   Else
      NombCta = ""
   End If
   
   IdCuenta = GetIdCuenta(NombCta, cod, DescCta, UltimoNivel)
   
   If IdCuenta = 0 Then
      MsgBeep vbExclamation
      Action = vbCancel
   
   ElseIf UltimoNivel = False Then
      MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
      Action = vbCancel
   
   Else
      
      Grid.TextMatrix(row, C_IDCUENTA) = IdCuenta
      Value = Format(cod, gFmtCodigoCta)
      Grid.TextMatrix(row, C_CUENTA) = DescCta
      Call FGrModRow(Grid, row, FGR_U, C_IDCUENTA, C_UPD)
      
   End If
   
End Sub

Private Sub Grid_BeforeEdit(ByVal row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   If Col <> C_CODCUENTA Or Grid.TextMatrix(row, C_DESC) = "" Then
      Exit Sub
   End If
   
   EdType = FEG_Edit
   
End Sub

Private Sub Grid_DblClick()
   Dim FrmPlan As FrmPlanCuentas
   Dim DescCta As String
   Dim CodCta As String
   Dim NombCuenta As String
   Dim row As Integer
   Dim IdCuenta As Long

   If Grid.Col <> C_SELCTA And Grid.Col <> C_CUENTA Then
      Exit Sub
   End If
   
   row = Grid.row
   
   Set FrmPlan = New FrmPlanCuentas

   If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta, True) = vbOK Then
      If DescCta <> "" Then
         Grid.TextMatrix(row, C_IDCUENTA) = IdCuenta
         Grid.TextMatrix(row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
         Grid.TextMatrix(row, C_CUENTA) = DescCta

         Grid.TextMatrix(row, C_UPD) = FGR_U
         
     End If

   End If
   Set FrmPlan = Nothing

End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyUpper(KeyAscii)
End Sub

Private Sub Op_VerAccess_Click()

   If Op_VerAccess Then
      Lb_File = "Base de Datos Remuneraciones"
      If UCase(Right(Tx_DbRemu, 4)) <> ".MDB" Then
         Tx_DbRemu = ""
      End If
   End If

   
End Sub

Private Sub Op_VerSQLServer_Click()

   If Op_VerSQLServer Then
      Lb_File = "Archivo Config. Remuneraciones"
      If UCase(Right(Tx_DbRemu, 4)) <> ".CFG" Then
         Tx_DbRemu = ""
      End If
   End If


End Sub
