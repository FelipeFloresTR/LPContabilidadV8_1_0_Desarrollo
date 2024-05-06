VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmConfigCtasLibVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Cuentas Libro de Ventas"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_DelAll 
      Caption         =   "&Eliminar"
      Height          =   800
      Left            =   12600
      Picture         =   "FrmConfigCtasLibVentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Eliminar TODAS las cuentas asignadas en esta  configuración"
      Top             =   3900
      Width           =   1155
   End
   Begin VB.CommandButton Bt_FmtImport 
      Caption         =   "Formato Imp."
      Height          =   315
      Left            =   12600
      TabIndex        =   12
      Top             =   6000
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   800
      Left            =   12600
      Picture         =   "FrmConfigCtasLibVentas.frx":0619
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   5160
      Width           =   1155
   End
   Begin TabDlg.SSTab Tab_Config 
      Height          =   6555
      Left            =   180
      TabIndex        =   0
      Top             =   1140
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   11562
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Cuentas Afecto"
      TabPicture(0)   =   "FrmConfigCtasLibVentas.frx":0C8C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Cb_LLenarDelGiro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Bt_LllenarDelGiro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GridAfecto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cuentas Exento"
      TabPicture(1)   =   "FrmConfigCtasLibVentas.frx":0CA8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridExento"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cuentas Total"
      TabPicture(2)   =   "FrmConfigCtasLibVentas.frx":0CC4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GridTotal"
      Tab(2).ControlCount=   1
      Begin FlexEdGrid3.FEd3Grid GridAfecto 
         Height          =   5715
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10081
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
      Begin VB.CommandButton Bt_LllenarDelGiro 
         Caption         =   "LLenar todo"
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.ComboBox Cb_LLenarDelGiro 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   915
      End
      Begin FlexEdGrid3.FEd3Grid GridExento 
         Height          =   5715
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10081
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
      Begin FlexEdGrid3.FEd3Grid GridTotal 
         Height          =   5715
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10081
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Del Giro:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   180
      Picture         =   "FrmConfigCtasLibVentas.frx":0CE0
      ScaleHeight     =   630
      ScaleWidth      =   690
      TabIndex        =   18
      Top             =   300
      Width           =   690
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   12600
      TabIndex        =   14
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   12600
      TabIndex        =   13
      Top             =   240
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
      Left            =   12540
      Picture         =   "FrmConfigCtasLibVentas.frx":135B
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Plan de Cuentas"
      Top             =   7260
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Bt_SelCuenta 
      Caption         =   "Cuentas"
      Height          =   795
      Left            =   12600
      Picture         =   "FrmConfigCtasLibVentas.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1155
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "Copiar a Excel"
      Height          =   795
      Left            =   12600
      Picture         =   "FrmConfigCtasLibVentas.frx":1CB7
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Copiar Excel"
      Top             =   2100
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   800
      Left            =   12600
      Picture         =   "FrmConfigCtasLibVentas.frx":226C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   1020
      TabIndex        =   15
      Top             =   240
      Width           =   11355
      Begin VB.ComboBox Cb_Buscar 
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1635
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
         Left            =   10800
         Picture         =   "FrmConfigCtasLibVentas.frx":28CE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar una cuenta"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Tx_Buscar 
         Height          =   315
         Left            =   6780
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   3915
      End
      Begin VB.ComboBox Cb_OrdenarPor 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar por:"
         Height          =   195
         Index           =   0
         Left            =   4020
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenar por:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmConfigCtasLibVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDENTIDAD = 0
Const C_RUT = 1
Const C_CODIGO = 2
Const C_NOMBRE = 3
Const C_IDDELGIRO = 4
Const C_DELGIRO = 5
Const C_IDCUENTA = 6
Const C_CODCUENTA = 7
Const C_CUENTA = 8
Const C_SELCTA = 9
Const C_AREANEG = 10
Const C_CODAREANEG = 11
Const C_CCOSTO = 12
Const C_CODCCOSTO = 13
Const C_EXIGEANEG = 14
Const C_EXIGECCOSTO = 15
Const C_UPD = 16

Const NCOLS = C_UPD

Const TAB_AFECTO = 0
Const TAB_EXENTO = 1
Const TAB_TOTAL = 2

Dim lBuscarRut As String
Dim lBuscarNombre As String
Dim lFirstConfigCtas  As Boolean
Dim lOldOrdenarPor As Integer
Dim lInLoad As Boolean
Dim lModificado As Boolean
Dim lMsgActivate As Boolean

Dim lCbCCostoAfecto As ClsCombo
Dim lCbANegAfecto As ClsCombo
Dim lCbCCostoExento As ClsCombo
Dim lCbANegExento As ClsCombo
Dim lCbCCostoTotal As ClsCombo
Dim lCbANegTotal As ClsCombo

'cuentas default
Dim lCodCtaAfecto As String
Dim lCodCtaExento As String
Dim lCodCtaTotal As String

Dim lHayCuentasDef As Boolean



Private Sub SetUpGrid(Grid As FEd3Grid, Optional ByVal ConDelGiro As Boolean = True)
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Grid.FixedCols = C_NOMBRE + 1
   
   Call FGrSetup(Grid, True)
      
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_CODIGO) = 1200
   Grid.ColWidth(C_NOMBRE) = 3500
   
   Grid.ColWidth(C_IDDELGIRO) = 0
   Grid.ColWidth(C_DELGIRO) = 560
   
   If Not ConDelGiro Then
      Grid.ColWidth(C_NOMBRE) = Grid.ColWidth(C_NOMBRE) + Grid.ColWidth(C_DELGIRO)
      Grid.ColWidth(C_DELGIRO) = 0
   End If
   
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CODCUENTA) = 1150
   Grid.ColWidth(C_CUENTA) = 3800
   Grid.ColWidth(C_SELCTA) = 300
   Grid.ColWidth(C_CODCCOSTO) = 0
   Grid.ColWidth(C_CODAREANEG) = 0
   Grid.ColWidth(C_CCOSTO) = 2000
   Grid.ColWidth(C_AREANEG) = 2000
   Grid.ColWidth(C_EXIGEANEG) = 0
   Grid.ColWidth(C_EXIGECCOSTO) = 0
   Grid.ColWidth(C_UPD) = 0
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_DELGIRO) = flexAlignCenterCenter
   Grid.ColAlignment(C_CCOSTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_AREANEG) = flexAlignLeftCenter

   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_CODIGO) = "Nombre Corto"
   Grid.TextMatrix(0, C_NOMBRE) = "Razón Social"
   If Grid.ColWidth(C_DELGIRO) > 0 Then
      Grid.TextMatrix(0, C_DELGIRO) = "Giro"
   End If
   Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_CCOSTO) = "Centro de Gestión"
   Grid.TextMatrix(0, C_AREANEG) = "Área de Negocio"
   
   Grid.Col = C_SELCTA
   Grid.Row = 0
   Set Grid.CellPicture = Bt_Cuentas.Picture
   
   Call FGrVRows(Grid, 1)
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()

   Select Case Tab_Config.Tab
      Case TAB_AFECTO
         Call FGr2Clip(GridAfecto, Me.Caption & " - Cuentas Afecto")
      Case TAB_EXENTO
         Call FGr2Clip(GridExento, Me.Caption & " - Cuentas Exento")
      Case TAB_TOTAL
         Call FGr2Clip(GridTotal, Me.Caption & " - Cuentas Total")
   End Select

   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Grid As FEd3Grid
   
   
   Select Case Tab_Config.Tab
      Case TAB_AFECTO
         Set Grid = GridAfecto
      Case TAB_EXENTO
         Set Grid = GridExento
      Case TAB_TOTAL
         Set Grid = GridTotal
   End Select

   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Grid.TextMatrix(Row, C_IDCUENTA) = ""
   Grid.TextMatrix(Row, C_CODCUENTA) = ""
   Grid.TextMatrix(Row, C_CUENTA) = ""
   
   Grid.TextMatrix(Row, C_UPD) = FGR_U
   
End Sub
Private Sub Bt_DelAll_Click()

   If MsgBox1("Este proceso eliminará TODOS los valores que las entidades ya tengan asignados." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   Call DelAllGrid(GridAfecto)
   Call DelAllGrid(GridExento)
   Call DelAllGrid(GridTotal)
   
   MsgBox1 "Ahora debe presionar el botón Aceptar para finalizar el proceso." & vbCrLf & vbCrLf & "Para anularlo, presione el botón Cancelar.", vbInformation
   

End Sub
Private Function DelAllGrid(Grid As FEd3Grid)
   Dim i As Integer, j As Integer
     
   For i = Grid.FixedRows To Grid.rows - 1
      For j = C_IDDELGIRO To C_CODCCOSTO
         Grid.TextMatrix(i, j) = ""
      Next j
      Grid.TextMatrix(i, C_UPD) = FGR_U
   Next i

End Function

Private Sub Bt_FmtImport_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewConfigCtasLibVentas
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Import_Click()
   Dim FName As String

   If MsgBox1("Este proceso reemplazará los valores que las entidades ya tengan asignados y no es posible cancelarlo después de realizado." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
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
   
   FName = FrmMain.Cm_ComDlg.Filename
      
   DoEvents
   
   Call ImpCtasLibVentas(FName)
      
End Sub

Private Sub Bt_LllenarDelGiro_Click()
   Dim Row As Integer

   If MsgBox1("¿Está seguro que desea cambiar el Giro de las Ventas para TODAS las entidades?", vbQuestion + vbYesNoCancel) <> vbYes Then
      Exit Sub
   End If
   
   For Row = GridAfecto.FixedRows To GridAfecto.rows - 1
      
      If Val(GridAfecto.TextMatrix(Row, C_IDENTIDAD)) > 0 Then
         GridAfecto.TextMatrix(Row, C_IDDELGIRO) = CbItemData(Cb_LLenarDelGiro)
         GridAfecto.TextMatrix(Row, C_DELGIRO) = FmtSiNo(Val(GridAfecto.TextMatrix(Row, C_IDDELGIRO)))
         GridAfecto.TextMatrix(Row, C_UPD) = FGR_U
         lModificado = True
      Else
         Exit For
      End If
      
   Next Row
      
End Sub

Private Sub Bt_OK_Click()

   If Not Valida(GridAfecto, "Afecto") Then
      Exit Sub
   End If
   If Not Valida(GridExento, "Exento") Then
      Exit Sub
   End If
   If Not Valida(GridTotal, "Total") Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Call SaveAll
   Me.MousePointer = vbDefault
   
   Unload Me
End Sub

Private Sub Bt_Search_Click()

   If lModificado Then
      If MsgBox1("Antes de buscar es necesario grabar los cambios realizados en la configuración." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNoCancel + vbQuestion) = vbYes Then
         Call SaveAll
      Else
         Exit Sub
      End If
   End If
   
   lBuscarRut = ""
   lBuscarNombre = ""

   If Trim(Tx_Buscar) <> "" Then
      If UCase(Cb_Buscar) = "RUT" Then
         lBuscarRut = Trim(Tx_Buscar)
      Else
         lBuscarNombre = Trim(Tx_Buscar)
      End If
      
   End If
   
   Call LoadAll
End Sub

Private Sub Cb_OrdenarPor_Click()
   Static InCbOrdenarPor As Boolean

   If lInLoad Then
      Exit Sub
   End If
   
   If InCbOrdenarPor Then
      Exit Sub
   End If
   
   InCbOrdenarPor = True
   
   If lModificado Then
      If MsgBox1("Antes de ordenar es necesario grabar los cambios realizados en la configuración." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNoCancel + vbQuestion) = vbYes Then
         Call SaveAll
         Call LoadAll
         lOldOrdenarPor = Cb_OrdenarPor.ListIndex
      Else
         Cb_OrdenarPor.ListIndex = lOldOrdenarPor
      End If
   Else
      Call LoadAll
      lOldOrdenarPor = Cb_OrdenarPor.ListIndex
   End If
   
   InCbOrdenarPor = False
   
End Sub

Private Sub Form_Activate()

   If lMsgActivate Then
      Exit Sub
   End If
   
   If lHayCuentasDef Then
      If lFirstConfigCtas = True Then
         MsgBox1 "Para esta primera configuración, el sistema propone cuentas, centros de gestión y áreas de negocio de acuerdo al uso de éstas en el libro de ventas de este año y del año anterior." & vbCrLf & vbCrLf & "Si las cuentas contables coinciden con las por omisión, no se asignan." & vbCrLf & vbCrLf & "Usted puede modificarlas si lo desea.", vbInformation
      Else
         MsgBox1 "El sistema propone cuentas, centros de gestión y áreas de negocio de acuerdo al uso de éstas en el libro de ventas de este año, sólo para las cuentas que están en blanco." & vbCrLf & vbCrLf & "Si las cuentas contables coinciden con las por omisión, no se asignan." & vbCrLf & vbCrLf & "Usted puede modificarlas si lo desea.", vbInformation
      End If
   End If
   
   lMsgActivate = True
   
End Sub

Private Sub Form_Load()
   Dim Rs As Recordset
   Dim i As Integer
   Dim Q1 As String
   
   lInLoad = True
   Call SetUpGrid(GridAfecto, True)
   Call SetUpGrid(GridExento, False)
   Call SetUpGrid(GridTotal, False)
   
   If LoadCodCuentasDefLibros(LIB_VENTAS, lCodCtaAfecto, lCodCtaExento, lCodCtaTotal) Then
   
      Call FillCuentasUtilizadas(LIB_VENTAS, LIBVENTAS_AFECTO, "CodCtaAfectoVta", "CodCCostoAfectoVta", "CodAreaNegAfectoVta", lCodCtaAfecto)
      Call FillCuentasUtilizadas(LIB_VENTAS, LIBVENTAS_EXENTO, "CodCtaExentoVta", "CodCCostoExentoVta", "CodAreaNegExentoVta", lCodCtaExento)
      Call FillCuentasUtilizadas(LIB_VENTAS, LIBVENTAS_TOTAL, "CodCtaTotalVta", "CodCCostoTotalVta", "CodAreaNegTotalVta", lCodCtaTotal)
      
      'vemos si es la primera vez que se hace esta configuración, para tomar las cuentas utilizadas en el año anterior sólo la primera vez
      Set Rs = OpenRs(DbMain, "SELECT Codigo FROM ParamEmpresa WHERE Tipo = 'CONFCTAVTA'  AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      If Rs.EOF Then
         Call FillCuentasUtilizadasAnoAnt(LIB_VENTAS, LIBVENTAS_AFECTO, "CodCtaAfectoVta", "CodCCostoAfectoVta", "CodAreaNegAfectoVta", lCodCtaAfecto)
         Call FillCuentasUtilizadasAnoAnt(LIB_VENTAS, LIBVENTAS_EXENTO, "CodCtaExentoVta", "CodCCostoExentoVta", "CodAreaNegExentoVta", lCodCtaExento)
         Call FillCuentasUtilizadasAnoAnt(LIB_VENTAS, LIBVENTAS_TOTAL, "CodCtaTotalVta", "CodCCostoTotalVta", "CodAreaNegTotalVta", lCodCtaTotal)
         
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES( 'CONFCTAVTA', 0, 'ConfigCtasImpLibVtaSIIAnoAnt', " & gEmpresa.id & "," & gEmpresa.Ano & ")"
         Call ExecSQL(DbMain, Q1)
   
      End If
      Call CloseRs(Rs)
      lHayCuentasDef = True
      
   Else
      lHayCuentasDef = False
      
   End If
   
   Call CbAddItem(Cb_OrdenarPor, "Nombre", 1)
   Call CbAddItem(Cb_OrdenarPor, "RUT", 2)
   Cb_OrdenarPor.ListIndex = 0   'nombre
   lOldOrdenarPor = 0
   
   Call CbAddItem(Cb_Buscar, "Nombre", 1)
   Call CbAddItem(Cb_Buscar, "RUT", 2)
   Cb_Buscar.ListIndex = 0   'nombre
   
   Call CbAddItem(Cb_LLenarDelGiro, "Si", VAL_SI)
   Call CbAddItem(Cb_LLenarDelGiro, "No", VAL_NO)

   Cb_LLenarDelGiro.ListIndex = 0   'vacío
   
   Set lCbCCostoAfecto = New ClsCombo
   Set lCbCCostoExento = New ClsCombo
   Set lCbCCostoTotal = New ClsCombo
   
   Set lCbANegAfecto = New ClsCombo
   Set lCbANegExento = New ClsCombo
   Set lCbANegTotal = New ClsCombo

   Call lCbCCostoAfecto.SetControl(GridAfecto.CbList(C_CCOSTO))
   Call lCbANegAfecto.SetControl(GridAfecto.CbList(C_AREANEG))
   
   Call lCbCCostoExento.SetControl(GridExento.CbList(C_CCOSTO))
   Call lCbANegExento.SetControl(GridExento.CbList(C_AREANEG))
   
   Call lCbCCostoTotal.SetControl(GridTotal.CbList(C_CCOSTO))
   Call lCbANegTotal.SetControl(GridTotal.CbList(C_AREANEG))
      
   Call FillCbGrid(lCbCCostoAfecto, lCbANegAfecto)
   Call FillCbGrid(lCbCCostoExento, lCbANegExento)
   Call FillCbGrid(lCbCCostoTotal, lCbANegTotal)
   
   Call LoadAll
   
   Tab_Config.Tab = TAB_AFECTO
   
   lInLoad = False

End Sub
Private Sub FillCbGrid(CbCCosto As ClsCombo, CbANeg As ClsCombo)
   Dim Q1 As String
   
   Q1 = "SELECT Descripcion, Codigo FROM CentroCosto WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbCCosto.AddItem(" ", " ")
   Call CbCCosto.FillCombo(DbMain, Q1, "")
      
   Q1 = "SELECT Descripcion, Codigo FROM AreaNegocio WHERE IdEmpresa = " & gEmpresa.id & " AND Vigente <> 0 ORDER BY Descripcion"
   Call CbANeg.AddItem(" ", " ")
   Call CbANeg.FillCombo(DbMain, Q1, "")
      
End Sub

Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String, Q2 As String, QWh As String, QSort As String, QEmp As String, OnEmp As String
   Dim QCta As String
   Dim TipoVal As String
   
   Q1 = "SELECT IdEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Codigo as CodEnt, Entidades.Nombre as NombEnt, Entidades.EsDelGiro, "
   Q1 = Q1 & " CentroCosto.Codigo as CodCCosto, CentroCosto.Descripcion as DescCCosto, AreaNegocio.Codigo as CodAreaNeg, AreaNegocio.Descripcion as DescAreaNeg, "
   Q1 = Q1 & " Cuentas.IdCuenta, Cuentas.Codigo as CodCuenta, Cuentas.Descripcion as Cuenta, Cuentas.Atrib" & ATRIB_AREANEG & ", Atrib" & ATRIB_CCOSTO
   Q1 = Q1 & " FROM ((Entidades LEFT JOIN Cuentas ON " ' Entidades.CodCtaAfecto = Cuentas.Codigo "
   
   QWh = " WHERE Entidades.Estado = " & EE_ACTIVO & " AND Clasif" & ENT_CLIENTE & " <> 0 "
   QWh = QWh & " AND Entidades.IdEmpresa = " & gEmpresa.id & " AND (Cuentas.Ano = " & gEmpresa.Ano & " OR Cuentas.Ano IS NULL)"
   
   Q2 = ""
   
   If lBuscarRut <> "" Then
      QWh = QWh & " AND Entidades.Rut = '" & vFmtCID(lBuscarRut, True) & "'"
   ElseIf lBuscarNombre <> "" Then
      QWh = QWh & " AND " & GenLike(DbMain, lBuscarNombre, "Entidades.Nombre")
   End If
      
   If Cb_OrdenarPor = "Nombre" Then
      QSort = " ORDER BY Entidades.Nombre "
   Else
      QSort = " ORDER BY Rut "
   End If
   
   OnEmp = " AND Entidades.IdEmpresa = Cuentas.IdEmpresa ) "
   
   QCta = " Entidades.CodCtaAfectoVta = Cuentas.Codigo "
   TipoVal = "Afecto"
   Q2 = " LEFT JOIN CentroCosto ON Entidades.CodCCosto" & TipoVal & "Vta = CentroCosto.Codigo  "
   Q2 = Q2 & " AND CentroCosto.IdEmpresa = Entidades.IdEmpresa  ) "
   Q2 = Q2 & " LEFT JOIN AreaNegocio ON Entidades.CodAreaNeg" & TipoVal & "Vta = AreaNegocio.Codigo "
   Q2 = Q2 & " AND AreaNegocio.IdEmpresa = Entidades.IdEmpresa  "
   Call LoadGrid(GridAfecto, Q1 & QCta & OnEmp & Q2 & QWh & QSort)
   
   QCta = " Entidades.CodCtaExentoVta = Cuentas.Codigo "
   TipoVal = "Exento"
   Q2 = " LEFT JOIN CentroCosto ON Entidades.CodCCosto" & TipoVal & "Vta = CentroCosto.Codigo "
   Q2 = Q2 & " AND CentroCosto.IdEmpresa = Entidades.IdEmpresa  ) "
   Q2 = Q2 & " LEFT JOIN AreaNegocio ON Entidades.CodAreaNeg" & TipoVal & "Vta = AreaNegocio.Codigo "
   Q2 = Q2 & " AND AreaNegocio.IdEmpresa = Entidades.IdEmpresa  "
   Call LoadGrid(GridExento, Q1 & QCta & OnEmp & Q2 & QWh & QSort)
   
   QCta = " Entidades.CodCtaTotalVta = Cuentas.Codigo  "
   QCta = QCta & " AND Entidades.IdEmpresa = Cuentas.IdEmpresa "
   TipoVal = "Total"
   Q2 = " LEFT JOIN CentroCosto ON Entidades.CodCCosto" & TipoVal & "Vta = CentroCosto.Codigo "
   Q2 = Q2 & " AND CentroCosto.IdEmpresa = Entidades.IdEmpresa  ) "
   Q2 = Q2 & " LEFT JOIN AreaNegocio ON Entidades.CodAreaNeg" & TipoVal & "Vta = AreaNegocio.Codigo "
   Q2 = Q2 & " AND AreaNegocio.IdEmpresa = Entidades.IdEmpresa  "
   Call LoadGrid(GridTotal, Q1 & QCta & OnEmp & Q2 & QWh & QSort)
   
End Sub
Private Sub LoadGrid(Grid As FEd3Grid, ByVal Q1 As String)
   Dim i As Integer
   Dim Rs As Recordset
   Dim ConDelGiro As Boolean
   
   If Grid.ColWidth(C_DELGIRO) Then
      ConDelGiro = True
   End If
   
   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)
      Grid.TextMatrix(i, C_CODIGO) = vFld(Rs("CodEnt"))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("NombEnt"))
      
      If ConDelGiro Then
         Grid.TextMatrix(i, C_IDDELGIRO) = vFld(Rs("EsDelGiro"))
         Grid.TextMatrix(i, C_DELGIRO) = FmtSiNo(Val(Grid.TextMatrix(i, C_IDDELGIRO)))
      End If
      
      Grid.TextMatrix(i, C_CODCCOSTO) = vFld(Rs("CodCCosto"))
      Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("DescCCosto"))
      Grid.TextMatrix(i, C_CODAREANEG) = vFld(Rs("CodAreaNeg"))
      Grid.TextMatrix(i, C_AREANEG) = vFld(Rs("DescAreaNeg"))
         
      Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("IdCuenta"))
      Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta((vFld(Rs("CodCuenta"))))
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Cuenta"))
      Grid.TextMatrix(i, C_SELCTA) = ">>"
      
      Grid.TextMatrix(i, C_EXIGEANEG) = vFld(Rs("Atrib" & ATRIB_AREANEG))
      Grid.TextMatrix(i, C_EXIGECCOSTO) = vFld(Rs("Atrib" & ATRIB_CCOSTO))
      
      i = i + 1
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_RUT
   Grid.Redraw = True
   

End Sub

Private Sub Form_Resize()
   Dim H As Integer
   
   Tab_Config.Height = Me.Height - Tab_Config.Top - 800
   H = Tab_Config.Height - GridAfecto.Top - 300
   GridAfecto.Height = H
   GridExento.Height = H
   GridTotal.Height = H
   
   Tab_Config.Width = Me.Width - Tab_Config.Left - Bt_SelCuenta.Width - 560
'   GridAfecto.Width = Tab_Config.Width - GridAfecto.Left - 120
   GridAfecto.Width = Tab_Config.Width - 120 - 120     'se pone directo 120 porque el tab mueve los objetos a la izquierda para ocultarlos
   GridExento.Width = GridAfecto.Width
   GridTotal.Width = GridAfecto.Width
   
   Bt_SelCuenta.Left = Me.Width - Bt_SelCuenta.Width - 340
   Bt_CopyExcel.Left = Bt_SelCuenta.Left
   Bt_Del.Left = Bt_SelCuenta.Left
   Bt_Import.Left = Bt_SelCuenta.Left
   Bt_FmtImport.Left = Bt_SelCuenta.Left
   
   Call FGrVRows(GridAfecto, 1)
   Call FGrVRows(GridExento, 1)
   Call FGrVRows(GridTotal, 1)
   
   GridAfecto.Col = C_RUT
   GridAfecto.ColSel = C_RUT
   GridExento.Col = C_RUT
   GridExento.ColSel = C_RUT
   GridTotal.Col = C_RUT
   GridTotal.ColSel = C_RUT
   
   GridAfecto.Redraw = True
End Sub
Private Sub Bt_SelCuenta_Click()
   Dim Row As Integer
   Dim Grid As FEd3Grid
   
   If Tab_Config.Tab = TAB_AFECTO Then
      Set Grid = GridAfecto
   ElseIf Tab_Config.Tab = TAB_EXENTO Then
      Set Grid = GridExento
   Else                    'TAB_total
      Set Grid = GridTotal
   End If
     
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call GridDblClick(Grid)
End Sub

Private Sub GridDblClick(Grid As Control)
   Dim FrmPlan As FrmPlanCuentas
   Dim DescCta As String
   Dim CodCta As String
   Dim NombCuenta As String
   Dim Row As Integer, Col As Integer
   Dim IdCuenta As Long
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col = C_CODCUENTA Or Col = C_CUENTA Or Col = C_SELCTA Then
   
      'Columna Cuenta
      Set FrmPlan = New FrmPlanCuentas
   
      If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta, True) = vbOK Then
         If DescCta <> "" Then
            Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
            Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
            Grid.TextMatrix(Row, C_CUENTA) = DescCta
            
            Grid.TextMatrix(Row, C_EXIGEANEG) = GetAtribCuenta(IdCuenta, ATRIB_AREANEG)
            Grid.TextMatrix(Row, C_EXIGECCOSTO) = GetAtribCuenta(IdCuenta, ATRIB_CCOSTO)
               
            Grid.TextMatrix(Row, C_UPD) = FGR_U
            
            lModificado = True
            
        End If

      End If
      Set FrmPlan = Nothing

   ElseIf Col = C_DELGIRO Then
      Grid.TextMatrix(Row, C_IDDELGIRO) = (Val(Grid.TextMatrix(Row, C_IDDELGIRO)) + 1) Mod 2
      Grid.TextMatrix(Row, C_DELGIRO) = FmtSiNo(Val(Grid.TextMatrix(Row, C_IDDELGIRO)))
      Grid.TextMatrix(Row, C_UPD) = FGR_U
      lModificado = True
      
   End If
End Sub

Private Sub SaveAll()
   
   Call SaveGrid(GridAfecto, "CodCtaAfectoVta", "Afecto")
   Call SaveGrid(GridExento, "CodCtaExentoVta", "Exento")
   Call SaveGrid(GridTotal, "CodCtaTotalVta", "Total")
   
   lModificado = False

End Sub
Private Sub SaveGrid(Grid As FEd3Grid, ByVal CodCtaFldName As String, ByVal TipoVal As String)
   Dim i As Integer
   Dim Q1 As String
   Dim ConDelGiro As Boolean
   
   If Grid.ColWidth(C_DELGIRO) Then
      ConDelGiro = True
   End If
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_IDENTIDAD) <> "" Then
         If Grid.TextMatrix(i, C_UPD) = FGR_U Then
            Q1 = "UPDATE Entidades SET " & CodCtaFldName & "= '" & ParaSQL(VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA))) & "' "
            If ConDelGiro Then
               Q1 = Q1 & ", EsDelGiro = " & Val(Grid.TextMatrix(i, C_IDDELGIRO))
            End If
            Q1 = Q1 & ", CodCCosto" & TipoVal & "Vta = '" & ParaSQL(Grid.TextMatrix(i, C_CODCCOSTO)) & "'"
            Q1 = Q1 & ", CodAreaNeg" & TipoVal & "Vta = '" & ParaSQL(Grid.TextMatrix(i, C_CODAREANEG)) & "'"
            Q1 = Q1 & " WHERE IdEntidad = " & Grid.TextMatrix(i, C_IDENTIDAD)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            Call ExecSQL(DbMain, Q1)
         End If
      End If
   Next i
   
   
End Sub

Private Sub GridAfecto_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   If Col = C_CCOSTO Then
      GridAfecto.TextMatrix(Row, C_CODCCOSTO) = lCbCCostoAfecto.ItemData
      GridAfecto.TextMatrix(Row, C_UPD) = FGR_U
   
   ElseIf Col = C_AREANEG Then
      GridAfecto.TextMatrix(Row, C_CODAREANEG) = lCbANegAfecto.ItemData
      GridAfecto.TextMatrix(Row, C_UPD) = FGR_U

   End If

   
End Sub
Private Sub GridAfecto_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Col = C_CCOSTO Or Col = C_AREANEG Then
      EdType = FlexEdGrid3.FEG_List
   End If
   
End Sub

Private Sub GridAfecto_DblClick()
   Call GridDblClick(GridAfecto)
   
End Sub

Private Sub GridExento_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   If Col = C_CCOSTO Then
      GridExento.TextMatrix(Row, C_CODCCOSTO) = lCbCCostoExento.ItemData
      GridExento.TextMatrix(Row, C_UPD) = FGR_U
   
   ElseIf Col = C_AREANEG Then
      GridExento.TextMatrix(Row, C_CODAREANEG) = lCbANegExento.ItemData
      GridExento.TextMatrix(Row, C_UPD) = FGR_U

   End If

End Sub

Private Sub GridExento_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)

   If Col = C_CCOSTO Or Col = C_AREANEG Then
      EdType = FlexEdGrid3.FEG_List
   End If
   
End Sub

Private Sub GridExento_DblClick()
   Call GridDblClick(GridExento)

End Sub

Private Sub GridTotal_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   If Col = C_CCOSTO Then
      GridTotal.TextMatrix(Row, C_CODCCOSTO) = lCbCCostoTotal.ItemData
      GridTotal.TextMatrix(Row, C_UPD) = FGR_U
   
   ElseIf Col = C_AREANEG Then
      GridTotal.TextMatrix(Row, C_CODAREANEG) = lCbANegTotal.ItemData
      GridTotal.TextMatrix(Row, C_UPD) = FGR_U

   End If

End Sub

Private Sub GridTotal_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)

   If Col = C_CCOSTO Or Col = C_AREANEG Then
      EdType = FlexEdGrid3.FEG_List
   End If

End Sub

Private Sub GridTotal_DblClick()
   Call GridDblClick(GridTotal)

End Sub
Private Function ImpCtasLibVentas(ByVal FName As String) As Integer
   Dim FNameLogImp As String
   Dim ConfigErr As Boolean
   Dim MaxReg As Integer
   Dim Fd As Long
   Dim l As Integer, i As Integer
   Dim p As Long
   Dim Buf As String
   Dim RazonSocial As String
   Dim NotValidRut As Boolean
   Dim RazonSocialProv As String
   Dim DelGiro As Integer
   Dim AuxCodCtaAfecto As String, AuxIdCtaAfecto As Long, AuxDescCtaAfecto As String
   Dim AuxCodCtaExento As String, AuxIdCtaExento As Long, AuxDescCtaExento As String
   Dim AuxCodCtaTotal As String, AuxIdCtaTotal As Long, AuxDescCtaTotal As String
   Dim NomCta As String, UltNivel As Boolean
   Dim NuevaEntidad As Boolean
   Dim Sep As String
   Dim RutProv As String
   Dim IdEntidad As Long
   Dim Q1 As String
   Dim CodANegAfecto As String, CodANegExento As String, CodANegTotal As String
   Dim CodCCostoAfecto As String, CodCCostoExento As String, CodCCostoTotal As String
   Dim AuxId As Long, NEntidades As Integer, NUpdate As Integer
   

   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibCtasLibComp-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxReg = 3000
   If gDbType = SQL_ACCESS Then
    If LineCount(FName, MaxReg) < 0 Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para la lista de cuentas del libro de ventas (máx. " & MaxReg & " documentos)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If ERR Then
      MsgErr FName
      ImpCtasLibVentas = -ERR
      Exit Function
   End If
   Sep = vbTab
   ConfigErr = False
         
   'Campos: 'Rut Proveedor' vbTab 'Razon Social' vbTab 'Prop. IVA (blanco, T, N, P)' vbTab 'Cod. Cta. Afecto' vbTab 'Cod. Cta. Exento' vbTab 'Cod. Cta. Total' vbTab 'ANegAfecto' vbTab 'ANegExento' vbTab 'ANegTotal' vbTab 'CCostoAfecto' vbTab 'CCostoExento' vbTab 'CCostoTotal'

         
   Do Until EOF(Fd)
         
      Line Input #Fd, Buf
      l = l + 1
      
               
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If
      
      p = 1
   
      'RUT Proveedor
      RutProv = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutProv) Then
         Call AddLogImp(FNameLogImp, FName, l, "RUT inválido.")
         ConfigErr = True
         GoTo NextRec
      End If

      NuevaEntidad = False
      IdEntidad = GetIdEntidad(RutProv, RazonSocial, NotValidRut)
      
      'Razón Social
      RazonSocialProv = Trim(NextField2(Buf, p, Sep))
      If RazonSocialProv = "" And IdEntidad = 0 Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta razón social, dado que el RUT no ha sido creado en el sistema.")
         ConfigErr = True
         GoTo NextRec
      End If
      
      If IdEntidad = 0 Then
         If Not AddEntidad(RutProv, RazonSocialProv, IdEntidad, ENT_CLIENTE) Then
            Call AddLogImp(FNameLogImp, FName, l, "Error al crear nueva entidad: " & RutProv & " " & RazonSocialProv)
            ConfigErr = True
            GoTo NextRec
         End If
         NuevaEntidad = True
         NEntidades = NEntidades + 1
      End If
      
      'Del Giro
      DelGiro = ValSiNo(Trim(NextField2(Buf, p, Sep)))
      
      If DelGiro <> VAL_SI And DelGiro <> VAL_NO Then
         Call AddLogImp(FNameLogImp, FName, l, "Indicador del Giro inválido.")
         ConfigErr = True
         GoTo NextRec
      End If
            
      
      'códigos cuentas
      
      'código cuenta Afecto
      AuxCodCtaAfecto = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      NomCta = ""
      
      If AuxCodCtaAfecto <> "" Then
         AuxIdCtaAfecto = GetIdCuenta(NomCta, AuxCodCtaAfecto, AuxDescCtaAfecto, UltNivel)
         If AuxIdCtaAfecto <= 0 Or Not UltNivel Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de cuenta Afecto inválido")
            ConfigErr = True
            GoTo NextRec
         End If
      Else
         AuxIdCtaAfecto = 0
         AuxDescCtaAfecto = ""
      End If
      
      'código cuenta Exento
      AuxCodCtaExento = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      NomCta = ""
      
      If AuxCodCtaExento <> "" Then
         AuxIdCtaExento = GetIdCuenta(NomCta, AuxCodCtaExento, AuxDescCtaExento, UltNivel)
         If AuxIdCtaExento <= 0 Or Not UltNivel Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de cuenta Exento inválido")
            ConfigErr = True
            GoTo NextRec
         End If
      Else
         AuxIdCtaExento = 0
         AuxDescCtaExento = ""
      End If
            
      'código cuenta Total
      AuxCodCtaTotal = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      NomCta = ""
      
      If AuxCodCtaTotal <> "" Then
         AuxIdCtaTotal = GetIdCuenta(NomCta, AuxCodCtaTotal, AuxDescCtaTotal, UltNivel)
         If AuxIdCtaTotal <= 0 Or Not UltNivel Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de cuenta Total inválido")
            ConfigErr = True
            GoTo NextRec
         End If
      Else
         AuxIdCtaTotal = 0
         AuxDescCtaTotal = ""
      End If
            
      NomCta = ""
      
      'Codigo Area Negocio Afecto (opcional)
      CodANegAfecto = Trim(NextField2(Buf, p))
      
      If CodANegAfecto <> "" Then
         AuxId = GetAreaNegocio(CodANegAfecto)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de área de negocio Afecto inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If
            
      'Codigo Area Negocio Exento (opcional)
      CodANegExento = Trim(NextField2(Buf, p))
      
      If CodANegExento <> "" Then
         AuxId = GetAreaNegocio(CodANegExento)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de área de negocio Exento inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If
            
      'Codigo Area Negocio Total (opcional)
      CodANegTotal = Trim(NextField2(Buf, p))
      
      If CodANegTotal <> "" Then
         AuxId = GetAreaNegocio(CodANegTotal)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de área de negocio Total inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If
            
      'Codigo Centro de Costo Afecto (opcional)
      CodCCostoAfecto = Trim(NextField2(Buf, p))
      
      If CodCCostoAfecto <> "" Then
         AuxId = GetCentroCosto(CodCCostoAfecto)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de centro de costo Afecto inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If
            
      'Codigo Centro de Costo Exento (opcional)
      CodCCostoExento = Trim(NextField2(Buf, p))
      
      If CodCCostoExento <> "" Then
         AuxId = GetCentroCosto(CodCCostoExento)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de centro de costo Exento inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If
            
      'Codigo Centro de Costo Total (opcional)
      CodCCostoTotal = Trim(NextField2(Buf, p))
      
      If CodCCostoTotal <> "" Then
         AuxId = GetCentroCosto(CodCCostoTotal)
         If AuxId <= 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Código de centro de costo Total inválido.")
            ConfigErr = True
            GoTo NextRec
         End If
      End If

      Q1 = "UPDATE Entidades SET"
      Q1 = Q1 & "  EsDelGiro = " & DelGiro
      Q1 = Q1 & ", CodCtaAfectoVta = '" & AuxCodCtaAfecto & "'"
      Q1 = Q1 & ", CodCtaExentoVta = '" & AuxCodCtaExento & "'"
      Q1 = Q1 & ", CodCtaTotalVta = '" & AuxCodCtaTotal & "'"
      Q1 = Q1 & ", CodAreaNegAfecto = '" & CodANegAfecto & "'"
      Q1 = Q1 & ", CodAreaNegExento = '" & CodANegExento & "'"
      Q1 = Q1 & ", CodAreaNegTotal = '" & CodANegTotal & "'"
      Q1 = Q1 & ", CodCCostoAfecto = '" & CodCCostoAfecto & "'"
      Q1 = Q1 & ", CodCCostoExento = '" & CodCCostoExento & "'"
      Q1 = Q1 & ", CodCCostoTotal = '" & CodCCostoTotal & "'"
      Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      
      Call ExecSQL(DbMain, Q1)
      
      NUpdate = NUpdate + 1
       
NextRec:
   Loop

   Close #Fd
   
   Call LoadAll
   
   If ConfigErr Then
      MsgBox1 "Se encontraron algunos registros con errores.", vbExclamation
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If

   Else
      MsgBox1 "Proceso de importación finalizado." & vbCrLf & vbCrLf & " - " & NEntidades & " nuevas" & vbCrLf & vbCrLf & " - " & NUpdate & " actualizaciones", vbInformation
   End If
   
End Function

Private Function Valida(Grid As Object, ByVal Tipo As String) As Boolean
   Dim i As Integer
   
   Valida = False
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDENTIDAD)) = 0 Then
         Exit For
      End If
      
      If Val(Grid.TextMatrix(i, C_IDCUENTA)) <> 0 Then
         
         If Val(Grid.TextMatrix(i, C_EXIGEANEG)) <> 0 And Grid.TextMatrix(i, C_CODAREANEG) = "" Then
            MsgBox1 "Falta definir Área de Negocio en Cuentas de " & Tipo & " para el proveedor " & Grid.TextMatrix(i, C_NOMBRE), vbExclamation
            Exit Function
         End If
      
         If Val(Grid.TextMatrix(i, C_EXIGECCOSTO)) <> 0 And Grid.TextMatrix(i, C_CODCCOSTO) = "" Then
            MsgBox1 "Falta definir Centro de Costo en Cuentas de " & Tipo & " para el proveedor " & Grid.TextMatrix(i, C_NOMBRE), vbExclamation
            Exit Function
         End If
      
      End If
     
   Next i
   
   Valida = True
   
End Function


