VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmExportEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Empresa para Solicitar Soporte"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   ForeColor       =   &H00C00000&
   Icon            =   "FrmExportEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Ls_Ano 
      Height          =   1035
      Left            =   7200
      TabIndex        =   8
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Frame Fr_Sort 
      Caption         =   "Ordenar por"
      Height          =   975
      Left            =   7260
      TabIndex        =   6
      Top             =   3660
      Width           =   1275
      Begin VB.OptionButton Op_SortRUT 
         Caption         =   "RUT"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   795
      End
      Begin VB.OptionButton Op_SortNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Bt_Export 
      Caption         =   "&Exportar"
      Height          =   735
      Left            =   7260
      MousePointer    =   99  'Custom
      Picture         =   "FrmExportEmp.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nueva empresa"
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7260
      TabIndex        =   4
      Top             =   540
      Width           =   1200
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   4635
      Left            =   1500
      TabIndex        =   0
      Top             =   540
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8176
      Cols            =   4
      Rows            =   20
      FixedCols       =   0
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   1
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "NOTA: el sistema genera un archivo compactado con el año seleccionado y el año anterior"
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1500
      TabIndex        =   9
      Top             =   5400
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Año:"
      Height          =   255
      Left            =   7260
      TabIndex        =   7
      Top             =   1920
      Width           =   795
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   570
      Index           =   0
      Left            =   420
      Picture         =   "FrmExportEmp.frx":03BB
      Top             =   540
      Width           =   570
   End
   Begin VB.Label La_demo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DEMO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002A01A6&
      Height          =   330
      Left            =   7200
      TabIndex        =   5
      Top             =   4860
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "FrmExportEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_NOMBRECORTO = 1
Const C_ID = 2
Const C_ESTADO = 3    'no Activo


Private Const C_NOMLARGO = 2
Private Const C_IDEMPRESA = 3
Private Const C_IDPERFIL = 4
Private Const C_PRIV = 5

Private Const C_FCIERRE = 2
Private Const C_FAPERTURA = 3

Dim lRc As Integer
Dim LsAno As ClsCombo

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Set LsAno = New ClsCombo
   Call LsAno.SetControl(Ls_Ano)

   Call SetUpGrid
   Call LoadAll
      
  ' Bt_Del.Enabled = ChkVMant(VMANT_2005) se dejo para todos =
   
   La_demo.Visible = gAppCode.Demo
   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_RUT) = 1500
   Grid.ColWidth(C_NOMBRECORTO) = 2500
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESTADO) = 1000
      
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRECORTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_NOMBRECORTO) = "Nombre Corto"
   Grid.TextMatrix(0, C_ESTADO) = "Activa"
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Q1 = "SELECT IdEmpresa, Rut, NombreCorto, Estado FROM Empresas"
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')"
   End If
   
   If Op_SortNombre Then
      Q1 = Q1 & " ORDER BY NombreCorto"
   Else
      Q1 = Q1 & " ORDER BY Val(Rut)"
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 1
   Grid.Rows = Row
   Do While Rs.EOF = False
      Grid.Rows = Row + 1
      
      Grid.TextMatrix(Row, C_RUT) = FmtCID(vFld(Rs("Rut")))
      Grid.TextMatrix(Row, C_NOMBRECORTO) = vFld(Rs("NombreCorto"))
      Grid.TextMatrix(Row, C_ESTADO) = IIf(vFld(Rs("Estado")) = 0, "Si", "No")
      
      Grid.Row = Row
      Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdEmpresa"))
      
      Row = Row + 1
      
      If gAppCode.NivProd = VER_5EMP And Row > 5 Then
         Exit Do
      End If
      
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
      
End Sub

Private Sub Grid_Click()
   Call LstAnosEmp
End Sub

Private Sub Grid_DblClick()
   Call Bt_Export_Click
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption)
   End If
End Sub

Private Sub Op_SortNombre_Click()

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub

Private Sub Op_SortRUT_Click()
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub

Private Sub LstAnosEmp()
   Dim Q1 As String
   Dim Ano As Integer, MaxAno As Integer
   Dim Rs As Recordset
   Dim UltAño As Integer
   Dim i As Integer
   Dim AnoTope As Integer
   Dim Row As Integer
   Dim idEmp As Long
   Dim Rut As String
      
   Call AddDebug("ExportEmp.LstAnosEmp: 1", 1)
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) = "" Then
      Exit Sub
   End If
   
   idEmp = Grid.TextMatrix(Row, C_ID)
   Rut = Grid.TextMatrix(Row, C_RUT)
      
   LsAno.Clear
   Ano = Year(Int(Now))
      
   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & idEmp
      
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MaxAno = vFld(Rs("MaxAno"))
   End If
   Call CloseRs(Rs)
   
   If MaxAno <= 0 Then
      MaxAno = Year(Now)
   End If
   
   Call AddDebug("ExportEmp.LstAnosEmp: 2 - " & MaxAno, 1)
   
   Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & idEmp
   Q1 = Q1 & " ORDER BY Ano DESC"
   Set Rs = OpenRs(DbMain, Q1)
   
   Call AddDebug("ExportEmp.LstAnosEmp: 3", 1)
   
   AnoTope = Year(Now) + 2
   
   For i = AnoTope To 2000 Step -1
      Call LsAno.AddItem(i, i, 0, 0)
      If i = MaxAno Then
         LsAno.ListIndex = LsAno.NewIndex
      End If
   
      Do Until Rs.EOF
         Ano = vFld(Rs("Ano"))
         
         If Ano > AnoTope Then ' años muy del futuro o tiene mal la fecha del computador
            Exit Do
         End If
         
         If i = Ano Then
            LsAno.List(LsAno.NewIndex) = Ano & " *"
            LsAno.Matrix(C_FCIERRE, LsAno.NewIndex) = vFld(Rs("FCierre"))
            LsAno.Matrix(C_FAPERTURA, LsAno.NewIndex) = vFld(Rs("FApertura"))
            Rs.MoveNext
            Exit Do
         ElseIf i > Ano Then
            Exit Do
         End If
      Loop
   
   Next i
   
   Call CloseRs(Rs)
   
   Call AddDebug("ExportEmp.LstAnosEmp: FIN", 1)

End Sub

Private Sub Bt_Export_Click()
   Static bExporting As Boolean
   Dim Rut As String, NCorto As String
   Dim IdEmpresa As Long
   Dim Row As Integer, Ano As Integer
   Dim FnEmpr As String, i As Integer, FnZip As String, Fn As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) = "" Then
      Exit Sub
   End If
   
   If LsAno.ListIndex < 0 Then
      Exit Sub
   End If
   
   If bExporting Then
      Exit Sub
   End If
   bExporting = True
   
   MousePointer = vbHourglass
   DoEvents
   
   IdEmpresa = Grid.TextMatrix(Row, C_ID)
   Rut = Grid.TextMatrix(Row, C_RUT)
   NCorto = Grid.TextMatrix(Row, C_NOMBRECORTO)

   Ano = Val(LsAno.ItemData)

'   FnEmpr = DbMain.Name
'   i = rInStr(FnEmpr, "\")
   
'   If i <= 0 Then
'      MousePointer = vbDefault
'      bExporting = False
'      Exit Sub
'   End If
'
'   FnEmpr = Mid(FnEmpr, i + 1)
   

   FnEmpr = vFmtRut(Rut) & ".mdb"
   
   i = Len(FnEmpr)
   FnZip = "LPContab_" & Left(FnEmpr, i - 4) & "_" & Format(Now, "yymmdd") & ".zip"
   
   On Error Resume Next
   FrmMain.Cm_FileDlg.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
   FrmMain.Cm_FileDlg.CancelError = True
   FrmMain.Cm_FileDlg.FileName = FnZip
   FrmMain.Cm_FileDlg.InitDir = W.TmpDir
   FrmMain.Cm_FileDlg.ShowSave
   
   If Err.Number Then
      MousePointer = vbDefault
      bExporting = False
      Exit Sub
   End If
   
   FnZip = FrmMain.Cm_FileDlg.FileName
   
   Fn = GenDbZip(FnEmpr, Ano, FnZip)
   i = rInStr(Fn, "\")
   
   If Len(Fn) > 0 Then
      If MsgBox1("Se generó el archivo" & vbCrLf & Fn & vbCrLf & vbCrLf & "¿ Desea abrir la carpeta del archivo ?", vbInformation Or vbYesNo) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", Left(Fn, i), "", "", SW_SHOW)
      End If
   End If
   
   MousePointer = vbDefault
   bExporting = False
   
End Sub

Public Function GenDbZip(ByVal FnEmpr As String, ByVal Ano As Integer, ByVal ZipFile As String) As String
   'Dim ZipFile As String
   Dim zOpt As ZipOPT_t
   Dim zFiles As ZIPnames_t
   Dim zFnc As ZIPUSERFUNCTIONS_t
   Dim Rc As Long, KBytes As Long
   Dim FileName As String
   Dim i As Integer, nFiles As Integer
   Dim FAnoAnt As String
   
   nFiles = 1
   zFiles.zFiles(0) = gDbPath & "\" & BD_COMUN
   
   nFiles = nFiles + 1
   zFiles.zFiles(1) = gDbPath & "\Empresas\" & Ano & "\" & FnEmpr
   
   FAnoAnt = gDbPath & "\Empresas\" & Ano - 1 & "\" & FnEmpr
   If ExistFile(FAnoAnt) Then
      nFiles = nFiles + 1
      zFiles.zFiles(2) = FAnoAnt
   End If
   
   zOpt.Date = vbNullString
   zOpt.flevel = Asc(9)  ' Compression Level (0 - 9)
   zOpt.szRootDir = gDbPath

   i = Len(FnEmpr)

   If Len(ZipFile) < 1 Then
      FileName = "LPContab_" & Left(FnEmpr, i - 4) & "_" & Format(Now, "yymmdd") & ".zip"
      ZipFile = W.TmpDir & "\" & FileName
   End If
   
   On Error Resume Next

   Rc = VBZip32(ZipFile, nFiles, zFiles, zOpt, zFnc)
   
   If Err Then
      MsgErr "No se puede generar el archivo " & ZipFile
      Exit Function
   End If
   
   If Rc Then
      Call AddLog("GenZip: Error " & Rc & " al generar el archivo " & ZipFile & ", DLL_Err=" & Err.LastDllError)
      MsgBox1 "Error " & Rc & " al generar el archivo " & ZipFile, vbCritical
   Else
      KBytes = FileLen(ZipFile) / 1024
'      Tx_XlsFile = Tx_XlsFile & vbCrLf & lZipFile & vbCrLf & "Tamaño: " & Format(KBytes, NUMFMT) & " KBytes"
      
      GenDbZip = ZipFile
   End If

End Function


