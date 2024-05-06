VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de LP Contabilidad"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9000
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   7680
      TabIndex        =   10
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton bt_Usuarios 
         Caption         =   "&Usuarios"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CommandButton bt_Perfiles 
         Caption         =   "&Perfiles"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":0EAE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton bt_UsuarioPrv 
         Caption         =   "Pr&ivilegios"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":1437
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1980
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Bt_Indices 
         Caption         =   "Indices"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":18D2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Valores e Índices"
         Top             =   1980
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Emp 
         Caption         =   "&Empresas"
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":1CDA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_DefRazonesFin 
         Caption         =   "&Razones Fin."
         Height          =   855
         Left            =   60
         Picture         =   "FrmMain.frx":2407
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   1155
      End
   End
   Begin VB.Frame Fr_Invisible 
      Caption         =   "Lo que no se vé"
      ForeColor       =   &H00FF0000&
      Height          =   1515
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      Begin MSComDlg.CommonDialog Cm_FileDlg 
         Left            =   240
         Top             =   900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Pc_Sel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   300
         Picture         =   "FrmMain.frx":2899
         ScaleHeight     =   110.527
         ScaleMode       =   0  'User
         ScaleWidth      =   150
         TabIndex        =   7
         Top             =   240
         Width           =   150
      End
      Begin MSComDlg.CommonDialog Cm_PrtDlg 
         Left            =   1740
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FlexEdGrid2.FEd2Grid FEd2Grid1 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
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
      Begin FlexEdGrid3.FEd3Grid FEd3Grid1 
         Height          =   435
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
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
      Begin MSComDlg.CommonDialog Cm_ComDlg 
         Left            =   840
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Pc_Access 
      DrawStyle       =   5  'Transparent
      Height          =   2835
      Left            =   1260
      Picture         =   "FrmMain.frx":2BFE
      ScaleHeight     =   2775
      ScaleWidth      =   6255
      TabIndex        =   9
      Top             =   60
      Width           =   6315
      Begin VB.Frame Frame2 
         Caption         =   "Invisibles"
         Height          =   1095
         Left            =   3420
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   1575
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   300
            TabIndex        =   17
            Top             =   420
            Width           =   675
         End
      End
      Begin VB.Label Lb_Descargando 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descargando actualizador..."
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Lb_Version 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   315
         Left            =   720
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label La_demo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00A67300&
         Height          =   330
         Left            =   60
         TabIndex        =   15
         Top             =   2340
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.PictureBox Pc_SQLServer 
      Height          =   2835
      Left            =   1260
      Picture         =   "FrmMain.frx":56F8
      ScaleHeight     =   2775
      ScaleWidth      =   6315
      TabIndex        =   18
      Top             =   60
      Width           =   6375
   End
   Begin VB.Menu M_Empresas 
      Caption         =   "&Empresas"
      Begin VB.Menu M_MantEmpresas 
         Caption         =   "&Mantención..."
      End
      Begin VB.Menu MSep_E1 
         Caption         =   "-"
      End
      Begin VB.Menu M_ListEmpresas 
         Caption         =   "&Listado de Empresas..."
      End
      Begin VB.Menu M_ContEmpresas 
         Caption         =   "&Control de Empresas..."
      End
      Begin VB.Menu MSep_E2 
         Caption         =   "-"
      End
      Begin VB.Menu M_ImpEmpHR 
         Caption         =   "Capturar Empresas desde HR..."
      End
      Begin VB.Menu M_ImpEmpFromAccess 
         Caption         =   "Importar Empresas desde LPContabilidad Access..."
      End
      Begin VB.Menu M_ImpEmpLpRemuFromAccess 
         Caption         =   "Capturar Empresas desde LPRemu"
      End
      Begin VB.Menu M_ImpEmpArchivo 
         Caption         =   "Capturar Empresas desde Archivo"
      End
      Begin VB.Menu MSep_E3 
         Caption         =   "-"
      End
      Begin VB.Menu M_RemoveEmpAno 
         Caption         =   "Eliminar Empresa-Año..."
      End
      Begin VB.Menu MSep_E4 
         Caption         =   "-"
      End
      Begin VB.Menu M_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu M_Utilitarios 
      Caption         =   "&Utilitarios"
      Visible         =   0   'False
      Begin VB.Menu M_Importar 
         Caption         =   "&Importar datos empresa desde HyperRenta..."
      End
   End
   Begin VB.Menu M_Valores 
      Caption         =   "&Valores"
      Begin VB.Menu M_Monedas 
         Caption         =   "&Monedas"
         Begin VB.Menu M_Equivalencias 
            Caption         =   "&Equivalencias..."
         End
         Begin VB.Menu M_ConfigMonedas 
            Caption         =   "&Configuración..."
         End
      End
      Begin VB.Menu M_Indices 
         Caption         =   "&Valores e Índices..."
      End
   End
   Begin VB.Menu MC__Config 
      Caption         =   "&Configuración"
      Begin VB.Menu MC_Oficina 
         Caption         =   "Datos Oficina..."
      End
      Begin VB.Menu MC_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MC_RazonesFin 
         Caption         =   "Razones Financieras..."
      End
      Begin VB.Menu MC_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MC_Usuarios 
         Caption         =   "&Usuarios..."
      End
      Begin VB.Menu MC_Perfiles 
         Caption         =   "&Perfiles..."
      End
      Begin VB.Menu MC_Privilegios 
         Caption         =   "Pr&ivilegios de Usuarios por Empresa..."
      End
      Begin VB.Menu MC_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MC_CambiarClave 
         Caption         =   "&Cambiar clave administrador..."
      End
      Begin VB.Menu MC_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MC_CrearUsrFisc 
         Caption         =   "Crear/Actualizar usuarios Fiscalizadores..."
      End
      Begin VB.Menu MC_HabilitarRecSql 
         Caption         =   "Habilitar recuperacion de datos SQL..."
      End
   End
   Begin VB.Menu M_Sistema 
      Caption         =   "&Sistema"
      Begin VB.Menu M_SetupPrt 
         Caption         =   "Preparar &Impresora..."
      End
      Begin VB.Menu Sep_S1 
         Caption         =   "-"
      End
      Begin VB.Menu M_Backup 
         Caption         =   "&Respaldos..."
      End
      Begin VB.Menu Sep_S2 
         Caption         =   "-"
      End
      Begin VB.Menu MC_Desbloquear 
         Caption         =   "Desbloquear Conexión..."
      End
      Begin VB.Menu Sep_S3 
         Caption         =   "-"
      End
      Begin VB.Menu MC_SolicCod 
         Caption         =   "Licenciar Producto..."
      End
      Begin VB.Menu MC_AutEquipos 
         Caption         =   "Ingresar Código de Licencia..."
         Visible         =   0   'False
      End
      Begin VB.Menu MC_Desactivar 
         Caption         =   "Desactivar Licencia..."
      End
      Begin VB.Menu Sep_S4 
         Caption         =   "-"
      End
      Begin VB.Menu MC_MantDB 
         Caption         =   "Mantención Base de Datos"
         Begin VB.Menu MC_Reparar 
            Caption         =   "Reparar..."
         End
         Begin VB.Menu MC_Compactar 
            Caption         =   "Compactar..."
         End
      End
      Begin VB.Menu MC_ActTxtTilde 
         Caption         =   "Actualizar Textos con Tilde..."
      End
   End
   Begin VB.Menu MH_Help 
      Caption         =   "Ayuda"
      Begin VB.Menu MH_Manual 
         Caption         =   "Manual LP Administrador..."
      End
      Begin VB.Menu MH_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu MH_HlpBackup 
         Caption         =   "Ayuda Respaldo..."
      End
      Begin VB.Menu MH_Export 
         Caption         =   "Exportar Empresa..."
      End
      Begin VB.Menu MH_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MH_DownLast 
         Caption         =   "Descargar actualización..."
      End
      Begin VB.Menu MH_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MH_RepError 
         Caption         =   "Reporte de problema..."
      End
      Begin VB.Menu MH_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MH_AcercaDe 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_DefRazonesFin_Click()
   Call MC_RazonesFin_Click
End Sub

Private Sub Bt_Emp_Click()
   Call M_MantEmpresas_Click
End Sub

Private Sub Bt_Indices_Click()
   Call M_Indices_Click
End Sub

Private Sub bt_Perfiles_Click()
   Call MC_Perfiles_Click
   
End Sub

Private Sub bt_UsuarioPrv_Click()
    Call MC_Privilegios_Click
End Sub

Private Sub bt_Usuarios_Click()
   Call MC_Usuarios_Click
   
End Sub


Private Sub Form_Load()

   Call AddDebug("FrmMain: LLegamos a Form_Load 1", 1)

   Set gFrmMain = Me
   
   Call AddDebug("FrmMain: LLegamos a Form_Load 2", 1)
   
   Call SetCaption
      
   Call AddDebug("FrmMain: Pasamos Caption", 1)
   
   'FEdGrid1.TextMatrix(0, 0) = "$1#2¿P" ' No borrar
   'FEdGrid1.TextMatrix(0, 0) = "" ' No borrar
   
   FEd2Grid1.TextMatrix(0, 0) = "$1#2¿P" ' No borrar
   FEd2Grid1.TextMatrix(0, 0) = "" ' No borrar
   
   FEd3Grid1.TextMatrix(0, 0) = "$7#3?F#" ' No borrar
   FEd3Grid1.TextMatrix(0, 0) = "" ' No borrar
   
   
 '  M_ListEmpresas.Enabled = ChkVMant(VMANT_2005) Se dejo para todos =
   
   M_RemoveEmpAno.Enabled = ChkPriv(PRV_ADM_SIS)
   
#If DATACON = 2 Then       'SQL
   MC_MantDB.Visible = False
   MSep_E3.Visible = False
   M_Backup.Visible = False
   Sep_S2.Visible = False
   MH_Export.Visible = False
   
#Else                      'Access
   MC_ActTxtTilde.Visible = False
   
#End If
   
'   Lb_Version = "Versión " & App.Major & "." & App.Minor
   Lb_Version = "Versión " & App.Major & "." & App.Minor & " " & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
      
   If gAppCode.Demo And Not APP_DEMO Then
      MsgBox1 "Este programa no está registrado y funcionará en modo demo.", vbInformation
   End If
   
'   If APP_DEMO Then     'si es la versión DEMO del programa (en duro)
'      MC_AutEquipos.Enabled = False
'   End If
    
    If gAppCode.Demo And Not APP_DEMO Then
    M_ImpEmpArchivo.Enabled = False
    
    Else
    M_ImpEmpArchivo.Enabled = True
    End If
   
   La_demo.Visible = APP_DEMO
   
   If Val(DBEngine.Version) > 3.51 Then
      MC_Reparar.Visible = False
   End If
   
   If gDbType = SQL_ACCESS Then
      Pc_SQLServer.Visible = False
      Pc_Access.Visible = True
      M_ImpEmpFromAccess.Visible = False
      '625532
      MC_HabilitarRecSql.Visible = False
      '625532
   Else
      Pc_SQLServer.Visible = True
      Pc_Access.Visible = False
      M_ImpEmpFromAccess.Visible = True
      '625532
      MC_HabilitarRecSql.Visible = True
      '625532
   End If

   
   Call AddDebug("FrmMain: Nos Vamos", 1)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Call AddDebug("FrmMain_Unload: 1", 1)

   Call CloseDb(DbMain)
   Call CheckRs(True)
   
   Call AddDebug("FrmMain_Unload: 3", 1)
   
   End
   
End Sub


Private Sub M_ImpEmpArchivo_Click()
Dim Frm As FrmEmpArchivo
   Dim Rc As Integer
   
   Set Frm = New FrmEmpArchivo
   Rc = Frm.FSelect
   Set Frm = Nothing
   
'   If Rc = vbOK Then
'      Dim FrmE As FrmMantEmpresa
'      Set FrmE = New FrmMantEmpresa
'      Call FrmE.FNew(0, gEmprHR.EmpConta.Rut, gEmprHR.EmpConta.NombreCorto)
'      Set FrmE = Nothing
'   End If
End Sub

#If DATACON = 2 Then       'SQL

Private Sub M_ImpEmpFromAccess_Click()
   'Call ImpListEmpFromAccess
    Dim Resp As Long
    Dim rutaOri, rutaDest As String
    Dim fso As Object
    Dim CopyFile As String
    Dim desteny As String


'    Dim Directorio As String
'    With CreateObject("WScript.Shell")
'    Directorio = .SpecialFolders("Mydocuments") & "\"
'    End With
    If Not gUsuario.Nombre = gAdmUser Then
        MsgBox1 "Este proceso solo lo debe hacer el usuario administ", vbExclamation
        Exit Sub
   End If


    gDbPathRespaldoTras = Replace(Replace(gDbPath, "Datos", "RespaldoTras"), "LPContabSQL", "LPContab") '"C:\HR\LPContabilidadProd\Software_mio\Contabilidad70\Respaldo"

    If Dir(gDbPathRespaldoTras, vbDirectory) = "" Then
        MkDir gDbPathRespaldoTras
    End If

    '3391062 ffv
    CopyFile = Replace(gDbPath, "LPContabSQL", "LPContab") '"directory\*" Or copyfile = "file location"
    'CopyFile = gDbPath '"directory\*" Or copyfile = "file location"
    '3391062 ffv

    desteny = gDbPathRespaldoTras '"final directory"
    Set fso = CreateObject("Scripting.FileSystemObject")
    'fso.CopyFolder copyfile, desteny, True
    'Set fso = Nothing


'    If Dir(gDbPathRespaldoTras & "\Empresas", vbDirectory) <> "" Then
'        If MsgBox1("El sistema ya tiene un respaldo" & vbNewLine & vbNewLine & "¿Lo desea sobrescribir?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'          fso.CopyFolder CopyFile, desteny, True
'        End If
'    Else
'        If MsgBox1("El sistema debe hacer un respaldo de la base de datos de Access, la demora dependera del tamaño de las bases." & vbNewLine & vbNewLine & "¿Desea Continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'           fso.CopyFolder CopyFile, desteny, True
'        Else
'           Exit Sub
'        End If
'    End If

    'MsgBox "Se creara un respaldo para el buen funcionamiento del traspaso" & vbNewLine & vbNewLine & "Favor esperar hasta que el respaldo termine.", vbInformation, "Sistema de Respaldo"
    MsgBox "Se creará un respaldo para el buen funcionamiento del traspaso" & vbNewLine & vbNewLine & "Este proceso puede tardar varios minutos, favor esperar hasta que finalice.", vbInformation, "Sistema de Respaldo"

    fso.CopyFolder CopyFile, desteny, True
    Set fso = Nothing


    gDbPathAux = gDbPath
    gDbPath = gDbPathRespaldoTras
    Call ActualizarBase
    Call ImpListEmpFromAccess
    Dim Frm As FrmSelEmpresasTras
    Dim Rc As Integer
    Set Frm = New FrmSelEmpresasTras
    Rc = Frm.FSelect
    Set Frm = Nothing
    
    'If Rc = vbOK Then
End Sub

#End If
Private Sub ActualizarBase()

    Dim oCarpeta, oCarpetaEmpresa, oCarpetaAno As Object
    Dim oArchivo, oArchivoEmpresa, oArchivoAno As Object
    Dim i As Integer
    Dim RutaEmpresas As String
    Dim fso, fsoEmpresa, fsoAno As Object
    Dim RutaEmpresa, RutaAno As String
    Dim archivoODir() As String
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oCarpeta = fso.GetFolder(gDbPathRespaldoTras)
    
    For Each oArchivo In oCarpeta.SubFolders
        
        Set fsoEmpresa = CreateObject("Scripting.FileSystemObject")
        RutaEmpresa = gDbPathRespaldoTras & "\" & oArchivo.Name
        archivoODir = Split(oArchivo.Name, ".")
        
        Set oCarpetaEmpresa = fsoEmpresa.GetFolder(RutaEmpresa)
        For Each oArchivoEmpresa In oCarpetaEmpresa.SubFolders
        
            Set fsoAno = CreateObject("Scripting.FileSystemObject")
            RutaAno = RutaEmpresa & "\" & oArchivoEmpresa.Name
            Set oCarpetaAno = fsoAno.GetFolder(RutaAno)
            For Each oArchivoAno In oCarpetaAno.Files
                Call OpenDbEmp(Replace(oArchivoAno.Name, ".mdb", ""), oArchivoEmpresa.Name)
                Call CorrigeBase
                'ffv 25-01-24
               ' Call CloseDb(DbMain)
            Next oArchivoAno
            Set fsoAno = Nothing
        
        
        Next oArchivoEmpresa
        Set fsoEmpresa = Nothing
       
     
    Next oArchivo
    Set fso = Nothing
End Sub


Private Sub M_ImpEmpHR_Click()
   Dim Frm As FrmEmpHR
   Dim Rc As Integer
   
   Set Frm = New FrmEmpHR
   Rc = Frm.FSelect
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Dim FrmE As FrmMantEmpresa
      Set FrmE = New FrmMantEmpresa
      Call FrmE.FNew(0, gEmprHR.EmpConta.Rut, gEmprHR.EmpConta.NombreCorto)
      Set FrmE = Nothing
   End If


End Sub

Private Sub M_ImpEmpLpRemuFromAccess_Click()
 'Call ImpListEmpLpRemuFromAccess
 
 Dim Frm As FrmEmpLpRemu
   Dim Rc As Boolean
   
   Set Frm = New FrmEmpLpRemu
    Frm.Show
   
   If Frm.lRc = vbCancel Then
    
    Unload Frm
  Set Frm = Nothing
    'Else

   End If
 
End Sub

Private Sub M_RemoveEmpAno_Click()
   Dim Frm As FrmResetEmprAno
   
   Set Frm = New FrmResetEmprAno
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MC_ActTxtTilde_Click()
#If DATACON = 2 Then
   Me.MousePointer = vbHourglass
      
   Call CorrigeTextosConAcentosScriptUTF8
   
   Me.MousePointer = vbDefault

#End If
   
End Sub

Private Sub MC_Compactar_Click()
   Dim ConnStr As String

#If DATACON = 1 Then

   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en el sistema." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   
   'ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
   If CompactDb2(DbMain, True, gComunConnStr) = 0 Then 'no hubo error
      If OpenDbAdm() = False Then
         End
      End If
   Else
      MsgBox1 "Problemas al tratar de compactar la base de datos.", vbExclamation + vbOKOnly
   End If
   
   Me.MousePointer = vbDefault
#End If

End Sub

Private Sub MC_CrearUsrFisc_Click()
   Dim Frm As FrmCrearUsrFiscalizador
   
   Set Frm = New FrmCrearUsrFiscalizador
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub MC_Desactivar_Click()
   Dim i As Integer, r As Integer, Buf As String

   If MsgBox1("ATENCIÓN" & vbCrLf & "¿Está seguro que desea desactivar TODAS las licencias que tiene habilitadas?" & vbCrLf & "Al desincribir este programa, éste funcionará sólo en modo DEMO, por lo tanto no tendrá acceso a toda la información.", vbYesNo Or vbExclamation) = vbNo Then
      Exit Sub
   End If
   
   On Error Resume Next
   
   r = 0
   For i = 1 To 100
   
      If GetIniString(gLicFile, PC_EQUIP, PC_NOM & i) <> "" Then
         r = r + 1
         
         Call SetIniString(gLicFile, PC_EQUIP, PC_AUT & i, FwEncrypt1("N" & r, KEY_CRYP + i * 155))
         
      End If
   Next i
      
   If r > 0 Then
      Buf = 717171
      Call SetIniString(gLicFile, PC_INFO, PC_NIV & 3, FwEncrypt1(Buf, KEY_CRYP + 3147))
      Buf = 868686
      Call SetIniString(gLicFile, PC_INFO, PC_NLIC & 3, FwEncrypt1(Buf, KEY_CRYP + 5043))
      Buf = "5RTREQTWWQ"
      Call SetIniString(gLicFile, PC_INFO, PC_NCOD & 1, FwEncrypt1(Buf, KEY_CRYP + 2345))
   End If
      
   gAppCode.Demo = True
   
   Call SetCaption

End Sub

Private Sub MC_Desbloquear_Click()
   Dim Frm As FrmDesbloquear
   
   Set Frm = New FrmDesbloquear
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

'625532 si codigo es 1 habilitara el boton REC documentos SQL en formulario FrmConfig "configuracion Inial" LPcontab si es  0 se encuentra deshabilitado
Private Sub MC_HabilitarRecSql_Click()
Dim Q1 As String
Dim Rs As Recordset

  Q1 = "SELECT Tipo, Codigo FROM PARAM "
  Q1 = Q1 & " WHERE Tipo = 'BTRECUSQL' "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
     If vFld(Rs("Codigo")) = 1 Then
     '  MsgBox1("ATENCIÓN" & vbNewLine & vbNewLine & "Boton se encuentra Activo:" & vbNewLine & vbNewLine & vbNewLine & "¿Desea Deshabilitar Boton?", vbQuestion + vbYesNo + vbDefaultButton2)
        If MsgBox1("ATENCIÓN" & vbNewLine & vbNewLine & "Boton se encuentra Habilitado:" & vbNewLine & vbNewLine & vbNewLine & "¿Desea Deshabilitar Boton?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         
         Q1 = ""
         Q1 = Q1 & " UPDATE PARAM set Codigo = '0'"
         Q1 = Q1 & " WHERE TIPO = 'BTRECUSQL' "
         Q1 = Q1 & " AND Codigo = '1'"
         
         Call ExecSQL(DbMain, Q1)
          
        End If
     ElseIf vFld(Rs("Codigo")) = 0 Then
         
        If MsgBox1("ATENCIÓN" & vbNewLine & vbNewLine & "Boton se encuentra Deshabilitado:" & vbNewLine & vbNewLine & vbNewLine & "¿Desea Habilitar Boton?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         
         Q1 = ""
         Q1 = Q1 & " UPDATE PARAM set Codigo = '1'"
         Q1 = Q1 & " WHERE TIPO = 'BTRECUSQL' "
         Q1 = Q1 & " AND Codigo = '0'"
         
         Call ExecSQL(DbMain, Q1)
          
        End If
     
     End If
   Else
      
      If MsgBox1("ATENCIÓN" & vbNewLine & vbNewLine & "Boton se encuentra Deshabilitado:" & vbNewLine & vbNewLine & vbNewLine & "¿Desea Habilitar Boton?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         
         Q1 = ""
         Q1 = Q1 & " INSERT INTO PARAM"
         Q1 = Q1 & " (TIPO,CODIGO,VALOR)"
         Q1 = Q1 & " VALUES"
         Q1 = Q1 & " ('BTRECUSQL',"
         Q1 = Q1 & " '1',"
         Q1 = Q1 & " 'BOTON RECUPERACION SQL')"
         
         Call ExecSQL(DbMain, Q1)
          
        End If
   
   End If
   
   Call CloseRs(Rs)
 

End Sub
'625532

Private Sub MC_Oficina_Click()
   Dim Frm As FrmOficina
   
   Set Frm = New FrmOficina
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MC_RazonesFin_Click()
   Dim Frm As FrmRazones
   
   Set Frm = New FrmRazones
   Call Frm.FDefinir
   Set Frm = Nothing

End Sub

Private Sub MC_Reparar_Click()
   Dim DbPath As String
   
#If DATACON = 1 Then
   If MsgBox1("Antes de realizar esta operación, verifique que no haya ningún usuario trabajando en el sistema." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   DbPath = DbMain.Name
   
   Call CloseDb(DbMain)
   
   If RepairDb(DbPath) Then
      If OpenDbAdm() = False Then
         End
      End If
      Me.MousePointer = vbDefault
   Else
      Unload Me
      End
   End If
#End If

End Sub

Private Sub MH_AcercaDe_Click()
   Dim Frm As FrmAbout
   
   Set Frm = New FrmAbout
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_Backup_Click()
   Dim Frm As FrmGenBackup
   
   Set Frm = New FrmGenBackup
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub M_ConfigMonedas_Click()
   Dim Frm As FrmMonedas
   
   Set Frm = New FrmMonedas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ContEmpresas_Click()
   Dim Frm As FrmRepContEmpresas
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmRepContEmpresas
   Frm.Show vbModal
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault

End Sub

Private Sub M_Equivalencias_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Frm.FEdit (0)
   Set Frm = Nothing

End Sub

Private Sub M_Indices_Click()
   Dim Frm As FrmIPC
   
   Set Frm = New FrmIPC
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_ListEmpresas_Click()
   Dim Frm As FrmPrtEmpresas
   
   MousePointer = vbHourglass
   Set Frm = New FrmPrtEmpresas
   Frm.Show vbModal
   Set Frm = Nothing
   MousePointer = vbDefault
   
End Sub

Private Sub M_MantEmpresas_Click()
   Dim Frm As FrmEmpresas
   
   Set Frm = New FrmEmpresas
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub


Private Sub MC_AutEquipos_Click()
   Dim Frm As FrmEquiposAut
   
   If APP_DEMO Then     'si es la versión DEMO del programa (en duro)
      MsgBox1 "Esta versión del programa siempre funciona en modo Demo." & vbCrLf & "Si usted ya tiene una licencia, baje la actualización desde el menú" & vbCrLf & "o desde el sitio web y luego ingrese el código de licencia.", vbInformation
      Exit Sub
   End If

   If gOficina.Rut = "" Then
      MsgBox1 "Debe ingresar el RUT de su empresa en el menú Configuración >> Datos Oficina.", vbExclamation
      Exit Sub
   End If
   
   Set Frm = New FrmEquiposAut
   Call Frm.Admin
   Set Frm = Nothing

   Call SetCaption

End Sub

Private Sub MC_Perfiles_Click()
   Dim Frm As FrmPerfiles
   
   Set Frm = New FrmPerfiles
   Call Frm.ShowPerfiles(True)
   Set Frm = Nothing
   
End Sub

Private Sub MC_Privilegios_Click()
   Dim Frm As FrmUsuarioPriv
   
   Set Frm = New FrmUsuarioPriv
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub M_Salir_Click()
   Unload Me
End Sub

Private Sub M_SetupPrt_Click()
   Dim CurrPrt As String
   Dim Rc As Integer
   
   If PrepararPrt(Cm_PrtDlg) Then
   
      Call SetIniString(gIniFile, "Config", "Printer", Printer.DeviceName)
   Else
      Call FindPrinter(GetIniString(gIniFile, "Config", "Printer"), True)
    
   End If
   
   'CurrPrt = Printer.DeviceName
   'Set Printer = FindPrinter(CurrPrt)

End Sub

Private Sub MC_SolicCod_Click()
   Dim Frm As FrmEquiposAut
   
   If gOficina.Rut = "" Then
      MsgBox1 "Debe ingresar la Razón Social y el RUT de su empresa en el menú Configuración >> Datos Oficina.", vbExclamation
      Exit Sub
   End If

   Set Frm = New FrmEquiposAut
   Call Frm.Solicitud
   Set Frm = Nothing

End Sub

Private Sub MC_Usuarios_Click()
   Dim Frm As FrmUsuarios
   
   Set Frm = New FrmUsuarios
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub



Private Sub MH_DownLast_Click()
   Static bDown As Boolean
   
   If bDown Then
      Exit Sub
   End If
   bDown = True

   Lb_Descargando.Visible = True
   MousePointer = vbHourglass
   DoEvents

   If Trim(gAppCode.Rut) = "" Then
      gAppCode.Rut = gOficina.Rut
   End If
   
   Call FwDownLast(Me, Cm_FileDlg, APP_DEMO)

   MousePointer = vbDefault
   Lb_Descargando.Visible = False
   DoEvents

   bDown = False

End Sub


Private Sub MH_HlpBackup_Click()
   Dim Frm As FrmBackup
   
   Set Frm = New FrmBackup
   Frm.Show vbModal
   Set Frm = Nothing
   

End Sub

Private Sub MH_Manual_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_LP_Administrador.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual del Administrador de LP Contabilidad, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual del Administrador de LP Contabilidad." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub MH_RepError_Click()
   Dim Frm As FrmRepError
   
   Set Frm = New FrmRepError
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub MC_CambiarClave_Click()
   Dim Frm As FrmCambioClave
   
   Set Frm = New FrmCambioClave
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub SetCaption()

   Me.Caption = "Administrador " & gLexContab & " - " & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")

   If gAppCode.Demo Then
      Me.Caption = Me.Caption & " - DEMO"
   End If


End Sub

Private Sub MH_Export_Click()
   Dim Frm As FrmExportEmp
   
   Set Frm = New FrmExportEmp
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

