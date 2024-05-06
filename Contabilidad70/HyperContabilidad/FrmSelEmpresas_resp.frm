VERSION 5.00
Begin VB.Form FrmSelEmpresas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Empresa"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "FrmSelEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Sort 
      Caption         =   "Ordenar por"
      Height          =   975
      Left            =   7260
      TabIndex        =   11
      Top             =   3120
      Width           =   1275
      Begin VB.OptionButton Op_SortNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Op_SortRUT 
         Caption         =   "RUT"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.ListBox Ls_Ano 
      Height          =   1035
      Left            =   7320
      TabIndex        =   1
      Top             =   1620
      Width           =   1155
   End
   Begin VB.ListBox Ls_Empresas 
      Height          =   4935
      Left            =   1620
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Default         =   -1  'True
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   0
      Left            =   420
      Picture         =   "FrmSelEmpresas.frx":000C
      Top             =   480
      Width           =   690
   End
   Begin VB.Label La_nEmp 
      AutoSize        =   -1  'True
      Caption         =   "000"
      Height          =   195
      Left            =   7200
      TabIndex        =   10
      Top             =   5520
      Width           =   270
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
      Left            =   7440
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "(*) con datos"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Año:"
      Height          =   195
      Index           =   2
      Left            =   7320
      TabIndex        =   7
      Top             =   1380
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   1
      Left            =   3060
      TabIndex        =   6
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT"
      Height          =   315
      Index           =   6
      Left            =   1620
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   0
      Left            =   2820
      TabIndex        =   4
      Top             =   480
      Width           =   2115
   End
End
Attribute VB_Name = "FrmSelEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_NOMLARGO = 2
Private Const C_IDEMPRESA = 3
Private Const C_IDPERFIL = 4
Private Const C_PRIV = 5

Private Const C_FCIERRE = 2
Private Const C_FAPERTURA = 3

Dim lRc As Integer
Dim lsEmpresa As ClsCombo
Dim LsAno As ClsCombo



Friend Function FSelect() As Integer
   Me.Show vbModal
   
   FSelect = lRc
End Function
Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub
Private Sub Bt_Sel_Click()
   Dim IdEmpresa As Long
   Dim Ano As Integer
   Dim Rut As String
   Dim Nombre As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 1", 1)
   
   If lsEmpresa.ListIndex < 0 Then
      Exit Sub
   End If
   
   If LsAno.ListIndex < 0 Then
      Exit Sub
   End If
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 2 - " & lsEmpresa.ListIndex & " - " & LsAno.ListIndex, 1)
   
   MousePointer = vbHourglass
   DoEvents
   
   IdEmpresa = Val(lsEmpresa.Matrix(C_IDEMPRESA))
   'Ano = LsAno.List(LsAno.ListIndex)
   Ano = Val(LsAno.ItemData)
   Rut = lsEmpresa.ItemData
   Nombre = lsEmpresa.List2(lsEmpresa.ListIndex)
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 3", 1)

   'Creo o chequeo base de datos de la empresa
   If CrearNuevoAno(IdEmpresa, Ano, Rut, Nombre) = False Then
      MousePointer = vbDefault
      Exit Sub
   End If
   
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: 4", 1)
   
   'ASIGNO DATOS A LA ESTRUCTURA
   gEmpresa.Rut = Rut
   gEmpresa.NombreCorto = Nombre
   gEmpresa.id = IdEmpresa
   gEmpresa.Ano = Ano
'   Debug.Print "gEmpresa.Ano =" & gEmpresa.Ano
   gEmpresa.FCierre = vFmt(LsAno.Matrix(C_FCIERRE))
   gEmpresa.FApertura = vFmt(LsAno.Matrix(C_FAPERTURA))
   
   'esto se hizo especial para la Asoc. de AFP porque tenían varias empresas con el mismo RUT.
   'El campo RutDisp sirve para imprimirlo en los membretes
   Q1 = "SELECT RutDisp FROM Empresas WHERE IdEmpresa = " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)
   gEmpresa.RutDisp = ""
   If Not Rs.EOF Then
      gEmpresa.RutDisp = FwDecrypt1(vFld(Rs("RutDisp")), KEY_CRYP + 10)
   End If
   Call CloseRs(Rs)
   
   gUsuario.idPerfil = lsEmpresa.Matrix(C_IDPERFIL)
   gUsuario.Priv = lsEmpresa.Matrix(C_PRIV)
      
   lRc = vbOK
   
   Unload Me
   
   
   Call AddDebug("FrmSelEmpresas.Bt_Sel_Click: FIN", 1)
   
End Sub

Private Sub Form_Load()

   lRc = vbCancel
   
   Call AddDebug("FrmSelEmpresas: Load", 1)
   
   Set LsAno = New ClsCombo
   Call LsAno.SetControl(Ls_Ano)
   
   Call AddDebug("FrmSelEmpresas: Después de ClsCombo", 1)
   
   If gVarIniFile.SelEmprPorRUT Then
      Op_SortRUT = True    'llama a FillList
   Else
      Op_SortNombre = True    'llama a FillList
   End If
   
   Call AddDebug("FrmSelEmpresas: Después de FillList", 1)
   
   La_demo.visible = gAppCode.Demo
   'Fr_Sort.Visible = Not gAppCode.Demo
   
End Sub

Private Sub FillList()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
      
   Set lsEmpresa = New ClsCombo
   Call lsEmpresa.SetControl(Ls_Empresas)
      
   If gUsuario.Nombre = gAdmUser Then
      Q1 = "SELECT Empresas.idEmpresa, Rut, NombreCorto, 0 as idPerfil, 0 as Privilegios"
      Q1 = Q1 & " FROM Empresas"
      Q1 = Q1 & " WHERE Estado = 0 "   'Empresas activas
      
      If gAppCode.Demo Then
         Q1 = Q1 & " AND RUT IN ('1','2','3')"
      End If
   Else
      Q1 = "SELECT Empresas.idEmpresa, Rut, NombreCorto, UsuarioEmpresa.idPerfil, Perfiles.Privilegios"
      Q1 = Q1 & " FROM (Empresas INNER JOIN UsuarioEmpresa ON Empresas.idEmpresa=UsuarioEmpresa.IdEmpresa)"
      Q1 = Q1 & " LEFT JOIN Perfiles ON UsuarioEmpresa.idPerfil = Perfiles.idPerfil"
      Q1 = Q1 & " WHERE UsuarioEmpresa.idUsuario = " & gUsuario.IdUsuario
      Q1 = Q1 & " AND Empresas.Estado = 0 "   'Empresas activas
   
      If gAppCode.Demo Then
         Q1 = Q1 & " AND RUT IN ('1','2','3')"
      End If
   
   End If
   
   If Op_SortNombre Then
      Q1 = Q1 & " ORDER BY NombreCorto, RUT"
   Else
'      Q1 = Q1 & " ORDER BY right('0' & RUT,8)"

      Q1 = Q1 & " ORDER BY right(" & SqlConcat(gDbType, "'00000000000'", "Rut") & ", 12)"  ' 19 mar 2019 - pam: falla el orden porque algunos tienen puntos y otros no
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      lsEmpresa.AddItem FmtStRut(vFld(Rs("Rut"))) & vbTab & vFld(Rs("NombreCorto"))
      lsEmpresa.ItemData(lsEmpresa.NewIndex) = vFld(Rs("Rut"))
      lsEmpresa.Matrix(C_NOMLARGO, lsEmpresa.NewIndex) = vFld(Rs("NombreCorto"))
      lsEmpresa.Matrix(C_IDEMPRESA, lsEmpresa.NewIndex) = vFld(Rs("idEmpresa"))
      lsEmpresa.Matrix(C_IDPERFIL, lsEmpresa.NewIndex) = vFld(Rs("idPerfil"))
      
      If gUsuario.Nombre = gAdmUser Then
         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = PRV_ADMIN
      Else
         lsEmpresa.Matrix(C_PRIV, lsEmpresa.NewIndex) = vFld(Rs("Privilegios"))
      End If
      
      Rs.MoveNext
      
      If gAppCode.Demo And lsEmpresa.ListCount >= 3 Then
         Exit Do
      End If
      
   Loop
   Call CloseRs(Rs)
   
   La_nEmp = Format(lsEmpresa.ListCount, NUMFMT)
   
'   If gAppCode.NivProd = VER_5EMP And lsEmpresa.ListCount > 5 Then
   If lsEmpresa.ListCount > gMaxEmpLicencia And Not gAppCode.Demo And gDbType = SQL_ACCESS Then
      MsgBox1 "Esta versión sólo permite trabajar con a lo más " & gMaxEmpLicencia & " empresas." & vbCrLf & vbCrLf & "Utilice el administrador para eliminar algunas empresas y así poder utilizar el sistema.", vbExclamation
      Bt_Sel.Enabled = False
   End If
   
End Sub



Private Sub Ls_Ano_DblClick()
   Call PostClick(Bt_Sel)
   
   Call AddDebug("FrmSelEmpresas.Ls_Ano_DblClick: FIN", 1)

End Sub

Private Sub Ls_Empresas_Click_old()
   Dim Q1 As String
   Dim Ano As Integer
   Dim Rs As Recordset
   Dim UltAño As Integer
   Dim i As Integer
   
   'Set LsAno = New ClsCombo
   'Call LsAno.SetControl(Ls_Ano)
   
   LsAno.Clear
   Ano = Year(Int(Now))
   
   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
      
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
            
      If vFld(Rs("MaxAno")) > 0 Then
          
         'agregamos un año más del último, para que lo pueda crear, si lo desea
         LsAno.AddItem vFld(Rs("MaxAno")) + 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.ListIndex = 0
            
         Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
         Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
         Q1 = Q1 & " ORDER BY Ano DESC"
         Call LsAno.FillCombo(DbMain, Q1, -1)
         
         UltAño = Val(LsAno.list(Ls_Ano.ListCount - 1))
         
         For i = 1 To 4
            LsAno.AddItem UltAño - i
            LsAno.ItemData(LsAno.NewIndex) = 0
         Next i
                
         If LsAno.ListCount > 1 Then
            LsAno.ListIndex = 1  'seleccionamos último año existente
         End If
   
      Else
         'No existe ningun año creado, ofrecemos: año-4, año-3, año-2, año-1, año y año+1
         LsAno.AddItem Ano + 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         
         LsAno.AddItem Ano
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.ListIndex = LsAno.NewIndex
         
         LsAno.AddItem Ano - 1
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 2
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 3
         LsAno.ItemData(LsAno.NewIndex) = 0
         LsAno.AddItem Ano - 4
         LsAno.ItemData(LsAno.NewIndex) = 0
               
      End If
      
   End If
   
   Call CloseRs(Rs)
   
End Sub

Private Sub Ls_Empresas_Click()
   Dim Q1 As String
   Dim Ano As Integer, MaxAno As Integer
   Dim Rs As Recordset
   Dim UltAño As Integer
   Dim i As Integer
   Dim AnoTope As Integer
      
   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 1", 1)
      
   LsAno.Clear
   Ano = Year(Int(Now))
   
   Q1 = "SELECT Max(Ano) as MaxAno FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
      
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MaxAno = vFld(Rs("MaxAno"))
   End If
   Call CloseRs(Rs)
   
   If MaxAno <= 0 Then
      MaxAno = Year(Now)
   End If
   
   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 2 - " & MaxAno, 1)
   
   Q1 = "SELECT Ano, FCierre, FApertura FROM EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa=" & lsEmpresa.Matrix(C_IDEMPRESA)
   Q1 = Q1 & " ORDER BY Ano DESC"
   Set Rs = OpenRs(DbMain, Q1)
   
   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: 3", 1)
   
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
            LsAno.list(LsAno.NewIndex) = Ano & " *"
            LsAno.Matrix(C_FCIERRE, LsAno.NewIndex) = vFld(Rs("FCierre"))
            LsAno.Matrix(C_FAPERTURA, LsAno.NewIndex) = vFld(Rs("FApertura"))
            Rs.MoveNext
            Exit Do
         Else ' If i > Ano Then 19 oct 2017 - pam: para evitar que entre en un loop infinito
            Exit Do
         End If
      Loop
   
   Next i
   
   Call CloseRs(Rs)
   
   Call AddDebug("FrmSelEmpresas.Ls_Empresas_Click: FIN", 1)

End Sub
Private Sub Ls_Empresas_DblClick()

   Call PostClick(Bt_Sel)
   
End Sub

Private Sub Op_SortNombre_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "0")
   gVarIniFile.SelEmprPorRUT = 0
   Me.MousePointer = vbDefault

End Sub

Private Sub Op_SortRUT_Click()

   Me.MousePointer = vbHourglass
   Call FillList
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", "1")
   gVarIniFile.SelEmprPorRUT = 1
   Me.MousePointer = vbDefault

End Sub



