VERSION 5.00
Begin VB.Form FrmidUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificación del Usuario - Administración"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "FrmidUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Ck_Link 
      Caption         =   "Crear &ícono en el Escritorio"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1470
      Width           =   2415
   End
   Begin VB.CommandButton Bt_Aceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox Tx_Clave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   3180
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Tx_Nombre 
      Height          =   315
      Left            =   1380
      MaxLength       =   15
      TabIndex        =   1
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label La_Demo 
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
      Left            =   3300
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Clave:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   2
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Index           =   4
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   585
   End
End
Attribute VB_Name = "FrmidUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRc As Integer
Private Sub Bt_Aceptar_Click()
   
   Call AddLog("FrmIdUsuario: Vamos a Valida Usuario", 1)
   
   If ValidaUsuario() = False Then
      Exit Sub
   End If
   
   Call AddLog("FrmIdUsuario: Paso Valida Usuario", 1)
      
   Call SetLastUser(gIniFile, gUsuario.Nombre)
   
   Call AddLog("FrmIdUsuario: Paso SetLastUser", 1)
   
   Call SetIniString(gIniFile, "Config", "LinkAdm", Abs(Ck_Link))

   Call AddLog("FrmIdUsuario: Paso SetIniString", 1)
   
   If Ck_Link Then
      Call CreateLnk("$Desktop\Administrador LPContabilidad ")
   End If
   
   Call AddLog("FrmIdUsuario: Paso CreateLnk", 1)
   
   lRc = vbOK
   
   Unload Me
End Sub
Public Function FShow() As Integer
   Me.Show vbModal
   FShow = lRc
   
   Call AddDebug("FShow=" & lRc)

End Function

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   
   Tx_Nombre = GetLastUser(gIniFile, gAdmUser)
   
   FrmidUsuario.Top = FrmStart.Top + (FrmStart.Height - FrmidUsuario.Height) / 2 - 500
   FrmidUsuario.Left = FrmStart.Left + FrmStart.Width - FrmidUsuario.Width
      
   'Ck_Link.Value = Abs(Val(GetIniString(gIniFile, "Config", "LinkAdm", "1")) <> 0)

   La_demo.Visible = gAppCode.Demo
   
   
   

End Sub
Private Function ValidaAdm_Old() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   ValidaAdm_Old = False
   
   'Q1 = "SELECT Clave,Usuario,idUsuario FROM Usuarios WHERE idUsuario=" & ID_ADMIN
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If Trim(vFld(Rs("Clave"))) <> GenClave(LCase(vFld(Rs("Usuario")) & Trim(Tx_Clave))) Then
         MsgBox1 "Clave incorrecta", vbExclamation
         Call CloseRs(Rs)
         Exit Function
         
      End If
      
      gUsuario.ClaveACtual = vFld(Rs("Clave"))
      gUsuario.Nombre = vFld(Rs("Usuario"))
      gUsuario.IdUsuario = vFld(Rs("idUsuario"))
      
   End If
   Call CloseRs(Rs)
    
   ValidaAdm_Old = True
   
End Function

Private Function ValidaUsuario() As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   DbMainDate = GetDbNow(DbMain)
   
   ValidaUsuario = False
   gUsuario.Nombre = LCase(Trim(Tx_Nombre))
     
   Q1 = "SELECT IdUsuario, Clave,PrivAdm, Activo, HabilitadoHasta FROM Usuarios "
   Q1 = Q1 & " WHERE Usuario = '" & gUsuario.Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then
      MsgBox1 "Usuario desconocido.", vbExclamation
      Call CloseRs(Rs)
      Exit Function
      
   End If
   
   If Trim(Rs("Clave")) <> GenClave(LCase(gUsuario.Nombre & Trim(Tx_Clave))) Then
      Debug.Print GenClave(LCase(gUsuario.Nombre & Trim(Tx_Clave)))
      
      If W.InDesign Then
         If MsgBox1("Modo diseño ¿Desea verificar la clave?", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
            MsgBox1 "Clave incorrecta.", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
      Else
         MsgBox1 "Clave incorrecta.", vbExclamation
         Call CloseRs(Rs)
         Exit Function
      End If
            
   End If
   
   If vFld(Rs("PrivAdm")) <> PRV_ADM And LCase(Tx_Nombre) <> LCase(gAdmUser) Then
      MsgBox1 "Este usuario no tiene privilegio de administrador.", vbExclamation
      Call CloseRs(Rs)
      Exit Function
   Else
      gUsuario.Priv = PRV_ADMIN

   End If
   
   If vFld(Rs("Activo")) = 0 Then
      MsgBox1 "Este usuario no está activo en el sistema.", vbExclamation
      Call CloseRs(Rs)
      Exit Function
      
   End If
      
   If vFld(Rs("HabilitadoHasta")) > 0 And Int(DbMainDate) > vFld(Rs("HabilitadoHasta")) Then
      MsgBox1 "Este usuario no está activo en el sistema.", vbExclamation
      Call ExecSQL(DbMain, "UPDATE Usuarios SET Activo = 0, HabilitadoHasta = NULL WHERE IdUsuario = " & vFld(Rs("IdUsuario")))
      Call CloseRs(Rs)
      Exit Function
      
   End If
      
   If Tx_Nombre = gAdmUser Then
      gUsuario.Priv = PRV_ADMIN
   
   End If
   
   
   gUsuario.IdUsuario = vFld(Rs("idUsuario"))
   gUsuario.ClaveACtual = vFld(Rs("Clave"))
   
   Call CloseRs(Rs)
   ValidaUsuario = True
   
End Function
Private Sub Form_Activate()

   On Error Resume Next

   If Tx_Nombre <> "" Then
      Tx_Clave.SetFocus
   End If

   
End Sub
