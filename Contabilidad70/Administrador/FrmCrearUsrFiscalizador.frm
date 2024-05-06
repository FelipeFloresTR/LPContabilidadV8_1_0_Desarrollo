VERSION 5.00
Begin VB.Form FrmCrearUsrFiscalizador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear o Actualizar Usuarios Fiscalizadores"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7740
      TabIndex        =   15
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7740
      TabIndex        =   14
      Top             =   960
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   1380
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.CheckBox Ch_Clave 
         Caption         =   "Modifica clave"
         Height          =   255
         Left            =   3060
         TabIndex        =   16
         Top             =   1980
         Width           =   2295
      End
      Begin VB.TextBox Tx_User 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox Tx_Clave1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1920
         Width           =   1035
      End
      Begin VB.TextBox Tx_Clave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox Tx_User 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   420
         Width           =   1395
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   3000
         Picture         =   "FrmCrearUsrFiscalizador.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   225
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   1740
         TabIndex        =   3
         Top             =   1440
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Empresa 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   3915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario 2:"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   12
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese clave:"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   11
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Repita  clave:"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   10
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario 1:"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Habilitados hasta:"
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   5
         Top             =   1500
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa:"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   240
      Picture         =   "FrmCrearUsrFiscalizador.frx":030A
      Top             =   600
      Width           =   750
   End
End
Attribute VB_Name = "FrmCrearUsrFiscalizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const USR_FISC1 = "fiscalizador1"
Const USR_FISC2 = "fiscalizador2"

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()

   If Valida() Then
      Call SaveAll
      Unload Me
   End If
   
   
End Sub

Private Sub Form_Load()

   Tx_User(1) = USR_FISC1
   Tx_User(2) = USR_FISC2
   
   Call SetTxDate(Tx_Hasta, DateAdd("m", 1, Now))
   
   Call LoadEmpresas
   
End Sub

Private Sub LoadEmpresas()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT IdEmpresa, Rut, NombreCorto, Estado FROM Empresas"
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')"
   End If
   
   Q1 = Q1 & " ORDER BY NombreCorto"
   
   Set Rs = OpenRs(DbMain, Q1)

   Do While Not Rs.EOF
      Call CbAddItem(Cb_Empresa, vFld(Rs("NombreCorto")), vFld(Rs("IdEmpresa")))
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
      
   Cb_Empresa.ListIndex = -1
   
End Sub
Private Function Valida() As Boolean

   Valida = False
      
   If CbItemData(Cb_Empresa) <= 0 Then
      MsgBox1 "Debe seleccionar la empresa.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Hasta) = "" Or GetTxDate(Tx_Hasta) < Int(Now) Then
      MsgBox1 "Fecha inválida.", vbExclamation
      Exit Function
   End If
      
   If LCase(Trim(Tx_Clave1)) <> LCase(Trim(Tx_Clave2)) Then
      MsgBox1 "Las claves son distintas.", vbExclamation + vbOKOnly
      Tx_Clave1.SetFocus
      Exit Function
      
   End If
     
      
   Valida = True
   
End Function
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim idPerfil As Integer
   Dim IdFisc1 As Long, IdFisc2 As Long
   Dim IdEmpresa As Long
   Dim Nombre As String
   Dim Clave As String, ClaveEnc1 As Long, ClaveEnc2 As Long
   Dim MsgClave As String
   
   Nombre = LCase(Trim(Tx_User(1)))
   Clave = LCase(Trim(Tx_Clave1))
   ClaveEnc1 = GenClave(LCase(Nombre & Clave))

   Nombre = LCase(Trim(Tx_User(2)))
   ClaveEnc2 = GenClave(LCase(Nombre & Clave))
   
   
   'creamos perfil Fiscalizador con sólo privilegio de Ver
   
   Q1 = "SELECT IdPerfil FROM Perfiles WHERE Nombre = 'Fiscalizador'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
   
      Q1 = "SELECT Max(idPerfil) as M FROM Perfiles"
      Set Rs = OpenRs(DbMain, Q1)
      idPerfil = vFld(Rs("M")) + 1
      Call CloseRs(Rs)

      Q1 = "INSERT INTO Perfiles (IdPerfil, Nombre, Privilegios, IdApp)"
      Q1 = Q1 & " VALUES (" & idPerfil & ",'Fiscalizador', " & PRV_VER_INFO & ", 0)"
      Call ExecSQL(DbMain, Q1)
      
   Else
      
      idPerfil = vFld(Rs("IdPerfil"))
      
   End If
   
   Call CloseRs(Rs)
   
   MsgClave = IIf(Ch_Clave <> 0, vbCrLf & vbCrLf & "Se asignará la nueva clave.", "")
   
   
   'Usuario Fiscalizador 1
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & USR_FISC1 & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
   
      Q1 = "INSERT INTO Usuarios (Usuario, Clave, NombreLargo, PrivAdm, Activo, HabilitadoHasta)"
      Q1 = Q1 & " VALUES ('" & USR_FISC1 & "', " & ClaveEnc1 & ", 'Fiscalizador 1', 0, -1, " & GetTxDate(Tx_Hasta) & ")"
   
      Call ExecSQL(DbMain, Q1)
      
      Call CloseRs(Rs)
      
      Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & USR_FISC1 & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdFisc1 = vFld(Rs("IdUsuario"))
      End If
         
   Else
   
      IdFisc1 = vFld(Rs("IdUsuario"))
      
      MsgBox1 "El usuario " & USR_FISC1 & " ya existe." & vbCrLf & vbCrLf & "Se asignará a la empresa seleccionada." & MsgClave, vbInformation
      
      Q1 = "UPDATE Usuarios SET PrivAdm = 0, Activo = -1, HabilitadoHasta = " & GetTxDate(Tx_Hasta)
      If Ch_Clave <> 0 Then
         Q1 = Q1 & ", Clave = " & ClaveEnc1
      End If
      Q1 = Q1 & " WHERE IdUsuario = " & IdFisc1
            
      Call ExecSQL(DbMain, Q1)
      
   End If
   
   Call CloseRs(Rs)
   
   'Usuario Fiscalizador 2
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & USR_FISC2 & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
   
      Q1 = "INSERT INTO Usuarios (Usuario, Clave, NombreLargo, PrivAdm, Activo, HabilitadoHasta)"
      Q1 = Q1 & " VALUES ('" & USR_FISC2 & "', " & ClaveEnc2 & ", 'Fiscalizador 2', 0, -1, " & GetTxDate(Tx_Hasta) & ")"
   
      Call ExecSQL(DbMain, Q1)
      
      Call CloseRs(Rs)
      
      Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & USR_FISC2 & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdFisc2 = vFld(Rs("IdUsuario"))
      End If
            
   Else
   
      IdFisc2 = vFld(Rs("IdUsuario"))
      
      MsgBox1 "El usuario " & USR_FISC2 & " ya existe." & vbCrLf & vbCrLf & "Se asignará a la empresa seleccionada." & MsgClave, vbInformation
            
      Q1 = "UPDATE Usuarios SET PrivAdm = 0, Activo = -1, HabilitadoHasta = " & GetTxDate(Tx_Hasta)
      If Ch_Clave <> 0 Then
         Q1 = Q1 & ", Clave = " & ClaveEnc2
      End If
      Q1 = Q1 & " WHERE IdUsuario = " & IdFisc2
      
      Call ExecSQL(DbMain, Q1)
      
   End If
   
   Call CloseRs(Rs)
  
   'Asignamos perfil a estos usuarios para la empresa seleccionada
   IdEmpresa = CbItemData(Cb_Empresa)
   
   'Fiscalizador 1
   'primero vemos si no está asignado a otra empresa
   Q1 = "SELECT IdEmpresa FROM UsuarioEmpresa "
   Q1 = Q1 & " WHERE IdUsuario = " & IdFisc1 & " AND IdEmpresa <> " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)

   If Not Rs.EOF Then    'está asignao a otra empresa, preguntamos si lo borremoas de esa empresa
      If MsgBox1("El usuario " & USR_FISC1 & " ya está asignado a la empresa " & cbItemText(Cb_Empresa, vFld(Rs("IdEmpresa"))) & "." & vbCrLf & vbCrLf & " ¿Desea eliminar esta asignación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         Q1 = " WHERE IdUsuario = " & IdFisc1 & " AND IdEmpresa = " & vFld(Rs("IdEmpresa"))
         Call DeleteSQL(DbMain, "UsuarioEmpresa", Q1)
      End If
   End If
   
   Call CloseRs(Rs)
   
   'ahora vemos si ya está asignado a esta empresa
   Q1 = "SELECT IdPerfil FROM UsuarioEmpresa "
   Q1 = Q1 & " WHERE IdUsuario = " & IdFisc1 & " AND IdEmpresa = " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then    'no lo está, lo asignamos
      Q1 = "INSERT INTO UsuarioEmpresa (IdUsuario, IdEmpresa, IdPerfil)"
      Q1 = Q1 & " VALUES( " & IdFisc1 & "," & IdEmpresa & "," & idPerfil & ")"
      
      Call ExecSQL(DbMain, Q1)
      
   Else              'está asignado, actualizamos el perfil por si acaso
      Q1 = "UPDATE UsuarioEmpresa SET IdPerfil = " & idPerfil
      Q1 = Q1 & " WHERE IdUsuario = " & IdFisc1 & " AND IdEmpresa = " & IdEmpresa
      
      Call ExecSQL(DbMain, Q1)

   End If
   
   Call CloseRs(Rs)
   
   'Fiscalizador 2
   'primero vemos si no está asignado a otra empresa
   Q1 = "SELECT IdEmpresa FROM UsuarioEmpresa "
   Q1 = Q1 & " WHERE IdUsuario = " & IdFisc2 & " AND IdEmpresa <> " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)

   If Not Rs.EOF Then    'está asignao a otra empresa, preguntamos si lo borremoas de esa empresa
      If MsgBox1("El usuario " & USR_FISC2 & " ya está asignado a la empresa " & cbItemText(Cb_Empresa, vFld(Rs("IdEmpresa"))) & "." & vbCrLf & vbCrLf & "¿Desea eliminar esta asignación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
         Q1 = " IdUsuario = " & IdFisc2 & " AND IdEmpresa = " & vFld(Rs("IdEmpresa"))
         Call ExecSQL(DbMain, "UsuarioEmpresa", Q1)
      End If
   End If
   
   Call CloseRs(Rs)
   
   'ahora vemos si ya está asignado a esta empresa
   Q1 = "SELECT IdPerfil FROM UsuarioEmpresa "
   Q1 = Q1 & " WHERE IdUsuario = " & IdFisc2 & " AND IdEmpresa = " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then    'no lo está, lo asignamos
      Q1 = "INSERT INTO UsuarioEmpresa (IdUsuario, IdEmpresa, IdPerfil)"
      Q1 = Q1 & " VALUES( " & IdFisc2 & "," & IdEmpresa & "," & idPerfil & ")"
      
      Call ExecSQL(DbMain, Q1)
      
   Else              'está asignado, actualizamos el perfil por si acaso
      Q1 = "UPDATE UsuarioEmpresa SET IdPerfil = " & idPerfil
      Q1 = Q1 & " WHERE IdUsuario = " & IdFisc2 & " AND IdEmpresa = " & IdEmpresa
      
      Call ExecSQL(DbMain, Q1)

   End If
   
   Call CloseRs(Rs)
   

End Sub
Private Sub Tx_Hasta_GotFocus()
   Call DtGotFocus(Tx_Hasta)
   
End Sub

Private Sub Tx_Hasta_LostFocus()
   
   If Trim$(Tx_Hasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Hasta)
      
End Sub

Private Sub Tx_Hasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Hasta)
   Set Frm = Nothing
   
End Sub

