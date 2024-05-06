VERSION 5.00
Begin VB.Form FrmMantUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de usuario"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8850
   Icon            =   "FrmMantUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7380
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7380
      TabIndex        =   7
      Top             =   480
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   1440
      TabIndex        =   9
      Top             =   420
      Width           =   5535
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   3000
         Picture         =   "FrmMantUsuario.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2400
         Width           =   225
      End
      Begin VB.CheckBox Ch_Activo 
         Caption         =   "Activo"
         Height          =   255
         Left            =   4020
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Tx_Clave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1920
         Width           =   1035
      End
      Begin VB.TextBox Tx_Clave1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1740
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1515
      End
      Begin VB.TextBox Tx_NombreLargo 
         Height          =   315
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox Ck_PrvAdm 
         Caption         =   "Administrador del Sistema"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   1980
         Width           =   2235
      End
      Begin VB.CheckBox Ch_Clave 
         Caption         =   "Modifica clave"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Indica que se desea cambiar la clave de este usuario por la ingresada en  el campo Ingrese Clave"
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Habilitado hasta:"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   16
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Repita  clave:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   1980
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese clave:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Largo:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1020
         Width           =   1050
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   420
      Picture         =   "FrmMantUsuario.frx":0316
      Top             =   540
      Width           =   750
   End
End
Attribute VB_Name = "FrmMantUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Oper As Integer
Dim OldNombre As String
Dim ModClave As Integer
Dim lidUsuario As Long
Dim lNombre As String
Dim lRc As Integer
Dim ClaveACtual As String

Private Sub Bt_Cancel_Click()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   lRc = vbCancel
   
   If Oper = OPER_EDIT Then
      Caption = "Editar usuario"
      Call FillForm
      
   Else
      Caption = "Nuevo usuario"
      
   End If
      
End Sub

Private Sub bt_OK_Click()

   If Valida() = False Then
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   Call SaveAll
   MousePointer = vbDefault
   
   lRc = vbOK
   lNombre = Tx_Nombre
   Unload Me
   
End Sub
Public Function FNew(Nombre As String, IdUsuario As Long) As Integer
  
   Oper = OPER_NEW
   Me.Show vbModal
   
   FNew = lRc
   Nombre = lNombre
   IdUsuario = lidUsuario
   
End Function

Public Function FEdit(Nombre As String, ByVal IdUsuario As Long) As Integer
   lidUsuario = IdUsuario
   Oper = OPER_EDIT

   Me.Show vbModal
   FEdit = lRc
   Nombre = lNombre
   
End Function

Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim t_Nombre As String
   Dim Clave As String
   Dim UserExist As Integer
   Dim StrClave As String
   
   t_Nombre = LCase(Trim(Tx_Nombre))
   Clave = LCase(Trim(Tx_Clave1))
   
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & t_Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      UserExist = True
      
   End If
   Call CloseRs(Rs)
   
   If Ch_Clave Or Oper = OPER_NEW Then
      StrClave = " Clave ='" & GenClave(LCase(t_Nombre & Clave)) & "',"
   End If
   
   Select Case Oper
      
      Case OPER_NEW
         If UserExist = True Then
            MsgBox1 "Este usuario ya existe.", vbExclamation
            Exit Sub
                     
         Else
            lidUsuario = AdvTbAddNew(DbMain, "Usuarios", "idUsuario", "Usuario", t_Nombre)
                       
         End If
      
      Case OPER_EDIT
         If t_Nombre <> OldNombre And UserExist = True Then    'cambió nombre por uno que ya existe
            MsgBox1 "Este usuario ya existe.", vbExclamation
            Exit Sub
            
         End If
         
      End Select
      
      Q1 = "UPDATE Usuarios SET " & StrClave & " Usuario = '" & t_Nombre & "'"
      Q1 = Q1 & ",NombreLargo='" & ParaSQL(Tx_NombreLargo) & "'"
      Q1 = Q1 & ",PrivAdm=" & Ck_PrvAdm
      Q1 = Q1 & ", Activo = " & IIf(Ch_Activo <> 0, -1, 0)
      Q1 = Q1 & ", HabilitadoHasta = " & GetTxDate(Tx_Hasta)
      Q1 = Q1 & " WHERE IdUsuario =" & lidUsuario
      Call ExecSQL(DbMain, Q1)
            
    
      
End Sub

Private Function Valida() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim t_Nombre As String
   Dim i As Integer
   
   Valida = False
   t_Nombre = LCase(Trim(Tx_Nombre))
  
   If t_Nombre = "" Then
      MsgBox1 "Debe ingresar nombre.", vbExclamation
      Tx_Nombre.SetFocus
      Exit Function
   End If
   
   If Trim(Tx_NombreLargo) = "" Then
      MsgBox1 "Debe ingresar nombre largo", vbExclamation
      Tx_NombreLargo.SetFocus
      Exit Function
      
   End If
   
   Q1 = "SELECT IdUsuario FROM Usuarios WHERE Usuario = '" & t_Nombre & "'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then    ' existe
      If Oper = OPER_NEW Then
         MsgBox1 "Este usuario ya existe.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
         
      ElseIf Oper = OPER_EDIT And t_Nombre <> OldNombre Then    'cambió nombre por uno que ya existe
         MsgBox1 "Este usuario ya existe.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
         
      End If
      
   End If
   Call CloseRs(Rs)
   
   If LCase(Trim(Tx_Clave1)) <> LCase(Trim(Tx_Clave2)) Then
      MsgBox1 "Las claves son distintas.", vbExclamation + vbOKOnly
      Tx_Clave1.SetFocus
      Exit Function
      
   End If
     
   Valida = True
   
End Function


Private Sub Tx_Clave1_Change()
   Ch_Clave.Value = 1
End Sub

Private Sub Tx_Clave1_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub


Private Sub Tx_Clave2_Change()
   Ch_Clave.Value = 1

End Sub

Private Sub Tx_Clave2_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
   
End Sub

Private Sub Tx_Nombre_KeyPress(KeyAscii As Integer)
   Call KeyUserId(KeyAscii)
   
End Sub
Private Sub FillForm()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Usuario, Clave, NombreLargo, PrivAdm, Activo, HabilitadoHasta FROM Usuarios "
   Q1 = Q1 & " WHERE idUsuario=" & lidUsuario
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      OldNombre = LCase(vFld(Rs("Usuario")))
      Tx_Nombre = LCase(vFld(Rs("Usuario")))
      ClaveACtual = vFld(Rs("Clave"))
      Tx_NombreLargo = vFld(Rs("NombreLargo"), True)
      Ck_PrvAdm.Value = vFld(Rs("PrivAdm"))
      Ch_Activo.Value = IIf(vFld(Rs("Activo")) <> 0, 1, 0)
      Call SetTxDate(Tx_Hasta, vFld(Rs("HabilitadoHasta")))
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

