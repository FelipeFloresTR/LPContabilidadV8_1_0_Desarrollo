VERSION 5.00
Begin VB.Form FrmCambioClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio clave"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "FrmCambioClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   1680
      TabIndex        =   5
      Top             =   540
      Width           =   4455
      Begin VB.TextBox Tx_ClaveActual 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2340
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Tx_Clave1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2340
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Tx_Clave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2340
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese su clave actual:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese nueva clave:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Repita nueva clave:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1755
      End
   End
   Begin VB.CommandButton bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   480
      Picture         =   "FrmCambioClave.frx":000C
      Top             =   600
      Width           =   795
   End
End
Attribute VB_Name = "FrmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRc As Integer

Private Sub bt_Cancel_Click()
   Unload Me
   
End Sub

Private Sub bt_OK_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Clave As String, QryWhere As String
   
   If Valida() = False Then
      Exit Sub
      
   End If
   
   Clave = LCase(Trim(Tx_Clave1))
   If Trim(gUsuario.ClaveACtual) <> "" Then
      QryWhere = " AND Clave='" & GenClave(LCase(gUsuario.Nombre & Trim(gUsuario.ClaveACtual))) & "'"
      
   Else
      QryWhere = " AND Clave='" & GenClave(LCase(gUsuario.Nombre & gUsuario.ClaveACtual)) & "'"
      
   End If
   
   Q1 = "UPDATE Usuarios SET Clave ='" & GenClave(LCase(gUsuario.Nombre & Clave)) & "'"
   Q1 = Q1 & " WHERE IdUsuario =" & gUsuario.IdUsuario
   Call ExecSQL(DbMain, Q1)
   
   gUsuario.ClaveACtual = GenClave(LCase(gUsuario.Nombre & Clave))
   
   MsgBox1 "Su cambio de clave ha sido realizado ", vbExclamation
   Unload Me
   
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   Caption = "Usuario " & gUsuario.Nombre
   
End Sub

Private Sub Tx_ClaveActual_GotFocus()
   Tx_ClaveActual = ""
   
End Sub

Private Function Valida() As Boolean

   Valida = False
   
   If gUsuario.ClaveACtual <> GenClave(LCase(gUsuario.Nombre) & Trim(Tx_ClaveActual)) Then
      MsgBox1 "Clave actual es incorrecta, intentelo nuevamente", vbExclamation
      Tx_ClaveActual.SetFocus
      Exit Function
      
   End If
    
   If Trim(Tx_Clave1) <> Trim(Tx_Clave2) Then
      MsgBox1 "Claves no son iguales ", vbExclamation
      Tx_Clave2.SetFocus
      Exit Function
      
   End If
   Valida = True
   
      
End Function
