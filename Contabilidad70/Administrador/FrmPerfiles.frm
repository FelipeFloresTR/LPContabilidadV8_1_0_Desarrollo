VERSION 5.00
Begin VB.Form FrmPerfiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de Perfiles"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   HelpContextID   =   3
   Icon            =   "FrmPerfiles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_MarcarTodo 
      Caption         =   "Marcar todo"
      Height          =   315
      Left            =   4980
      TabIndex        =   9
      Top             =   4440
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9420
      TabIndex        =   8
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perfil"
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   480
      Width           =   3495
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmPerfiles.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox Tx_Perfil 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   6
         Top             =   300
         Width           =   2295
      End
      Begin VB.ListBox Ls_Perfil 
         Height          =   2790
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   2295
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Nuevo"
         Height          =   675
         Left            =   2520
         MousePointer    =   99  'Custom
         Picture         =   "FrmPerfiles.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo usuario"
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton Bt_Ren 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   2520
         MousePointer    =   99  'Custom
         Picture         =   "FrmPerfiles.frx":088E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Editar usuario"
         Top             =   1860
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Privilegios"
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Index           =   1
      Left            =   4980
      TabIndex        =   1
      Top             =   480
      Width           =   5595
      Begin VB.ListBox Ls_Priv 
         Height          =   3210
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   300
         Width           =   5355
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   360
      Picture         =   "FrmPerfiles.frx":0E61
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "FrmPerfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrivChanged As Integer
Private IdxPerfil As Integer
Private lNoIncPrivMaestros As Boolean

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub
Private Sub Bt_Del_Click()
   Dim Rs As Recordset
   Dim Q1 As String
      
   If Ls_Perfil.ListCount <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
      
   'vemos si hay usuarios con este perfil
   Q1 = "SELECT IdUsuario FROM UsuarioEmpresa WHERE IdPerfil = " & Ls_Perfil.ItemData(IdxPerfil)
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      MsgBeep vbExclamation
      MsgBox "No es posible borrar este perfil. Hay usuarios que lo utilizan.", vbExclamation + vbOKOnly
      Tx_Perfil.SetFocus
      Call CloseRs(Rs)
      Exit Sub
   End If
   
   Call CloseRs(Rs)
   
   If MsgBox("¿Está seguro que desea borrar este perfil?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   Call DeleteSQL(DbMain, "Perfiles", " WHERE IdPerfil = " & Ls_Perfil.ItemData(IdxPerfil))
   
   Ls_Perfil.RemoveItem IdxPerfil
   If Ls_Perfil.ListCount > 0 Then
      Ls_Perfil.ListIndex = 0
   End If
   Tx_Perfil.SetFocus
   
End Sub


Private Sub Bt_MarcarTodo_Click()
   Dim i As Integer
   
   For i = 0 To Ls_Priv.ListCount - 1
      Ls_Priv.Selected(i) = True
   Next i
   
   PrivChanged = True
   Ls_Priv.SetFocus

End Sub

Private Sub Bt_Ren_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Idx As Integer
   Dim idPerfil As Integer
   Dim Perfil As String
   Dim i As Integer
   
   If Ls_Perfil.ListIndex < 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Perfil = Trim(Tx_Perfil)
   If Perfil = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   For i = 0 To Ls_Perfil.ListCount - 1
      If Perfil = Ls_Perfil.List(i) Then
         MsgBeep vbExclamation
         MsgBox "Este perfil ya existe.", vbOKOnly + vbExclamation
         Tx_Perfil.SetFocus
         Exit Sub
      End If
   Next i
      
   Call CloseRs(Rs)
         
   idPerfil = Ls_Perfil.ItemData(IdxPerfil)
   
   Call ExecSQL(DbMain, "UPDATE Perfiles SET Nombre ='" & ParaSQL(Perfil) & "' WHERE IdPerfil = " & idPerfil)
      
   Ls_Perfil.List(Ls_Perfil.ListIndex) = Perfil
   Ls_Perfil.SetFocus
   
End Sub

Private Sub Bt_New_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim idPerfil As Integer
   Dim Perfil As String
   Dim i As Integer
   
   Perfil = Trim(Tx_Perfil)
   If Perfil = "" Then
      MsgBox1 "Debe ingresar el nombre de un perfil", vbExclamation
      Tx_Perfil.SetFocus
      Exit Sub
   End If
   
   For i = 0 To Ls_Perfil.ListCount - 1
      If UCase(Perfil) = UCase(Ls_Perfil.List(i)) Then
         MsgBeep vbExclamation
         MsgBox "Este perfil ya existe", vbOKOnly + vbExclamation
         Tx_Perfil.SetFocus
         Exit Sub
      End If
   Next i
      
   Call CloseRs(Rs)
      
   idPerfil = GetMaxTableId("IdPerfil", "Perfiles", "")
   
   Call ExecSQL(DbMain, "INSERT INTO Perfiles (IdPerfil, IdApp, Nombre, Privilegios) VALUES(" & idPerfil & "," & gApp & ",'" & ParaSQL(Perfil) & "', 0)")
   Ls_Perfil.AddItem Perfil
   Ls_Perfil.ListIndex = Ls_Perfil.NewIndex
   Ls_Perfil.ItemData(Ls_Perfil.ListIndex) = idPerfil

   Call ClearPriv
   Ls_Perfil.SetFocus
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SavePerfil
End Sub

Private Sub Ls_Perfil_Click()

   IdxPerfil = Ls_Perfil.ListIndex
   
   If Ls_Perfil.ListCount > 0 Then
      Tx_Perfil = Ls_Perfil
      Call CheckPriv
   End If
   
End Sub

Private Sub Form_Load()
   FormPos Me

   'Ls_Priv.Enabled = False
   
   Call LoadPriv
   Call LoadPerfiles
   
End Sub
Private Sub LoadPerfiles()
   Dim Q1 As String
   
   Q1 = "SELECT IdPerfil, Nombre FROM Perfiles WHERE IdApp = " & gApp & " ORDER BY Nombre"
   
   Call LoadDbList(Ls_Perfil, Q1, -1)
      
   If Ls_Perfil.ListCount > 0 Then
      IdxPerfil = 0
   Else
      IdxPerfil = -1
   End If
   
End Sub

Private Sub LoadPriv()
   Dim i As Integer
   
   Ls_Priv.Clear
   
   For i = 0 To UBound(gPrivilegios)
      Ls_Priv.AddItem gPrivilegios(i)
      Ls_Priv.ItemData(Ls_Priv.NewIndex) = 2 ^ i
   Next i
  
   PrivChanged = False
   
End Sub

Private Sub CheckPriv()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   Call ClearPriv
   
   If IdxPerfil < 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT Privilegios FROM Perfiles WHERE IdPerfil = " & Ls_Perfil.ItemData(IdxPerfil)
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then
      For i = 0 To Ls_Priv.ListCount - 1
         If TienePrivilegio(Ls_Priv.ItemData(i), Rs(0)) = True Then
            Ls_Priv.Selected(i) = True
         Else
            Ls_Priv.Selected(i) = False
         End If
      Next i
   End If
   
   Call CloseRs(Rs)
   
   If Ls_Priv.ListCount > 0 Then
      Ls_Priv.TopIndex = 0
      Ls_Priv.ListIndex = 0
   End If
End Sub

Private Sub ClearPriv()
   Dim i As Integer
   
   For i = 0 To Ls_Priv.ListCount - 1
      Ls_Priv.Selected(i) = False
   Next i
   
   If Ls_Priv.ListCount > 0 Then
      Ls_Priv.TopIndex = 0
      Ls_Priv.ListIndex = 0
   End If
End Sub

Private Function GetPriv() As Long
   Dim i As Integer
   Dim Priv As Long

   Priv = 0
   For i = 0 To Ls_Priv.ListCount - 1
      If Ls_Priv.Selected(i) = True Then
         Priv = Priv + Ls_Priv.ItemData(i)
      End If
   Next i

   GetPriv = Priv
End Function

Private Sub SavePerfil()
   Dim Priv As Long
   
   If IdxPerfil < 0 Then
      Exit Sub
   End If
   
   If Ls_Perfil.ListCount = 0 Then
      Exit Sub
   End If
      
   If PrivChanged = True Then
      Priv = GetPriv()
     
      Call ExecSQL(DbMain, "UPDATE Perfiles SET Privilegios = " & Priv & " WHERE IdPerfil = " & Ls_Perfil.ItemData(IdxPerfil))
   
   End If
         
   PrivChanged = False

End Sub

Private Sub Ls_Priv_ItemCheck(Item As Integer)
   
   If LCase(Ls_Perfil) = "(todo)" Then
      Ls_Priv.Selected(Item) = True
      If Ls_Priv.Visible = True Then
         MsgBeep vbExclamation
      End If
   Else
      PrivChanged = True
   End If
   
End Sub

Private Sub Ls_Priv_LostFocus()
   MousePointer = vbHourglass
   Call SavePerfil
   MousePointer = vbDefault
End Sub
Public Sub ShowPerfiles(ByVal NoIncPrivMaestros As Boolean)

   lNoIncPrivMaestros = NoIncPrivMaestros
   Load Me
   Me.Show vbModal

End Sub

