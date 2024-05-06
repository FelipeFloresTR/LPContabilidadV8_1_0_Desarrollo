VERSION 5.00
Begin VB.Form FrmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "FrmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton bt_Del 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   3960
      Picture         =   "FrmUsuarios.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2580
      Width           =   975
   End
   Begin VB.ListBox Ls_Usuario 
      Height          =   3570
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   2355
   End
   Begin VB.CommandButton Bt_New 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   3960
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsuarios.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo usuario"
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton Bt_Mod 
      Caption         =   "&Modificar"
      Height          =   735
      Left            =   3960
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsuarios.frx":088E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Editar usuario"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   300
      Picture         =   "FrmUsuarios.frx":0E61
      Top             =   540
      Width           =   945
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Del_Click()
   Dim Rs As Recordset
   Dim Q1 As String

   If Ls_Usuario.ListCount = 0 Then
      Exit Sub
      
   End If
   
   If Ls_Usuario.ListIndex = -1 Then
      Exit Sub
      
   End If
  
   MousePointer = vbHourglass
   
   If MsgBox1("¿Está seguro que desea eliminar al usuario " & Ls_Usuario.Text & " ?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      MousePointer = vbDefault
      Exit Sub
      
   End If
            
   Call DeleteSQL(DbMain, "Usuarios", " WHERE IdUsuario = " & Ls_Usuario.ItemData(Ls_Usuario.ListIndex))
   Ls_Usuario.RemoveItem Ls_Usuario.ListIndex
   If Ls_Usuario.ListCount > 0 Then
      Ls_Usuario.ListIndex = 0
      
   End If
   
   MousePointer = vbDefault

End Sub


Private Sub Bt_Mod_Click()
   Dim IdUsuario As Long
   Dim Nombre As String
   Dim Frm As FrmMantUsuario
   
   If Ls_Usuario.ListCount = 0 Then
      Exit Sub
      
   End If
   
   If Ls_Usuario.ListIndex = -1 Then
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass

   IdUsuario = Ls_Usuario.ItemData(Ls_Usuario.ListIndex)
   Nombre = Ls_Usuario.Text
   
   Set Frm = New FrmMantUsuario
   If Frm.FEdit(Nombre, IdUsuario) = vbOK Then
      Ls_Usuario.List(Ls_Usuario.ListIndex) = Nombre
      
   End If
   
   Set Frm = Nothing
   MousePointer = vbDefault

End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmMantUsuario
   Dim Nombre As String
   Dim IdUsuario As Long
   
   MousePointer = vbHourglass
   Set Frm = New FrmMantUsuario
   If Frm.FNew(Nombre, IdUsuario) = vbOK Then
      Ls_Usuario.AddItem Nombre
      Ls_Usuario.ItemData(Ls_Usuario.NewIndex) = IdUsuario
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
   
End Sub
Private Sub Form_Load()
   Dim Q1 As String

   Q1 = "SELECT Usuario,idUsuario FROM Usuarios WHERE Usuario<>'Administ'"
   Call FillCombo(Ls_Usuario, DbMain, Q1, -1)
   
   If gAppCode.Demo Then
      Bt_New.Enabled = False
   End If
   
End Sub
Private Sub Ls_Usuario_DblClick()
   Call Bt_Mod_Click
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
   
End Sub
