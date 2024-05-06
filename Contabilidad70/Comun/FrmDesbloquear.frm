VERSION 5.00
Begin VB.Form FrmDesbloquear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloquear Conexión"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmDesbloquear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Update 
      Caption         =   "Actualizar"
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Desbloquear"
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox Ls_Usr 
      Height          =   2400
      Left            =   1320
      TabIndex        =   0
      Top             =   780
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuarios de Windows conectados:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   540
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "FrmDesbloquear.frx":000C
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "FrmDesbloquear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String, i As Integer

   i = Ls_Usr.ListIndex
   If i < 0 Then
      Exit Sub
   End If

   If MsgBox1("¿ Está seguro que desea desbloquear la conexión del usuario '" & Ls_Usr.List(i) & "' ?", vbYesNo Or vbDefaultButton2 Or vbQuestion) = vbNo Then
      Exit Sub
   End If

   If Ls_Usr.ListIndex >= 0 Then
'      Q1 = "DELETE * FROM PcUsr WHERE Usr = '" & Ls_Usr & "'"
      Q1 = " WHERE Usr = '" & Ls_Usr.List(i) & "' And Pid=" & Ls_Usr.ItemData(i)
      Call DeleteSQL(DbMain, "PcUsr", Q1)
      Call AddLog("Desconecta usuario: '" & Ls_Usr.List(i) & ", pid=" & Ls_Usr.ItemData(i))
      Ls_Usr.RemoveItem i
      
   End If

End Sub

Private Sub Bt_Update_Click()

   MousePointer = vbHourglass
   DoEvents

   Call LoadAll
   
   MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   
   Ls_Usr.Clear
   
   DoEvents
   
   Q1 = "SELECT Usr, Pid FROM PcUsr ORDER BY Usr"
   Call FillCombo(Ls_Usr, DbMain, Q1, -1)

End Sub
