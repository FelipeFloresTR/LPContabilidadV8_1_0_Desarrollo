VERSION 5.00
Begin VB.Form FrmGrupo 
   Caption         =   "Nuevo/Editar Grupo"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6900
      TabIndex        =   2
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   6900
      TabIndex        =   1
      Top             =   540
      Width           =   1155
   End
   Begin VB.Frame Fr_Datos 
      Height          =   1275
      Left            =   1500
      TabIndex        =   4
      Top             =   480
      Width           =   5055
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   0
         Top             =   540
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   540
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   420
      Picture         =   "FrmGrupo.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "FrmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lOper As Integer

Dim lId As Long
Dim lNombre As String

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub bt_OK_Click()
   
   If Not Valida() Then
      Exit Sub
   End If
   
   Call SaveAll

   lRc = vbOK
   Unload Me
End Sub

Private Function Valida() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   Valida = False
   
   lNombre = ParaSQL(Tx_Nombre)
   If lNombre = "" Then
      MsgBox1 "Nombre inválido.", vbExclamation
      Exit Function
   End If
   
   Q1 = "SELECT IdGrupo FROM AFGrupos WHERE NombGrupo = '" & lNombre & "' AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      Call MsgBox1("Este grupo ya existe.", vbExclamation)
      Call CloseRs(Rs)
      Exit Function
   End If
   Call CloseRs(Rs)
   
   Valida = True
   
End Function

Public Function FNew(id As Long) As Integer

   lOper = O_NEW
   lNombre = ""
   lId = 0
   
   Me.Show vbModal
   id = lId
   FNew = lRc
   
End Function
Public Function FEdit(ByVal id As Long, Nombre As String) As Integer

   lOper = O_EDIT
   lId = id
   lNombre = Nombre
   
   Me.Show vbModal
   FEdit = lRc
   Nombre = lNombre
   
End Function

Private Sub Form_Load()
   Tx_Nombre = lNombre
   
   If lOper = O_NEW Then
      Me.Caption = "Nuevo Grupo"
   Else
      Me.Caption = "Editar Grupo"
   End If
   
   Fr_Datos.Caption = Me.Caption
   
End Sub

Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset

   If lOper = O_NEW Then
   
      lId = AdvTbAddNew(DbMain, "AFGrupos", "IdGrupo", "NombGrupo", ParaSQL(lNombre))
      
      If lId > 0 Then
         Q1 = "UPDATE AFGrupos SET IdEmpresa = " & gEmpresa.id & " WHERE IdGrupo = " & lId
         Call ExecSQL(DbMain, Q1)
      End If
      
   Else
      Q1 = "UPDATE AFGrupos SET NombGrupo = '" & ParaSQL(lNombre) & "' WHERE IdGrupo = " & lId
      Call ExecSQL(DbMain, Q1)
      
   End If
      
End Sub
