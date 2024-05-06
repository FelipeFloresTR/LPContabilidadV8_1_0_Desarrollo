VERSION 5.00
Begin VB.Form FrmMarkActFijo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marcar Activo Fijo para Exportar"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8715
   Icon            =   "FrmMarkActFijo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   1500
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      Begin VB.TextBox Tx_Desc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4875
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   5160
         TabIndex        =   0
         Top             =   2460
         Width           =   1155
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NOTA: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   8
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Activo Fijo:"
         Height          =   255
         Left            =   420
         TabIndex        =   7
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Esta opción se utiliza solamente cuando un activo fijo no ha sido exportado al año siguiente."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   1260
         Width           =   5835
      End
      Begin VB.Label Label3 
         Caption         =   $"FrmMarkActFijo.frx":000C
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   420
         TabIndex        =   5
         Top             =   1800
         Width           =   5835
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   300
      Picture         =   "FrmMarkActFijo.frx":009A
      ScaleHeight     =   780
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   480
      Width           =   750
   End
End
Attribute VB_Name = "FrmMarkActFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdActivo As Long

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String
   
   If lIdActivo <= 0 Or Tx_Desc = "" Then
      MsgBox1 "Activo fijo no encontrado.", vbExclamation
      Exit Sub
   End If
   
   Q1 = "UPDATE MovActivoFijo SET FExported = 0 WHERE IdActFijo = " & lIdActivo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Unload Me

End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim Rs As Recordset

   Q1 = "SELECT Descrip FROM MovActivoFijo WHERE IdActFijo = " & lIdActivo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Tx_Desc = vFld(Rs("Descrip"))
   End If
   
   Call CloseRs(Rs)
      
End Sub

Public Sub FEdit(ByVal IdActivo As Long)

   lIdActivo = IdActivo
   
   Me.Show vbModal
   
End Sub

