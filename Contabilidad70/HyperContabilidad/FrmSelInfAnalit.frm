VERSION 5.00
Begin VB.Form FrmSelInfAnalit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes Analíticos"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "FrmSelInfAnalit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   540
      Picture         =   "FrmSelInfAnalit.frx":000C
      ScaleHeight     =   705
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   600
      Width           =   675
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Index           =   0
      Left            =   1740
      TabIndex        =   4
      Top             =   540
      Width           =   2595
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Analítico por Cuenta"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   1
         Top             =   780
         Width           =   1815
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Analítico por Entidad"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   0
         Top             =   360
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FrmSelInfAnalit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INFO_ENTIDAD = 1
Const INFO_CUENTA = 2

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As FrmInfAnalitico
   
   Set Frm = New FrmInfAnalitico
   
   Me.MousePointer = vbHourglass
   If Op_Informes(INFO_ENTIDAD).Value <> 0 Then
      Call Frm.FViewPorEntidad(0)
   Else
      Call Frm.FViewPorCuenta(0)
   End If
   
   Set Frm = Nothing
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Op_Informes(INFO_ENTIDAD) = True
   
End Sub

Private Sub Op_Informes_DblClick(Index As Integer)

   Call PostClick(Bt_Sel)
   
End Sub
